#!/usr/bin/env pwsh
#requires -Version 7.0

<#
.SYNOPSIS
Produces a new BACPAC from an existing one with a table removed, schema and data.
The source file is opened read-only and never modified.

.DESCRIPTION
Takes a BACPAC in, writes a BACPAC out, minus one table. No database involved. The source
keeps its bytes.

WHY THIS IS NOT A DacFx SCRIPT - the constraint that shapes everything below:

  DacFx cannot write a BACPAC from a file. Its whole public surface offers exactly two
  ways to produce one, and neither starts from an existing package:

    DacServices.ExportBacpac   writes a BACPAC from a LIVE DATABASE
    BacPackage                 exposes only Load, Unpack and Dispose - no save, no pack

  Confirmed on the shipped assembly (170.2.70), on the published API reference for versions
  140 through 162, and against the DacFx repository, which ships SqlPackage and the
  Build.Sql SDK but not the closed-source runtime. Even the far simpler request of
  excluding tables at export time is still an open backlog item (DacFx issue #233).

  So the only supported DacFx route is a round trip through a database - ImportBacpac,
  DROP TABLE, ExportBacpac - which needs a SQL server, runs for hours on a large file, and
  begins with the very import that tends to be the problem in the first place.

  This script therefore rewrites the archive directly. DacFx is still used where it is
  authoritative: it validates the package on the way in, and it VALIDATES THE RESULT on the
  way out. That output gate is not decoration - it is what catches an incomplete removal.

WHAT A COMPLETE REMOVAL ACTUALLY INVOLVES - each step is required:

  1. model.xml    the table element AND every top-level element it owns. A clustered index
                  is declared as a SIBLING of the table, not as a child, so removing only
                  the table leaves a dangling index and DacFx then rejects the package with
                  "Could not load schema model from package". Ownership is resolved through
                  the model's own DefiningTable / IndexedObject / Parent relationships,
                  because constraints are named independently of their table.
  2. model.xml    refuse the run when anything OUTSIDE the table still references it (a
                  foreign key on another table, a view, a procedure body). Removing those
                  would silently change other objects, so this is a stop, not a warning.
  3. Data parts   the Data/<schema>.<table>/*.BCP entries.
  4. _rels/.rels  the Relationship entries targeting those data parts, two per part.
  5. Origin.xml   drop the table from DataPhaseTables, and recompute the SHA256 of the new
                  model.xml into Checksums. That checksum is a plain SHA256 over the
                  model.xml bytes, verified against the original before being rewritten.

.PARAMETER BacpacPath
BACPAC to read. Opened read-only, never modified.

.PARAMETER OutputPath
Where to write the new BACPAC. Defaults to <name>.without-<schema>.<table>.bacpac next to
the source.

.PARAMETER SchemaName
Schema of the table to remove. Default: dbo. Ignored when -TableName already carries one.

.PARAMETER TableName
Table to remove. Accepts "Table", "schema.Table" and "[schema].[Table]" alike, so the
qualified name can be pasted straight out of a log line or a report. Default: WhoIsActive,
the collection table of the widely used sp_WhoIsActive procedure, whose query_plan column
nests XML deeper than the 128 levels an import accepts - the case this script was built for.

.PARAMETER ListTables
List the tables held by the package, with their data part count and compressed size, then
stop. No output file is written. Use it to find the exact name to pass to -TableName.

.PARAMETER DacFxPath
Folder holding the .NET 8 build of Microsoft.SqlServer.Dac.dll. Auto-detected from the
sqlpackage dotnet tool and from the NuGet cache when omitted. The assemblies under
"Program Files\Microsoft SQL Server\<version>\DAC\bin" target .NET Framework and CANNOT be
loaded by PowerShell 7 - they fail with a type initializer error.

.PARAMETER SkipValidation
Skip the DacFx load of the source and of the result. Only for the case where no .NET 8
DacFx build is reachable. The output is then unverified.

.PARAMETER CompressionLevel
Deflate level for the rewritten archive: Optimal (default), Fastest or NoCompression. Every
entry is decompressed and recompressed on the way through, which dominates the runtime on
large files. Fastest trades output size for wall-clock.

.PARAMETER LogFile
Optional path to also append the console output to. The parent folder is created if needed.

.PARAMETER Force
Overwrite an existing output file.

.EXAMPLE
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac' -ListTables
# Which tables are in there, and what does each cost.

.EXAMPLE
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac' -TableName '[dbo].[WhoIsActive]' -WhatIf
# Qualified name accepted as written. Reports the removal set and any external reference,
# writes nothing.

.EXAMPLE
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac'
# Writes mydb.without-dbo.WhoIsActive.bacpac and validates it with DacFx.

.EXAMPLE
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac' -TableName AuditLog -CompressionLevel Fastest -Force

.NOTES
Self-contained: PowerShell 7 and, for the validation step, any .NET 8 DacFx build - the
"sqlpackage" dotnet tool is the easiest source.

Rewriting a BACPAC is not supported by Microsoft. Import the result against an empty,
disposable target database before trusting it with anything.

MIT licensed. See LICENSE.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$BacpacPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SchemaName = 'dbo',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TableName = 'WhoIsActive',

    [Parameter()]
    [string]$DacFxPath,

    [Parameter()]
    [switch]$ListTables,

    [Parameter()]
    [switch]$SkipValidation,

    [Parameter()]
    [ValidateSet('Optimal', 'Fastest', 'NoCompression')]
    [string]$CompressionLevel = 'Optimal',

    [Parameter()]
    [string]$LogFile,

    [Parameter()]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region logging -------------------------------------------------------------------------

$script:LogFilePath = ''
if ($LogFile) {
    $logFolder = Split-Path -Parent $LogFile
    if ($logFolder -and -not (Test-Path -LiteralPath $logFolder)) {
        # -WhatIf:$false on purpose: the log folder is not the state change being confirmed,
        # and a -WhatIf run must still be able to record what it would have done.
        $null = New-Item -ItemType Directory -Path $logFolder -Force -WhatIf:$false
    }
    $script:LogFilePath = $LogFile
}

function Write-ToolLog {
    <#
        Console is the product here, so the output goes to the host deliberately. A copy is
        appended to -LogFile when one was given.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '',
        Justification = 'Command-line tool: the console output is the deliverable.')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string]$Message,
        [Parameter()] [ValidateSet('INFO', 'WARN', 'ERROR')] [string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
    if ($script:LogFilePath) {
        # -WhatIf:$false on purpose: writing the log is not the state change being confirmed.
        # Without it, a -WhatIf run logs nothing and prints an "Add Content" line per entry.
        Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding utf8 -WhatIf:$false
    }
}

#endregion

#region helpers -------------------------------------------------------------------------

function Find-DacFxFolder {
    <#
        Locates a .NET 8 DacFx build. The "Program Files\Microsoft SQL Server\<v>\DAC\bin"
        assemblies are deliberately NOT considered: they target .NET Framework and throw a
        type initializer error under PowerShell 7 (missing
        System.Diagnostics.Eventing.EventDescriptor). The sqlpackage dotnet tool ships the
        net8.0 build, which loads cleanly.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $candidates = @(
        Join-Path $HOME '.dotnet/tools/.store/microsoft.sqlpackage/*/microsoft.sqlpackage/*/tools/net8.0/any'
        Join-Path $HOME '.nuget/packages/microsoft.sqlserver.dacfx/*/lib/net8.0'
    )
    # The file name is part of the -Path pattern on purpose: Get-ChildItem returns nothing
    # when -Filter is combined with a wildcarded -Path.
    foreach ($pattern in $candidates) {
        $hit = Get-ChildItem -Path (Join-Path $pattern 'Microsoft.SqlServer.Dac.dll') -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending | Select-Object -First 1
        if ($hit) { return $hit.DirectoryName }
    }
    return $null
}

function Import-DacFx {
    <#
        Loads DacFx with an AssemblyResolve hook, because Microsoft.SqlServer.Dac.dll pulls
        a dozen siblings that are not on the probing path of this process.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$Folder
    )

    $dacDll = Join-Path $Folder 'Microsoft.SqlServer.Dac.dll'
    if (-not (Test-Path -LiteralPath $dacDll -PathType Leaf)) { return $false }

    $resolver = [System.ResolveEventHandler] {
        param($eventSender, $resolveArgs)
        $null = $eventSender
        $simpleName = ([System.Reflection.AssemblyName]::new($resolveArgs.Name)).Name
        $candidate  = Join-Path $Folder "$simpleName.dll"
        if (Test-Path -LiteralPath $candidate) { return [System.Reflection.Assembly]::LoadFrom($candidate) }
        return $null
    }
    [System.AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
    $null = [System.Reflection.Assembly]::LoadFrom($dacDll)
    $extDll = Join-Path $Folder 'Microsoft.SqlServer.Dac.Extensions.dll'
    if (Test-Path -LiteralPath $extDll) { $null = [System.Reflection.Assembly]::LoadFrom($extDll) }
    return $true
}

function Test-DacPackageModel {
    <#
        Loads a package's schema model through DacFx. This is the real gate: an incomplete
        removal leaves unresolvable references and the constructor throws
        "Could not load schema model from package".
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [string]$PackagePath
    )

    try {
        $model = [Microsoft.SqlServer.Dac.Model.TSqlModel]::new($PackagePath)
        try {
            $scope = [Microsoft.SqlServer.Dac.Model.DacQueryScopes]::UserDefined
            $count = @($model.GetObjects($scope)).Count
            return @{ IsValid = $true; ObjectCount = $count; Reason = $null }
        }
        finally { $model.Dispose() }
    }
    catch {
        $inner = $_.Exception
        while ($inner.InnerException) { $inner = $inner.InnerException }
        return @{ IsValid = $false; ObjectCount = 0; Reason = $inner.Message }
    }
}

function Test-NameBelongsToTable {
    <#
        True when a DAC element name is the table itself or something it owns, e.g.
        "[dbo].[WhoIsActive]" or "[dbo].[WhoIsActive].[cx_collection_time]".
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string]$Name,
        [Parameter(Mandatory)] [string]$Qualified
    )

    if ([string]::IsNullOrEmpty($Name)) { return $false }
    return ($Name -ieq $Qualified -or
            $Name.StartsWith("$Qualified.", [System.StringComparison]::OrdinalIgnoreCase))
}

function ConvertTo-TableIdentifier {
    <#
        Normalises however the caller wrote the table name into a schema and a table.

        All four of these name the same object, and all four are accepted, because the
        qualified form is what the model itself displays and is therefore what gets copied
        out of a log line or a report:

            WhoIsActive                 uses -SchemaName
            dbo.WhoIsActive             schema taken from the string
            [dbo].[WhoIsActive]         schema taken from the string
            [WhoIsActive]               uses -SchemaName

        A closing bracket inside an identifier is doubled by SQL Server, so "[a]]b]" is the
        single name a]b - the regex accepts that and unescapes it.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$DefaultSchema
    )

    $bracketed = [regex]::Matches($Name, '\[((?:[^\]]|\]\])*)\]')
    if ($bracketed.Count -ge 2) {
        return @{
            Schema = $bracketed[0].Groups[1].Value -replace '\]\]', ']'
            Table  = $bracketed[1].Groups[1].Value -replace '\]\]', ']'
        }
    }
    if ($bracketed.Count -eq 1) {
        return @{ Schema = $DefaultSchema; Table = $bracketed[0].Groups[1].Value -replace '\]\]', ']' }
    }
    if ($Name.Contains('.')) {
        $parts = $Name.Split('.', 2)
        return @{ Schema = $parts[0]; Table = $parts[1] }
    }
    return @{ Schema = $DefaultSchema; Table = $Name }
}

function Get-ElementOwnerName {
    <#
        Returns the name of the table a top-level element belongs to, or an empty string.

        Ownership is explicit in the DAC model, through one relationship per element type:

            SqlPrimaryKeyConstraint / SqlUniqueConstraint      DefiningTable
            SqlCheckConstraint / SqlDefaultConstraint          DefiningTable
            SqlForeignKeyConstraint                            DefiningTable  (the table HOLDING the key)
            SqlIndex / SqlStatistic                            IndexedObject
            SqlDmlTrigger                                      Parent

        This matters because constraints are named independently of their table:
        [dbo].[PK_Orders] belongs to [dbo].[Orders] but its name says nothing about it.
        Without this lookup, every table carrying a primary key would be reported as having
        an external reference and the removal would be refused.

        ForeignTable is deliberately NOT treated as ownership: it is the table a foreign key
        POINTS AT. A key on another table pointing at the one being removed is exactly the
        case that must stop the run.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlNode]$Element
    )

    foreach ($relationshipName in @('DefiningTable', 'IndexedObject', 'Parent')) {
        $reference = $Element.SelectSingleNode(
            "./*[local-name()='Relationship' and @Name='$relationshipName']" +
            "/*[local-name()='Entry']/*[local-name()='References' and @Name]")
        if ($reference) { return $reference.GetAttribute('Name') }
    }
    return ''
}

function Get-ArchiveEntryContent {
    <#
        Reads one ZIP entry fully into a byte array. Used only for the three small control
        parts (model.xml, Origin.xml, _rels/.rels) - data parts are streamed, never buffered.
    #>
    [CmdletBinding()]
    [OutputType([byte[]])]
    param(
        [Parameter(Mandatory)] $Archive,
        [Parameter(Mandatory)] [string]$EntryName
    )

    $entry = $Archive.Entries | Where-Object { $_.FullName -ieq $EntryName } | Select-Object -First 1
    if (-not $entry) { throw "Entry not found in the archive: $EntryName" }

    $buffer = [System.IO.MemoryStream]::new()
    $stream = $entry.Open()
    try { $stream.CopyTo($buffer) } finally { $stream.Dispose() }
    return $buffer.ToArray()
}

function Save-XmlToByteArray {
    <#
        Serialises a document without re-indenting: these files are consumed by SqlPackage,
        not read by a human, and Indent would rewrite whitespace the model was stored with.
        The BOM is preserved per part - model.xml has none, Origin.xml has one - because the
        checksum is computed over the exact bytes.
    #>
    [CmdletBinding()]
    [OutputType([byte[]])]
    param(
        [Parameter(Mandatory)] [System.Xml.XmlDocument]$Document,
        [Parameter()] [switch]$WithBom
    )

    $settings = [System.Xml.XmlWriterSettings]::new()
    $settings.Encoding           = [System.Text.UTF8Encoding]::new([bool]$WithBom)
    $settings.Indent             = $false
    $settings.OmitXmlDeclaration = $false

    $buffer = [System.IO.MemoryStream]::new()
    $writer = [System.Xml.XmlWriter]::Create($buffer, $settings)
    try { $Document.Save($writer) } finally { $writer.Dispose() }
    return $buffer.ToArray()
}

function ConvertTo-Sha256Hex {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] [byte[]]$Bytes
    )

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
}

function Get-ZipArchiveFormat {
    <#
        Reports whether an archive uses the classic 32-bit ZIP layout or ZIP64, by looking
        for the ZIP64 end-of-central-directory record and its locator in the file tail.

        This matters: a 32-bit archive caps entry size, total size and offsets at 4 GB and
        the entry count at 65535. SqlPackage emits ZIP64 by itself once a BACPAC passes
        those limits, so ZIP64 is legitimate - but a rewrite must not promote a 32-bit
        archive to ZIP64 gratuitously, which is why the caller compares before and after.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        # The end-of-central-directory sits in the last 22 bytes plus up to 64 KB of comment.
        $tailLength = [int][Math]::Min([long]70000, $stream.Length)
        $null = $stream.Seek(-$tailLength, [System.IO.SeekOrigin]::End)

        $tail = [byte[]]::new($tailLength)
        $filled = 0
        while ($filled -lt $tailLength) {
            $read = $stream.Read($tail, $filled, $tailLength - $filled)
            if ($read -le 0) { break }
            $filled += $read
        }

        # PK\006\006 = ZIP64 end of central directory, PK\006\007 = its locator.
        foreach ($signature in @(([byte[]](0x50, 0x4B, 0x06, 0x06)), ([byte[]](0x50, 0x4B, 0x06, 0x07)))) {
            for ($i = $filled - $signature.Length; $i -ge 0; $i--) {
                $found = $true
                for ($j = 0; $j -lt $signature.Length; $j++) {
                    if ($tail[$i + $j] -ne $signature[$j]) { $found = $false; break }
                }
                if ($found) { return 'ZIP64' }
            }
        }
        return 'ZIP32'
    }
    finally { $stream.Dispose() }
}

#endregion

#region setup ---------------------------------------------------------------------------

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$BacpacPath = (Resolve-Path -LiteralPath $BacpacPath).Path

# Accept "Table", "schema.Table" and "[schema].[Table]" alike. Without this, pasting the
# qualified name straight out of a log line silently searches for a table literally called
# "[dbo].[WhoIsActive]" and reports it as missing.
$identifier = ConvertTo-TableIdentifier -Name $TableName -DefaultSchema $SchemaName
$SchemaName = $identifier.Schema
$TableName  = $identifier.Table

if ($ListTables) {
    $tableReport = @()
    $listZip = [System.IO.Compression.ZipFile]::OpenRead($BacpacPath)
    try {
        $dataSize  = @{}
        $dataParts = @{}
        foreach ($entry in $listZip.Entries) {
            if (-not $entry.FullName.StartsWith('Data/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $folder = ($entry.FullName -split '/')[1]
            if (-not $dataSize.ContainsKey($folder)) { $dataSize[$folder] = [long]0; $dataParts[$folder] = 0 }
            $dataSize[$folder] += $entry.CompressedLength
            $dataParts[$folder]++
        }

        $listModelBytes = Get-ArchiveEntryContent -Archive $listZip -EntryName 'model.xml'
        $listDoc = [System.Xml.XmlDocument]::new()
        $listDoc.Load([System.IO.MemoryStream]::new($listModelBytes))

        foreach ($node in $listDoc.SelectNodes("//*[local-name()='Element' and @Type='SqlTable']")) {
            $name  = $node.GetAttribute('Name')
            $split = ConvertTo-TableIdentifier -Name $name -DefaultSchema 'dbo'
            $key   = "$($split.Schema).$($split.Table)"
            $tableReport += [pscustomobject]@{
                Table        = $name
                DataParts    = $(if ($dataParts.ContainsKey($key)) { $dataParts[$key] } else { 0 })
                CompressedMB = $(if ($dataSize.ContainsKey($key)) { [math]::Round($dataSize[$key] / 1MB, 1) } else { 0 })
            }
        }
    }
    finally { $listZip.Dispose() }

    Write-ToolLog "Tables in $BacpacPath : $($tableReport.Count)"
    $tableReport | Sort-Object CompressedMB -Descending
    return
}

if (-not $OutputPath) {
    $sourceFolder = Split-Path -Parent $BacpacPath
    $sourceName   = [System.IO.Path]::GetFileNameWithoutExtension($BacpacPath)
    $OutputPath   = Join-Path $sourceFolder "$sourceName.without-$SchemaName.$TableName.bacpac"
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

if ([System.IO.Path]::GetFullPath($BacpacPath) -ieq $OutputPath) {
    $reason = 'Output path is the source file. The source is never modified in place.'
    Write-ToolLog $reason -Level ERROR
    throw $reason
}
if ((Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    $reason = "Output already exists: $OutputPath (use -Force to overwrite)"
    Write-ToolLog $reason -Level ERROR
    throw $reason
}

$qualifiedName   = "[$SchemaName].[$TableName]"
$dataPartPrefix  = "Data/$SchemaName.$TableName/"
$relTargetPrefix = "/Data/$SchemaName.$TableName/"
$sourceSizeMb    = [math]::Round((Get-Item -LiteralPath $BacpacPath).Length / 1MB, 1)
$sourceZipFormat = Get-ZipArchiveFormat -Path $BacpacPath
$stopwatch       = [System.Diagnostics.Stopwatch]::StartNew()

Write-ToolLog '========================================================================'
Write-ToolLog 'New-BacpacWithoutTable'
Write-ToolLog "  Source : $BacpacPath ($sourceSizeMb MB)"
Write-ToolLog "  Format : $sourceZipFormat"
Write-ToolLog "  Table  : $qualifiedName"
Write-ToolLog "  Output : $OutputPath"
Write-ToolLog "  Deflate: $CompressionLevel"
Write-ToolLog '========================================================================'

#endregion

#region DacFx availability ---------------------------------------------------------------

$dacFxReady  = $false
$sourceCheck = @{ IsValid = $false; ObjectCount = 0; Reason = $null }

if ($SkipValidation) {
    Write-ToolLog 'Validation skipped by request. The result will not be checked by DacFx.' -Level WARN
}
else {
    if (-not $DacFxPath) { $DacFxPath = Find-DacFxFolder }
    if (-not $DacFxPath) {
        $reason = 'No .NET 8 DacFx build found. Install the sqlpackage dotnet tool (dotnet tool install -g microsoft.sqlpackage), pass -DacFxPath, or run with -SkipValidation.'
        Write-ToolLog $reason -Level ERROR
        throw $reason
    }
    $dacFxReady = Import-DacFx -Folder $DacFxPath
    if (-not $dacFxReady) {
        $reason = "Microsoft.SqlServer.Dac.dll not found under: $DacFxPath"
        Write-ToolLog $reason -Level ERROR
        throw $reason
    }
    Write-ToolLog "  DacFx  : $DacFxPath"

    $sourceCheck = Test-DacPackageModel -PackagePath $BacpacPath
    if (-not $sourceCheck.IsValid) {
        $reason = "DacFx cannot load the source package: $($sourceCheck.Reason)"
        Write-ToolLog $reason -Level ERROR
        throw $reason
    }
    Write-ToolLog "  Source model loaded by DacFx: $($sourceCheck.ObjectCount) user-defined object(s)."
}

#endregion

#region read the control parts ------------------------------------------------------------

Write-ToolLog 'Reading model.xml, Origin.xml and _rels/.rels...'

$entryNames    = @()
$dataPartCount = 0
$dataPartBytes = [long]0

$zip = [System.IO.Compression.ZipFile]::OpenRead($BacpacPath)
try {
    $entryNames  = @($zip.Entries | ForEach-Object { $_.FullName })
    $modelBytes  = Get-ArchiveEntryContent -Archive $zip -EntryName 'model.xml'
    $originBytes = Get-ArchiveEntryContent -Archive $zip -EntryName 'Origin.xml'
    $relsBytes   = Get-ArchiveEntryContent -Archive $zip -EntryName '_rels/.rels'

    foreach ($entry in $zip.Entries) {
        if ($entry.FullName.StartsWith($dataPartPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $dataPartCount++
            $dataPartBytes += $entry.CompressedLength
        }
    }
}
finally { $zip.Dispose() }

Write-ToolLog "  Entries: $($entryNames.Count)  model.xml: $([math]::Round($modelBytes.Length / 1MB, 1)) MB  _rels/.rels: $([math]::Round($relsBytes.Length / 1MB, 1)) MB"

#endregion

#region model.xml -------------------------------------------------------------------------

$modelDoc = [System.Xml.XmlDocument]::new()
$modelDoc.PreserveWhitespace = $true
$modelDoc.Load([System.IO.MemoryStream]::new($modelBytes))

# Top-level only: columns and anything nested inside the table element leave with their
# parent. What has to be found here are the SIBLINGS that belong to the table, and they
# come in two shapes.
#   by name        [dbo].[WhoIsActive].[cx_collection_time] - an index carrying the table
#                  in its own name
#   by ownership   [dbo].[PK_Orders] - a constraint named independently, tied to its table
#                  only through a DefiningTable / IndexedObject / Parent relationship
# Missing the second shape is not a cosmetic gap: it makes every table with a primary key
# look externally referenced, and the run would be refused.
$removalSet = @()
foreach ($node in $modelDoc.SelectNodes("//*[local-name()='Element' and @Name]")) {
    if ($node.ParentNode.LocalName -ne 'Model') { continue }

    $belongsByName  = Test-NameBelongsToTable -Name $node.GetAttribute('Name') -Qualified $qualifiedName
    $belongsByOwner = (Get-ElementOwnerName -Element $node) -ieq $qualifiedName

    if ($belongsByName -or $belongsByOwner) { $removalSet += $node }
}

if ($removalSet.Count -eq 0) {
    $reason = "Table $qualifiedName was not found in the model. Nothing written."
    Write-ToolLog $reason -Level ERROR
    throw $reason
}

Write-ToolLog "Removal set - $($removalSet.Count) top-level element(s):"
foreach ($node in $removalSet) {
    Write-ToolLog "    $($node.GetAttribute('Type'))  $($node.GetAttribute('Name'))"
}
Write-ToolLog "Data parts for the table: $dataPartCount ($([math]::Round($dataPartBytes / 1MB, 1)) MB compressed)"

foreach ($node in $removalSet) { $null = $node.ParentNode.RemoveChild($node) }

# Anything outside the table that still names it would be left dangling. Removing those
# would change other objects, so this is a refusal rather than a warning.
$externalReferences = @()
foreach ($node in $modelDoc.SelectNodes("//*[local-name()='References' and @Name]")) {
    if (-not (Test-NameBelongsToTable -Name $node.GetAttribute('Name') -Qualified $qualifiedName)) { continue }
    $owner = $node.ParentNode
    while ($owner -and -not ($owner.LocalName -eq 'Element' -and $owner.GetAttribute('Name'))) {
        $owner = $owner.ParentNode
    }
    $externalReferences += $(if ($owner) { "$($owner.GetAttribute('Type')) $($owner.GetAttribute('Name'))" } else { '(orphan reference)' })
}

if ($externalReferences.Count -gt 0) {
    Write-ToolLog "$($externalReferences.Count) object(s) outside $qualifiedName still reference it. Removing the table would break them:" -Level ERROR
    foreach ($owner in ($externalReferences | Select-Object -Unique | Select-Object -First 20)) {
        Write-ToolLog "    $owner" -Level ERROR
    }
    throw "Cannot remove $qualifiedName - $($externalReferences.Count) external reference(s) would be left dangling. Nothing written."
}
Write-ToolLog '  No external reference to the table remains.'

$newModelBytes = Save-XmlToByteArray -Document $modelDoc    # original model.xml carries no BOM
$newChecksum   = ConvertTo-Sha256Hex -Bytes $newModelBytes

#endregion

#region _rels/.rels -----------------------------------------------------------------------

$relsDoc = [System.Xml.XmlDocument]::new()
$relsDoc.PreserveWhitespace = $true
$relsDoc.Load([System.IO.MemoryStream]::new($relsBytes))

$relsRemoved = 0
foreach ($node in @($relsDoc.SelectNodes("//*[local-name()='Relationship' and @Target]"))) {
    if ($node.GetAttribute('Target').StartsWith($relTargetPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        $null = $node.ParentNode.RemoveChild($node)
        $relsRemoved++
    }
}
Write-ToolLog "_rels/.rels: removed $relsRemoved relationship(s)."
$newRelsBytes = Save-XmlToByteArray -Document $relsDoc

#endregion

#region Origin.xml ------------------------------------------------------------------------

$originDoc = [System.Xml.XmlDocument]::new()
$originDoc.PreserveWhitespace = $true
$originDoc.Load([System.IO.MemoryStream]::new($originBytes))

$originNs = [System.Xml.XmlNamespaceManager]::new($originDoc.NameTable)
$originNs.AddNamespace('d', 'http://schemas.microsoft.com/sqlserver/dac/Serialization/2012/02')

$checksumNode = $originDoc.SelectSingleNode('//d:Checksums/d:Checksum', $originNs)
if (-not $checksumNode) {
    $reason = 'Origin.xml carries no model.xml checksum. Refusing to guess at its format.'
    Write-ToolLog $reason -Level ERROR
    throw $reason
}

# Sanity check before trusting the recomputation: the stored value must be a plain SHA256
# over the ORIGINAL model.xml bytes. If it is not, the format is not what is assumed here.
$originalChecksum = ConvertTo-Sha256Hex -Bytes $modelBytes
if ($checksumNode.InnerText -ine $originalChecksum) {
    $reason = "Origin.xml checksum does not match a SHA256 of the original model.xml. Stored=$($checksumNode.InnerText) Computed=$originalChecksum. Refusing to rewrite it."
    Write-ToolLog $reason -Level ERROR
    throw $reason
}
Write-ToolLog '  Origin.xml checksum verified against the original model.xml.'
$checksumNode.InnerText = $newChecksum

$dataPhaseNode = $originDoc.SelectSingleNode('//d:DataPhaseTables', $originNs)
if ($dataPhaseNode) {
    $before = @($dataPhaseNode.InnerText -split ',' | Where-Object { $_.Trim() })
    $after  = @($before | Where-Object { $_.Trim() -ine "$SchemaName.$TableName" })
    Write-ToolLog "  Origin.xml DataPhaseTables: $($before.Count) -> $($after.Count) entries."
    $dataPhaseNode.InnerText = $after -join ','
}
else {
    Write-ToolLog '  Origin.xml has no DataPhaseTables section.'
}

$newOriginBytes = Save-XmlToByteArray -Document $originDoc -WithBom   # Origin.xml carries a BOM

#endregion

if (-not $PSCmdlet.ShouldProcess($OutputPath, "Write a BACPAC without $qualifiedName")) {
    Write-ToolLog 'WhatIf: nothing written.'
    return
}

#region rebuild the archive ---------------------------------------------------------------

Write-ToolLog 'Writing the new archive (data parts are streamed, never buffered)...'

$level   = [System.IO.Compression.CompressionLevel]::$CompressionLevel
$copied  = 0
$dropped = 0

if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }

$sourceZip = [System.IO.Compression.ZipFile]::OpenRead($BacpacPath)
try {
    $outputZip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $sourceZip.Entries) {
            $name = $entry.FullName
            if ($name.StartsWith($dataPartPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $dropped++
                continue
            }

            $target = $outputZip.CreateEntry($name, $level)
            $targetStream = $target.Open()
            try {
                switch -Regex ($name) {
                    '^model\.xml$'   { $targetStream.Write($newModelBytes,  0, $newModelBytes.Length);  break }
                    '^Origin\.xml$'  { $targetStream.Write($newOriginBytes, 0, $newOriginBytes.Length); break }
                    '^_rels/\.rels$' { $targetStream.Write($newRelsBytes,   0, $newRelsBytes.Length);   break }
                    default {
                        $sourceStream = $entry.Open()
                        try { $sourceStream.CopyTo($targetStream) } finally { $sourceStream.Dispose() }
                    }
                }
            }
            finally { $targetStream.Dispose() }

            $copied++
            if ($copied % 500 -eq 0) {
                Write-ToolLog "    $copied entries written ($($stopwatch.Elapsed.ToString('hh\:mm\:ss')))..."
            }
        }
    }
    finally { $outputZip.Dispose() }
}
finally { $sourceZip.Dispose() }

$outputSizeMb = [math]::Round((Get-Item -LiteralPath $OutputPath).Length / 1MB, 1)
Write-ToolLog "  Entries written: $copied   data parts dropped: $dropped"

# A 32-bit source must stay 32-bit. .NET only emits ZIP64 when a limit is actually crossed,
# but recompression can change sizes, so the outcome is verified rather than assumed.
$outputZipFormat = Get-ZipArchiveFormat -Path $OutputPath
Write-ToolLog "  Archive format: source $sourceZipFormat -> output $outputZipFormat"
if ($sourceZipFormat -eq 'ZIP32' -and $outputZipFormat -eq 'ZIP64') {
    Write-ToolLog 'The output was promoted to ZIP64 while the source was a classic 32-bit archive.' -Level ERROR
    Write-ToolLog "  Left in place for inspection: $OutputPath" -Level ERROR
    throw 'Refusing to hand over a ZIP64 archive built from a 32-bit source. Retry with -CompressionLevel Optimal, or inspect the entry count and total size.'
}

#endregion

#region validate the result ----------------------------------------------------------------

if ($dacFxReady) {
    Write-ToolLog 'Validating the result with DacFx...'
    $resultCheck = Test-DacPackageModel -PackagePath $OutputPath
    if (-not $resultCheck.IsValid) {
        Write-ToolLog "The produced BACPAC does not load: $($resultCheck.Reason)" -Level ERROR
        Write-ToolLog "  Left in place for inspection: $OutputPath" -Level ERROR
        throw "Validation failed - the produced BACPAC is not usable: $($resultCheck.Reason)"
    }
    Write-ToolLog "  DacFx loaded the result: $($resultCheck.ObjectCount) user-defined object(s) (source had $($sourceCheck.ObjectCount))."
}
else {
    Write-ToolLog 'Result NOT validated - no DacFx load was performed.' -Level WARN
}

#endregion

$stopwatch.Stop()

Write-ToolLog '========================================================================'
Write-ToolLog "Written: $OutputPath"
Write-ToolLog "  Size    : $outputSizeMb MB (source $sourceSizeMb MB)"
Write-ToolLog "  Format  : $outputZipFormat (source $sourceZipFormat)"
Write-ToolLog "  Removed : $qualifiedName, $($removalSet.Count) model element(s), $dropped data part(s), $relsRemoved relationship(s)"
Write-ToolLog "  Elapsed : $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
Write-ToolLog "  Source unchanged: $BacpacPath"
Write-ToolLog '========================================================================'
Write-ToolLog 'Rewriting a BACPAC is not supported by Microsoft. Import the result against an empty, disposable target database before trusting it.' -Level WARN
