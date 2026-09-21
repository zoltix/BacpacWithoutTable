#!/usr/bin/env pwsh
#requires -Version 7.0

<#
.SYNOPSIS
Checks a BACPAC offline: archive integrity, CRC of every part, model checksum, schema model
load, and internal consistency. Read-only, no database, no server.

.DESCRIPTION
Answers "is this file intact and coherent?" without importing it. Eight checks, each reported
separately so a failure says what is wrong rather than just that something is.

  1  ARCHIVE     the ZIP opens and its central directory is readable
  1b COMPAT      whether the OLD .NET Framework SqlPackage can read it at all - see below
  2  PARTS       model.xml, Origin.xml, [Content_Types].xml and _rels/.rels are present
  3  CRC         every entry decompresses and matches its stored CRC32
  4  CHECKSUM    Origin.xml holds a SHA256 of model.xml that matches the model.xml bytes
  5  MODEL       model.xml parses, and DacFx loads the schema model from the package
  6  DATA        every data folder corresponds to a table that still exists in the model and
                 is declared in DataPhaseTables
  7  RELS        the _rels/.rels relationships and the TableData parts agree, both directions

WHAT EACH CHECK CATCHES - they fail for different reasons and are not interchangeable:

  COMPAT is not about the file being sound - it is about who can read it. Past 4 GB a BACPAC
  is ZIP64, and the .NET Framework SqlPackage (the x86 build from DacFramework.msi under
  "Program Files\Microsoft SQL Server\<version>\DAC\bin") reads through System.IO.Packaging,
  whose ZIP reader has no ZIP64 support. It fails in ZipIOLocalFileBlock.Validate with

      System.IO.FileFormatException: File contains corrupted data.

  which names the wrong culprit: the package is intact, the reader cannot parse it. This is
  dotnet/runtime issue 94899, fixed only for .NET 9. Every check here reads through
  System.IO.Compression, which handles ZIP64 - so without COMPAT this tool would return a
  clean verdict on a file SqlPackage then refuses. The answer is always the .NET 8 build,
  the "sqlpackage" dotnet tool. Note the trap: the x86 build sits under "Program Files",
  not "Program Files (x86)", so its location gives no hint of its architecture.

  CRC is the only one that detects physical corruption: a truncated copy, a bad transfer, a
  failing disk. It is also the expensive one, because it decompresses the whole archive.

  CHECKSUM detects a model.xml that was altered without its checksum being updated. The DacFx
  load catches it too, but only that one: BacPackage.Load does NOT verify the checksum, a
  package with a deliberately zeroed one still loads. This check therefore stands on its own
  under -SkipModelLoad, and it prints both values, where DacFx only says they differ.

  MODEL catches a schema model that no longer resolves, typically an element removed while
  something still references it. A package can pass ARCHIVE, PARTS and CRC and still fail here.

  DATA and RELS catch an incomplete edit: data parts left behind for a table that is gone from
  the model, or relationships pointing at parts that no longer exist. RELS covers only
  TableData-*.BCP parts - BlobData-<guid>.BIN parts hold oversized binary column values and
  are referenced from inside the BCP data, never from .rels.

.PARAMETER BacpacPath
One or more BACPAC files to check. Accepts pipeline input and wildcards.

.PARAMETER Quick
Skip the CRC check. Everything else still runs. Use it for a fast structural verdict on a
large file - a 15 GB package takes minutes to decompress in full.

.PARAMETER DacFxPath
Folder holding the .NET 8 build of Microsoft.SqlServer.Dac.dll. Auto-detected from the
sqlpackage dotnet tool and from the NuGet cache when omitted. The assemblies under
"Program Files\Microsoft SQL Server\<version>\DAC\bin" target .NET Framework and CANNOT be
loaded by PowerShell 7 - they fail with a type initializer error.

.PARAMETER SkipModelLoad
Skip the DacFx part of check 5. model.xml is still parsed and the table list still built.

.PARAMETER LogFile
Optional path to also append the console output to. The parent folder is created if needed.

.EXAMPLE
./Test-Bacpac.ps1 -BacpacPath 'C:\backup\mydb.bacpac'

.EXAMPLE
./Test-Bacpac.ps1 -BacpacPath 'C:\backup\*.bacpac' -Quick
# Structural verdict on a whole folder, without decompressing anything.

.EXAMPLE
Get-ChildItem C:\backup\*.bacpac | ./Test-Bacpac.ps1 | Where-Object Verdict -ne 'OK'
# Full check on every file, keeping only what failed.

.NOTES
Self-contained: PowerShell 7 and, for the MODEL check, any .NET 8 DacFx build - the
"sqlpackage" dotnet tool is the easiest source.

This is a file-level verdict, not a guarantee that the package will import. Only an actual
import against a disposable target proves that.

MIT licensed. See LICENSE.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('FullName')]
    [string[]]$BacpacPath,

    [Parameter()]
    [switch]$Quick,

    [Parameter()]
    [string]$DacFxPath,

    [Parameter()]
    [switch]$SkipModelLoad,

    [Parameter()]
    [string]$LogFile
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    #region logging -----------------------------------------------------------------------

    $script:LogFilePath = ''
    if ($LogFile) {
        $logFolder = Split-Path -Parent $LogFile
        if ($logFolder -and -not (Test-Path -LiteralPath $logFolder)) {
            $null = New-Item -ItemType Directory -Path $logFolder -Force
        }
        $script:LogFilePath = $LogFile
    }

    function Write-ToolLog {
        <#
            Console is the product here, so the output goes to the host deliberately. A copy
            is appended to -LogFile when one was given.
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
            Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding utf8
        }
    }

    #endregion

    #region helpers -----------------------------------------------------------------------

    function Find-DacFxFolder {
        <#
            Locates a .NET 8 DacFx build. The "Program Files\Microsoft SQL Server\<v>\DAC\bin"
            assemblies are deliberately NOT considered: they target .NET Framework and throw a
            type initializer error under PowerShell 7. The file name is part of the -Path
            pattern because Get-ChildItem returns nothing when -Filter meets a wildcarded path.
        #>
        [CmdletBinding()]
        [OutputType([string])]
        param()

        $candidates = @(
            Join-Path $HOME '.dotnet/tools/.store/microsoft.sqlpackage/*/microsoft.sqlpackage/*/tools/net8.0/any'
            Join-Path $HOME '.nuget/packages/microsoft.sqlserver.dacfx/*/lib/net8.0'
        )
        foreach ($pattern in $candidates) {
            $hit = Get-ChildItem -Path (Join-Path $pattern 'Microsoft.SqlServer.Dac.dll') -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | Select-Object -First 1
            if ($hit) { return $hit.DirectoryName }
        }
        return $null
    }

    function Import-DacFx {
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

    function Get-ZipArchiveFormat {
        <#
            Classic 32-bit ZIP or ZIP64, read from the end-of-central-directory records.
            SqlPackage emits ZIP64 by itself past 4 GB, so ZIP64 is not a fault, but knowing
            which one a file uses decides whether the old x86 SqlPackage can open it.
        #>
        [CmdletBinding()]
        [OutputType([string])]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $stream = [System.IO.File]::OpenRead($Path)
        try {
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

    function Get-ArchiveEntryContent {
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

    #endregion

    #region DacFx -------------------------------------------------------------------------

    $dacFxReady = $false
    if (-not $SkipModelLoad) {
        if (-not $DacFxPath) { $DacFxPath = Find-DacFxFolder }
        if ($DacFxPath) { $dacFxReady = Import-DacFx -Folder $DacFxPath }
        if (-not $dacFxReady) {
            Write-ToolLog 'No .NET 8 DacFx build found - the MODEL check will be partial. Install the sqlpackage dotnet tool, pass -DacFxPath, or -SkipModelLoad to silence this.' -Level WARN
        }
    }

    #endregion

    $overall = @()
}

process {
    foreach ($path in $BacpacPath) {
        foreach ($resolved in @(Resolve-Path -Path $path -ErrorAction SilentlyContinue)) {

            $file     = $resolved.Path
            $fileName = Split-Path $file -Leaf
            $sizeMb   = [math]::Round((Get-Item -LiteralPath $file).Length / 1MB, 1)
            $watch    = [System.Diagnostics.Stopwatch]::StartNew()
            $results  = [ordered]@{}

            Write-ToolLog '========================================================================'
            Write-ToolLog "Test-Bacpac: $fileName ($sizeMb MB)"
            Write-ToolLog '========================================================================'

            $zip = $null
            try {
                # --- 1. ARCHIVE + COMPAT ----------------------------------------------------
                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
                    $entryCount = $zip.Entries.Count
                    $zipFormat  = Get-ZipArchiveFormat -Path $file
                    $results['ARCHIVE'] = 'PASS'
                    Write-ToolLog "  [PASS] ARCHIVE  - $entryCount entries, $zipFormat"

                    # See the COMPAT paragraph in the help: past 4 GB a package is ZIP64 and
                    # the x86 SqlPackage reports it as corrupt when it simply cannot read it.
                    if ($zipFormat -eq 'ZIP64' -or $sizeMb -gt 4096) {
                        $results['COMPAT'] = 'WARN'
                        Write-ToolLog "  [WARN] COMPAT   - $zipFormat / $sizeMb MB: the .NET Framework (x86) SqlPackage cannot read this and reports 'File contains corrupted data'. Use the .NET 8 build - the 'sqlpackage' dotnet tool - not the one under Program Files\Microsoft SQL Server\<version>\DAC\bin." -Level WARN
                    }
                    else {
                        $results['COMPAT'] = 'PASS'
                        Write-ToolLog "  [PASS] COMPAT   - $zipFormat under 4 GB, readable by either SqlPackage build"
                    }
                }
                catch {
                    # A file still being written is not a damaged file. Saying "corrupt" about
                    # an export in flight would send someone hunting a fault that is not there.
                    if ($_.Exception.Message -match 'being used by another process') {
                        $results['ARCHIVE'] = 'LOCKED'
                        Write-ToolLog '  [LOCKED] ARCHIVE - the file is open in another process, probably still being written. Not checked.' -Level WARN
                        throw 'locked'
                    }
                    $results['ARCHIVE'] = 'FAIL'
                    Write-ToolLog "  [FAIL] ARCHIVE  - $($_.Exception.Message)" -Level ERROR
                    throw 'archive'
                }

                # --- 2. PARTS ---------------------------------------------------------------
                $required = @('model.xml', 'Origin.xml', '[Content_Types].xml', '_rels/.rels')
                $present  = @($zip.Entries | ForEach-Object { $_.FullName })
                $missing  = @($required | Where-Object { $part = $_; -not ($present | Where-Object { $_ -ieq $part }) })
                if ($missing.Count -eq 0) {
                    $results['PARTS'] = 'PASS'
                    Write-ToolLog '  [PASS] PARTS    - model.xml, Origin.xml, [Content_Types].xml, _rels/.rels'
                }
                else {
                    $results['PARTS'] = 'FAIL'
                    Write-ToolLog "  [FAIL] PARTS    - missing: $($missing -join ', ')" -Level ERROR
                    throw 'parts'
                }

                # --- 3. CRC -----------------------------------------------------------------
                if ($Quick) {
                    $results['CRC'] = 'SKIP'
                    Write-ToolLog '  [SKIP] CRC      - -Quick requested'
                }
                else {
                    # Reading an entry to its end makes ZipArchive verify the stored CRC32 and
                    # throw on mismatch. A 1 MB sink keeps memory flat whatever the file size.
                    $badEntries = @()
                    $checked = 0
                    $sink = [byte[]]::new(1MB)
                    foreach ($entry in $zip.Entries) {
                        try {
                            $stream = $entry.Open()
                            try { while ($stream.Read($sink, 0, $sink.Length) -gt 0) { } }
                            finally { $stream.Dispose() }
                        }
                        catch { $badEntries += "$($entry.FullName): $($_.Exception.Message)" }

                        $checked++
                        if ($checked % 2000 -eq 0) {
                            Write-ToolLog "         ... $checked/$entryCount entries verified ($($watch.Elapsed.ToString('hh\:mm\:ss')))"
                        }
                    }
                    if ($badEntries.Count -eq 0) {
                        $results['CRC'] = 'PASS'
                        Write-ToolLog "  [PASS] CRC      - $checked entries decompressed, all CRC32 correct"
                    }
                    else {
                        $results['CRC'] = 'FAIL'
                        Write-ToolLog "  [FAIL] CRC      - $($badEntries.Count) corrupt entry/entries:" -Level ERROR
                        foreach ($bad in ($badEntries | Select-Object -First 10)) {
                            Write-ToolLog "             $bad" -Level ERROR
                        }
                    }
                }

                # --- 4. CHECKSUM ------------------------------------------------------------
                $modelBytes  = Get-ArchiveEntryContent -Archive $zip -EntryName 'model.xml'
                $originBytes = Get-ArchiveEntryContent -Archive $zip -EntryName 'Origin.xml'

                $originDoc = [System.Xml.XmlDocument]::new()
                $originDoc.Load([System.IO.MemoryStream]::new($originBytes))
                $originNs = [System.Xml.XmlNamespaceManager]::new($originDoc.NameTable)
                $originNs.AddNamespace('d', 'http://schemas.microsoft.com/sqlserver/dac/Serialization/2012/02')

                $checksumNode = $originDoc.SelectSingleNode('//d:Checksums/d:Checksum', $originNs)
                if (-not $checksumNode) {
                    $results['CHECKSUM'] = 'WARN'
                    Write-ToolLog '  [WARN] CHECKSUM - Origin.xml carries no model.xml checksum' -Level WARN
                }
                else {
                    $computed = ConvertTo-Sha256Hex -Bytes $modelBytes
                    if ($checksumNode.InnerText -ieq $computed) {
                        $results['CHECKSUM'] = 'PASS'
                        Write-ToolLog "  [PASS] CHECKSUM - model.xml matches Origin.xml ($($computed.Substring(0, 16))...)"
                    }
                    else {
                        $results['CHECKSUM'] = 'FAIL'
                        Write-ToolLog '  [FAIL] CHECKSUM - model.xml does NOT match the checksum in Origin.xml' -Level ERROR
                        Write-ToolLog "             declared: $($checksumNode.InnerText)" -Level ERROR
                        Write-ToolLog "             computed: $computed" -Level ERROR
                    }
                }

                # --- 5. MODEL ---------------------------------------------------------------
                $modelDoc = [System.Xml.XmlDocument]::new()
                try { $modelDoc.Load([System.IO.MemoryStream]::new($modelBytes)) }
                catch {
                    $results['MODEL'] = 'FAIL'
                    Write-ToolLog "  [FAIL] MODEL    - model.xml is not well-formed XML: $($_.Exception.Message)" -Level ERROR
                    throw 'model'
                }

                $modelTables = @{}
                foreach ($node in $modelDoc.SelectNodes("//*[local-name()='Element' and @Type='SqlTable']")) {
                    $name = $node.GetAttribute('Name')
                    $bits = [regex]::Matches($name, '\[((?:[^\]]|\]\])*)\]')
                    if ($bits.Count -ge 2) {
                        $modelTables["$($bits[0].Groups[1].Value).$($bits[1].Groups[1].Value)"] = $true
                    }
                }

                if ($dacFxReady) {
                    try {
                        $model = [Microsoft.SqlServer.Dac.Model.TSqlModel]::new($file)
                        try {
                            $objectCount = @($model.GetObjects([Microsoft.SqlServer.Dac.Model.DacQueryScopes]::UserDefined)).Count
                            $results['MODEL'] = 'PASS'
                            Write-ToolLog "  [PASS] MODEL    - DacFx loaded the schema: $objectCount user-defined object(s), $($modelTables.Count) table(s)"
                        }
                        finally { $model.Dispose() }
                    }
                    catch {
                        $inner = $_.Exception
                        while ($inner.InnerException) { $inner = $inner.InnerException }
                        $results['MODEL'] = 'FAIL'
                        Write-ToolLog "  [FAIL] MODEL    - DacFx refused the schema model: $($inner.Message)" -Level ERROR
                    }
                }
                else {
                    $results['MODEL'] = 'SKIP'
                    Write-ToolLog "  [SKIP] MODEL    - no DacFx load; model.xml is well-formed, $($modelTables.Count) table(s)"
                }

                # --- 6. DATA ----------------------------------------------------------------
                $dataFolders = @{}
                foreach ($entry in $zip.Entries) {
                    if (-not $entry.FullName.StartsWith('Data/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                    $dataFolders[($entry.FullName -split '/')[1]] = $true
                }

                $declared = @()
                $dataPhaseNode = $originDoc.SelectSingleNode('//d:DataPhaseTables', $originNs)
                if ($dataPhaseNode) {
                    $declared = @($dataPhaseNode.InnerText -split ',' | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() })
                }

                # A declared table with no data folder is normal: it was empty at export time.
                # The faults are a data folder for a table the model no longer has, and a data
                # folder that nothing declares.
                $orphanData = @($dataFolders.Keys | Where-Object { -not $modelTables.ContainsKey($_) })
                $undeclared = @($dataFolders.Keys | Where-Object { $declared -notcontains $_ })

                if ($orphanData.Count -eq 0 -and $undeclared.Count -eq 0) {
                    $results['DATA'] = 'PASS'
                    Write-ToolLog "  [PASS] DATA     - $($dataFolders.Count) data folder(s), all backed by a table in the model and declared in DataPhaseTables"
                }
                else {
                    $results['DATA'] = 'FAIL'
                    if ($orphanData.Count -gt 0) {
                        Write-ToolLog "  [FAIL] DATA     - $($orphanData.Count) data folder(s) for table(s) absent from the model:" -Level ERROR
                        foreach ($item in ($orphanData | Select-Object -First 10)) { Write-ToolLog "             $item" -Level ERROR }
                    }
                    if ($undeclared.Count -gt 0) {
                        Write-ToolLog "  [FAIL] DATA     - $($undeclared.Count) data folder(s) missing from DataPhaseTables:" -Level ERROR
                        foreach ($item in ($undeclared | Select-Object -First 10)) { Write-ToolLog "             $item" -Level ERROR }
                    }
                }

                # --- 7. RELS ----------------------------------------------------------------
                $relsBytes = Get-ArchiveEntryContent -Archive $zip -EntryName '_rels/.rels'
                $relsDoc = [System.Xml.XmlDocument]::new()
                $relsDoc.Load([System.IO.MemoryStream]::new($relsBytes))

                # Only TableData parts are declared in _rels/.rels. BlobData-<guid>.BIN parts
                # hold oversized binary column values and are referenced from inside the BCP
                # data instead, never from .rels - verified on an untouched 15 GB export:
                # 18324 TableData parts, 18324 relationships, and 126 BlobData parts with no
                # relationship at all. Requiring coverage for those reports healthy packages
                # as broken, which is exactly what an earlier version of this check did.
                $tableDataNames = @{}
                $blobDataCount  = 0
                $otherDataParts = @()
                foreach ($entry in $zip.Entries) {
                    if (-not $entry.FullName.StartsWith('Data/', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                    if ($entry.Name -like 'TableData-*')     { $tableDataNames["/$($entry.FullName)"] = $true }
                    elseif ($entry.Name -like 'BlobData-*')  { $blobDataCount++ }
                    elseif ($entry.Name)                     { $otherDataParts += $entry.FullName }
                }

                $relTargets = @{}
                foreach ($node in $relsDoc.SelectNodes("//*[local-name()='Relationship' and @Target]")) {
                    $target = $node.GetAttribute('Target')
                    if ($target.StartsWith('/Data/', [System.StringComparison]::OrdinalIgnoreCase)) { $relTargets[$target] = $true }
                }

                $danglingRels = @($relTargets.Keys | Where-Object { -not $tableDataNames.ContainsKey($_) })
                $unlinkedData = @($tableDataNames.Keys | Where-Object { -not $relTargets.ContainsKey($_) })

                if ($danglingRels.Count -eq 0 -and $unlinkedData.Count -eq 0) {
                    $results['RELS'] = 'PASS'
                    Write-ToolLog "  [PASS] RELS     - $($relTargets.Count) relationship(s) match $($tableDataNames.Count) TableData part(s); $blobDataCount BlobData part(s) carry no relationship by design"
                }
                else {
                    $results['RELS'] = 'FAIL'
                    if ($danglingRels.Count -gt 0) {
                        Write-ToolLog "  [FAIL] RELS     - $($danglingRels.Count) relationship(s) point at a part that does not exist:" -Level ERROR
                        foreach ($item in ($danglingRels | Select-Object -First 10)) { Write-ToolLog "             $item" -Level ERROR }
                    }
                    if ($unlinkedData.Count -gt 0) {
                        Write-ToolLog "  [FAIL] RELS     - $($unlinkedData.Count) TableData part(s) that no relationship declares:" -Level ERROR
                        foreach ($item in ($unlinkedData | Select-Object -First 10)) { Write-ToolLog "             $item" -Level ERROR }
                    }
                }

                if ($otherDataParts.Count -gt 0) {
                    Write-ToolLog "  [WARN] RELS     - $($otherDataParts.Count) data part(s) follow neither the TableData nor the BlobData naming, and were not checked:" -Level WARN
                    foreach ($item in ($otherDataParts | Select-Object -First 5)) { Write-ToolLog "             $item" -Level WARN }
                }
            }
            catch {
                # A thrown stage name means a fatal check already logged its own detail; any
                # other exception is unexpected and is recorded as such.
                if ($_.Exception.Message -notin @('archive', 'parts', 'model', 'locked')) {
                    Write-ToolLog "  [FAIL] unexpected error: $($_.Exception.Message)" -Level ERROR
                    $results['ERROR'] = 'FAIL'
                }
            }
            finally {
                if ($zip) { $zip.Dispose() }
            }

            $watch.Stop()
            $failed = @($results.GetEnumerator() | Where-Object { $_.Value -eq 'FAIL' })
            $verdict =
                if ($results.Contains('ARCHIVE') -and $results['ARCHIVE'] -eq 'LOCKED') { 'LOCKED (not checked)' }
                elseif ($failed.Count -eq 0) { 'OK' }
                else { 'CORRUPT / INCONSISTENT' }

            Write-ToolLog "  VERDICT: $verdict   ($($watch.Elapsed.ToString('hh\:mm\:ss')))"
            if ($failed.Count -gt 0) {
                Write-ToolLog "$fileName -> $verdict (failed: $(($failed | ForEach-Object { $_.Key }) -join ', '))" -Level ERROR
            }

            $overall += [pscustomobject]@{
                File     = $fileName
                SizeMB   = $sizeMb
                Verdict  = $verdict
                Archive  = $(if ($results.Contains('ARCHIVE'))  { $results['ARCHIVE'] }  else { '-' })
                Compat   = $(if ($results.Contains('COMPAT'))   { $results['COMPAT'] }   else { '-' })
                Parts    = $(if ($results.Contains('PARTS'))    { $results['PARTS'] }    else { '-' })
                Crc      = $(if ($results.Contains('CRC'))      { $results['CRC'] }      else { '-' })
                Checksum = $(if ($results.Contains('CHECKSUM')) { $results['CHECKSUM'] } else { '-' })
                Model    = $(if ($results.Contains('MODEL'))    { $results['MODEL'] }    else { '-' })
                Data     = $(if ($results.Contains('DATA'))     { $results['DATA'] }     else { '-' })
                Rels     = $(if ($results.Contains('RELS'))     { $results['RELS'] }     else { '-' })
                Elapsed  = $watch.Elapsed.ToString('hh\:mm\:ss')
            }
        }
    }
}

end {
    Write-ToolLog '========================================================================'
    Write-ToolLog "Checked $($overall.Count) file(s)."
    Write-ToolLog '========================================================================'
    $overall
}
