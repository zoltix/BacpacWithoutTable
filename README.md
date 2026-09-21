# bacpac-tools

Two PowerShell scripts for working on a SQL Server or Azure SQL `.bacpac` **as a file**, with
no database and no server involved.

| Script | What it does |
|---|---|
| `New-BacpacWithoutTable.ps1` | A BACPAC in, a new BACPAC out, minus one table, schema and data. The source is opened read-only and keeps its bytes. |
| `Test-Bacpac.ps1` | Checks a BACPAC offline in eight checks and returns a verdict per file. |

Requires PowerShell 7. For the validation steps, any .NET 8 build of DacFx — the `sqlpackage`
dotnet tool is the easiest source:

```bash
dotnet tool install --global microsoft.sqlpackage
```

Both scripts auto-detect it. No other dependency, nothing to import, nothing to install.

---

## Why these exist

A BACPAC that contains a table whose data cannot be imported is unusable, and there is no
supported way to take that table out again. The common case is a monitoring table: the
collection table of the widely used `sp_WhoIsActive` procedure stores query plans in an XML
column, and those plans routinely nest deeper than the 128 levels an import accepts. One such
table makes the whole package fail.

Re-exporting is often impossible — the source database has moved on, or the export is an
archived backup. That leaves editing the package itself.

---

## `New-BacpacWithoutTable.ps1`

### Why it is not built on DacFx

It would be, if DacFx could do it. It cannot, and that was verified rather than assumed:

| Checked | Result |
|---|---|
| `BacPackage` on the shipped assembly (170.2.70) | only `Load`, `Unpack`, `Dispose`. No save, no pack |
| Published API reference, versions 140 to 162 | the same three methods |
| `DacPackage.Load` on a BACPAC | refuses: *Cannot create a DAC package from a file that contains exported data* |
| `BuildPackage` from a BACPAC model | emits a DACPAC: 4 KB out of a 242 KB source, **no `Data/` at all** |

DacFx writes a BACPAC from a **live database** and nowhere else, through `ExportBacpac`. The
only supported route is therefore a round trip through a database — `ImportBacpac`,
`DROP TABLE`, `ExportBacpac` — which needs a SQL server, runs for hours on a large file, and
begins with the very import that is usually the problem. Even the far simpler request of
excluding tables at export time is still an open backlog item upstream
([DacFx issue #233](https://github.com/microsoft/DacFx/issues/233)).

So the archive is rewritten directly, and DacFx is kept where it is authoritative: it loads
the source to prove it is a real DAC package, and **it loads the result**. That output gate is
the safety net, not decoration — it is what catches the dangling-index bug described below.

### What a complete removal involves

Each step is required. Skipping any one of them produces a file that looks fine and is not.

| Step | Why |
|---|---|
| Remove the table **and every top-level element it owns** | A clustered index is a *sibling* of the table, not a child. Removing only the table leaves it dangling and DacFx then rejects the package with `Could not load schema model from package` |
| Resolve ownership through the model, not through names | `[dbo].[PK_Orders]` belongs to `[dbo].[Orders]` but its name says nothing. Ownership is read from the `DefiningTable`, `IndexedObject` and `Parent` relationships. Without this, every table with a primary key would be refused |
| Refuse when anything **outside** still references the table | A foreign key on another table, a view, a procedure body. Removing those would silently change other objects, so it is a stop, not a warning |
| Drop the `Data/<schema>.<table>/*.BCP` parts | The data is what the table costs |
| Drop the matching `_rels/.rels` relationships | Two per data part |
| Fix `Origin.xml` | Remove the table from `DataPhaseTables`, and recompute the SHA256 of the new `model.xml` into `Checksums` |

The stored checksum is verified against the original `model.xml` before it is rewritten. If it
does not match a plain SHA256 of those bytes, the format is not what this script assumes and
the run stops rather than guessing.

`ForeignTable` is deliberately not treated as ownership: it is the table a foreign key *points
at*, which is precisely the case that must stop the run.

### The 32-bit ZIP constraint

A classic ZIP caps entry size, total size and offsets at 4 GB, and the entry count at 65535.
SqlPackage emits ZIP64 by itself once a BACPAC passes those limits, so ZIP64 is legitimate.
What must not happen is a 32-bit archive being promoted to ZIP64 by the rewrite. The script
reads the end-of-central-directory records of both files and **fails** if a 32-bit source
produced a ZIP64 output.

### Usage

```powershell
# What is in there? Table list with data part count and compressed size
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac' -ListTables

# Look first: reports the removal set and any external reference, writes nothing
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac' -WhatIf

# Produce the file
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac'

# Another table, faster deflate, overwrite
./New-BacpacWithoutTable.ps1 -BacpacPath 'C:\backup\mydb.bacpac' `
    -TableName AuditLog -CompressionLevel Fastest -Force
```

The default table is `dbo.WhoIsActive`, the case the script was written for. Any other table
goes through `-TableName`, written whichever way is at hand:

| Written as | Resolves to |
|---|---|
| `-TableName AuditLog` | `[dbo].[AuditLog]`, schema from `-SchemaName` |
| `-TableName audit.AuditLog` | `[audit].[AuditLog]` |
| `-TableName '[audit].[AuditLog]'` | `[audit].[AuditLog]` |

The qualified form matters: it is what the model and the logs display, so it is what gets
pasted back in. Taken literally it would name a table called `[dbo].[AuditLog]` and be reported
as missing, hence the normalisation.

Every entry is decompressed and recompressed on the way through, which dominates the runtime:
about 40 seconds for 165 MB, about 9 minutes for 15 GB. `-CompressionLevel Fastest` trades
output size for wall-clock. Data parts are streamed with a fixed buffer and never held in
memory, so file size does not drive memory use.

### Measured on real exports

Removing `dbo.WhoIsActive` from a 165 MB export:

| | |
|---|---|
| Model elements removed | 2 (the table and its clustered index) |
| Data parts dropped | 2 |
| `_rels/.rels` relationships removed | 4 |
| `DataPhaseTables` | 552 → 551 entries |
| User-defined objects | 3030 → 3028 |
| Archive format | ZIP32 → ZIP32 |
| Source SHA256 | unchanged |

Removing a 8.7 GB table from a 15 GB ZIP64 export:

| | |
|---|---|
| Model elements removed | 3 (table, index, primary key) |
| Data parts dropped | 3073 |
| Relationships removed | 6146 |
| Output | 15266 MB → 6268 MB |
| Elapsed | 8 min 37 s |

Refusing to remove a table referenced elsewhere: 8 external references found — two foreign keys
held by other tables and a view — and nothing written.

---

## `Test-Bacpac.ps1`

Is this file intact and coherent? Answered without importing it, without a server, read-only.

### Why not just use SqlPackage

Because it cannot. SqlPackage offers seven actions — `Extract`, `DeployReport`, `DriftReport`,
`Publish`, `Script`, `Export`, `Import` — and every one of them either reads from or writes to a
database. `Import` is the only one that accepts a BACPAC, and it needs a target connection; its
properties are all about *how* to import, there is no dry-run and no validate-only switch. It
also connects to the target **before** reading the package, so pointing it at a bogus server
fails on the connection and tells you nothing about the file.

### The eight checks

| Check | Catches |
|---|---|
| `ARCHIVE` | the ZIP does not open, or its central directory is unreadable |
| `COMPAT` | the package is ZIP64 or over 4 GB, so the **old x86 SqlPackage cannot read it** — see below |
| `PARTS` | `model.xml`, `Origin.xml`, `[Content_Types].xml` or `_rels/.rels` missing |
| `CRC` | **physical corruption** — a truncated copy, a bad transfer, a failing disk. Every entry is decompressed and its stored CRC32 verified |
| `CHECKSUM` | `model.xml` altered without its `Origin.xml` SHA256 being updated |
| `MODEL` | a schema model that no longer resolves, typically an element removed while something still references it |
| `DATA` | data parts left for a table absent from the model, or a data folder missing from `DataPhaseTables` |
| `RELS` | `_rels/.rels` and the `TableData` parts disagreeing, in either direction |

`CRC` is the only one that detects physical corruption, and the only expensive one. `-Quick`
skips it and keeps the seven others.

A file held by another process — an export still being written — is reported as
`LOCKED (not checked)`, never as corrupt.

### `COMPAT`, and the error that blames the wrong thing

There are two SqlPackage builds, and only one of them can read a large BACPAC.

| Build | Location | Architecture | Reads ZIP64 |
|---|---|---|---|
| .NET Framework, from `DacFramework.msi` | `Program Files\Microsoft SQL Server\<version>\DAC\bin` | x86 | **no** |
| .NET 8, the `sqlpackage` dotnet tool | `~/.dotnet/tools` | x64 | yes |

The x86 build reads packages through `System.IO.Packaging`, whose ZIP reader has no ZIP64
support. Past 4 GB a BACPAC is necessarily ZIP64, and that build fails with:

```text
System.IO.FileFormatException: File contains corrupted data.
   at MS.Internal.IO.Zip.ZipIOLocalFileBlock.Validate(...)
```

**The package is intact. The reader cannot parse it.** The message names the wrong culprit and
sends people looking for a damaged file that does not exist. This is
[dotnet/runtime issue 94899](https://github.com/dotnet/runtime/issues/94899), fixed only for
.NET 9. Microsoft's own
[troubleshooting guidance](https://learn.microsoft.com/en-us/sql/tools/sqlpackage/troubleshooting-issues-and-performance-with-sqlpackage)
opens by recommending the .NET build over the `DacFramework.msi` one.

Note the trap: the x86 build sits under `Program Files`, not `Program Files (x86)`, so its
location gives no hint of its architecture.

Every check here reads through `System.IO.Compression`, which handles ZIP64. Without `COMPAT`
the tool would hand out a clean verdict on a file SqlPackage then refuses.

### `RELS` covers only `TableData` parts

This distinction was learned the hard way. A first version required every data part to be
declared in `_rels/.rels` and duly reported five untouched, Microsoft-produced exports as
broken. They were not: `BlobData-<guid>.BIN` parts hold oversized binary column values and are
referenced from inside the BCP data, never from `.rels`. Measured on an untouched 15 GB export:

| | |
|---|---|
| `TableData` parts | 18324 |
| Relationships in `_rels/.rels` | 18324 |
| `BlobData` parts | 126 |
| `BlobData` parts declared in `.rels` | 0 |

A data part following neither naming convention is reported as a warning rather than silently
ignored, since it would mean the format has moved.

### Usage

```powershell
# Full check, one file
./Test-Bacpac.ps1 -BacpacPath 'C:\backup\mydb.bacpac'

# Structural verdict on a whole folder, nothing decompressed
./Test-Bacpac.ps1 -BacpacPath 'C:\backup\*.bacpac' -Quick

# Full check on every file, keeping only what failed
Get-ChildItem C:\backup\*.bacpac | ./Test-Bacpac.ps1 | Where-Object Verdict -ne 'OK'
```

One object is returned per file, so results can be filtered, sorted or exported.

### Verified against deliberate damage

A validator that always answers OK is worth nothing, so it was checked against known-bad files.

| File | Verdict | Which check caught it |
|---|---|---|
| Original 165 MB export | `OK` | everything passes, 14 s |
| Output of `New-BacpacWithoutTable.ps1` | `OK` | everything passes |
| 4 KB of junk written at offset 50 MB | `CORRUPT` | `CRC`, naming the exact entry — every other check still passed |
| One byte added to `model.xml`, checksum left alone | `CORRUPT` | `CHECKSUM` and `MODEL`, printing both hashes |

Then against a real fleet: a `-Quick` sweep of ten files from 165 MB to 35 GB took about
90 seconds in total.

---

## Limits, stated plainly

**Rewriting a BACPAC is not supported by Microsoft.** These scripts produce a file that DacFx
accepts, which is a strong signal and not a guarantee. Import the result against an empty,
disposable target database before trusting it with anything.

A `Test-Bacpac` verdict is a file-level verdict. It says the package is internally coherent, not
that the import will succeed — data-level failures, collation mismatches and the XML nesting
limit that motivated all this are invisible at file level.

Removing a table does not make a large package readable by the x86 SqlPackage unless the result
drops under 4 GB. A 15 GB export trimmed to 6 GB is still ZIP64.

## Licence

MIT. See [LICENSE](LICENSE).
