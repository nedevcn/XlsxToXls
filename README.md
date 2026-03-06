# Nedev.FileConverters.XlsxToXls

![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg) ![NuGet](https://img.shields.io/nuget/v/Nedev.FileConverters.XlsxToXls)

A high-performance **XLSX → XLS** converter library as part of the `Nedev.FileConverters` ecosystem. Targets **.NET 8.0** and **.NET Standard 2.1** with **zero third-party dependencies** (core package dependency is optional and pulled in via NuGet). It reads Office Open XML (`.xlsx`) workbooks and writes Excel 97–2003 binary (`.xls`, BIFF8) using only built-in BCL types.

---

## Features

- **Library + CLI** — core converter available as DLL and command‑line tool (see below).
- **Core package integration** — the library implements `IFileConverter` and is decorated with `FileConverterAttribute` so it plugs into `Nedev.FileConverters.Core`’s discovery/DI helpers. Both the DLL and the CLI automatically register the converter via `ServiceCollectionExtensions.AddFileConverter`.

- **Zero third-party dependencies** — uses only `System.IO.Compression`, `System.Xml`, `System.Buffers`, and core .NET types.
- **Performance-oriented** — `ArrayPool<byte>` for buffers, streaming `XmlReader` for XLSX, `Span<byte>` for BIFF output to minimize allocations.
- **Multi‑targeted** — builds for `net8.0` and `netstandard2.1`; see Build instructions above.

---

## API

### Conversion

| Method | Description |
|--------|-------------|
| `XlsxToXlsConverter.Convert(Stream xlsxStream, Stream xlsStream)` | Converts from a readable XLSX stream to a writable XLS stream. |
| `XlsxToXlsConverter.ConvertFile(string xlsxPath, string xlsPath)` | Converts a file to another file by path. |

### Example

```csharp
using Nedev.FileConverters.XlsxToXls;

// Stream-based
using var xlsx = File.OpenRead("input.xlsx");
using var xls = File.Create("output.xls");
XlsxToXlsConverter.Convert(xlsx, xls);

// File-based
XlsxToXlsConverter.ConvertFile("input.xlsx", "output.xls");
```

### CLI usage

```bash
# build the CLI project first (see Build section)
cd src/Nedev.FileConverters.XlsxToXls.Cli
dotnet run -- ../path/to/input.xlsx ../path/to/output.xls
# or after publishing:
# XlsxToXls.Cli.exe input.xlsx output.xls
```

The CLI version uses the core package’s abstraction; it registers `XlsxToXlsConverter.FileConverterAdapter` with `ServiceCollection` and resolves an `IFileConverter` instance, demonstrating seamless plugin-style integration.

---

## Supported (Conversion Completeness)

### Workbook & sheets

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Multiple worksheets | `xl/workbook.xml` + rels | BOUNDSHEET, separate sheet streams |
| Sheet names | `name` on `<sheet>` | Truncated to 31 chars in BIFF |
| Codepage | — | 1252 (Latin) |

### Cell data

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Numbers | `<v>` with number format | NUMBER |
| Text (shared strings) | `t="s"` + SST | LABELSST / SST + CONTINUE |
| Inline / direct text | `t="str"`, `t="inlineStr"` | LABEL |
| Empty cells | `<c>` without value | BLANK |
| Booleans | `t="b"` | BOOLERR (boolean) |
| Errors | `t="e"` (#DIV/0!, #N/A, etc.) | BOOLERR (error), mapped to BIFF codes |
| Unicode | UTF-8 in XLSX | 16-bit in LABEL / SST |
| Formulas (basic) | `<f>` (formula) + cached `<v>` | FORMULA (+ STRING record for string results); limited parser (refs/areas, basic operators, a few functions) |

### Cell formatting (from `xl/styles.xml`)

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Fonts | `fonts/font` | FONT |
| Number formats | `numFmts/numFmt` | FORMAT |
| Cell XFs | `cellXfs/xf` | XF (style + cell XFs), cell `s` → XF index |
| Minimum fonts | — | At least 4 fonts ensured |

### Rows, columns & layout

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Used range | Computed from rows/cells | DIMENSION |
| Default column width | — | DEFCOLWIDTH (8) |
| Column width / visibility | `<col>` (width, hidden) | COLINFO |
| Row height / visibility | `<row>` (ht, hidden) | ROW |
| Merged cells | `<mergeCells>` | MERGEDCELLS |

### Sheet-level settings

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Freeze panes | `sheetViews/sheetView/pane` | WINDOW2 + PANE |
| Horizontal page breaks | `rowBreaks/brk` | HORIZONTALPAGEBREAKS |
| Vertical page breaks | `colBreaks/brk` | VERTICALPAGEBREAKS |
| Page setup | `pageSetup` (orientation, scale, fitToWidth/Height) | PAGESETUP |
| Margins | `pageMargins` | LEFTMARGIN, RIGHTMARGIN, TOPMARGIN, BOTTOMMARGIN |
| Print area | `definedName` Print_Area / _xlnm.Print_Area in workbook.xml | NAME (Lbl) + ptgArea3D |
| Print titles (rows/cols) | `definedName` Print_Titles / _xlnm.Print_Titles in workbook.xml | NAME (Lbl) + ptgArea3D |

### Hyperlinks, comments & data validation

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Cell/range hyperlinks (URLs) | `<hyperlink ref="..." r:id="...">` + sheet rels | HYPERLINK (URL moniker) |
| Cell comments (notes) | `commentsN.xml` (authors + commentList) | NOTE + OBJ + TXO/CONTINUE (cell, author, text) |
| Data validation | `dataValidations` / `dataValidation` (sqref, type, formula1/2) | DATAVALIDATIONS + DATAVALIDATION; **list** type with explicit comma-separated list supported (formula as tStr); other types written with flags/ranges/prompt/error strings, simple formulas compiled to RPN when possible |

### Shared string table (SST)

- Large SSTs are split across **SST + CONTINUE** records (BIFF record data &lt; 8224 bytes).

---

## Not supported (current limitations)

- **Formulas (advanced)** — only a subset of Excel formulas is compiled (no full Excel function set; shared formula edge cases may be imperfect).
- **Data validation (advanced)** — explicit lists and simple formulas/ranges are supported; complex formulas/functions and edge cases may not compile even though basic RPN emission is attempted.
- **Conditional formatting** — not implemented.
- **Charts, images, drawings** — not implemented.
- **Threaded comments** — only legacy comments (commentsN.xml) are read.

---

## XLS limits applied

| Limit | Value | Behavior |
|-------|--------|----------|
| Max rows | 65,536 | No truncation; out-of-range may produce invalid BIFF. |
| Max columns | 256 (A–IV) | No truncation. |
| Sheet name length | 31 | Truncated. |

---

## Build

From the repository root:

```bash
cd src/Nedev.FileConverters.XlsxToXls
dotnet build
cd ../Nedev.FileConverters.XlsxToXls.Cli
dotnet build
```

Outputs are produced for both target frameworks, e.g.:

```
src/Nedev.FileConverters.XlsxToXls/bin/Debug/net8.0/Nedev.FileConverters.XlsxToXls.dll
src/Nedev.FileConverters.XlsxToXls/bin/Debug/netstandard2.1/Nedev.FileConverters.XlsxToXls.dll
src/Nedev.FileConverters.XlsxToXls.Cli/bin/Debug/net8.0/Nedev.FileConverters.XlsxToXls.Cli.dll
```


---

## Project layout

```
Nedev.FileConverters.XlsxToXls/
├── src/
│   ├── Nedev.FileConverters.XlsxToXls.csproj
│   ├── XlsxToXlsConverter.cs   # Public API + BIFF orchestration
│   └── Internal/
│       ├── BiffWriter.cs      # BIFF8 record writing
│       ├── OleCompoundWriter.cs
│       ├── StylesData.cs
│       ├── StylesReader.cs    # xl/styles.xml
│       └── XlsxReader.cs      # XLSX read (sheets, cells, comments, hyperlinks, etc.)
└── README.md
```

---

## License

This project is licensed under the **MIT** license. See [LICENSE](LICENSE) for details.

## NuGet

The library is published as `Nedev.FileConverters.XlsxToXls` (version 0.1.0 at time of writing) and depends on `Nedev.FileConverters.Core` when that package is installed.
