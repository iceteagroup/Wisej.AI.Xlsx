# Wisej.AI.XlsX

High fidelity extraction of Excel workbooks (XLSX, XLSM, XLTX, XLTM, Flat OPC) into **Markdown or plain text, one page per sheet**, built to feed language model context and retrieval pipelines in [Wisej.AI](https://wisej.com).

```csharp
var result = new XlsxExtractor().Extract("sales.xlsx");

Console.WriteLine(result.Text);                 // the whole workbook, with "<!-- Page n -->" separators
Console.WriteLine(result.Pages[1].SheetName);   // the sheet of page 2
Console.WriteLine(result.Pages[1].Text);        // its text
File.WriteAllBytes(result.Images[0].FileName, result.Images[0].Data);
```

## What is extracted

| Content | Output |
|---|---|
| Values | Every value as Excel displays it, using its number format: `1,234.50`, `($5.00)`, `12.5%`, `1.23E+04`, `1 3/4`, `15K`, dates, times and elapsed times (`[h]:mm`) |
| Sheets | A `## Sheet name` heading per page; hidden and very hidden sheets are skipped unless requested |
| Ranges | Each block of cells, separated from the others by empty rows and columns, becomes a pipe table with its address (`Range A1:D20`); empty columns are removed |
| Headers | The first row when it contains labels, otherwise the column letters; a row of labels above several blocks (forms, price lists) heads every block below it |
| Excel tables | `Table: Name (A1:D20)` with the table's own header |
| Titles and notes | Cells alone on their rows (titles, subtitles, remarks) become paragraphs |
| Merged cells | The value is repeated in every covered cell (default) or written once |
| Hyperlinks | `[text](url)`, including `HYPERLINK()` formulas |
| Formulas | Optional, after the value: `6 (=SUM(B3:B4))`; formulas never calculated are written as `=formula` |
| Comments | Threaded comments with their replies, and notes, listed at the end of the sheet |
| Charts | The cached series data as a table, with the chart title, type, axis titles and anchor cell (including Office 2016 chart types and chart sheets) |
| Pictures | `![alt text](image3.png)` for floating pictures and pictures placed in cells, the bytes in `Result.Images`, an optional `ImageDescriber` callback (vision model, OCR) called once per distinct picture |
| Text boxes | The text of shapes and text boxes |
| Metadata | Core, application and custom properties, and the names of all the sheets |
| Robustness | Malformed URIs, missing parts and Windows path separators are repaired; Strict conformance namespaces are normalized; worksheets and shared strings are streamed, so large sheets are never loaded as a whole |

## Number formats

`NumberFormatter` implements the Excel number format language: up to four sections (positive, negative, zero, text), conditions (`[<1000]`), digit placeholders (`0 # ?`), thousands separators and scaling (`#,##0,"K"`), percentages, scientific and engineering notation, fractions with free or fixed denominators, literal text, currency and locale tags (`[$€-407]`, `[$-F800]`), dates and times with fractions of seconds, elapsed time, the 1900 and 1904 date systems, and the 1900 leap year bug of Excel. Numbers are rounded half away from zero on their 15 significant digits, as Excel does, so `2.675` formatted with `0.00` is `2.68`.

The built-in short date and date time formats, which Excel displays with the regional settings of the machine, are written as ISO 8601 (`2026-10-14`) so that a language model never has to guess between day and month. `Culture` sets the separators and the month and day names.

## Options

| Option | Default | Description |
|---|---|---|
| `Format` | `Markdown` | `Markdown` or `PlainText` |
| `PageHeaderFormat` | `<!-- Page {0} -->` | Line introducing each page in `Text`; `{0}` page, `{1}` total, `{2}` sheet name; empty for none |
| `Culture` | `InvariantCulture` | Separators and month and day names |
| `RowsPerPage` | `0` | Splits large sheets into pages of at most this many table rows, repeating the header |
| `MaxRowsPerSheet` | `0` | Writes at most this many table rows per sheet, followed by the number of rows omitted |
| `TableHeaders` | `Auto` | `Auto`, `FirstRow` or `ColumnLetters` for ranges that are not Excel tables |
| `MergedCells` | `Repeat` | Repeat merged values in every covered cell, or leave them `Empty` |
| `IncludeCellRanges`, `IncludeRowNumbers` | `true`, `false` | Trace values back to their cells |
| `IncludeFormulas`, `IncludeHyperlinks`, `IncludeComments`, `IncludeCharts`, `IncludeTextBoxes`, `IncludeImages`, `IncludeImageData` | | Fine grained control of the content |
| `IncludeHiddenSheets`, `IncludeHiddenRows`, `IncludeHiddenColumns` | `false` | Include what the author hid (rows hidden by a filter are hidden rows) |
| `SkipDecorativeImages` | `true` | Skip the pictures marked as decorative |
| `ImageDescriber` | `null` | `Func<XlsxImage, string>` that describes an image |

## Architecture

```
src/Wisej.AI.XlsX
  XlsxExtractor            Public entry point
  Packaging/               Secure package access (no DTD, no resolver), Strict namespace normalization, repair
  Reading/                 Workbook, streamed worksheets and shared strings, number formats, styles, tables,
                           hyperlinks, comments, drawings, charts, pictures in cells (rich values), metadata
  Model/                   Format independent workbook model
  Layout/                  Splits every sheet into tables and paragraphs, chooses the headers
  Rendering/               Markdown and plain text renderers, pagination
tests/Wisej.AI.XlsX.Tests  Unit tests, building workbooks from raw SpreadsheetML
```

The library targets **.NET Standard 2.0** and **.NET 8**, so it runs on .NET Framework 4.6.2+ and every modern .NET, on Windows, Linux and macOS. Its only dependency is `DocumentFormat.OpenXml` 3.1.0, the same version Wisej.AI uses. The assembly is strong-named (`extension.snk`), so it can be consumed by the strong-named Wisej.AI assembly on .NET Framework.

## Limitations

- Legacy binary `.xls` files, binary `.xlsb` workbooks and encrypted workbooks are not supported (an `InvalidDataException` explains why).
- Values are the results cached by the application that saved the workbook: formulas are not recalculated.
- Pivot tables are written as the cells they produce; their definition and cache are not read. Conditional formatting, data validation, sparklines and slicers are ignored.
- The layout of a sheet is inferred from empty rows and columns; sheets that use borders or colors alone to separate their blocks may be grouped differently than they appear.
