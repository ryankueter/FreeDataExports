# FreeDataExports

> **⚠️ Note:** FreeDataExports is being replaced by [FreeDataExportsv2](https://github.com/ryankueter/FreeDataExportsv2)
> and will no longer be maintained. New projects should use FreeDataExportsv2, which offers a
> cleaner API, richer formatting, charts, images, Excel tables, and identical XLSX/ODS output
> from the same code.

A lightweight .NET library for exporting data to `.xlsx`, `.ods`, and `.csv` files —
no Excel or LibreOffice installation required. Available on
[NuGet](https://www.nuget.org/packages/FreeDataExports).

---

## Table of Contents

1. [Supported Formats & Platforms](#supported-formats--platforms)
2. [Initialization](#initialization)
3. [Excel Export (XLSX 2019)](#excel-export-xlsx-2019)
4. [LibreOffice Export (ODS 1.3)](#libreoffice-export-ods-13)
5. [CSV Export](#csv-export)
6. [Column Widths](#column-widths)
7. [Tab Colors](#tab-colors)
8. [Format Code Overrides](#format-code-overrides)
9. [Error Handling](#error-handling)
10. [Saving / Getting Bytes](#saving--getting-bytes)
11. [Formatting Reference](#formatting-reference)
12. [Data Types Reference](#data-types-reference)

---

## Supported Formats & Platforms

### File formats

| Format | Extension | Spec |
|---|---|---|
| Office Open XML Spreadsheet | `.xlsx` | Excel 2019 / OOXML |
| OpenDocument Spreadsheet | `.ods` | ODF v1.3 |
| Comma-Separated Values | `.csv` | RFC 4180 |

### .NET platform support

- .NET 6 through .NET 10
- .NET Standard 2.0
- .NET Core 2.x – 3.1
- .NET Framework 4.6.1 – 4.8
- Mono 5.4, 6.4
- Xamarin.iOS 10.14, 12.16
- Xamarin.Android 8.0, 10.0
- Universal Windows Platform 10.0.16299
- Unity 2018.1

See the full compatibility matrix at
[dotnet.microsoft.com/platform/dotnet-standard](https://dotnet.microsoft.com/platform/dotnet-standard#versions).

---

## Initialization

FreeDataExports exposes a shared `IDataWorkbook` interface so you can write format-agnostic
code and switch between XLSX and ODS with a single line change.

```csharp
using FreeDataExports;

// Shared interface — lets you write the same code for either format
IDataWorkbook workbook = new DataExport().CreateXLSX2019();
IDataWorkbook workbook = new DataExport().CreateODSv1_3();
```

> **Note:** Tab color strings and column-width units differ between XLSX and ODS.
> See [Tab Colors](#tab-colors) and [Column Widths](#column-widths) for details.

---

## Excel Export (XLSX 2019)

```csharp
using FreeDataExports;

// Create a new XLSX workbook
var workbook = new DataExport().CreateXLSX2019();

// Optional metadata
workbook.CreatedBy = "Jane Doe";
workbook.FontSize  = 11;

// Add worksheets
var orders    = workbook.AddWorksheet("Orders");
var inventory = workbook.AddWorksheet("Inventory");

// Header row
orders.AddRow()
    .AddCell("OrderId",    DataType.String)
    .AddCell("Item",       DataType.String)
    .AddCell("Units",      DataType.String)
    .AddCell("Price",      DataType.String)
    .AddCell("OrderDate",  DataType.String)
    .AddCell("SalesAssoc", DataType.String)
    .AddCell("Delivered",  DataType.String);

// Data rows
foreach (var o in Orders)
{
    orders.AddRow()
        .AddCell(o.OrderId,         DataType.Number)
        .AddCell(o.Item,            DataType.String)
        .AddCell(o.Units,           DataType.Number)
        .AddCell(o.Price,           DataType.Currency)
        .AddCell(o.OrderDate,       DataType.LongDate)
        .AddCell(o.SalesAssociate,  DataType.String)
        .AddCell(o.Delivered,       DataType.Boolean);
}

// Inventory sheet
inventory.AddRow()
    .AddCell("ItemId", DataType.String)
    .AddCell("Item",   DataType.String)
    .AddCell("Number", DataType.String)
    .AddCell("Price",  DataType.String);

foreach (var i in Inventory)
{
    inventory.AddRow()
        .AddCell(i.ItemId,  DataType.Number)
        .AddCell(i.Item,    DataType.String)
        .AddCell(i.Number,  DataType.Number)
        .AddCell(i.Price,   DataType.Currency);
}

// Optional — tab colors (RGB hex, no #)
orders.TabColor    = "0000FF00"; // green
inventory.TabColor = "000000FF"; // blue

// Optional — column widths (character units)
orders.ColumnWidths("10", "10", "5", "10", "28.5", "15", "10");
inventory.ColumnWidths("10", "10", "10", "10");

// Optional — format overrides and error worksheet
workbook.Format(DataType.DateTime24, @"m/d/yy\ h:mm;@");
workbook.AddErrorsWorksheet();

// Save
await workbook.SaveAsync("Orders.xlsx");
```

---

## LibreOffice Export (ODS 1.3)

The ODS API mirrors the XLSX API — the main differences are the tab color format and
column width units. See [Tab Colors](#tab-colors) and [Column Widths](#column-widths).

```csharp
using FreeDataExports;

// Create a new ODS workbook
var workbook = new DataExport().CreateODSv1_3();

// Optional metadata
workbook.CreatedBy = "Jane Doe";
workbook.FontSize  = 11;

// Add worksheets
var orders    = workbook.AddWorksheet("Orders");
var inventory = workbook.AddWorksheet("Inventory");

// Header row
orders.AddRow()
    .AddCell("OrderId",    DataType.String)
    .AddCell("Item",       DataType.String)
    .AddCell("Units",      DataType.String)
    .AddCell("Price",      DataType.String)
    .AddCell("OrderDate",  DataType.String)
    .AddCell("SalesAssoc", DataType.String)
    .AddCell("Delivered",  DataType.String);

// Data rows
foreach (var o in Orders)
{
    orders.AddRow()
        .AddCell(o.OrderId,         DataType.Number)
        .AddCell(o.Item,            DataType.String)
        .AddCell(o.Units,           DataType.Number)
        .AddCell(o.Price,           DataType.Currency)
        .AddCell(o.OrderDate,       DataType.LongDate)
        .AddCell(o.SalesAssociate,  DataType.String)
        .AddCell(o.Delivered,       DataType.Boolean);
}

// Inventory sheet
inventory.AddRow()
    .AddCell("ItemId", DataType.String)
    .AddCell("Item",   DataType.String)
    .AddCell("Number", DataType.String)
    .AddCell("Price",  DataType.String);

foreach (var i in Inventory)
{
    inventory.AddRow()
        .AddCell(i.ItemId,  DataType.Number)
        .AddCell(i.Item,    DataType.String)
        .AddCell(i.Number,  DataType.Number)
        .AddCell(i.Price,   DataType.Currency);
}

// Optional — tab colors (CSS hex, with #)
orders.TabColor    = "#00FF00"; // green
inventory.TabColor = "#0000FF"; // blue

// Optional — column widths (inches or centimeters — specify the unit)
orders.ColumnWidths(".8in", "1in", ".5in", "1in", "1.5in", "1in", "1in");
inventory.ColumnWidths(".8in", "1in", ".8in", "1in");

// Optional — format overrides and error worksheet
workbook.Format(DataType.Currency, "symbol=$,language=en,country=US,decimals=2");
workbook.AddErrorsWorksheet();

// Save
await workbook.SaveAsync("Orders.ods");
```

---

## CSV Export

```csharp
using FreeDataExports;

var csv = new DataExport().CreateCsv();

// Header
csv.AddRow("OrderId", "Item", "Units", "Price", "OrderDate", "SalesAssoc", "Delivered");

// Data
foreach (var o in Orders)
{
    csv.AddRow(o.OrderId, o.Item, o.Units, o.Price,
               o.OrderDate, o.SalesAssociate, o.Delivered);
}

await csv.SaveAsync("Orders.csv");
```

`AddRow` returns `this` so calls can be chained:

```csharp
csv.AddRow("A", "B", "C")
   .AddRow(1,   2,   3);
```

### CsvFile methods

| Method | Description |
|---|---|
| `AddRow(params object[] values)` | Appends a row; returns `this` for chaining |
| `GetBytes()` | Returns the CSV as a `byte[]` |
| `GetBytesAsync()` | Async version of `GetBytes()` |
| `Save(path)` | Synchronous save to a file path |
| `SaveAsync(path)` | Asynchronous save to a file path |

---

## Column Widths

Set column widths after adding all rows. Returns the worksheet for chaining.

```csharp
// XLSX — character-width units (same as Excel's column width)
orders.ColumnWidths("10", "10", "5", "10", "28.5", "15", "10");

// ODS — physical units (inches or centimeters, include the unit suffix)
orders.ColumnWidths(".8in", "1in", ".5in", "1in", "1.5in", "1in", "1in");
orders.ColumnWidths("2cm", "2.5cm", "1.5cm", "2cm");
```

> **XLSX vs ODS:** XLSX column widths are specified in Excel character units. ODS column
> widths require a physical measurement with a unit suffix (`in` or `cm`). These values
> are **not interchangeable** — specify the appropriate format for the workbook type you
> are creating.

---

## Tab Colors

Tab color format differs between XLSX and ODS:

| Format | Syntax | Example (green) |
|---|---|---|
| XLSX | RGB hex, no `#`, 8 characters | `"0000FF00"` |
| ODS | CSS hex with leading `#` | `"#00FF00"` |

```csharp
// XLSX
orders.TabColor = "0000FF00"; // green
orders.TabColor = "000000FF"; // blue

// ODS
orders.TabColor = "#00FF00";  // green
orders.TabColor = "#0000FF";  // blue
```

---

## Format Code Overrides

Override the default format code for any `DataType` before writing data:

```csharp
// XLSX — standard Excel format strings
workbook.Format(DataType.DateTime24, @"m/d/yy\ h:mm;@");
workbook.Format(DataType.Currency,   @"""$""#,##0.00");

// ODS — key=value descriptor strings
workbook.Format(DataType.Currency, "symbol=$,language=en,country=US,decimals=2");
workbook.Format(DataType.Decimal,  "decimals=8");
```

Overrides apply to all worksheets in the workbook.

---

## Error Handling

Every `AddCell` call is wrapped in a `try/catch`. When a type conversion fails:

- The offending cell is left empty.
- The error is recorded internally.

### Options

```csharp
// Auto-create a red "Errors" worksheet on save — only added when errors actually occur
workbook.AddErrorsWorksheet();

// Or retrieve errors as a string at any time
string report = workbook.GetErrors();
if (!string.IsNullOrEmpty(report))
    Console.WriteLine(report);
```

---

## Saving / Getting Bytes

All three workbook types (`XLSX`, `ODS`, `CSV`) share the same save API:

```csharp
// Synchronous — file path
workbook.Save(path);

// Asynchronous — file path
await workbook.SaveAsync(path);

// In-memory byte array (e.g. for an HTTP response body)
byte[] bytes = workbook.GetBytes();

// Asynchronous byte array
byte[] bytes = await workbook.GetBytesAsync();
```

---

## Formatting Reference

### XLSX format strings

These are Excel number-format codes stored in the `.xlsx` package. Many contain escape
characters and are written as C# verbatim or literal strings.

| Data Type | Format string |
|---|---|
| `Currency` | `@"""$""#,##0.00"` |
| `Percent` | `@"0.00%"` |
| `DateTime` | `@"[$-409]m/d/yyyy\ h:mm:ss\ AM/PM;@"` |
| `Decimal` | `"0.0000"` |
| `Date` | `@"m/d/yyyy"` |
| `Time` | `@"h:mm\ AM/PM"` |
| `LongDate` | `@"[$-F800]dddd\,\ mmmm\ dd\,\ yyyy"` |
| `DateTime24` | `@"m/d/yyyy\ h:mm:ss;@"` |
| `Time24` | `@"h:mm:ss;@"` |

### ODS format descriptors

ODS files use key=value descriptor strings rather than Excel format codes.

| Data Type | Descriptor string |
|---|---|
| `Currency` | `"symbol=$,language=en,country=US,decimals=2"` |
| `Currency` (short) | `"decimals=2"` |
| `Decimal` | `"decimals=4"` |

---

## Data Types Reference

| Value | Description |
|---|---|
| `String` | Plain text |
| `Number` | Numeric (general format) |
| `Currency` | Currency symbol with two decimal places |
| `Percent` | Percentage with two decimal places |
| `Decimal` | Fixed decimal (four places by default) |
| `DateTime` | Date and 12-hour time |
| `DateTime24` | Date and 24-hour time |
| `Date` | Short date only |
| `LongDate` | Full written date (e.g. Thursday, April 16, 2026) |
| `Time` | 12-hour time only |
| `Time24` | 24-hour time only |
| `Boolean` | `TRUE` / `FALSE` |

---

## License

This project is licensed under the MIT License.
See the [LICENSE](../LICENSE) file for details.
