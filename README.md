# Excel Validator

[![NuGet](https://img.shields.io/nuget/v/excel-validator?style=flat-square)](https://www.nuget.org/packages/excel-validator)
[![Build Status](https://dev.azure.com/devopspractices1/Space%20Game%20-%20web%20-%20Tests/_apis/build/status/sheddy123.ExcelValidation?branchName=main)](https://dev.azure.com/devopspractices1/Space%20Game%20-%20web%20-%20Tests/_build/latest?definitionId=15&branchName=main)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue?style=flat-square)](LICENSE)

Schema validation for `.xlsx` worksheets. Declare the columns you expect and the type each one holds, hand it a workbook, and get back a structured list of what is wrong and exactly where, or an annotated copy of the workbook with the bad cells highlighted for whoever sent it to you.

Built on [ClosedXML](https://github.com/ClosedXML/ClosedXML) (MIT), so there is no commercial licensing restriction to worry about.

```csharp
using ExcelValidator;

var schema = new ExcelSchema()
    .Column("ID", ExcelCellType.Integer)
    .Column("Username", ExcelCellType.Text, maxLength: 50)
    .Column("Email", ExcelCellType.Text, pattern: @"^[^@\s]+@[^@\s]+\.[^@\s]+$");

var result = new ExcelSheetValidator().Validate(File.ReadAllBytes("people.xlsx"), schema);

if (!result.IsValid)
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine(error);   // B4: 'Username' is required but the cell is empty.
    }
}
```

## Install

```
dotnet add package excel-validator
```

Targets `netstandard2.0` and `net8.0`, so it runs on .NET Framework 4.6.1+, .NET Core 2.0+, and all modern .NET.

| File type | Format  | Excel version  |
| :-------- | :------ | :------------- |
| `.xlsx`   | OpenXML | 2007 and newer |

The legacy `.xls` format is not supported.

## Defining a schema

Each column gets a name, a type, and optionally some constraints:

```csharp
var schema = new ExcelSchema()
    .Column("ID", ExcelCellType.Integer)
    .Column("Signed up", ExcelCellType.DateTime)
    .Column("Notes", ExcelCellType.Text, required: false)
    .Column("Status", allowedValues: new[] { "Active", "Pending" })
    .Column("Code", ExcelCellType.Text, minLength: 3, maxLength: 10);
```

If you only care that the headers line up, skip the types:

```csharp
var schema = ExcelSchema.FromHeaders("ID", "Username", "Email");
```

By default a column the schema doesn't declare is an error. To ignore extra columns:

```csharp
schema.AllowUnexpectedColumns = true;
```

### Column types

| `ExcelCellType` | Accepts |
| :-------------- | :------ |
| `Text`          | any non-empty value |
| `Integer`       | a whole number within `int` range |
| `Long`          | a whole number within `long` range |
| `Decimal`       | a number within `decimal` range |
| `Double`        | a number within `double` range |
| `Boolean`       | an Excel boolean, or the text `true`/`false` |
| `DateTime`      | an Excel date, or a parseable date string |
| `Guid`          | anything `Guid.TryParse` accepts |

A cell satisfies its type whether Excel stored the value natively or as text a cell holding the number `34` and one holding the string `"34"` both satisfy `Integer`, since which one a file uses depends on the tool that wrote it rather than on whether the data is right. Text is parsed with the invariant culture, so results don't change with the server's locale.

## Reading the result

`Validate` does not throw on invalid *data* it returns what it found. It throws only when the file itself can't be read; see [Errors vs exceptions](#errors-vs-exceptions).

```csharp
result.IsValid            // true when Errors is empty
result.Errors             // every problem, header errors first, then row by row
result.MissingColumns     // schema columns absent from the sheet
result.UnexpectedColumns  // sheet columns absent from the schema
result.RowsValidated      // data rows examined, blank rows excluded
result.ColumnsAreValid    // the header row alone
result.RowsAreValid       // the data cells alone
result.Truncated          // whether MaxErrors cut the list short
```

Each `ValidationError` locates itself as precisely as the problem allows:

```csharp
foreach (var error in result.Errors)
{
    error.Kind;        // ValidationErrorKind.TypeMismatch
    error.Message;     // "'abc' is not a valid Integer value for column 'ID'."
    error.ColumnName;  // "ID"
    error.Address;     // "A4"   (null for a column that's missing entirely)
    error.Row;         // 4
    error.Column;      // 1
    error.Value;       // "abc"
}
```

Branch on `Kind`, not on `Message` the wording is meant for humans and may change between releases.

| `ValidationErrorKind`  | Meaning |
| :--------------------- | :------ |
| `MissingColumn`        | the schema declares a column the sheet doesn't have |
| `UnexpectedColumn`     | the sheet has a column the schema doesn't declare |
| `DuplicateColumn`      | the same header appears twice |
| `EmptyHeader`          | a header cell is blank |
| `RequiredValueMissing` | a required cell is empty |
| `TypeMismatch`         | the value isn't readable as the column's type |
| `LengthOutOfRange`     | the text is shorter or longer than allowed |
| `ValueNotAllowed`      | the value isn't in `AllowedValues` |
| `PatternMismatch`      | the value doesn't match `Pattern` |

Only the error that explains a cell is reported: a value of the wrong type isn't also reported for failing the length and pattern rules it was never going to satisfy.

## Getting the result as JSON

For returning a validation outcome over an HTTP API, logging it, or handing it to a front end,
serialize the result:

```csharp
var result = validator.Validate(bytes, schema);

string json = result.ToJson();                 // pretty-printed
string compact = result.ToJson(indented: false);
// Or, equivalently: ExcelValidationJson.Serialize(result)
```

Keys are camelCase and enum values are written as names, so the output stays readable:

```json
{
  "isValid": false,
  "columnsAreValid": true,
  "rowsAreValid": false,
  "rowsValidated": 5,
  "truncated": false,
  "missingColumns": [],
  "unexpectedColumns": [],
  "errors": [
    {
      "kind": "typeMismatch",
      "message": "'N/A' is not a valid Integer value for column 'EmpId'.",
      "columnName": "EmpId",
      "address": "A3",
      "row": 3,
      "column": 1,
      "value": "N/A"
    }
  ]
}
```

Keys whose value is null (the address of a missing column, say) are omitted rather than written as
`null`.

## Annotating a workbook

To hand the file back to whoever sent it, with the problems marked in place:

```csharp
var result = validator.Validate(bytes, schema);

if (!result.IsValid)
{
    byte[] annotated = validator.Annotate(bytes, result);
    File.WriteAllBytes("people-errors.xlsx", annotated);
}
```

Every offending cell is filled red and given a comment explaining the problem. Your input array is not modified. Column-level errors land on the header cell, except for a column that's missing entirely, which has no cell to mark.

## Options

```csharp
var options = new ExcelValidationOptions
{
    WorksheetName = "Data",       // by name; otherwise WorksheetPosition (1-based, default 1)
    HeaderRow = 3,                // when the sheet has a title block above the table
    CaseSensitiveHeaders = true,  // default false
    TrimValues = false,           // default true; when true, a cell of spaces counts as empty
    MaxErrors = 100,              // default 1000; stop early and set result.Truncated
};

var result = validator.Validate(bytes, schema, options);
```

`MaxErrors` exists so a wholly malformed 100k-row import can't allocate an error per cell. Set it to `int.MaxValue` to collect everything.

## Errors vs exceptions

Invalid data is a **result**, not an exception that's the normal case this library exists to report on. `ExcelValidationException` is thrown only when there's nothing to validate: the bytes aren't an `.xlsx` file, the file is corrupt or password-protected, or the requested worksheet doesn't exist.

```csharp
try
{
    var result = validator.Validate(bytes, schema);
    // inspect result.Errors
}
catch (ExcelValidationException ex)
{
    // the file itself couldn't be read
}
```

## Dependency injection

`ExcelSheetValidator` is stateless and thread-safe, so register it once:

```csharp
services.AddSingleton<IExcelSheetValidator, ExcelSheetValidator>();
```

## Migrating from 1.x

Version 2.0 is a clean break. The 1.x API is gone rather than deprecated, because the two things that most needed fixing the EPPlus non-commercial licence and the mutable per-call state were both baked into its shape.

**Why the rewrite:** 1.x ran on EPPlus 5, which is licensed [Polyform Noncommercial](https://polyformproject.org/licenses/noncommercial/1.0.0/), and it set `LicenseContext = NonCommercial` on your behalf. Anyone using `excel-validator` 1.x in a commercial product was relying on a licence they never agreed to and likely didn't qualify for. 2.0 runs on ClosedXML under MIT, so the question doesn't arise.

What changed:

| 1.x | 2.0 |
| :-- | :-- |
| `new ValidateExcelSheet(model)`, read `.IsValidFile` | `validator.Validate(bytes, schema)` returns the result |
| `ExcelValidationModel` as both input and output | `ExcelSchema` in, `ExcelValidationResult` out |
| `ValidationType = "Normal"` / `"Data Validation"` | one `Validate` method; the schema says what to check |
| `DataType = "int32"` (string, resolved by reflection) | `ExcelCellType.Integer` (enum, checked at compile time) |
| `ErrorComment`, one concatenated string | `Errors`, a list with `Kind`, `Address`, `Row`, `Column`, `Value` |
| Sheet mutated in place, returned via `UpdatedSheet` | `Annotate(bytes, result)` returns a new workbook |
| Exceptions swallowed into `ErrorComment` | unreadable files throw `ExcelValidationException` |

Before:

```csharp
var model = new ExcelValidationModel
{
    ExcelFile = bytes,
    DataValidation = new Dictionary<string, DataValidationModel>
    {
        { "ID", new DataValidationModel { DataType = "int32", InputType = "Text" } },
        { "Email", new DataValidationModel { DataType = "string", InputType = "Email" } },
    },
    ValidationType = CustomNames.Data_Validation,
};

var validator = new ValidateExcelSheet(model);
if (!validator.IsValidFile.RowIsValid) { /* parse ErrorComment */ }
```

After:

```csharp
var schema = new ExcelSchema()
    .Column("ID", ExcelCellType.Integer)
    .Column("Email", ExcelCellType.Text, pattern: @"^[^@\s]+@[^@\s]+\.[^@\s]+$");

var result = new ExcelSheetValidator().Validate(bytes, schema);
foreach (var error in result.Errors) { /* structured */ }
```

`InputType` has no equivalent: it was carried on the model but never actually checked against anything. Use `Pattern` or `AllowedValues` for what it implied.

## Try it

`ExcelValidator.Sample` is a runnable tour: it builds a workbook, validates it, prints every kind
of error, and writes an annotated copy you can open in Excel.

```
dotnet run --project ExcelValidator.Sample
```

In Visual Studio, set it as the startup project and press F5. (`ExcelValidator` itself is a class
library and has no entry point, so it can't be started directly.)

## Building

```
dotnet build
dotnet test
dotnet pack -c Release
```

## Licence

MIT — see [LICENSE](LICENSE).
