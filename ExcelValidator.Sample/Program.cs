using ClosedXML.Excel;
using ExcelValidator;

// A runnable tour of the library. Press F5, or: dotnet run --project ExcelValidator.Sample
//
// Nothing here reads a checked-in file: the sample writes its own workbook first, so what is
// being validated is visible in the code rather than hidden in a fixture.

var validator = new ExcelSheetValidator();
var outputDirectory = Path.Combine(AppContext.BaseDirectory, "output");
Directory.CreateDirectory(outputDirectory);

// The columns we expect. This is the same schema the README opens with.
var schema = new ExcelSchema()
    .Column("ID", ExcelCellType.Integer)
    .Column("Username", ExcelCellType.Text, maxLength: 12)
    .Column("Email", ExcelCellType.Text, pattern: @"^[^@\s]+@[^@\s]+\.[^@\s]+$")
    .Column("Signed up", ExcelCellType.DateTime)
    .Column("Status", allowedValues: new[] { "Active", "Pending" })
    .Column("Notes", ExcelCellType.Text, required: false);

Heading("1. A workbook that satisfies the schema");

var goodWorkbook = BuildWorkbook(sheet =>
{
    WriteHeaders(sheet);
    WriteRow(sheet, 2, 1, "ada", "ada@example.com", new DateTime(2024, 1, 15), "Active", "founder");
    WriteRow(sheet, 3, 2, "grace", "grace@example.com", new DateTime(2024, 3, 2), "Pending", null);
});

Report(validator.Validate(goodWorkbook, schema));

Heading("2. A workbook with one of every kind of problem");

var badWorkbook = BuildWorkbook(sheet =>
{
    WriteHeaders(sheet);
    sheet.Cell(1, 7).Value = "Nickname";                    // UnexpectedColumn

    // A good row, to show only the bad cells get reported.
    WriteRow(sheet, 2, 1, "ada", "ada@example.com", new DateTime(2024, 1, 15), "Active", null);

    WriteRow(sheet, 3, "not-a-number", "grace", "grace@example.com", new DateTime(2024, 3, 2), "Active", null);
    //          TypeMismatch: "not-a-number" is not an Integer

    WriteRow(sheet, 4, 3, "a-username-far-too-long", "linus@example.com", new DateTime(2024, 4, 1), "Active", null);
    //                     LengthOutOfRange: longer than maxLength 12

    WriteRow(sheet, 5, 4, "linus", "not-an-email", new DateTime(2024, 5, 9), "Active", null);
    //                                PatternMismatch

    WriteRow(sheet, 6, 5, "edsger", "edsger@example.com", new DateTime(2024, 6, 1), "Banana", null);
    //                                                                              ValueNotAllowed

    WriteRow(sheet, 7, 6, null, "alan@example.com", new DateTime(2024, 7, 4), "Pending", null);
    //                    RequiredValueMissing

    // Row 8 left blank entirely, to show blank rows are skipped rather than reported.
    WriteRow(sheet, 9, 7, "barbara", "barbara@example.com", new DateTime(2024, 8, 8), "Active", null);
});

var result = validator.Validate(badWorkbook, schema);
Report(result);

Heading("3. Handing the file back with the problems marked");

var annotated = validator.Annotate(badWorkbook, result);
var annotatedPath = Path.Combine(outputDirectory, "people-with-errors.xlsx");
File.WriteAllBytes(annotatedPath, annotated);

Console.WriteLine($"Wrote {annotatedPath}");
Console.WriteLine("Every bad cell is filled red and carries a comment explaining why.");
Console.WriteLine("Open it in Excel to see what the person who sent you the file would see.");

Heading("4. An unreadable file is an exception, not a result");

try
{
    validator.Validate(new byte[] { 0x00, 0x01, 0x02 }, schema);
    Console.WriteLine("unreachable");
}
catch (ExcelValidationException ex)
{
    Console.WriteLine($"Caught ExcelValidationException as expected:");
    Console.WriteLine($"  {ex.Message}");
}

Console.WriteLine();
Console.WriteLine("Invalid data is reported through the result. Only a file that cannot be read at");
Console.WriteLine("all throws, so a bad upload never needs a try/catch to handle normally.");
Console.WriteLine();

// ---------------------------------------------------------------------------------------------

static void Report(ExcelValidationResult result)
{
    Console.WriteLine($"IsValid          {result.IsValid}");
    Console.WriteLine($"RowsValidated    {result.RowsValidated}");
    Console.WriteLine($"ColumnsAreValid  {result.ColumnsAreValid}");
    Console.WriteLine($"RowsAreValid     {result.RowsAreValid}");

    if (result.MissingColumns.Count > 0)
    {
        Console.WriteLine($"MissingColumns   {string.Join(", ", result.MissingColumns)}");
    }

    if (result.UnexpectedColumns.Count > 0)
    {
        Console.WriteLine($"UnexpectedCols   {string.Join(", ", result.UnexpectedColumns)}");
    }

    if (result.Errors.Count == 0)
    {
        Console.WriteLine("\nNo errors.");
        return;
    }

    Console.WriteLine($"\n{result.Errors.Count} error(s):\n");
    Console.WriteLine($"  {"CELL",-6} {"KIND",-22} {"COLUMN",-11} DETAIL");
    Console.WriteLine($"  {new string('-', 6)} {new string('-', 22)} {new string('-', 11)} {new string('-', 40)}");

    foreach (var error in result.Errors)
    {
        Console.WriteLine($"  {error.Address ?? "-",-6} {error.Kind,-22} {error.ColumnName ?? "-",-11} {error.Message}");
    }
}

static void Heading(string title)
{
    Console.WriteLine();
    Console.WriteLine(new string('=', 94));
    Console.WriteLine(title);
    Console.WriteLine(new string('=', 94));
    Console.WriteLine();
}

static void WriteHeaders(IXLWorksheet sheet)
{
    string[] headers = { "ID", "Username", "Email", "Signed up", "Status", "Notes" };
    for (var i = 0; i < headers.Length; i++)
    {
        sheet.Cell(1, i + 1).Value = headers[i];
        sheet.Cell(1, i + 1).Style.Font.Bold = true;
    }
}

static void WriteRow(IXLWorksheet sheet, int row, object? id, string? username, string? email, DateTime signedUp, string status, string? notes)
{
    if (id is int i)
    {
        sheet.Cell(row, 1).Value = i;
    }
    else if (id is string s)
    {
        sheet.Cell(row, 1).Value = s;
    }

    if (username is not null)
    {
        sheet.Cell(row, 2).Value = username;
    }

    if (email is not null)
    {
        sheet.Cell(row, 3).Value = email;
    }

    sheet.Cell(row, 4).Value = signedUp;
    sheet.Cell(row, 5).Value = status;

    if (notes is not null)
    {
        sheet.Cell(row, 6).Value = notes;
    }
}

static byte[] BuildWorkbook(Action<IXLWorksheet> configure)
{
    using var book = new XLWorkbook();
    var sheet = book.AddWorksheet("People");
    configure(sheet);

    using var stream = new MemoryStream();
    book.SaveAs(stream);
    return stream.ToArray();
}
