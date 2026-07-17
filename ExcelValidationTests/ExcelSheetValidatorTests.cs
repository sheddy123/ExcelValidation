using System;
using System.IO;
using System.Linq;
using ExcelValidator;
using Xunit;

namespace ExcelValidationTests
{
    public class ExcelSheetValidatorTests
    {
        private static readonly string[] PeopleHeaders = { "ID", "Username", "Email" };

        private readonly ExcelSheetValidator _validator = new();

        private static ExcelSchema PeopleSchema() => new ExcelSchema()
            .Column("ID", ExcelCellType.Integer)
            .Column("Username", ExcelCellType.Text)
            .Column("Email", ExcelCellType.Text);

        [Fact]
        public void Validate_WorkbookMatchingSchema_IsValid()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 1, "ada", "ada@example.com" },
                new object?[] { 2, "grace", "grace@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            Assert.True(result.IsValid);
            Assert.Empty(result.Errors);
            Assert.Equal(2, result.RowsValidated);
            Assert.True(result.ColumnsAreValid);
            Assert.True(result.RowsAreValid);
        }

        [Fact]
        public void Validate_HeaderMissingFromWorksheet_ReportsMissingColumn()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "ID", "Username" },
                new object?[] { 1, "ada" });

            var result = _validator.Validate(workbook, PeopleSchema());

            Assert.False(result.IsValid);
            Assert.False(result.ColumnsAreValid);
            Assert.Equal(new[] { "Email" }, result.MissingColumns);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.MissingColumn, error.Kind);
            Assert.Equal("Email", error.ColumnName);
        }

        [Fact]
        public void Validate_WorksheetHasColumnSchemaDoesNotDeclare_ReportsUnexpectedColumn()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "ID", "Username", "Email", "Nickname" },
                new object?[] { 1, "ada", "ada@example.com", "addie" });

            var result = _validator.Validate(workbook, PeopleSchema());

            Assert.False(result.IsValid);
            Assert.Equal(new[] { "Nickname" }, result.UnexpectedColumns);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.UnexpectedColumn, error.Kind);
            Assert.Equal("D1", error.Address);
        }

        [Fact]
        public void Validate_AllowUnexpectedColumns_ReportsThemWithoutFailing()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "ID", "Username", "Email", "Nickname" },
                new object?[] { 1, "ada", "ada@example.com", "addie" });

            var schema = PeopleSchema();
            schema.AllowUnexpectedColumns = true;

            var result = _validator.Validate(workbook, schema);

            Assert.True(result.IsValid);
            Assert.Equal(new[] { "Nickname" }, result.UnexpectedColumns);
        }

        [Fact]
        public void Validate_EmptyCellInRequiredColumn_ReportsRequiredValueMissing()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 1, null, "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.RequiredValueMissing, error.Kind);
            Assert.Equal("Username", error.ColumnName);
            Assert.Equal("B2", error.Address);
            Assert.Equal(2, error.Row);
            Assert.False(result.RowsAreValid);
        }

        [Fact]
        public void Validate_WhitespaceOnlyCell_CountsAsEmpty()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 1, "   ", "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.RequiredValueMissing, error.Kind);
        }

        [Fact]
        public void Validate_EmptyCellInOptionalColumn_IsValid()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 1, null, "ada@example.com" });

            var schema = new ExcelSchema()
                .Column("ID", ExcelCellType.Integer)
                .Column("Username", ExcelCellType.Text, required: false)
                .Column("Email", ExcelCellType.Text);

            var result = _validator.Validate(workbook, schema);

            Assert.True(result.IsValid);
        }

        [Fact]
        public void Validate_TextWhereIntegerExpected_ReportsTypeMismatch()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { "not-a-number", "ada", "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.TypeMismatch, error.Kind);
            Assert.Equal("A2", error.Address);
            Assert.Equal("not-a-number", error.Value);
        }

        [Fact]
        public void Validate_FractionalNumberWhereIntegerExpected_ReportsTypeMismatch()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 1.5, "ada", "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.TypeMismatch, error.Kind);
        }

        [Fact]
        public void Validate_IntegerStoredAsTextOrAsNumber_BothSatisfyIntegerColumn()
        {
            // Whether a producer wrote 34 as a number or as the string "34" is an artifact of the
            // tool that wrote the file, not a statement about whether the data is correct.
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 34, "ada", "ada@example.com" },
                new object?[] { "35", "grace", "grace@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            Assert.True(result.IsValid);
        }

        [Fact]
        public void Validate_ValueTooLong_ReportsLengthOutOfRange()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { 1, "a-very-long-username", "ada@example.com" });

            var schema = new ExcelSchema()
                .Column("ID", ExcelCellType.Integer)
                .Column("Username", ExcelCellType.Text, maxLength: 8)
                .Column("Email", ExcelCellType.Text);

            var result = _validator.Validate(workbook, schema);

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.LengthOutOfRange, error.Kind);
        }

        [Fact]
        public void Validate_ValueOutsideAllowedValues_ReportsValueNotAllowed()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "Status" },
                new object?[] { "Pending" },
                new object?[] { "Banana" });

            var schema = new ExcelSchema()
                .Column("Status", allowedValues: new[] { "Pending", "Active" });

            var result = _validator.Validate(workbook, schema);

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.ValueNotAllowed, error.Kind);
            Assert.Equal("Banana", error.Value);
        }

        [Fact]
        public void Validate_ValueNotMatchingPattern_ReportsPatternMismatch()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "Email" },
                new object?[] { "ada@example.com" },
                new object?[] { "not-an-email" });

            var schema = new ExcelSchema()
                .Column("Email", pattern: @"^[^@\s]+@[^@\s]+\.[^@\s]+$");

            var result = _validator.Validate(workbook, schema);

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.PatternMismatch, error.Kind);
            Assert.Equal("A3", error.Address);
        }

        [Fact]
        public void Validate_TypeMismatch_DoesNotAlsoReportDownstreamRules()
        {
            // A value of the wrong type would fail the length and pattern checks too; reporting all
            // three would bury the one that explains the problem.
            var workbook = TestWorkbook.WithRows(
                new[] { "Age" },
                new object?[] { "abc" });

            var schema = new ExcelSchema()
                .Column("Age", ExcelCellType.Integer, maxLength: 1, pattern: @"^\d+$");

            var result = _validator.Validate(workbook, schema);

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.TypeMismatch, error.Kind);
        }

        [Theory]
        [InlineData(ExcelCellType.Boolean, "true")]
        [InlineData(ExcelCellType.Boolean, "FALSE")]
        [InlineData(ExcelCellType.Guid, "6f9619ff-8b86-d011-b42d-00cf4fc964ff")]
        [InlineData(ExcelCellType.DateTime, "2024-01-31")]
        [InlineData(ExcelCellType.Decimal, "3.14")]
        [InlineData(ExcelCellType.Double, "-2.5e3")]
        [InlineData(ExcelCellType.Long, "9223372036854775806")]
        public void Validate_ValueMatchingDeclaredType_IsValid(ExcelCellType type, string value)
        {
            var workbook = TestWorkbook.WithRows(new[] { "Value" }, new object?[] { value });

            var result = _validator.Validate(workbook, new ExcelSchema().Column("Value", type));

            Assert.True(result.IsValid, $"expected '{value}' to be a valid {type}");
        }

        [Theory]
        [InlineData(ExcelCellType.Boolean, "yes")]
        [InlineData(ExcelCellType.Guid, "not-a-guid")]
        [InlineData(ExcelCellType.DateTime, "the 4th")]
        [InlineData(ExcelCellType.Decimal, "3.1.4")]
        [InlineData(ExcelCellType.Integer, "99999999999")]
        public void Validate_ValueNotMatchingDeclaredType_ReportsTypeMismatch(ExcelCellType type, string value)
        {
            var workbook = TestWorkbook.WithRows(new[] { "Value" }, new object?[] { value });

            var result = _validator.Validate(workbook, new ExcelSchema().Column("Value", type));

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.TypeMismatch, error.Kind);
        }

        [Fact]
        public void Validate_NativeExcelDateAndBoolean_SatisfyTheirColumns()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "When", "Flag" },
                new object?[] { new DateTime(2024, 1, 31), true });

            var schema = new ExcelSchema()
                .Column("When", ExcelCellType.DateTime)
                .Column("Flag", ExcelCellType.Boolean);

            var result = _validator.Validate(workbook, schema);

            Assert.True(result.IsValid);
        }

        [Fact]
        public void Validate_HeaderCasingDiffers_MatchesByDefault()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "id", "USERNAME", "eMaIl" },
                new object?[] { 1, "ada", "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());

            Assert.True(result.IsValid);
        }

        [Fact]
        public void Validate_HeaderCasingDiffers_FailsWhenCaseSensitivityRequested()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "id", "Username", "Email" },
                new object?[] { 1, "ada", "ada@example.com" });

            var result = _validator.Validate(
                workbook,
                PeopleSchema(),
                new ExcelValidationOptions { CaseSensitiveHeaders = true });

            Assert.Contains(result.Errors, e => e.Kind == ValidationErrorKind.MissingColumn && e.ColumnName == "ID");
        }

        [Fact]
        public void Validate_DuplicateHeader_ReportsDuplicateColumn()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "ID", "ID" },
                new object?[] { 1, 2 });

            var result = _validator.Validate(workbook, new ExcelSchema().Column("ID", ExcelCellType.Integer));

            Assert.Contains(result.Errors, e => e.Kind == ValidationErrorKind.DuplicateColumn);
        }

        [Fact]
        public void Validate_EmptyHeaderCellBetweenColumns_ReportsEmptyHeader()
        {
            var workbook = TestWorkbook.Build(sheet =>
            {
                sheet.Cell(1, 1).Value = "ID";
                // B1 deliberately left blank.
                sheet.Cell(1, 3).Value = "Email";
                sheet.Cell(2, 1).Value = 1;
                sheet.Cell(2, 3).Value = "ada@example.com";
            });

            var schema = new ExcelSchema()
                .Column("ID", ExcelCellType.Integer)
                .Column("Email", ExcelCellType.Text);

            var result = _validator.Validate(workbook, schema);

            var error = Assert.Single(result.Errors);
            Assert.Equal(ValidationErrorKind.EmptyHeader, error.Kind);
            Assert.Equal("B1", error.Address);
        }

        [Fact]
        public void Validate_BlankRow_IsSkippedRatherThanReportedAsMissingValues()
        {
            var workbook = TestWorkbook.Build(sheet =>
            {
                sheet.Cell(1, 1).Value = "ID";
                sheet.Cell(2, 1).Value = 1;
                // Row 3 left entirely blank.
                sheet.Cell(4, 1).Value = 2;
            });

            var result = _validator.Validate(workbook, new ExcelSchema().Column("ID", ExcelCellType.Integer));

            Assert.True(result.IsValid);
            Assert.Equal(2, result.RowsValidated);
        }

        [Fact]
        public void Validate_HeaderRowOption_ReadsHeadersFromThatRow()
        {
            var workbook = TestWorkbook.Build(sheet =>
            {
                sheet.Cell(1, 1).Value = "Monthly report";
                sheet.Cell(3, 1).Value = "ID";
                sheet.Cell(4, 1).Value = 1;
            });

            var result = _validator.Validate(
                workbook,
                new ExcelSchema().Column("ID", ExcelCellType.Integer),
                new ExcelValidationOptions { HeaderRow = 3 });

            Assert.True(result.IsValid);
            Assert.Equal(1, result.RowsValidated);
        }

        [Fact]
        public void Validate_WorksheetSelectedByName_ValidatesThatSheet()
        {
            var workbook = TestWorkbook.Build(
                sheet =>
                {
                    sheet.Cell(1, 1).Value = "ID";
                    sheet.Cell(2, 1).Value = "not-a-number";
                },
                sheetName: "Data");

            var result = _validator.Validate(
                workbook,
                new ExcelSchema().Column("ID", ExcelCellType.Integer),
                new ExcelValidationOptions { WorksheetName = "Data" });

            Assert.Single(result.Errors);
        }

        [Fact]
        public void Validate_WorksheetNameNotInWorkbook_Throws()
        {
            var workbook = TestWorkbook.WithRows(new[] { "ID" }, new object?[] { 1 });

            var ex = Assert.Throws<ExcelValidationException>(() => _validator.Validate(
                workbook,
                new ExcelSchema().Column("ID"),
                new ExcelValidationOptions { WorksheetName = "Nope" }));

            Assert.Contains("Nope", ex.Message, StringComparison.Ordinal);
            Assert.Contains(TestWorkbook.DefaultSheetName, ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Validate_WorksheetPositionBeyondWorkbook_Throws()
        {
            var workbook = TestWorkbook.WithRows(new[] { "ID" }, new object?[] { 1 });

            Assert.Throws<ExcelValidationException>(() => _validator.Validate(
                workbook,
                new ExcelSchema().Column("ID"),
                new ExcelValidationOptions { WorksheetPosition = 5 }));
        }

        [Fact]
        public void Validate_ContentThatIsNotAWorkbook_ThrowsExcelValidationException()
        {
            var notAWorkbook = new byte[] { 0x00, 0x01, 0x02, 0x03 };

            var ex = Assert.Throws<ExcelValidationException>(
                () => _validator.Validate(notAWorkbook, PeopleSchema()));

            Assert.NotNull(ex.InnerException);
        }

        [Fact]
        public void Validate_MaxErrors_StopsEarlyAndFlagsTruncation()
        {
            var rows = Enumerable.Range(0, 50)
                .Select(_ => new object?[] { "not-a-number", "ada", "ada@example.com" })
                .ToArray();
            var workbook = TestWorkbook.WithRows(PeopleHeaders, rows);

            var result = _validator.Validate(
                workbook,
                PeopleSchema(),
                new ExcelValidationOptions { MaxErrors = 5 });

            Assert.Equal(5, result.Errors.Count);
            Assert.True(result.Truncated);
        }

        [Fact]
        public void Validate_UnderACommaDecimalCulture_StillReadsDotDecimals()
        {
            // Excel stores numbers culture-neutrally, so validation must not depend on the culture
            // the server happens to run under.
            var workbook = TestWorkbook.WithRows(new[] { "Price" }, new object?[] { "3.14" });
            var schema = new ExcelSchema().Column("Price", ExcelCellType.Decimal);

            var original = System.Globalization.CultureInfo.CurrentCulture;
            try
            {
                System.Globalization.CultureInfo.CurrentCulture =
                    new System.Globalization.CultureInfo("de-DE");
                var result = _validator.Validate(workbook, schema);
                Assert.True(result.IsValid);
            }
            finally
            {
                System.Globalization.CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public void Validate_Stream_DoesNotDisposeTheCallersStream()
        {
            var workbook = TestWorkbook.WithRows(PeopleHeaders, new object?[] { 1, "ada", "ada@example.com" });
            using var stream = new MemoryStream(workbook);

            _validator.Validate(stream, PeopleSchema());

            Assert.True(stream.CanRead);
        }

        [Fact]
        public void Validate_SameWorkbookTwice_ProducesTheSameResult()
        {
            // The 1.x design mutated shared state while walking the sheet, so a second run saw the
            // leftovers of the first. Validation must be a pure function of its inputs.
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { "not-a-number", "ada", "ada@example.com" });
            var schema = PeopleSchema();

            var first = _validator.Validate(workbook, schema);
            var second = _validator.Validate(workbook, schema);

            Assert.Equal(first.Errors.Count, second.Errors.Count);
            Assert.Equal(first.Errors[0].Address, second.Errors[0].Address);
        }

        [Fact]
        public void Validate_NullArguments_Throw()
        {
            var workbook = TestWorkbook.WithRows(new[] { "ID" }, new object?[] { 1 });

            Assert.Throws<ArgumentNullException>(() => _validator.Validate((byte[])null!, PeopleSchema()));
            Assert.Throws<ArgumentNullException>(() => _validator.Validate((Stream)null!, PeopleSchema()));
            Assert.Throws<ArgumentNullException>(() => _validator.Validate(workbook, null!));
        }

        [Fact]
        public void Annotate_FillsBadCellsRedAndCommentsThem()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { "not-a-number", "ada", "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());
            var annotated = _validator.Annotate(workbook, result);

            TestWorkbook.Read(annotated, sheet =>
            {
                var bad = sheet.Cell("A2");
                Assert.Equal(ClosedXML.Excel.XLColor.Red, bad.Style.Fill.BackgroundColor);
                Assert.True(bad.HasComment);

                var good = sheet.Cell("B2");
                Assert.NotEqual(ClosedXML.Excel.XLColor.Red, good.Style.Fill.BackgroundColor);
                Assert.False(good.HasComment);
            });
        }

        [Fact]
        public void Annotate_DoesNotModifyTheCallersArray()
        {
            var workbook = TestWorkbook.WithRows(
                PeopleHeaders,
                new object?[] { "not-a-number", "ada", "ada@example.com" });
            var copy = (byte[])workbook.Clone();

            var result = _validator.Validate(workbook, PeopleSchema());
            _validator.Annotate(workbook, result);

            Assert.Equal(copy, workbook);
        }

        [Fact]
        public void Annotate_UnexpectedColumn_AnnotatesTheHeaderCell()
        {
            var workbook = TestWorkbook.WithRows(
                new[] { "ID", "Username", "Email", "Nickname" },
                new object?[] { 1, "ada", "ada@example.com", "addie" });

            var result = _validator.Validate(workbook, PeopleSchema());
            var annotated = _validator.Annotate(workbook, result);

            TestWorkbook.Read(annotated, sheet => Assert.True(sheet.Cell("D1").HasComment));
        }

        [Fact]
        public void Annotate_ValidWorkbook_LeavesItUnmarked()
        {
            var workbook = TestWorkbook.WithRows(PeopleHeaders, new object?[] { 1, "ada", "ada@example.com" });

            var result = _validator.Validate(workbook, PeopleSchema());
            var annotated = _validator.Annotate(workbook, result);

            TestWorkbook.Read(annotated, sheet => Assert.False(sheet.Cell("A2").HasComment));
        }
    }
}
