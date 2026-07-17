using System.Text.Json;
using ExcelValidator;
using Xunit;

namespace ExcelValidationTests
{
    public class ExcelValidationJsonTests
    {
        private static readonly string[] Headers = { "ID", "Email" };

        private readonly ExcelSheetValidator _validator = new();

        private static ExcelSchema Schema() => new ExcelSchema()
            .Column("ID", ExcelCellType.Integer)
            .Column("Email", ExcelCellType.Text);

        [Fact]
        public void ToJson_ValidResult_ReportsValidAndEmptyErrors()
        {
            var workbook = TestWorkbook.WithRows(Headers, new object?[] { 1, "ada@example.com" });
            var result = _validator.Validate(workbook, Schema());

            using var doc = JsonDocument.Parse(result.ToJson());
            var root = doc.RootElement;

            Assert.True(root.GetProperty("isValid").GetBoolean());
            Assert.Equal(0, root.GetProperty("errors").GetArrayLength());
            Assert.Equal(1, root.GetProperty("rowsValidated").GetInt32());
        }

        [Fact]
        public void ToJson_ErrorResult_SerializesEachErrorWithReadableEnumAndAddress()
        {
            var workbook = TestWorkbook.WithRows(Headers, new object?[] { "not-a-number", "ada@example.com" });
            var result = _validator.Validate(workbook, Schema());

            using var doc = JsonDocument.Parse(result.ToJson());
            var root = doc.RootElement;

            Assert.False(root.GetProperty("isValid").GetBoolean());

            var errors = root.GetProperty("errors");
            Assert.Equal(1, errors.GetArrayLength());

            var error = errors[0];
            // Enum written as a readable camelCase name, not a number.
            Assert.Equal("typeMismatch", error.GetProperty("kind").GetString());
            Assert.Equal("A2", error.GetProperty("address").GetString());
            Assert.Equal("ID", error.GetProperty("columnName").GetString());
            Assert.Equal("not-a-number", error.GetProperty("value").GetString());
            Assert.Equal(2, error.GetProperty("row").GetInt32());
        }

        [Fact]
        public void ToJson_OmitsNullValuedKeys()
        {
            // A missing-column error has no cell, so its address/row/column are null and should not
            // appear in the JSON rather than showing up as null.
            var workbook = TestWorkbook.WithRows(new[] { "ID" }, new object?[] { 1 });
            var result = _validator.Validate(workbook, Schema());

            using var doc = JsonDocument.Parse(result.ToJson());
            var error = doc.RootElement.GetProperty("errors")[0];

            Assert.Equal("missingColumn", error.GetProperty("kind").GetString());
            Assert.False(error.TryGetProperty("address", out _));
            Assert.False(error.TryGetProperty("row", out _));
        }

        [Fact]
        public void ToJson_Compact_ProducesSingleLine()
        {
            var workbook = TestWorkbook.WithRows(Headers, new object?[] { 1, "ada@example.com" });
            var result = _validator.Validate(workbook, Schema());

            var compact = result.ToJson(indented: false);

            Assert.DoesNotContain('\n', compact);
        }

        [Fact]
        public void Serialize_NullResult_Throws()
        {
            Assert.Throws<System.ArgumentNullException>(() => ExcelValidationJson.Serialize(null!));
        }
    }
}
