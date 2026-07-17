using System;
using System.Linq;
using ExcelValidator;
using Xunit;

namespace ExcelValidationTests
{
    public class ExcelSchemaTests
    {
        [Fact]
        public void Column_ChainedCalls_BuildColumnsInOrder()
        {
            var schema = new ExcelSchema()
                .Column("ID", ExcelCellType.Integer)
                .Column("Email");

            Assert.Equal(new[] { "ID", "Email" }, schema.Columns.Select(c => c.Name));
            Assert.Equal(ExcelCellType.Integer, schema.Columns[0].Type);
            Assert.Equal(ExcelCellType.Text, schema.Columns[1].Type);
            Assert.True(schema.Columns[1].Required);
        }

        [Fact]
        public void FromHeaders_BuildsRequiredTextColumns()
        {
            var schema = ExcelSchema.FromHeaders("ID", "Email");

            Assert.All(schema.Columns, c =>
            {
                Assert.Equal(ExcelCellType.Text, c.Type);
                Assert.True(c.Required);
            });
        }

        [Fact]
        public void Column_DuplicateNameRegardlessOfCasing_Throws()
        {
            var schema = new ExcelSchema().Column("ID");

            Assert.Throws<ArgumentException>(() => schema.Column("id"));
        }

        [Fact]
        public void Schema_IsEnumerable()
        {
            var schema = new ExcelSchema().Column("ID").Column("Email");

            Assert.Equal(2, schema.Count());
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void ColumnRule_WithoutAName_Throws(string? name)
        {
            Assert.Throws<ArgumentException>(() => new ColumnRule(name!));
        }

        [Fact]
        public void ColumnRule_TrimsItsName()
        {
            Assert.Equal("ID", new ColumnRule("  ID  ").Name);
        }

        [Fact]
        public void ColumnRule_InvalidPattern_ThrowsWhereItIsSet()
        {
            // Better to fail here than on whichever cell first reaches the regex. ThrowsAny because
            // .NET raises RegexParseException, an ArgumentException subclass.
            Assert.ThrowsAny<ArgumentException>(() => new ColumnRule("Email") { Pattern = "([unclosed" });
        }
    }

    public class ExcelValidationOptionsTests
    {
        [Fact]
        public void Defaults_AreTheCommonCase()
        {
            var options = new ExcelValidationOptions();

            Assert.Null(options.WorksheetName);
            Assert.Equal(1, options.WorksheetPosition);
            Assert.Equal(1, options.HeaderRow);
            Assert.False(options.CaseSensitiveHeaders);
            Assert.True(options.TrimValues);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        public void RowAndPositionNumbers_AreOneBased(int value)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelValidationOptions { HeaderRow = value });
            Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelValidationOptions { WorksheetPosition = value });
            Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelValidationOptions { MaxErrors = value });
        }
    }
}
