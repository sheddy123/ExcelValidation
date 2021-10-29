using System;
using Xunit;
using ExcelValidator;
using ExcelValidationTests.Methods;
using System.IO;
using Microsoft.Extensions.FileProviders;
using ExcelValidator.Model;
using static ExcelValidator.Model.ValidateExcelSheet;

namespace ExcelValidationTests
{
    public class ValidateExcelTest
    {
        [Theory]
        [ClassData(typeof(StubData))]
        public void Test1(ExcelValidationModel stubData)
        {
            var validator = new ValidateExcelSheet(stubData);
            Assert.True(validator.IsValidFile.RowIsValid);
            Assert.True(validator.IsValidFile.ColumnIsValid);
        }
    }
}
