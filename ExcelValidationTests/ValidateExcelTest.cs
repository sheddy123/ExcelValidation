using System;
using Xunit;
using ExcelValidator;
using ExcelValidationTests.Methods;
using System.IO;
using Microsoft.Extensions.FileProviders;
using ExcelValidator.Model;
using static ExcelValidator.Model.ValidateExcelSheet;
using OfficeOpenXml;

namespace ExcelValidationTests
{
    public class ValidateExcelTest
    {
        [Theory]
        [ClassData(typeof(StubData))]
        public void Test1(ExcelValidationModel stubData)
        {
            var validator = new ValidateExcelSheet(stubData);
            //var stream = new System.IO.MemoryStream();
            //using (var pck = new ExcelPackage(stream))
            //{
            //    var wds = pck.Workbook.Worksheets.Add("Worksheets-Name", validator.IsValidFile.UpdatedSheet.Workbook.Worksheets[0]);
            //    var filepaths = "C:\\Users\\iodom\\source\\repos\\ExcelValidation\\ExcelValidationTests\\Files\\";
            //    string fullPath = Path.Combine(filepaths, "Dmtest.xlsx");
            //    FileInfo fi = new FileInfo(fullPath);
            //    pck.SaveAs(fi);
            //}
            Assert.True(validator.IsValidFile.RowIsValid);
            Assert.True(validator.IsValidFile.ColumnIsValid);
        }
    }
}
