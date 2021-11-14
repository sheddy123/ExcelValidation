/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/
using System;
using Xunit;
using ExcelValidator;
using ExcelValidationTests.Methods;
using System.IO;
using Microsoft.Extensions.FileProviders;
using ExcelValidator.Model;
using static ExcelValidator.Model.ValidateExcelSheet;
using OfficeOpenXml;
using System.ComponentModel;
using System.Linq;

namespace ExcelValidationTests
{
    public class ValidateExcelTest
    {
        [Theory]
        [ClassData(typeof(StubData))]
        public void Test1(ExcelValidationModel stubData)
        {
            //if (stubData.DataValidation != null)
            //{
            //    var key = stubData.DataValidation.Keys.Take(1).First();
                
            //    var key2 = stubData.DataValidation.Keys.Skip(0).Take(1);
            //    var type = Type.GetType("System.Double");
            //    var dat = Type.GetType("System.DateTime");

            //    var dd = Convert.ChangeType("34", type);
            //    var isValid1 = TypeDescriptor.GetConverter(type).IsValid("34");
            //    var isValid2 = TypeDescriptor.GetConverter(dat).IsValid("34");
            //    var dde = Convert.ChangeType("34", dat) ?? null;
            //    var dhh = type.Name;
            //}
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
            Assert.False(validator.IsValidFile.RowIsValid);
            Assert.True(validator.IsValidFile.ColumnIsValid);
        }
    }
}
