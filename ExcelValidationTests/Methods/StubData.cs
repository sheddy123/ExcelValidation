/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/
using ExcelValidator.Model;
using Microsoft.Extensions.FileProviders;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelValidator.Model.ValidateExcelSheet;

namespace ExcelValidationTests.Methods
{
    public class StubData : IEnumerable<object[]>
    {

        private byte[] ReturnFile()
        {
            var filepaths = "C:\\Users\\iodom\\source\\repos\\ExcelValidation\\ExcelValidationTests\\Files\\DemoFiles.xlsx";
            //var filepaths = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), "Files", "DemoFile")).Root;
            byte[] fileByte = System.IO.File.ReadAllBytes(filepaths);

            return fileByte;
        }
        private Dictionary<string, DataValidationModel> DataValidationStub()
        {
            
            Dictionary<string, DataValidationModel> stub = new Dictionary<string, DataValidationModel>();
            stub.Add("ID", new DataValidationModel() { DataType = "int32", InputType = "Text" } );
            stub.Add("Username", new DataValidationModel() { DataType = "string", InputType = "Text" });
            stub.Add("Email", new DataValidationModel() { DataType = "string", InputType = "Email" });
            
            return stub;
        }
        public IEnumerator<object[]> GetEnumerator()
        {
            yield return new object[] {
                new ExcelValidationModel { ExcelFile = ReturnFile(), HeaderColumns = new List<string>(){ "ID","Username","Email" },ValidationType = CustomNames.NormalVal }
            };
            #region PR#7 Data Validation of Excel Rows and Columns
            yield return new object[] {
                new ExcelValidationModel { ExcelFile = ReturnFile(), DataValidation = DataValidationStub(), ValidationType = CustomNames.Data_Validation }
            };
            #endregion
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}

