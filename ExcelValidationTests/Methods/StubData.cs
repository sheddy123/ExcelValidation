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

        public IEnumerator<object[]> GetEnumerator()
        {
            yield return new object[] {
                new ExcelValidationModel { ExcelFile = ReturnFile(), HeaderColumns = new List<string>(){ "ID","Username","Email" } }
            };
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}

