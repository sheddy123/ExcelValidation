using ExcelValidator.Interfaces;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelValidator.Model
{
    public partial class ValidateExcelSheet : IExcelValidator
    {
        private ExcelValidationModel _excelFile { get; }

        public ExcelValidationModel IsValidFile {
            get => _excelFile;
        }
        /// <summary>
        /// Takes an instance of an excel file and validates all the fields
        /// </summary>
        /// <param name="ExcelSheet"></param>
        /// <param name=""></param>
        public ValidateExcelSheet(ExcelValidationModel excelFile) => _excelFile = ValidateExcel(excelFile);
        

        public ExcelValidationModel ValidateExcel(ExcelValidationModel excelFile)
        {
            var headerIsValid = ValidationMethods.ValidateExcelHeader(excelFile);
            if (!headerIsValid)
                return new ExcelValidationModel { ErrorComment = CustomErrors.InvalidRows };

            //var rowsValid
            
            
            return new ExcelValidationModel { Comment = CustomErrors.ValidRows };
        }

     
    }
}
