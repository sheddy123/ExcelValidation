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

        public ExcelValidationModel IsValidFile
        {
            get => _excelFile;
        }
        /// <summary>
        /// Takes an instance of an excel file and validates all the fields
        /// </summary>
        /// <param name="ExcelSheet"></param>
        /// <param name=""></param>
        public ValidateExcelSheet(ExcelValidationModel excelFile) => _excelFile = ValidateExcel(excelFile);


        /// <summary>
        /// Takes in ExcelValidationModel that validates rows and columns
        /// </summary>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public ExcelValidationModel ValidateExcel(ExcelValidationModel excelFile)
        {
            //Validates the column(s)
            excelFile.ColumnIsValid = ValidationMethods.ValidateExcelColumns(excelFile);

            //Validates the row(s)
            excelFile.RowIsValid = ValidationMethods.ValidateExcelRows(excelFile);

            return excelFile;
        }


    }
}
