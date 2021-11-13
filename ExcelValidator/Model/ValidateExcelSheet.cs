/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/
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
        #region private property field excelFile used for modifications of data
        private ExcelValidationModel _excelFile { get; }
        #endregion

        #region Property to set the condition of the file from true or false
        public ExcelValidationModel IsValidFile
        {
            get => _excelFile;
        }
        #endregion

        #region Constructor that takes in the excel file model [ExcelValidationModel]
        /// <summary>
        /// Takes an instance of an excel file and validates all the fields
        /// </summary>
        /// <param name="ExcelSheet"></param>
        /// <param name=""></param>
        public ValidateExcelSheet(ExcelValidationModel excelFile) => _excelFile = ValidateExcel(excelFile);
        #endregion

        #region Method for Validating Excel Sheet
        /// <summary>
        /// Takes in ExcelValidationModel that validates rows and columns
        /// </summary>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public ExcelValidationModel ValidateExcel(ExcelValidationModel excelFile)
        {
            try
            {
                switch (excelFile.ValidationType)
                {
                    case CustomNames.NormalVal:
                        //Validates the column(s)
                        excelFile.ColumnIsValid = ValidationMethods.ValidateExcelColumns(excelFile);
                        //Validates the row(s)
                        excelFile.RowIsValid = ValidationMethods.ValidateExcelRows(excelFile); break;

                    case CustomNames.Data_Validation:
                        excelFile.DataValidation = (Dictionary<string, DataValidationModel>)excelFile.DataValidation;
                        //Validates the column(s)
                        excelFile.ColumnIsValid = ValidationMethods.DataValidateExcelColumns(excelFile);
                        //Validates the row(s)
                        excelFile.RowIsValid = ValidationMethods.DataTypeValidateExcelRows(excelFile); break;
                    default: break;
                }
                return excelFile;
            }
            catch (Exception ex)
            {
                return new ExcelValidationModel { ErrorComment = ex.Message };
            }
        }
        #endregion

    }
}
