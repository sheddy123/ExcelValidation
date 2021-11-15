/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelValidator.Model
{
    public partial class ValidateExcelSheet
    {
        public class ExcelValidationModel
        {
            public string ColumnName { get; }

            public ExcelWorksheet UpdatedSheet { get; set; }

            private List<HashSet<string>> _addRowEntriesList = new List<HashSet<string>>();

            public List<HashSet<string>> AddRowEntriesList
            {
                get => _addRowEntriesList;
                set
                {
                    _addRowEntriesList = value;
                }
            }


            public int Row { get; set; }

            public int Column { get; set; }

            public string Comment { get; set; }

            private bool _isValidRow;

            private bool _isValidColumn;

            private string _errorComment;

            public bool RowIsValid
            {
                get => _isValidRow;
                set
                {
                    _isValidRow = value;
                }
            }

            public bool ColumnIsValid
            {
                get => _isValidColumn;
                set
                {
                    _isValidColumn = value;
                }
            }

            public int EndRow { get; set; }

            public int EndColumn { get; set; }

            public string ErrorComment
            {
                get => _errorComment;
                set
                {
                    _errorComment += value;
                    
                }
            }

            public byte[] ExcelFile { get; set; }

            private List<string> _headerColumns;

            public List<string> HeaderColumns
            {
                get => _headerColumns;
                set
                {
                    _headerColumns = value;
                    _headerColumns = _headerColumns.ConvertAll(x => x.ToLowerInvariant());
                }
            }

            private string _mismatchedRows;

            public string MismatchedColumns { get => _mismatchedRows; set { _mismatchedRows = value; } }

            #region PR#7 Data Validation of Excel Rows and Columns

            private DataValidationModel _validationType;
            public DataValidationModel ValidationModel
            {
                get => _validationType; 
                set
                {
                    _validationType = value;
                    //var dataColumnKey = _dataValidation.Keys.Skip((Column - 1)).Take(1).First();
                    //_validationType = _dataValidation[dataColumnKey];
                }
            }

            private string _typeValidate;
            public string ValidationType { get => _typeValidate; set => _typeValidate = value; }

            private Dictionary<string, DataValidationModel> _dataValidation;
            public Dictionary<string, DataValidationModel> DataValidation
            {
                get => _dataValidation;
                set
                {
                    _dataValidation = value;

                    //var type = Type.GetType($"System.{Helpers.Helpers.UpperCaseFirst(_dataValidation[dataColumnKey].DataType)}");

                }
            }
            public string ColumnValidation { get; set; }
            #endregion
        }
    }
}
