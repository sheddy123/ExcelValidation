/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/
using OfficeOpenXml;
using System.Collections.Generic;

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
                set {
                    _errorComment = value;
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
        }
    }
}
