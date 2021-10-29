using System.Collections.Generic;

namespace ExcelValidator.Model
{
    public partial class ValidateExcelSheet
    {
        public class ExcelValidationModel
        {
            public string ColumnName { get; }

            public int Row { get; }

            public int Column { get; }

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

                    if (!_isValidRow)
                        _errorComment = _errorComment + " and " + CustomErrors.InvalidRows;

                }
            }

            public bool ColumnIsValid
            {
                get => _isValidColumn;
                set
                {
                    _isValidColumn = value;

                    if (!_isValidColumn)
                        _errorComment =  CustomErrors.InvalidColumns;
                }
            }

            public string ErrorComment
            {
                get => _errorComment;
                set => _errorComment = value;
            }

            public byte[] ExcelFile { get; set; }

            public List<object> HeaderColumns { get; set; }
        }
    }
}
