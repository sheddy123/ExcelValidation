/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/

using System;
using System.ComponentModel;
using static ExcelValidator.Model.ValidateExcelSheet;

namespace ExcelValidator.Model
{
    public class DataValidationModel
    {
        #region PR#7 Data Validation of Excel Rows and Columns
        private string _dataType;
        public string DataType
        {
            get => _dataType;
            set
            {
                _dataType = value;
                _dataType = Helpers.Helpers.UpperCaseFirst(_dataType);
            }
        }
        private bool _typeIsValid;
        public bool TypeIsValid
        {
            get => _typeIsValid;
            set => _typeIsValid = value;

        }
        public string MaxLength { get; set; }
        public string MinLength { get; set; }
        public string InputType { get; set; }

        private string _currentValue;
        private bool _isValid;
        public string CurrentValue
        {
            get => _currentValue;
            set
            {
                var type = Type.GetType($"System.{_dataType}");
                _currentValue = value;
                _typeIsValid = ((type == null) ? false : TypeDescriptor.GetConverter(type).IsValid(_currentValue));
                _isValid = (type == null) ? false : true;
            }
        }
        public bool IsValid { get => _isValid; }
      
        #endregion
    }
}
