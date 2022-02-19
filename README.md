# ExcelValidation

[![Build Status](https://dev.azure.com/devopspractices1/Space%20Game%20-%20web%20-%20Tests/_apis/build/status/sheddy123.ExcelValidation?branchName=main)](https://dev.azure.com/devopspractices1/Space%20Game%20-%20web%20-%20Tests/_build/latest?definitionId=15&branchName=main)
![Nuget](https://img.shields.io/nuget/v/excel-validator?style=plastic)

[![Board Status](https://dev.azure.com/devopspractices1/edf82c24-f3b4-4b8d-b4d8-c9d8226cdd76/5092ddc7-b118-4e90-ab83-6a0055a75ea7/_apis/work/boardbadge/f8bfb2aa-fa17-49b0-b903-6521b0552c3d?columnOptions=1)](https://dev.azure.com/devopspractices1/edf82c24-f3b4-4b8d-b4d8-c9d8226cdd76/_boards/board/t/5092ddc7-b118-4e90-ab83-6a0055a75ea7/Microsoft.RequirementCategory/)


Data Validation of ExcelSheet, with create, read, modify, delete Data validations . Types of validations supported: Integer (whole in Excel), Decimal, List, Date, Time, Any and Custom.  Strongly typed interface for each validation type

# Continuous integration
| Branch | Build Status |
| :---   | :---:        |          
| `master` | [![Build Status](https://dev.azure.com/devopspractices1/Space%20Game%20-%20web%20-%20Tests/_apis/build/status/sheddy123.ExcelValidation?branchName=main)](https://dev.azure.com/devopspractices1/Space%20Game%20-%20web%20-%20Tests/_build/latest?definitionId=15&branchName=main) |


# Supported file formats and versions
| File Type 	| Container Format |	File Format |	Excel Version(s) |
| :---        | :---             |  :---        | :---             |
| .xlsx       |	ZIP, CFB+ZIP     |	OpenXml     |	2007 and newer   |

# Installation
It is recommended to use NuGet through the VS Package Manager Console Install-Package <package> or using the VS "Manage NuGet Packages..." extension.

Install the ExcelValidator base package: 
| Console Terminal | Command |  
| :---             | :---:   |
| Package Manager  | `Install-Package excel-validator -Version 1.0.0` |
| .NET Cli         | `dotnet add package excel-validator --version 1.0.0` |
| PackageReference | `<PackageReference Include="excel-validator" Version="1.0.0" />`|
| Paket CLI        | `paket add excel-validator --version 1.0.0` |
| Script & Interactive | `#r "nuget: excel-validator, 1.0.0"` |
| Cake                  | // Install excel-validator as a Cake Addin `#addin nuget:?package=excel-validator&version=1.0.0` // Install excel-validator as a Cake Tool `#tool nuget:?package=excel-validator&version=1.0.0` |
  
  
  # How to use
  
  ```c#
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
```
# `MODELS`
- `ExcelValidationModel:` This object contains properties used for validating the excel file. Snippet of `ExcelValidationModel`

```c#        
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
            
            private DataValidationModel _validationType;
            public DataValidationModel ValidationModel
            {
                get => _validationType; 
                set
                {
                    _validationType = value;
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
                }
            }
            public string ColumnValidation { get; set; }
            #endregion
        }
    
```
`DataValidationModel`
```c#
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

```
- `ValidateExcel:` This method takes in an `ExcelValidationModel` object as parameter.  


