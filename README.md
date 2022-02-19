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
  
  ```
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
  


