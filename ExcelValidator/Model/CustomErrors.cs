/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/

namespace ExcelValidator.Model
{
    public partial class ValidateExcelSheet
    {
        public class CustomErrors
        {
            public const string InvalidColumns = "The header columns are invalid";
            public const string ValidColumns = "The columns are valid";
            public const string InvalidRows = "The header rows are invalid";
            public const string ValidRows = "The header rows are valid";
            
        }
        
        public class CustomNames
        {
            public const string NormalVal = "Normal";
            public const string Data_Validation = "Data Validation";
          
        }

    }
}
