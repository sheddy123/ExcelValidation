/*==============================================*\
|    Created By                                   |
|     Odom Ifeanyi Shadrach v1.0                  |
|            11/11/2021                           |
|                                                 |
|                                                 |
/================================================*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelValidator.Model.ValidateExcelSheet;

namespace ExcelValidator.Interfaces
{
    public interface IExcelValidator
    {
        ExcelValidationModel ValidateExcel(ExcelValidationModel excelFile);
    }
}
