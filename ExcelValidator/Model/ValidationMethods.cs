using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static ExcelValidator.Model.ValidateExcelSheet;

namespace ExcelValidator.Model
{
    public static class ValidationMethods
    {
        /// <summary>
        /// Displays an error for an invalid cell 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        private static bool SetError(ExcelRange cell, ExcelValidationModel model)
        {
            var fill = cell[model.Row, model.Column].Style.Fill;
            fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
            cell.AddComment(model.ErrorComment, " !!");

            return false;
        }

        /// <summary>
        /// Validate each cell in the excel file
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        private static bool ValidateText(ExcelRange cell, ExcelValidationModel model)
        {
            bool result = true;
            model.ErrorComment = string.Format("{0} is empty", model.ColumnName);

            if (cell[model.Row, model.Column].Value != null)
            {
                //check if cell value has a value
                if (string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    result = SetError(cell, model);
            }
            else
                result = SetError(cell, model);

            return result;
        }

        /// <summary>
        /// Convert byte array to a specific excel object
        /// </summary>
        /// <param name="arrBytes"></param>
        /// <returns></returns>
        private static ExcelPackage ByteArrayToObject(byte[] arrBytes)
        {
            using (MemoryStream memStream = new MemoryStream(arrBytes))
                return new ExcelPackage(memStream);

        }
        //{
        //    ExcelPackage package = new ExcelPackage(memStream);
        //return package;
        //}

        private static HashSet<string> GetHeaderColumns(this ExcelWorksheet sheet)
        {
            HashSet<string> columnNames = new HashSet<string>();
            foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
                columnNames.Add(firstRowCell.Text);
            return columnNames;
        }
       
        /// <summary>
        /// Validates top excel headers
        /// </summary>
        /// <param name="excelFileByte"></param>
        /// <param name="headers"></param>
        /// <returns></returns>
        public static bool ValidateExcelHeader(ExcelValidationModel excelSheet)
        {
            HashSet<object> headerEntries = new HashSet<object>(excelSheet.HeaderColumns);
            var excelFile2 = ByteArrayToObject(excelSheet.ExcelFile);
            var listColumnHeaders = excelFile2.Workbook.Worksheets[0].GetHeaderColumns();
            
            return headerEntries.SetEquals(listColumnHeaders);
        }
    }
}
