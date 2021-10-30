using OfficeOpenXml;
using System;
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
        private static bool SetError(ExcelRange cell, ExcelValidationModel model, string errorComment)
        {
            var fill = cell[model.Row, model.Column].Style.Fill;
            fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
            cell.AddComment(errorComment, " !!");

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
            var errorComment = string.Format("\n\n\n {0} is empty", cell[1, model.Column].Value.ToString());

            if (cell[model.Row, model.Column].Value != null)
            {
                //check if cell value has a value
                if (string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    result = SetError(cell, model, errorComment);
                
            }
            else
                result = SetError(cell, model, errorComment);
            


            return result;
        }

        /// <summary>
        /// Convert byte array to a specific excel object
        /// </summary>
        /// <param name="arrBytes"></param>
        /// <returns></returns>
        private static ExcelPackage ByteArrayToObject(byte[] arrBytes)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (MemoryStream memStream = new MemoryStream(arrBytes))
                    return new ExcelPackage(memStream);
            }
            catch (Exception ex)
            {

            }
            return null;
        }

        /// <summary>
        /// Get columns in topmost row
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static HashSet<string> GetHeaderColumns(this ExcelWorksheet sheet)
        {
            HashSet<string> columnNames = new HashSet<string>();

            foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
                columnNames.Add(firstRowCell.Text.ToLower());
            return columnNames;
        }

        /// <summary>
        /// Reads from second row
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static ExcelValidationModel GetHeaderRows(this ExcelWorksheet sheet)
        {
            List<HashSet<string>> rowEntriesList = new List<HashSet<string>>();
            HashSet<string> rowEntries = new HashSet<string>();
            ExcelValidationModel model = new ExcelValidationModel();
            model.Row = sheet.Dimension.Start.Row + 1;
            model.EndRow = sheet.Dimension.End.Row;
            model.Column = sheet.Dimension.Start.Column;
            model.EndColumn = sheet.Dimension.End.Column;

            while (model.Row <= model.EndRow)
            {

                if (model.Column > model.EndColumn)
                {
                    model.AddRowEntriesList.Add(rowEntries);
                    rowEntries = new HashSet<string>();
                    model.Row++;
                    model.Column = sheet.Dimension.Start.Column;
                }

                if (model.Row > model.EndRow)
                    break;

                ValidateText(sheet.Cells, model);
                rowEntries.Add(sheet.Cells[model.Row, model.Column].Value.ToString());
                model.Column++;
            }

            return model;
        }


        /// <summary>
        /// Validates top excel columns
        /// </summary>
        /// <param name="excelFileByte"></param>
        /// <param name="headers"></param>
        /// <returns></returns>
        public static bool ValidateExcelColumns(ExcelValidationModel excelSheet)
        {
            HashSet<string> headerEntries = new HashSet<string>(excelSheet.HeaderColumns);
            var excelFile2 = ByteArrayToObject(excelSheet.ExcelFile);
            var listColumnHeaders = excelFile2.Workbook.Worksheets[0].GetHeaderColumns();
            
            headerEntries.SymmetricExceptWith(listColumnHeaders);

            excelSheet.MismatchedColumns = string.Join(",", headerEntries.OrderBy(key => key).ToList());

            return headerEntries.SetEquals(listColumnHeaders);
        }

        /// <summary>
        /// Validates top excel rows
        /// </summary>
        /// <param name="excelSheet"></param>
        /// <returns></returns>
        public static bool ValidateExcelRows(ExcelValidationModel excelSheet)
        {
            var excelFile2 = ByteArrayToObject(excelSheet.ExcelFile);

            if (excelFile2.Workbook.Worksheets[0].Dimension.Rows <= 1)
                return true;

            var listColumnHeaders = excelFile2.Workbook.Worksheets[0].GetHeaderRows();

            return true;
        }
    }
}
