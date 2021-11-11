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
using System.IO;
using System.Linq;
using static ExcelValidator.Model.ValidateExcelSheet;

namespace ExcelValidator.Model
{
   public static class ValidationMethods
   {

       #region SetError method to color cell or field and indicate the label in the case of errors
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
       #endregion

       #region Method for validating Text
       /// <summary>
       /// Validate each cell in the excel file
       /// </summary>
       /// <param name="cell"></param>
       /// <param name="model"></param>
       /// <returns></returns>
       private static bool ValidateText(ExcelRange cell, ExcelValidationModel model)
       {
           bool result = true;
           var errorComment = string.Format("\n\n\n {0} is empty", cell[1, model.Column].Value?.ToString());

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
       #endregion

       #region Convert File to Byte
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
               throw;
           }
       }
       #endregion

       #region Get Header Columns on excel file
       /// <summary>
       /// Get columns in topmost row
       /// </summary>
       /// <param name="sheet"></param>
       /// <returns></returns>
       private static HashSet<string> GetHeaderColumns(this ExcelWorksheet sheet, ExcelValidationModel excelSheet)
       {
           HashSet<string> columnNames = new HashSet<string>();
           excelSheet.Column = 1;
           excelSheet.Row = 1;
           excelSheet.ColumnIsValid = true;
           foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
           {
               var result = ValidateText(sheet.Cells, excelSheet);
               if (!result)
               {
                   excelSheet.ColumnIsValid = false;
                   excelSheet.ErrorComment = $"{CustomErrors.InvalidColumns} at {firstRowCell.Address}";
               }
               else
                   columnNames.Add(firstRowCell.Text.ToLower());

               excelSheet.Column++;
           }
           excelSheet.UpdatedSheet = sheet;

           return columnNames;
       }
       #endregion

       #region Get Header Rows from Excel Sheet
       /// <summary>
       /// Reads from second row
       /// </summary>
       /// <param name="sheet"></param>
       /// <returns></returns>
       private static ExcelValidationModel GetHeaderRows(this ExcelWorksheet sheet, ExcelValidationModel excelSheet)
       {
           List<HashSet<string>> rowEntriesList = new List<HashSet<string>>();
           HashSet<string> rowEntries = new HashSet<string>();
           excelSheet.Row = sheet.Dimension.Start.Row + 1;
           excelSheet.EndRow = sheet.Dimension.End.Row;
           excelSheet.Column = sheet.Dimension.Start.Column;
           excelSheet.EndColumn = sheet.Dimension.End.Column;
           excelSheet.RowIsValid = true;
           while (excelSheet.Row <= excelSheet.EndRow)
           {

               if (excelSheet.Column > excelSheet.EndColumn)
               {
                   excelSheet.AddRowEntriesList.Add(rowEntries);
                   rowEntries = new HashSet<string>();
                   excelSheet.Row++;
                   excelSheet.Column = sheet.Dimension.Start.Column;
               }

               if (excelSheet.Row > excelSheet.EndRow)
                   break;

               var result = ValidateText(sheet.Cells, excelSheet);

               if (!result)
               {
                   excelSheet.RowIsValid = false;
                   excelSheet.ErrorComment = $"{CustomErrors.InvalidRows} at row {excelSheet.Row} column {excelSheet.Column} or Address: {sheet.Cells.Address}\n\n";
               }
               else
                   rowEntries.Add(sheet.Cells[excelSheet.Row, excelSheet.Column].Value.ToString());

               excelSheet.Column++;
           }
           excelSheet.UpdatedSheet = sheet;

           return excelSheet;
       }
       #endregion

       #region Validate Excel Columns
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
           var listColumnHeaders = excelFile2.Workbook.Worksheets[0].GetHeaderColumns(excelSheet);

           headerEntries.SymmetricExceptWith(listColumnHeaders);

           excelSheet.MismatchedColumns = string.Join(",", headerEntries.OrderBy(key => key).ToList());
           excelSheet.AddRowEntriesList.Add(listColumnHeaders);

           return excelSheet.ColumnIsValid;
       }
       #endregion

       #region Validate Excel Rows
       /// <summary>
       /// Validates top excel rows
       /// </summary>
       /// <param name="excelSheet"></param>
       /// <returns></returns>
       public static bool ValidateExcelRows(ExcelValidationModel excelSheet)
       {
           if (excelSheet.UpdatedSheet.Workbook.Worksheets[0].Dimension.Rows <= 1)
               return true;

           excelSheet.UpdatedSheet.Workbook.Worksheets[0].GetHeaderRows(excelSheet);

           return excelSheet.RowIsValid;
       }
       #endregion
   }
}
