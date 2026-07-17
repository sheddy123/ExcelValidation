using System;
using System.IO;
using ClosedXML.Excel;

namespace ExcelValidationTests
{
    /// <summary>
    /// Builds .xlsx files in memory. Tests construct exactly the workbook they need rather than
    /// depending on a checked-in fixture, so what a test asserts is visible in the test itself.
    /// </summary>
    internal static class TestWorkbook
    {
        public const string DefaultSheetName = "Sheet1";

        /// <summary>
        /// Builds a workbook from a header row followed by data rows. A null cell is left blank.
        /// </summary>
        public static byte[] WithRows(string[] headers, params object?[][] rows) =>
            Build(sheet =>
            {
                for (var c = 0; c < headers.Length; c++)
                {
                    sheet.Cell(1, c + 1).Value = headers[c];
                }

                for (var r = 0; r < rows.Length; r++)
                {
                    for (var c = 0; c < rows[r].Length; c++)
                    {
                        sheet.Cell(r + 2, c + 1).Value = ToCellValue(rows[r][c]);
                    }
                }
            });

        /// <summary>Builds a workbook by writing directly to the worksheet.</summary>
        public static byte[] Build(Action<IXLWorksheet> configure, string sheetName = DefaultSheetName)
        {
            using var book = new XLWorkbook();
            var sheet = book.AddWorksheet(sheetName);
            configure(sheet);

            using var stream = new MemoryStream();
            book.SaveAs(stream);
            return stream.ToArray();
        }

        /// <summary>Opens a workbook's first worksheet for inspection.</summary>
        public static void Read(byte[] workbook, Action<IXLWorksheet> assert)
        {
            using var stream = new MemoryStream(workbook);
            using var book = new XLWorkbook(stream);
            assert(book.Worksheet(1));
        }

        private static XLCellValue ToCellValue(object? value) => value switch
        {
            null => Blank.Value,
            string s => s,
            int i => i,
            long l => l,
            double d => d,
            decimal m => m,
            bool b => b,
            DateTime dt => dt,
            _ => value.ToString() ?? string.Empty,
        };
    }
}
