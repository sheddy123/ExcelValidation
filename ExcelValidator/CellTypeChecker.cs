using System;
using System.Globalization;
using ClosedXML.Excel;

namespace ExcelValidator
{
    /// <summary>
    /// Decides whether a cell's value satisfies an <see cref="ExcelCellType"/>.
    /// </summary>
    /// <remarks>
    /// A cell can hold a value either natively (Excel stores it as a number, date, or boolean) or as
    /// text that happens to look like one. Both are accepted: a column typed
    /// <see cref="ExcelCellType.Integer"/> matches a numeric cell holding 34 and a text cell holding
    /// "34" alike, since which one a file uses depends on how it was produced rather than on whether
    /// the data is correct. Text is parsed with the invariant culture so that validation does not
    /// depend on the server's locale.
    /// </remarks>
    internal static class CellTypeChecker
    {
        public static bool Matches(XLCellValue value, ExcelCellType expected)
        {
            switch (expected)
            {
                case ExcelCellType.Text:
                    return true;

                case ExcelCellType.Integer:
                    return TryGetDecimal(value, out var i) && IsWhole(i) && i >= int.MinValue && i <= int.MaxValue;

                case ExcelCellType.Long:
                    return TryGetDecimal(value, out var l) && IsWhole(l) && l >= long.MinValue && l <= long.MaxValue;

                case ExcelCellType.Decimal:
                    return TryGetDecimal(value, out _);

                case ExcelCellType.Double:
                    return TryGetDouble(value, out _);

                case ExcelCellType.Boolean:
                    return value.IsBoolean || bool.TryParse(AsText(value), out _);

                case ExcelCellType.DateTime:
                    return value.IsDateTime
                        || DateTime.TryParse(
                            AsText(value),
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out _);

                case ExcelCellType.Guid:
                    return Guid.TryParse(AsText(value), out _);

                default:
                    throw new ArgumentOutOfRangeException(nameof(expected), expected, "Unknown cell type.");
            }
        }

        private static bool IsWhole(decimal value) => decimal.Truncate(value) == value;

        /// <summary>
        /// Reads the value as a decimal. Excel holds all numbers as doubles, so a native numeric cell
        /// can carry a magnitude decimal cannot represent; that overflow means "not a decimal" rather
        /// than an error.
        /// </summary>
        private static bool TryGetDecimal(XLCellValue value, out decimal number)
        {
            if (value.IsNumber)
            {
                try
                {
                    number = (decimal)value.GetNumber();
                    return true;
                }
                catch (OverflowException)
                {
                    number = 0;
                    return false;
                }
            }

            if (value.IsText)
            {
                return decimal.TryParse(
                    value.GetText().Trim(),
                    NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture,
                    out number);
            }

            number = 0;
            return false;
        }

        private static bool TryGetDouble(XLCellValue value, out double number)
        {
            if (value.IsNumber)
            {
                number = value.GetNumber();
                return true;
            }

            if (value.IsText)
            {
                return double.TryParse(
                    value.GetText().Trim(),
                    NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.InvariantCulture,
                    out number);
            }

            number = 0;
            return false;
        }

        private static string AsText(XLCellValue value) =>
            value.IsText ? value.GetText().Trim() : value.ToString(CultureInfo.InvariantCulture).Trim();
    }
}
