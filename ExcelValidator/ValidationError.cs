using System.Globalization;

namespace ExcelValidator
{
    /// <summary>
    /// A single problem found in a worksheet, located as precisely as the problem allows:
    /// column-level errors carry only a <see cref="ColumnName"/>, cell-level errors also carry
    /// a <see cref="Row"/>, <see cref="Column"/>, and <see cref="Address"/>.
    /// </summary>
    public sealed class ValidationError
    {
        internal ValidationError(
            ValidationErrorKind kind,
            string message,
            string? columnName = null,
            int? row = null,
            int? column = null,
            string? address = null,
            string? value = null)
        {
            Kind = kind;
            Message = message;
            ColumnName = columnName;
            Row = row;
            Column = column;
            Address = address;
            Value = value;
        }

        /// <summary>Why this error was raised.</summary>
        public ValidationErrorKind Kind { get; }

        /// <summary>
        /// A human-readable description. Intended for logs and end users; the exact wording may
        /// change between releases, so branch on <see cref="Kind"/> instead.
        /// </summary>
        public string Message { get; }

        /// <summary>The header name of the column involved, if known.</summary>
        public string? ColumnName { get; }

        /// <summary>The 1-based row number, for cell-level errors.</summary>
        public int? Row { get; }

        /// <summary>The 1-based column number, for cell-level errors.</summary>
        public int? Column { get; }

        /// <summary>The cell reference such as <c>"B4"</c>, for cell-level errors.</summary>
        public string? Address { get; }

        /// <summary>The cell's value as displayed in Excel, for cell-level errors.</summary>
        public string? Value { get; }

        /// <summary>Returns the message, prefixed with the cell address when there is one.</summary>
        public override string ToString() =>
            Address is null
                ? Message
                : string.Format(CultureInfo.InvariantCulture, "{0}: {1}", Address, Message);
    }
}
