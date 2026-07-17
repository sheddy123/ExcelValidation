using System;

namespace ExcelValidator
{
    /// <summary>
    /// Controls how a workbook is read and how much of it is checked. The defaults suit the
    /// common case: first worksheet, headers in row 1, case-insensitive header matching.
    /// </summary>
    public sealed class ExcelValidationOptions
    {
        private int _worksheetPosition = 1;
        private int _headerRow = 1;
        private int _maxErrors = 1000;

        /// <summary>
        /// The worksheet to validate, by name. When null, <see cref="WorksheetPosition"/> is used instead.
        /// </summary>
        public string? WorksheetName { get; set; }

        /// <summary>
        /// The 1-based position of the worksheet to validate. Ignored when
        /// <see cref="WorksheetName"/> is set. Defaults to 1.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
        public int WorksheetPosition
        {
            get => _worksheetPosition;
            set
            {
                if (value < 1)
                {
                    throw new ArgumentOutOfRangeException(nameof(value), value, "Worksheet positions are 1-based.");
                }

                _worksheetPosition = value;
            }
        }

        /// <summary>
        /// The 1-based row holding the column headers. Data is read from the row after it. Defaults to 1.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
        public int HeaderRow
        {
            get => _headerRow;
            set
            {
                if (value < 1)
                {
                    throw new ArgumentOutOfRangeException(nameof(value), value, "Row numbers are 1-based.");
                }

                _headerRow = value;
            }
        }

        /// <summary>
        /// Whether header names must match the schema's casing exactly. Defaults to <see langword="false"/>.
        /// </summary>
        public bool CaseSensitiveHeaders { get; set; }

        /// <summary>
        /// Whether leading and trailing whitespace is trimmed from cell text before it is checked.
        /// Defaults to <see langword="true"/>, so a cell holding only spaces counts as empty.
        /// </summary>
        public bool TrimValues { get; set; } = true;

        /// <summary>
        /// Stop after this many errors and set <see cref="ExcelValidationResult.Truncated"/>. Defaults to
        /// 1000, which keeps a badly broken 100k-row import from allocating an error per cell. Set to
        /// <see cref="int.MaxValue"/> to collect everything.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
        public int MaxErrors
        {
            get => _maxErrors;
            set
            {
                if (value < 1)
                {
                    throw new ArgumentOutOfRangeException(nameof(value), value, "At least one error must be collectable.");
                }

                _maxErrors = value;
            }
        }
    }
}
