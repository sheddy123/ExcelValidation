using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelValidator
{
    /// <summary>
    /// Describes one expected column: its header name, the type of value it holds, and any
    /// additional constraints on that value.
    /// </summary>
    public sealed class ColumnRule
    {
        private Regex? _compiledPattern;
        private string? _pattern;

        /// <summary>
        /// Creates a rule for a column with the given header name.
        /// </summary>
        /// <param name="name">The column's header text. Matched against the worksheet case-insensitively
        /// unless <see cref="ExcelValidationOptions.CaseSensitiveHeaders"/> is set.</param>
        /// <exception cref="ArgumentException"><paramref name="name"/> is null, empty, or whitespace.</exception>
        public ColumnRule(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("A column name is required.", nameof(name));
            }

            Name = name.Trim();
        }

        /// <summary>The column's header text.</summary>
        public string Name { get; }

        /// <summary>The type of value the column holds. Defaults to <see cref="ExcelCellType.Text"/>.</summary>
        public ExcelCellType Type { get; set; } = ExcelCellType.Text;

        /// <summary>
        /// Whether every cell in the column must be non-empty. Defaults to <see langword="true"/>.
        /// When <see langword="false"/>, empty cells are skipped and the remaining constraints
        /// are only applied to cells that do have a value.
        /// </summary>
        public bool Required { get; set; } = true;

        /// <summary>Minimum length of the cell's text, or <see langword="null"/> for no minimum.</summary>
        public int? MinLength { get; set; }

        /// <summary>Maximum length of the cell's text, or <see langword="null"/> for no maximum.</summary>
        public int? MaxLength { get; set; }

        /// <summary>
        /// The set of values the column accepts, or <see langword="null"/> to accept any value of the
        /// declared <see cref="Type"/>. Compared case-insensitively.
        /// </summary>
        public IReadOnlyCollection<string>? AllowedValues { get; set; }

        /// <summary>
        /// A regular expression the cell's text must match, or <see langword="null"/> for no pattern check.
        /// </summary>
        /// <exception cref="ArgumentException">The value is not a valid regular expression.</exception>
        public string? Pattern
        {
            get => _pattern;
            set
            {
                // Compile eagerly so an invalid pattern surfaces where it was set, not on the first
                // cell that happens to reach the check.
                _compiledPattern = value is null
                    ? null
                    : new Regex(value, RegexOptions.CultureInvariant);
                _pattern = value;
            }
        }

        internal Regex? CompiledPattern => _compiledPattern;
    }
}
