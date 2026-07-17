using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelValidator
{
    /// <summary>
    /// The set of columns a worksheet is expected to contain. Build one with the fluent
    /// <see cref="Column"/> method, or with <see cref="FromHeaders(IEnumerable{string})"/> when you
    /// only care that the headers line up.
    /// </summary>
    /// <example>
    /// <code>
    /// var schema = new ExcelSchema()
    ///     .Column("ID", ExcelCellType.Integer)
    ///     .Column("Username", ExcelCellType.Text, maxLength: 50)
    ///     .Column("Email", ExcelCellType.Text, pattern: @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
    /// </code>
    /// </example>
    public sealed class ExcelSchema : IEnumerable<ColumnRule>
    {
        private readonly List<ColumnRule> _columns = new List<ColumnRule>();

        /// <summary>
        /// Whether the worksheet may contain columns the schema does not declare.
        /// Defaults to <see langword="false"/>, which reports them as
        /// <see cref="ValidationErrorKind.UnexpectedColumn"/>.
        /// </summary>
        public bool AllowUnexpectedColumns { get; set; }

        /// <summary>The declared columns, in the order they were added.</summary>
        public IReadOnlyList<ColumnRule> Columns => _columns;

        /// <summary>
        /// Builds a schema that only checks header names, with every column typed as
        /// <see cref="ExcelCellType.Text"/> and required.
        /// </summary>
        public static ExcelSchema FromHeaders(IEnumerable<string> headers)
        {
            if (headers is null)
            {
                throw new ArgumentNullException(nameof(headers));
            }

            var schema = new ExcelSchema();
            foreach (var header in headers)
            {
                schema.Add(new ColumnRule(header));
            }

            return schema;
        }

        /// <inheritdoc cref="FromHeaders(IEnumerable{string})"/>
        public static ExcelSchema FromHeaders(params string[] headers) => FromHeaders((IEnumerable<string>)headers);

        /// <summary>
        /// Declares a column.
        /// </summary>
        /// <param name="name">The column's header text.</param>
        /// <param name="type">The type of value the column holds.</param>
        /// <param name="required">Whether every cell in the column must be non-empty.</param>
        /// <param name="minLength">Minimum length of the cell's text, if any.</param>
        /// <param name="maxLength">Maximum length of the cell's text, if any.</param>
        /// <param name="allowedValues">The only values the column accepts, if restricted.</param>
        /// <param name="pattern">A regular expression the cell's text must match, if any.</param>
        /// <returns>This schema, so calls can be chained.</returns>
        /// <exception cref="ArgumentException">The schema already declares a column with this name.</exception>
        public ExcelSchema Column(
            string name,
            ExcelCellType type = ExcelCellType.Text,
            bool required = true,
            int? minLength = null,
            int? maxLength = null,
            IReadOnlyCollection<string>? allowedValues = null,
            string? pattern = null)
        {
            return Add(new ColumnRule(name)
            {
                Type = type,
                Required = required,
                MinLength = minLength,
                MaxLength = maxLength,
                AllowedValues = allowedValues,
                Pattern = pattern,
            });
        }

        /// <summary>
        /// Adds a pre-built <see cref="ColumnRule"/>.
        /// </summary>
        /// <returns>This schema, so calls can be chained.</returns>
        /// <exception cref="ArgumentException">The schema already declares a column with this name.</exception>
        public ExcelSchema Add(ColumnRule rule)
        {
            if (rule is null)
            {
                throw new ArgumentNullException(nameof(rule));
            }

            if (_columns.Any(c => string.Equals(c.Name, rule.Name, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException(
                    $"The schema already declares a column named '{rule.Name}'.", nameof(rule));
            }

            _columns.Add(rule);
            return this;
        }

        /// <inheritdoc/>
        public IEnumerator<ColumnRule> GetEnumerator() => _columns.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
