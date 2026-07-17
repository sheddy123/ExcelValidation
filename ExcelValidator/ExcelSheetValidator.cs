using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;

namespace ExcelValidator
{
    /// <summary>
    /// The default <see cref="IExcelSheetValidator"/>. Stateless and thread-safe; create one and
    /// reuse it, or register it as a singleton.
    /// </summary>
    /// <example>
    /// <code>
    /// var schema = new ExcelSchema()
    ///     .Column("ID", ExcelCellType.Integer)
    ///     .Column("Email", ExcelCellType.Text);
    ///
    /// var result = new ExcelSheetValidator().Validate(File.ReadAllBytes("people.xlsx"), schema);
    /// foreach (var error in result.Errors)
    /// {
    ///     Console.WriteLine(error);
    /// }
    /// </code>
    /// </example>
    public sealed class ExcelSheetValidator : IExcelSheetValidator
    {
        /// <inheritdoc/>
        public ExcelValidationResult Validate(byte[] workbook, ExcelSchema schema, ExcelValidationOptions? options = null)
        {
            if (workbook is null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            using (var stream = new MemoryStream(workbook, writable: false))
            {
                return Validate(stream, schema, options);
            }
        }

        /// <inheritdoc/>
        public ExcelValidationResult Validate(Stream workbook, ExcelSchema schema, ExcelValidationOptions? options = null)
        {
            if (workbook is null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (schema is null)
            {
                throw new ArgumentNullException(nameof(schema));
            }

            options ??= new ExcelValidationOptions();

            using (var book = Open(workbook))
            {
                var sheet = ResolveWorksheet(book, options);
                return ValidateWorksheet(sheet, schema, options);
            }
        }

        /// <inheritdoc/>
        public byte[] Annotate(byte[] workbook, ExcelValidationResult result, ExcelValidationOptions? options = null)
        {
            if (workbook is null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            options ??= new ExcelValidationOptions();

            using (var source = new MemoryStream(workbook, writable: false))
            using (var book = Open(source))
            {
                var sheet = ResolveWorksheet(book, options);

                foreach (var error in result.Errors)
                {
                    var cell = LocateCell(sheet, error, options);
                    if (cell is null)
                    {
                        continue;
                    }

                    cell.Style.Fill.SetBackgroundColor(XLColor.Red);
                    cell.GetComment().AddText(error.Message);
                }

                // SaveAs needs a writable, seekable stream, and the caller's array is neither.
                using (var destination = new MemoryStream())
                {
                    book.SaveAs(destination);
                    return destination.ToArray();
                }
            }
        }

        private static IXLCell? LocateCell(IXLWorksheet sheet, ValidationError error, ExcelValidationOptions options)
        {
            if (error.Row.HasValue && error.Column.HasValue)
            {
                return sheet.Cell(error.Row.Value, error.Column.Value);
            }

            // A column-level error has no cell of its own. Fall back to the header cell so the reader
            // is pointed at the right column; a column that is missing entirely has no cell at all.
            if (error.ColumnName is null)
            {
                return null;
            }

            var comparison = options.CaseSensitiveHeaders
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            foreach (var cell in HeaderCells(sheet, options))
            {
                if (string.Equals(cell.GetFormattedString().Trim(), error.ColumnName, comparison))
                {
                    return cell;
                }
            }

            return null;
        }

        /// <summary>
        /// The cell's reference, such as "B4". ClosedXML declares this nullable, but an address
        /// obtained from a real cell always renders.
        /// </summary>
        private static string Address(IXLCell cell) => cell.Address.ToString() ?? string.Empty;

        private static IEnumerable<IXLCell> HeaderCells(IXLWorksheet sheet, ExcelValidationOptions options)
        {
            var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            for (var column = 1; column <= lastColumn; column++)
            {
                yield return sheet.Cell(options.HeaderRow, column);
            }
        }

        private static XLWorkbook Open(Stream stream)
        {
            try
            {
                return new XLWorkbook(stream);
            }
            catch (Exception ex) when (!(ex is ExcelValidationException))
            {
                throw new ExcelValidationException(
                    "The stream could not be read as an .xlsx workbook. It may be corrupt, empty, password-protected, or in the older .xls format.",
                    ex);
            }
        }

        private static IXLWorksheet ResolveWorksheet(XLWorkbook book, ExcelValidationOptions options)
        {
            if (options.WorksheetName is not null)
            {
                if (!book.Worksheets.TryGetWorksheet(options.WorksheetName, out var named))
                {
                    throw new ExcelValidationException(
                        $"The workbook has no worksheet named '{options.WorksheetName}'. It contains: {string.Join(", ", WorksheetNames(book))}.");
                }

                return named;
            }

            if (options.WorksheetPosition > book.Worksheets.Count)
            {
                throw new ExcelValidationException(
                    $"The workbook has no worksheet at position {options.WorksheetPosition.ToString(CultureInfo.InvariantCulture)}; it has {book.Worksheets.Count.ToString(CultureInfo.InvariantCulture)}.");
            }

            return book.Worksheet(options.WorksheetPosition);
        }

        private static IEnumerable<string> WorksheetNames(XLWorkbook book)
        {
            foreach (var sheet in book.Worksheets)
            {
                yield return sheet.Name;
            }
        }

        private static ExcelValidationResult ValidateWorksheet(
            IXLWorksheet sheet,
            ExcelSchema schema,
            ExcelValidationOptions options)
        {
            var errors = new ErrorCollector(options.MaxErrors);
            var headers = ReadHeaders(sheet, options, errors);
            var matched = MatchColumns(schema, headers, options, errors, out var missing, out var unexpected);

            var rowsValidated = errors.IsFull
                ? 0
                : ValidateRows(sheet, matched, options, errors);

            return new ExcelValidationResult(
                errors.ToList(),
                missing,
                unexpected,
                rowsValidated,
                errors.IsFull);
        }

        /// <summary>
        /// Reads the header row into a map of name to column number, reporting empty and duplicate
        /// headers as it goes.
        /// </summary>
        private static Dictionary<string, HeaderColumn> ReadHeaders(
            IXLWorksheet sheet,
            ExcelValidationOptions options,
            ErrorCollector errors)
        {
            var comparer = options.CaseSensitiveHeaders
                ? StringComparer.Ordinal
                : StringComparer.OrdinalIgnoreCase;
            var headers = new Dictionary<string, HeaderColumn>(comparer);

            foreach (var cell in HeaderCells(sheet, options))
            {
                var name = cell.GetFormattedString().Trim();

                if (name.Length == 0)
                {
                    // A blank cell inside the header row means a column nothing can be matched against.
                    // Trailing blanks are excluded already, since LastColumnUsed bounds the scan.
                    errors.Add(new ValidationError(
                        ValidationErrorKind.EmptyHeader,
                        "The header cell is empty, so this column cannot be matched to the schema.",
                        row: cell.Address.RowNumber,
                        column: cell.Address.ColumnNumber,
                        address: Address(cell)));
                    continue;
                }

                if (headers.ContainsKey(name))
                {
                    errors.Add(new ValidationError(
                        ValidationErrorKind.DuplicateColumn,
                        $"The column '{name}' appears more than once in the header row.",
                        columnName: name,
                        row: cell.Address.RowNumber,
                        column: cell.Address.ColumnNumber,
                        address: Address(cell),
                        value: name));
                    continue;
                }

                headers.Add(name, new HeaderColumn(cell.Address.ColumnNumber, Address(cell)));
            }

            return headers;
        }

        /// <summary>
        /// Pairs each schema rule with the worksheet column holding it, and reports the columns on
        /// either side that have no partner.
        /// </summary>
        private static List<MatchedColumn> MatchColumns(
            ExcelSchema schema,
            Dictionary<string, HeaderColumn> headers,
            ExcelValidationOptions options,
            ErrorCollector errors,
            out IReadOnlyList<string> missing,
            out IReadOnlyList<string> unexpected)
        {
            var matched = new List<MatchedColumn>(schema.Columns.Count);
            var missingColumns = new List<string>();
            var claimed = new HashSet<string>(
                options.CaseSensitiveHeaders ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase);

            foreach (var rule in schema.Columns)
            {
                if (headers.TryGetValue(rule.Name, out var header))
                {
                    matched.Add(new MatchedColumn(rule, header.ColumnNumber));
                    claimed.Add(rule.Name);
                }
                else
                {
                    missingColumns.Add(rule.Name);
                    errors.Add(new ValidationError(
                        ValidationErrorKind.MissingColumn,
                        $"The worksheet has no column named '{rule.Name}'.",
                        columnName: rule.Name));
                }
            }

            var unexpectedColumns = new List<string>();
            foreach (var header in headers)
            {
                if (claimed.Contains(header.Key))
                {
                    continue;
                }

                unexpectedColumns.Add(header.Key);

                if (!schema.AllowUnexpectedColumns)
                {
                    errors.Add(new ValidationError(
                        ValidationErrorKind.UnexpectedColumn,
                        $"The schema does not declare a column named '{header.Key}'.",
                        columnName: header.Key,
                        row: options.HeaderRow,
                        column: header.Value.ColumnNumber,
                        address: header.Value.Address));
                }
            }

            missing = missingColumns;
            unexpected = unexpectedColumns;
            return matched;
        }

        private static int ValidateRows(
            IXLWorksheet sheet,
            List<MatchedColumn> columns,
            ExcelValidationOptions options,
            ErrorCollector errors)
        {
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
            var firstDataRow = options.HeaderRow + 1;
            var rowsValidated = 0;

            for (var row = firstDataRow; row <= lastRow; row++)
            {
                if (errors.IsFull)
                {
                    break;
                }

                if (sheet.Row(row).IsEmpty())
                {
                    // A blank row is padding, not a row of missing values. Excel readily leaves these
                    // behind after a user deletes content, and reporting one error per column for each
                    // would bury the real errors.
                    continue;
                }

                rowsValidated++;

                foreach (var column in columns)
                {
                    if (errors.IsFull)
                    {
                        break;
                    }

                    ValidateCell(sheet.Cell(row, column.ColumnNumber), column.Rule, options, errors);
                }
            }

            return rowsValidated;
        }

        private static void ValidateCell(
            IXLCell cell,
            ColumnRule rule,
            ExcelValidationOptions options,
            ErrorCollector errors)
        {
            var value = cell.Value;
            var text = cell.GetFormattedString();
            if (options.TrimValues)
            {
                text = text.Trim();
            }

            var address = Address(cell);

            if (value.IsBlank || text.Length == 0)
            {
                if (rule.Required)
                {
                    errors.Add(new ValidationError(
                        ValidationErrorKind.RequiredValueMissing,
                        $"'{rule.Name}' is required but the cell is empty.",
                        columnName: rule.Name,
                        row: cell.Address.RowNumber,
                        column: cell.Address.ColumnNumber,
                        address: address,
                        value: string.Empty));
                }

                // Nothing further is meaningful about an empty optional cell.
                return;
            }

            if (!CellTypeChecker.Matches(value, rule.Type))
            {
                errors.Add(new ValidationError(
                    ValidationErrorKind.TypeMismatch,
                    $"'{text}' is not a valid {rule.Type} value for column '{rule.Name}'.",
                    columnName: rule.Name,
                    row: cell.Address.RowNumber,
                    column: cell.Address.ColumnNumber,
                    address: address,
                    value: text));

                // The remaining rules describe a value of the right type, so checking them against a
                // value of the wrong type would only produce noise on top of the real error.
                return;
            }

            if (rule.MinLength.HasValue && text.Length < rule.MinLength.Value)
            {
                errors.Add(new ValidationError(
                    ValidationErrorKind.LengthOutOfRange,
                    $"'{rule.Name}' must be at least {rule.MinLength.Value.ToString(CultureInfo.InvariantCulture)} characters, but '{text}' is {text.Length.ToString(CultureInfo.InvariantCulture)}.",
                    columnName: rule.Name,
                    row: cell.Address.RowNumber,
                    column: cell.Address.ColumnNumber,
                    address: address,
                    value: text));
            }
            else if (rule.MaxLength.HasValue && text.Length > rule.MaxLength.Value)
            {
                errors.Add(new ValidationError(
                    ValidationErrorKind.LengthOutOfRange,
                    $"'{rule.Name}' must be at most {rule.MaxLength.Value.ToString(CultureInfo.InvariantCulture)} characters, but '{text}' is {text.Length.ToString(CultureInfo.InvariantCulture)}.",
                    columnName: rule.Name,
                    row: cell.Address.RowNumber,
                    column: cell.Address.ColumnNumber,
                    address: address,
                    value: text));
            }

            if (rule.AllowedValues is not null && !Contains(rule.AllowedValues, text))
            {
                errors.Add(new ValidationError(
                    ValidationErrorKind.ValueNotAllowed,
                    $"'{text}' is not an accepted value for column '{rule.Name}'. Accepted: {string.Join(", ", rule.AllowedValues)}.",
                    columnName: rule.Name,
                    row: cell.Address.RowNumber,
                    column: cell.Address.ColumnNumber,
                    address: address,
                    value: text));
            }

            if (rule.CompiledPattern is not null && !rule.CompiledPattern.IsMatch(text))
            {
                errors.Add(new ValidationError(
                    ValidationErrorKind.PatternMismatch,
                    $"'{text}' does not match the expected format for column '{rule.Name}'.",
                    columnName: rule.Name,
                    row: cell.Address.RowNumber,
                    column: cell.Address.ColumnNumber,
                    address: address,
                    value: text));
            }
        }

        private static bool Contains(IReadOnlyCollection<string> allowed, string text)
        {
            foreach (var candidate in allowed)
            {
                if (string.Equals(candidate, text, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private readonly struct MatchedColumn
        {
            public MatchedColumn(ColumnRule rule, int columnNumber)
            {
                Rule = rule;
                ColumnNumber = columnNumber;
            }

            public ColumnRule Rule { get; }

            public int ColumnNumber { get; }
        }

        /// <summary>A column found in the header row. The address is kept because ClosedXML will not
        /// let us build one from a row and column outside a worksheet.</summary>
        private readonly struct HeaderColumn
        {
            public HeaderColumn(int columnNumber, string address)
            {
                ColumnNumber = columnNumber;
                Address = address;
            }

            public int ColumnNumber { get; }

            public string Address { get; }
        }

        /// <summary>
        /// Accumulates errors up to a cap, so a wholly malformed file cannot exhaust memory.
        /// </summary>
        private sealed class ErrorCollector
        {
            private readonly List<ValidationError> _errors = new List<ValidationError>();
            private readonly int _max;

            public ErrorCollector(int max) => _max = max;

            public bool IsFull => _errors.Count >= _max;

            public void Add(ValidationError error)
            {
                if (!IsFull)
                {
                    _errors.Add(error);
                }
            }

            public List<ValidationError> ToList() => _errors;
        }
    }
}
