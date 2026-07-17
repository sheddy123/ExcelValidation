using System.Collections.Generic;
using System.Linq;

namespace ExcelValidator
{
    /// <summary>
    /// The outcome of validating a worksheet. Immutable; a new one is returned by each
    /// <see cref="IExcelSheetValidator.Validate(byte[], ExcelSchema, ExcelValidationOptions)"/> call.
    /// </summary>
    public sealed class ExcelValidationResult
    {
        internal ExcelValidationResult(
            IReadOnlyList<ValidationError> errors,
            IReadOnlyList<string> missingColumns,
            IReadOnlyList<string> unexpectedColumns,
            int rowsValidated,
            bool truncated)
        {
            Errors = errors;
            MissingColumns = missingColumns;
            UnexpectedColumns = unexpectedColumns;
            RowsValidated = rowsValidated;
            Truncated = truncated;
        }

        /// <summary>Whether the worksheet satisfied the schema completely.</summary>
        public bool IsValid => Errors.Count == 0;

        /// <summary>Every problem found, in the order encountered: header errors first, then row by row.</summary>
        public IReadOnlyList<ValidationError> Errors { get; }

        /// <summary>Schema columns whose header was not found in the worksheet.</summary>
        public IReadOnlyList<string> MissingColumns { get; }

        /// <summary>Worksheet columns the schema did not declare.</summary>
        public IReadOnlyList<string> UnexpectedColumns { get; }

        /// <summary>The number of data rows examined, not counting the header row.</summary>
        public int RowsValidated { get; }

        /// <summary>
        /// Whether validation stopped early because <see cref="ExcelValidationOptions.MaxErrors"/>
        /// was reached, meaning <see cref="Errors"/> is not the complete set.
        /// </summary>
        public bool Truncated { get; }

        /// <summary>Whether the header row satisfied the schema.</summary>
        public bool ColumnsAreValid => !Errors.Any(IsHeaderError);

        /// <summary>Whether every data cell satisfied its column's rule.</summary>
        public bool RowsAreValid => !Errors.Any(e => !IsHeaderError(e));

        private static bool IsHeaderError(ValidationError error) =>
            error.Kind == ValidationErrorKind.MissingColumn
            || error.Kind == ValidationErrorKind.UnexpectedColumn
            || error.Kind == ValidationErrorKind.DuplicateColumn
            || error.Kind == ValidationErrorKind.EmptyHeader;
    }
}
