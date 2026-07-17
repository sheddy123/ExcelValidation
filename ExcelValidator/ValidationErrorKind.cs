namespace ExcelValidator
{
    /// <summary>
    /// Identifies why a <see cref="ValidationError"/> was raised. Switch on this rather than
    /// matching on <see cref="ValidationError.Message"/>, whose wording is not part of the API contract.
    /// </summary>
    public enum ValidationErrorKind
    {
        /// <summary>The schema declares a column that the worksheet's header row does not contain.</summary>
        MissingColumn = 0,

        /// <summary>The worksheet has a column the schema does not declare, and the schema does not allow extras.</summary>
        UnexpectedColumn = 1,

        /// <summary>The header row contains the same column name more than once.</summary>
        DuplicateColumn = 2,

        /// <summary>A cell in the header row is empty.</summary>
        EmptyHeader = 3,

        /// <summary>A cell in a required column is empty.</summary>
        RequiredValueMissing = 4,

        /// <summary>A cell's value cannot be read as the column's declared <see cref="ExcelCellType"/>.</summary>
        TypeMismatch = 5,

        /// <summary>A cell's text length falls outside the column's configured minimum or maximum.</summary>
        LengthOutOfRange = 6,

        /// <summary>A cell's value is not among the column's configured allowed values.</summary>
        ValueNotAllowed = 7,

        /// <summary>A cell's value does not match the column's configured regular expression.</summary>
        PatternMismatch = 8,
    }
}
