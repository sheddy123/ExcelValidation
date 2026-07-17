namespace ExcelValidator
{ 
    /// <summary>
    /// The kind of value a column is expected to hold.
    /// </summary>
    public enum ExcelCellType
    {
        /// <summary>Any value. Every non-empty cell satisfies this type.</summary>
        Text = 0,

        /// <summary>A whole number in the range of <see cref="int"/>.</summary>
        Integer = 1,

        /// <summary>A whole number in the range of <see cref="long"/>.</summary>
        Long = 2,

        /// <summary>A number in the range and precision of <see cref="decimal"/>.</summary>
        Decimal = 3,

        /// <summary>A number in the range of <see cref="double"/>.</summary>
        Double = 4,

        /// <summary>A boolean, either a real Excel boolean or the text "true"/"false".</summary>
        Boolean = 5,

        /// <summary>A date, either a real Excel date or a parseable date string.</summary>
        DateTime = 6,

        /// <summary>A <see cref="System.Guid"/> in any of the formats <c>Guid.TryParse</c> accepts.</summary>
        Guid = 7,
    }
}
