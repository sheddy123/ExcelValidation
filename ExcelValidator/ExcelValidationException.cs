using System;

namespace ExcelValidator
{
    /// <summary>
    /// Thrown when a workbook cannot be read at all — it is corrupt, is not an .xlsx file, or does
    /// not contain the requested worksheet. Contrast with a validation <em>failure</em>, which is
    /// reported through <see cref="ExcelValidationResult"/> rather than thrown.
    /// </summary>
    public class ExcelValidationException : Exception
    {
        /// <summary>Creates an exception with a default message.</summary>
        public ExcelValidationException()
            : base("The workbook could not be read.")
        {
        }

        /// <summary>Creates an exception with the given message.</summary>
        public ExcelValidationException(string message)
            : base(message)
        {
        }

        /// <summary>Creates an exception with the given message and underlying cause.</summary>
        public ExcelValidationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
