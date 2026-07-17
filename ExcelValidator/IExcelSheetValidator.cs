using System.IO;

namespace ExcelValidator
{
    /// <summary>
    /// Validates .xlsx worksheets against an <see cref="ExcelSchema"/>.
    /// </summary>
    /// <remarks>
    /// Implementations are stateless and safe to call from multiple threads, and safe to register
    /// as a singleton in a DI container.
    /// </remarks>
    public interface IExcelSheetValidator
    {
        /// <summary>
        /// Validates a workbook read from a stream.
        /// </summary>
        /// <param name="workbook">A readable stream positioned at the start of an .xlsx file. Not disposed by this method.</param>
        /// <param name="schema">The columns to expect.</param>
        /// <param name="options">How to read the workbook, or null for the defaults.</param>
        /// <returns>The errors found, if any.</returns>
        /// <exception cref="System.ArgumentNullException"><paramref name="workbook"/> or <paramref name="schema"/> is null.</exception>
        /// <exception cref="ExcelValidationException">The workbook cannot be opened, or the requested worksheet does not exist.</exception>
        ExcelValidationResult Validate(Stream workbook, ExcelSchema schema, ExcelValidationOptions? options = null);

        /// <summary>
        /// Validates a workbook held in memory.
        /// </summary>
        /// <param name="workbook">The bytes of an .xlsx file.</param>
        /// <param name="schema">The columns to expect.</param>
        /// <param name="options">How to read the workbook, or null for the defaults.</param>
        /// <returns>The errors found, if any.</returns>
        /// <exception cref="System.ArgumentNullException"><paramref name="workbook"/> or <paramref name="schema"/> is null.</exception>
        /// <exception cref="ExcelValidationException">The workbook cannot be opened, or the requested worksheet does not exist.</exception>
        ExcelValidationResult Validate(byte[] workbook, ExcelSchema schema, ExcelValidationOptions? options = null);

        /// <summary>
        /// Returns a copy of the workbook with every cell named in <paramref name="result"/> filled red
        /// and given a comment describing the problem. Column-level errors, which have no cell address,
        /// are annotated on the header cell where one exists and otherwise skipped.
        /// </summary>
        /// <param name="workbook">The bytes of the same .xlsx file that produced <paramref name="result"/>.</param>
        /// <param name="result">The result to draw onto the workbook.</param>
        /// <param name="options">The same options used to validate, or null for the defaults.</param>
        /// <returns>The bytes of the annotated workbook. The input is not modified.</returns>
        /// <exception cref="System.ArgumentNullException"><paramref name="workbook"/> or <paramref name="result"/> is null.</exception>
        /// <exception cref="ExcelValidationException">The workbook cannot be opened, or the requested worksheet does not exist.</exception>
        byte[] Annotate(byte[] workbook, ExcelValidationResult result, ExcelValidationOptions? options = null);
    }
}
