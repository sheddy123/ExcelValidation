namespace ExcelValidator
{
    /// <summary>
    /// Convenience methods on <see cref="ExcelValidationResult"/>.
    /// </summary>
    public static class ExcelValidationResultExtensions
    {
        /// <summary>
        /// Renders this result as JSON. Equivalent to
        /// <see cref="ExcelValidationJson.Serialize(ExcelValidationResult, bool)"/>.
        /// </summary>
        /// <param name="result">The result to serialize.</param>
        /// <param name="indented">Whether to pretty-print. Defaults to <see langword="true"/>.</param>
        /// <returns>The JSON text.</returns>
        public static string ToJson(this ExcelValidationResult result, bool indented = true) =>
            ExcelValidationJson.Serialize(result, indented);
    }
}
