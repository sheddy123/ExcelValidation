using System;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelValidator
{
    /// <summary>
    /// Renders an <see cref="ExcelValidationResult"/> as JSON, for returning validation outcomes
    /// over an HTTP API, writing them to a log, or handing them to a front end.
    /// </summary>
    /// <remarks>
    /// The JSON uses camelCase keys and writes enum values as their names (for example
    /// <c>"typeMismatch"</c>) rather than numbers, so the output stays readable and stable even if
    /// the underlying enum values are renumbered. Keys whose value is null are omitted.
    /// </remarks>
    public static class ExcelValidationJson
    {
        private static readonly JsonSerializerOptions Indented = CreateOptions(writeIndented: true);
        private static readonly JsonSerializerOptions Compact = CreateOptions(writeIndented: false);

        /// <summary>
        /// Serializes a validation result to JSON.
        /// </summary>
        /// <param name="result">The result to serialize.</param>
        /// <param name="indented">Whether to pretty-print. Defaults to <see langword="true"/>.</param>
        /// <returns>The JSON text.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="result"/> is null.</exception>
        public static string Serialize(ExcelValidationResult result, bool indented = true)
        {
            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            var dto = new ResultDto
            {
                IsValid = result.IsValid,
                ColumnsAreValid = result.ColumnsAreValid,
                RowsAreValid = result.RowsAreValid,
                RowsValidated = result.RowsValidated,
                Truncated = result.Truncated,
                MissingColumns = result.MissingColumns.ToArray(),
                UnexpectedColumns = result.UnexpectedColumns.ToArray(),
                Errors = result.Errors.Select(e => new ErrorDto
                {
                    Kind = e.Kind,
                    Message = e.Message,
                    ColumnName = e.ColumnName,
                    Address = e.Address,
                    Row = e.Row,
                    Column = e.Column,
                    Value = e.Value,
                }).ToArray(),
            };

            return JsonSerializer.Serialize(dto, indented ? Indented : Compact);
        }

        private static JsonSerializerOptions CreateOptions(bool writeIndented) => new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = writeIndented,
            // The default encoder escapes quotes, apostrophes, and other characters for safe HTML
            // embedding, which turns a message like 'N/A' into ''N/A''. This output is a
            // data payload, not HTML, so the relaxed encoder keeps it readable. Anything consuming
            // it as JSON is responsible for encoding it if it later lands in an HTML page.
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
        };

        // A fixed shape is used rather than serializing ExcelValidationResult directly, so the JSON
        // contract does not shift if the result type gains internal or computed members later.
        private sealed class ResultDto
        {
            public bool IsValid { get; set; }

            public bool ColumnsAreValid { get; set; }

            public bool RowsAreValid { get; set; }

            public int RowsValidated { get; set; }

            public bool Truncated { get; set; }

            public string[] MissingColumns { get; set; } = Array.Empty<string>();

            public string[] UnexpectedColumns { get; set; } = Array.Empty<string>();

            public ErrorDto[] Errors { get; set; } = Array.Empty<ErrorDto>();
        }

        private sealed class ErrorDto
        {
            public ValidationErrorKind Kind { get; set; }

            public string Message { get; set; } = string.Empty;

            public string? ColumnName { get; set; }

            public string? Address { get; set; }

            public int? Row { get; set; }

            public int? Column { get; set; }

            public string? Value { get; set; }
        }
    }
}
