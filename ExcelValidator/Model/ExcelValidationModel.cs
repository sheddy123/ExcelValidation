using System.Collections.Generic;

namespace ExcelValidator.Model
{
    public partial class ValidateExcelSheet
    {
        public class ExcelValidationModel
        {
            public string ColumnName { get; }

            public int Row { get; }

            public int Column { get; }

            public string Comment { get; }

            public string  ErrorComment { get; set; }

            public byte[] ExcelFile { get; }

            public List<object> HeaderColumns { get; }

        }
        

    }
}
