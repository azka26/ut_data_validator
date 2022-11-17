using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;

namespace UTDataValidator
{
    public class ExcelDataDefinition
    {
        public ExcelDataDefinition(string cellValue, int rowNumber, ExcelWorksheet sheet)
        {
            string[] configs = cellValue.Split(';');
            for (int i = 0; i < configs.Length; i++)
            {
                configs[i] = configs[i].Trim();
            }

            Table = configs[0].Split(':')[1].Trim();
            WorksheetName = sheet.Name;
            CellValue = cellValue;
            RowNumber = rowNumber;
        }
        public string Table { get; }
        public string WorksheetName { get; }
        public int RowNumber { get; }
        public string CellValue { get; }
        public Dictionary<string, ColumnDefinition> ColumnMapping { get; set; } = new Dictionary<string, ColumnDefinition>();
        public DataTable Data { get; set; }
    }
}
