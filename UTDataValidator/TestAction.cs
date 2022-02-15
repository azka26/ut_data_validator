using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace UTDataValidator
{
    public class TestAction
    {
        public TestAction(string cellValue, int rowNumber, ExcelWorksheet sheet)
        {
            string[] configs = cellValue.Split(';');
            for (int i = 0; i < configs.Length; i++)
            {
                configs[i] = configs[i].Trim();
            }

            ActionName = configs[0].Split(':')[1].Trim();
            CellValue = cellValue;
            WorksheetName = sheet.Name;
            RowNumber = rowNumber;
            string loopValue = sheet.Cells[rowNumber + 1, 1].GetValue<string>();
            if (loopValue == null || !loopValue.ToLower().StartsWith("loop"))
            {
                throw new Exception($"Loop definition not found on sheet {sheet.Name} column 1, row {rowNumber + 1}.");
            }

            Loop = Convert.ToInt32(loopValue.Split(':')[1].Trim());
            string columnKey = sheet.Cells[rowNumber + 2, 1].GetValue<string>();
            string columnValue = sheet.Cells[rowNumber + 2, 2].GetValue<string>();
            if (string.IsNullOrEmpty(columnKey) || string.IsNullOrEmpty(columnValue))
            {
                throw new Exception($"key, value definition not found on sheet {sheet.Name} column 1 and 2, row {rowNumber + 1}.");
            }
        }

        public string ActionName { get; }
        public string WorksheetName { get; }
        public int Loop { get; }
        public int RowNumber { get; }
        public string CellValue { get; }
        public Dictionary<string, string> Parameters { get; internal set; }
    }
}
