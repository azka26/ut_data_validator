using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace UTDataValidator
{
    public class ExcelValidator 
    {
        private readonly string _excelPath;
        private readonly string _worksheetInitData;
        private readonly string _worksheetExpectedData;
        private readonly Func<ExcelDataDefinition, DataTable> _readTableFunction;
        private readonly Action<IEnumerable<ExcelDataDefinition>> _processInitData;
        public ExcelTestDefinition InitialData { get; private set; }
        public ExcelTestDefinition ExpectedData { get; private set; }

        public ExcelValidator(string excelPath, string worksheetInitData, string worksheetExpectedData, Func<ExcelDataDefinition, DataTable> readTableFunction, Action<IEnumerable<ExcelDataDefinition>> processInitData)
        {
            _excelPath = excelPath;
            _worksheetInitData = worksheetInitData;
            _worksheetExpectedData = worksheetExpectedData;
            _readTableFunction = readTableFunction;
            _processInitData = processInitData;
            ReadExcel();
        }

        public void Validate()
        {
            foreach (ExcelDataDefinition data in ExpectedData.ExcelDataDefinitions)
            {
                DataTable actual = _readTableFunction(data);
                CompareDataTable(data, actual);
            }
        }

        public void ExecuteAction(Action<TestAction> action) 
        { 
            if (ExpectedData.TestAction != null)
            {
                action(ExpectedData.TestAction);
            }
        }

        private void CompareDataTable(ExcelDataDefinition expected, DataTable actual)
        {
            Assert.AreEqual(expected.Data.Rows.Count, actual.Rows.Count, $"Different row count on table {expected.Table}.");
            for (int i = 0; i < expected.Data.Rows.Count; i++)
            {
                DataRow rowExpected = expected.Data.Rows[i];
                DataRow rowActual = actual.Rows[i];

                foreach (ColumnDefinition item in expected.ColumnMapping.Values)
                {
                    if (!item.NeedValidation) continue;

                    Assert.AreEqual(
                        rowExpected[item.ColumnName] == DBNull.Value || string.IsNullOrEmpty(rowExpected[item.ColumnName].ToString()),
                        rowActual[item.ColumnName] == DBNull.Value || string.IsNullOrEmpty(rowActual[item.ColumnName].ToString()),
                        $"Different value on table {expected.Table} column {item.ColumnName} row {i + 1}."
                    );
                    
                    if (rowExpected[item.ColumnName] != DBNull.Value && rowActual[item.ColumnName] != DBNull.Value)
                    {
                        Assert.AreEqual(
                            rowExpected[item.ColumnName], 
                            rowActual[item.ColumnName], 
                            $"Different value on table {expected.Table} column {item.ColumnName} row {i + 1}."
                        );
                    }
                }
            }
        }

        private void ReadExcel()
        {
            FileInfo fi = new FileInfo(_excelPath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet initData = excelPackage.Workbook.Worksheets[_worksheetInitData];
                ExcelWorksheet expectedData = excelPackage.Workbook.Worksheets[_worksheetExpectedData];
                CollectInitData(initData);
                CollectExpectedData(expectedData);
            }
        }

        private IEnumerable<ExcelDataDefinition> ReadDefinition(ExcelWorksheet sheet, out TestAction testAction)
        {
            testAction = null;

            Dictionary<string, ExcelDataDefinition> testDataMap = new Dictionary<string, ExcelDataDefinition>();
            int counterBlank = 0;
            int row = 1;
            while (counterBlank < 2)
            {
                ExcelRange cell = sheet.Cells[row, 1];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    counterBlank++;
                    continue;
                }

                counterBlank = 0;
                string value = cell.GetValue<string>();
                if (value.ToLower().StartsWith("table:"))
                {
                    ExcelDataDefinition testData = new ExcelDataDefinition(value, row, sheet);
                    if (testDataMap.ContainsKey(testData.Table.ToLower()))
                    {
                        throw new Exception($"Duplicate table name on worksheet {sheet.Name} between row {testData.RowNumber} and {testDataMap[testData.Table.ToLower()].RowNumber}.");
                    }

                    testDataMap.Add(testData.Table.ToLower(), testData);
                    ReadColumnDefinition(testData, sheet);
                    testData.Data = _readTableFunction(testData);
                }

                if (value.ToLower().StartsWith("action:"))
                {
                    if (testAction != null)
                    {
                        throw new Exception($"Duplicate action on worksheet {sheet.Name} between row {row} and {testAction.RowNumber}.");
                    }

                    testAction = new TestAction(value, row, sheet);
                    ReadActionParameters(testAction, sheet);
                }
                
                row++;
            }

            return testDataMap.Values.ToList();
        }

        private void ReadActionParameters(TestAction action, ExcelWorksheet sheet)
        {
            int row = action.RowNumber + 3;
            int keyColumn = 1;
            int valueColumn = 2;
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            while (true)
            {
                ExcelRange cell = sheet.Cells[row, keyColumn];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    break;
                }

                if (parameters.ContainsKey(cell.GetValue<string>().ToLower()))
                {
                    throw new Exception($"Duplicate key action parameter on worksheet {sheet.Name}, column 1, row {row}.");
                }

                ExcelRange cellValue = sheet.Cells[row, valueColumn];
                parameters.Add(cell.GetValue<string>().ToLower(), (cellValue?.GetValue<string>()??"").ToLower());

                row++;
            }
            action.Parameters = parameters;
        }
        private void ReadColumnDefinition(ExcelDataDefinition testData, ExcelWorksheet sheet)
        {
            int column = 1;
            int row = testData.RowNumber + 1;

            while (true)
            {
                ExcelRange cell = sheet.Cells[row, column];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    break;
                }

                ColumnDefinition columnDefinition = new ColumnDefinition(cell.GetValue<string>(), column);
                if (testData.ColumnMapping.ContainsKey(columnDefinition.ColumnName.ToLower()))
                {
                    throw new Exception($"Duplicate column name on worksheet {sheet.Name} for table {testData.Table} column {columnDefinition.ColumnIndex} and {testData.ColumnMapping[columnDefinition.ColumnName].ColumnIndex}.");
                }

                testData.ColumnMapping.Add(columnDefinition.ColumnName.ToLower(), columnDefinition);
                column++;
            }
        }

        private void ReadExcelData(ExcelDataDefinition testData, ExcelWorksheet sheet)
        {
            testData.Data.Clear();
            foreach (DataColumn column in testData.Data.Columns)
            {
                column.AllowDBNull = true;
            }

            int row = testData.RowNumber + 2;
            while(true)
            {
                ExcelRange cell = sheet.Cells[row, 1];
                if (cell != null && !string.IsNullOrEmpty(cell.GetValue<string>()) && cell.GetValue<string>().ToLower().StartsWith("table:"))
                {
                    break;
                }

                bool hasValue = false;
                foreach (ColumnDefinition columnDefinition in testData.ColumnMapping.Values)
                {
                    cell = sheet.Cells[row, columnDefinition.ColumnIndex];
                    if (cell != null && !string.IsNullOrEmpty(cell.GetValue<string>()))
                    {
                        hasValue = true;
                    }
                }
                if (!hasValue)
                {
                    break;
                }

                DataRow dataRow = testData.Data.NewRow();
                testData.Data.Rows.Add(dataRow);
                foreach (ColumnDefinition columnDefinition in testData.ColumnMapping.Values)
                {
                    if (!testData.Data.Columns.Contains(columnDefinition.ColumnName))
                    {
                        throw new Exception($"Column {columnDefinition.ColumnName} on table {testData.Table} not found.");
                    }
                    
                    cell = sheet.Cells[row, columnDefinition.ColumnIndex];
                    if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()) || cell.GetValue<string>().ToLower() == "null")
                    {
                        continue;
                    }

                    DataColumn column = testData.Data.Columns[columnDefinition.ColumnName];
                    Type dataType = column.DataType;

                    try
                    {
                        if (dataType == typeof(DateTime))
                        {
                            DateTime dateValue = cell.GetValue<DateTime>();
                            dataRow[columnDefinition.ColumnName] = dateValue;
                        }
                        else 
                        {
                            dataRow[columnDefinition.ColumnName] = Convert.ChangeType(cell.Value, dataType);
                        }
                    }
                    catch
                    {
                        throw new Exception($"Failed set value to DataTable on table {testData.Table} column {columnDefinition.ColumnName} excel row {row}.");
                    }
                }

                row++;
            }
        }

        private void CollectInitData(ExcelWorksheet sheet)
        {
            TestAction testAction = null;
            IEnumerable<ExcelDataDefinition> testData = ReadDefinition(sheet, out testAction);
            foreach (ExcelDataDefinition item in testData)
            {
                ReadExcelData(item, sheet);
            }
            _processInitData(testData);
            InitialData = new ExcelTestDefinition()
            {
                TestAction = testAction,
                ExcelDataDefinitions = testData
            };
        }

        private void CollectExpectedData(ExcelWorksheet sheet)
        {
            TestAction testAction = null;
            IEnumerable<ExcelDataDefinition> testData = ReadDefinition(sheet, out testAction);
            foreach (ExcelDataDefinition item in testData)
            {
                ReadExcelData(item, sheet);
            }
            ExpectedData = new ExcelTestDefinition()
            {
                TestAction = testAction,
                ExcelDataDefinitions = testData
            };
        }
    }
}
