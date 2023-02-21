// using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        public static IAssertion DefaultAssertion { get; set; }
        
        private readonly string _excelPath;
        private readonly string _worksheetInitData;
        private readonly string _worksheetExpectedData;
        private readonly IEventExcelValidator _eventExcelValidator;
        public ExcelTestDefinition InitialData { get; private set; }
        public ExcelTestDefinition ExpectedData { get; private set; }
        private readonly IAssertion _assert;
        
        public ExcelValidator(IAssertion assertion, string excelPath, string worksheetInitData, string worksheetExpectedData, IEventExcelValidator eventExcelValidator)
        {
            _excelPath = excelPath;
            _worksheetInitData = worksheetInitData;
            _worksheetExpectedData = worksheetExpectedData;
            _eventExcelValidator = eventExcelValidator;
            _assert = assertion;
            ReadExcel();
        }
        
        public ExcelValidator(string excelPath, string worksheetInitData, string worksheetExpectedData, IEventExcelValidator eventExcelValidator)
        {
            _excelPath = excelPath;
            _worksheetInitData = worksheetInitData;
            _worksheetExpectedData = worksheetExpectedData;
            _eventExcelValidator = eventExcelValidator;
            _assert = DefaultAssertion;
            ReadExcel();
        }

        public void Validate()
        {
            foreach (ExcelDataDefinition data in ExpectedData.ExcelDataDefinitions)
            {
                DataTable actual = _eventExcelValidator.ReadTable(data);
                CompareDataTable(data, actual);
            }
        }

        public void ExecuteAction() 
        {
            if (ExpectedData.TestActions != null)
            {
                foreach (TestAction action in ExpectedData.TestActions)
                {
                    for (int i = 0; i < action.Loop; i++)
                    {
                        _eventExcelValidator.ExecuteAction(action);
                    }
                }
            }
        }

        private void CompareDataTable(ExcelDataDefinition expected, DataTable actual)
        {
            _assert.AreEqual(expected.Data.Rows.Count, actual.Rows.Count, $"Different row count on table {expected.Table}.");
            for (var i = 0; i < expected.Data.Rows.Count; i++)
            {
                var rowExpected = expected.Data.Rows[i];
                var rowActual = actual.Rows[i];

                foreach (var item in expected.ColumnMapping.Values)
                {
                    if (!item.NeedValidation) continue;

                    _assert.AreEqual(
                        rowExpected[item.ColumnName] == DBNull.Value || string.IsNullOrEmpty(rowExpected[item.ColumnName].ToString()),
                        rowActual[item.ColumnName] == DBNull.Value || string.IsNullOrEmpty(rowActual[item.ColumnName].ToString()),
                        $"Different value on table {expected.Table} column {item.ColumnName} row {i + 1}."
                    );
                    
                    if (rowExpected[item.ColumnName] != DBNull.Value && rowActual[item.ColumnName] != DBNull.Value)
                    {
                        _assert.AreEqual(
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
            if (!File.Exists(_excelPath))
            {
                throw new Exception($"Excel file on path = '{_excelPath}' not found.");
            }

            FileInfo fi = new FileInfo(_excelPath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet initData = excelPackage.Workbook.Worksheets[_worksheetInitData];
                ExcelWorksheet expectedData = excelPackage.Workbook.Worksheets[_worksheetExpectedData];

                if (initData == null)
                {
                    throw new Exception($"Sheet {_worksheetInitData} not found on excel {_excelPath}.");
                }

                if (expectedData == null)
                {
                    throw new Exception($"Sheet {_worksheetExpectedData} not found on excel {_excelPath}.");
                }

                CollectInitData(initData);
                CollectExpectedData(expectedData);
            }
        }

        private IEnumerable<ExcelDataDefinition> ReadDefinition(ExcelWorksheet sheet)
        {
            Dictionary<string, ExcelDataDefinition> testDataMap = new Dictionary<string, ExcelDataDefinition>();
            int counterBlank = 0;
            int row = 1;
            while (counterBlank < 2)
            {
                ExcelRange cell = sheet.Cells[row, 1];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    counterBlank++;
                    row++;
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
                    testData.Data = _eventExcelValidator.ReadTable(testData);
                }
                
                row++;
            }

            return testDataMap.Values.ToList();
        }

        private IEnumerable<TestAction> ReadActions(ExcelWorksheet sheet)
        {
            List<TestAction> testActions = new List<TestAction>();
            int counterBlank = 0;
            int row = 1;
            while (counterBlank < 2)
            {
                ExcelRange cell = sheet.Cells[row, 1];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    counterBlank++;
                    row++;
                    continue;
                }

                counterBlank = 0;
                string value = cell.GetValue<string>();
                
                if (value.ToLower().StartsWith("action:"))
                {
                    TestAction testAction = new TestAction(value, row, sheet);
                    testActions.Add(testAction);
                    ReadActionParameters(testAction, sheet);
                }

                row++;
            }

            return testActions;
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
                parameters.Add(cell.GetValue<string>().ToLower(), (cellValue?.GetValue<string>()??""));

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
                        Object value = null;
                        if (!ConvertDataType(dataType, cell, out value))
                        {
                            throw new Exception($"Failed set value to DataTable on table {testData.Table} column {columnDefinition.ColumnName} excel row {row}.");
                        }

                        column.ReadOnly = false;
                        dataRow[columnDefinition.ColumnName] = Convert.ChangeType(value, dataType);
                    }
                    catch (Exception e)
                    {
                        throw new Exception($"Failed set value to DataTable on table {testData.Table} column {columnDefinition.ColumnName} excel row {row}.", e);
                    }
                }

                row++;
            }
        }

        private bool ConvertDataType(Type dataType, ExcelRange cell, out object value)
        {
            value = null;
            if (_eventExcelValidator.ConvertType(dataType, cell, out value))
            {
                return true;
            }
            
            if (dataType == typeof(DateTime))
            {
                value = cell.GetValue<DateTime>();
                return true;
            }
            
            if (dataType == typeof(TimeSpan))
            {
                var timeSpan = cell.GetValue<TimeSpan>();
                value = new TimeSpan(timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
                return true;
            }
            
            if (dataType == typeof(int))
            {
                value = cell.GetValue<int>();
                return true;
            }
            
            if (dataType == typeof(Int32))
            {
                value = cell.GetValue<Int32>();
                return true;
            }
            
            if (dataType == typeof(Int64))
            {
                value = cell.GetValue<Int64>();
                return true;
            }
            
            if (dataType == typeof(long))
            {
                value = cell.GetValue<long>();
                return true;
            }
            
            if (dataType == typeof(double))
            {
                value = cell.GetValue<double>();
                return true;
            }
            
            if (dataType == typeof(string))
            {
                value = cell.GetValue<string>();
                return true;
            }

            return false;
        }

        private void CollectInitData(ExcelWorksheet sheet)
        {
            IEnumerable<ExcelDataDefinition> testData = ReadDefinition(sheet);

            foreach (ExcelDataDefinition item in testData)
            {
                ReadExcelData(item, sheet);
            }

            _eventExcelValidator.InitData(testData);
            InitialData = new ExcelTestDefinition()
            {
                ExcelDataDefinitions = testData
            };
        }

        private void CollectExpectedData(ExcelWorksheet sheet)
        {
            IEnumerable<ExcelDataDefinition> testData = ReadDefinition(sheet);
            IEnumerable<TestAction> testActions = ReadActions(sheet);

            foreach (ExcelDataDefinition item in testData)
            {
                ReadExcelData(item, sheet);
            }

            ExpectedData = new ExcelTestDefinition()
            {
                TestActions = testActions,
                ExcelDataDefinitions = testData
            };
        }
    }
}
