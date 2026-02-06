using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using Microsoft.Extensions.DependencyInjection;

namespace UTDataValidator
{
    public abstract class ExcelValidationUTBase : IEventExcelValidator, IAssertion
    {
        public abstract bool ConvertType(Type type, ExcelRange excelRange, out object outputValue);
        public abstract void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition);
        public abstract DataTable ReadTable(ExcelDataDefinition excelDataDefinition);
        public abstract void AreEqual<T>(T expected, T actual, string message);
        public abstract void IsTrue(bool condition, string message);
        public abstract void CustomValidate(string validationName, DataRow expected, DataRow actual, string tableName, string columnName, int excelRowNumber);
        protected abstract void ResetConnection();
        protected abstract string GetConnectionString();

        private UTValidationActivity GetUTActivity(ExcelWorksheet worksheet)
        {
            var mapRow = new Dictionary<string, int>()
            {
                { nameof(UTValidationActivity.TestCase), 0 },
                { nameof(UTValidationActivity.ExpectedSheet), 0 },
                { nameof(UTValidationActivity.ErrorValidation), 0 }
            };

            var counterBlank = 0;
            for (var rowNumber = 1; counterBlank < 2; rowNumber++)
            {
                var cellTitle = worksheet.Cells[rowNumber, 1]?.GetValue<string>()?.Trim()?.ToLower() ?? string.Empty;
                var cellValue = worksheet.Cells[rowNumber, 2]?.GetValue<string>()?.Trim()?.ToLower() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(cellTitle) && string.IsNullOrWhiteSpace(cellValue))
                {
                    counterBlank++;
                    continue;
                }

                if (cellTitle.IsTableInfo())
                {
                    break;
                }

                foreach (var map in mapRow)
                {
                    var key = map.Key;
                    if (cellTitle == key.ToLower())
                    {
                        mapRow[key] = rowNumber;
                        break;
                    }
                }

                counterBlank = 0;
            }

            var result = new UTValidationActivity();
            result.TestCase = mapRow[nameof(UTValidationActivity.TestCase)] > 0
                ? worksheet.Cells[mapRow[nameof(UTValidationActivity.TestCase)], 2]?.GetValue<string>() ?? string.Empty
                : string.Empty;

            result.ExpectedSheet = mapRow[nameof(UTValidationActivity.ExpectedSheet)] > 0
                ? worksheet.Cells[mapRow[nameof(UTValidationActivity.ExpectedSheet)], 2]?.GetValue<string>() ?? string.Empty
                : string.Empty;

            if (mapRow[nameof(UTValidationActivity.ErrorValidation)] > 0)
            {
                for (var rowNumber = mapRow[nameof(UTValidationActivity.ErrorValidation)]; true; rowNumber++)
                {
                    var cellTitle = worksheet.Cells[rowNumber, 1]?.GetValue<string>()?.Trim()?.ToLower() ?? string.Empty;
                    var cellValue = worksheet.Cells[rowNumber, 2]?.GetValue<string>()?.Trim()?.ToLower() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(cellTitle) && string.IsNullOrWhiteSpace(cellValue))
                    {
                        break;
                    }

                    if (!string.IsNullOrWhiteSpace(cellTitle) && cellTitle != nameof(UTValidationActivity.ErrorValidation).ToLower())
                    {
                        break;
                    }

                    if (cellTitle == nameof(UTValidationActivity.ErrorValidation) && string.IsNullOrEmpty(cellValue))
                    {
                        continue;
                    }

                    if (string.IsNullOrEmpty(cellValue))
                    {
                        break;
                    }

                    result.ErrorValidation.Add(worksheet.Cells[rowNumber, 2].GetValue<string>().Trim());
                }
            }

            return result;
        }

        protected async Task RunTestAsync(
            FileInfo excelFile,
            string initSheetName,
            string expectedSheetName,
            Func<IServiceProvider> createServiceProvider,
            Func<IServiceProvider, Task> action)
        {
            if (excelFile == null)
            {
                throw new ArgumentNullException(nameof(excelFile));
            }

            if (!excelFile.Exists)
            {
                throw new FileNotFoundException($"File not found: {excelFile.FullName}");
            }

            if (string.IsNullOrEmpty(initSheetName))
            {
                throw new ArgumentNullException(nameof(initSheetName));
            }

            if (string.IsNullOrEmpty(expectedSheetName))
            {
                throw new ArgumentNullException(nameof(expectedSheetName));
            }

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            var validator = new ExcelValidator(this, excelFile.FullName, initSheetName, expectedSheetName, this);
            ResetConnection();

            var serviceProvider = createServiceProvider();
            using (var scope = serviceProvider.CreateScope())
            {
                var provider = scope.ServiceProvider;
                await action.Invoke(provider);
            }
            ResetConnection();

            validator.Validate();
            ResetConnection();
        }

        protected async Task RunTestAsync(
            FileInfo excelFile,
            string initSheetName,
            string expectedSheetName,
            Func<IServiceProvider> createServiceProvider,
            Func<IServiceProvider, UTContext, Task> action)
        {
            if (excelFile == null)
            {
                throw new ArgumentNullException(nameof(excelFile));
            }

            if (!excelFile.Exists)
            {
                throw new FileNotFoundException($"File not found: {excelFile.FullName}");
            }

            if (string.IsNullOrEmpty(initSheetName))
            {
                throw new ArgumentNullException(nameof(initSheetName));
            }

            if (string.IsNullOrEmpty(expectedSheetName))
            {
                throw new ArgumentNullException(nameof(expectedSheetName));
            }

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            var validator = new ExcelValidator(this, excelFile.FullName, initSheetName, expectedSheetName, this);
            ResetConnection();

            var utContext = new UTContext();
            var serviceProvider = createServiceProvider();
            using (var scope = serviceProvider.CreateScope())
            {
                var provider = scope.ServiceProvider;
                await action.Invoke(provider, utContext);
            }
            ResetConnection();

            var excelPackage = new ExcelPackage(excelFile);
            var worksheet = excelPackage.Workbook.Worksheets[initSheetName];
            var utAct = GetUTActivity(worksheet);
            validator.Validate(utAct, utContext);
            ResetConnection();
        }
        
        protected async Task RunTestAsync(
            FileInfo excelFile,
            string initSheetName,
            Func<IServiceProvider> createServiceProvider,
            Func<IServiceProvider, UTContext, Task> action)
        {
            ResetConnection();
            
            #region Validate Parameters
            if (excelFile == null)
            {
                throw new ArgumentNullException(nameof(excelFile));
            }

            if (!excelFile.Exists)
            {
                throw new FileNotFoundException($"File not found: {excelFile.FullName}");
            }

            if (string.IsNullOrEmpty(initSheetName))
            {
                throw new ArgumentNullException(nameof(initSheetName));
            }

            if (createServiceProvider == null)
            {
                throw new ArgumentNullException(nameof(createServiceProvider));
            }

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }
            #endregion

            #region Get Activity
            var excelPackage = new ExcelPackage(excelFile);
            var worksheet = excelPackage.Workbook.Worksheets[initSheetName];
            var utAct = GetUTActivity(worksheet);
            if (utAct == null)
            {
                throw new Exception($"UTActivity not found in file: \"{excelFile.FullName}\", worksheet: {initSheetName}.");
            }

            if (string.IsNullOrEmpty(utAct.ExpectedSheet))
            {
                throw new Exception($"ExpectedSheet not found in file: \"{excelFile.FullName}\", worksheet: {initSheetName}.");
            }

            if (excelPackage.Workbook.Worksheets[utAct.ExpectedSheet] == null)
            {
                throw new Exception($"ExpectedSheet with name = \"{utAct.ExpectedSheet}\" not found in file: \"{excelFile.FullName}\".");
            }
            #endregion

            var validator = new ExcelValidator(this, excelFile.FullName, initSheetName, utAct.ExpectedSheet, this);
            ResetConnection();
            
            var utContext = new UTContext();
            var serviceProvider = createServiceProvider();
            using (var scope = serviceProvider.CreateScope())
            {
                var provider = scope.ServiceProvider;
                await action.Invoke(provider, utContext);
            }
            ResetConnection();
            
            validator.Validate(utAct, utContext);
            ResetConnection();
        }

        protected async Task ValidateExcelAsync(
            FileInfo excelFile,
            string sheetToCheck)
        {
            if (excelFile == null)
            {
                throw new ArgumentNullException(nameof(excelFile));
            }

            if (!excelFile.Exists)
            {
                throw new FileNotFoundException($"File not found: {excelFile.FullName}");
            }

            if (string.IsNullOrEmpty(sheetToCheck))
            {
                throw new ArgumentNullException(nameof(sheetToCheck));
            }

            _ = new ExcelValidator(this, excelFile.FullName, sheetToCheck, sheetToCheck, this);
        }
    }
}