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

        private UTValidationActivity GetUTManualActivity(int startRow, ExcelWorksheet worksheet)
        {
            var result = new UTValidationActivity();
            result.ExecutionType = ExecutionType.MANUAL;
            result.ErrorValidation = new List<string>();
            result.ExpectedData = string.Empty;

            if (worksheet.Cells[startRow, 1].GetValue<string>().ToLower() != "expecteddata")
            {
                return result;
            }

            if (worksheet.Cells[startRow + 1, 1].GetValue<string>().ToLower() != "errorvalidation")
            {
                return result;
            }

            result.ExpectedData = worksheet.Cells[startRow, 2].GetValue<string>();
            var counterBlank = 0;
            for (var errorValidationRow = startRow + 1; counterBlank < 2; errorValidationRow++)
            {
                if (worksheet.Cells[errorValidationRow, 1] != null && !string.IsNullOrWhiteSpace(worksheet.Cells[errorValidationRow, 1].GetValue<string>()) && worksheet.Cells[errorValidationRow, 1].GetValue<string>().ToLower() != "errorvalidation")
                {
                    break;
                }

                ExcelRange cell = worksheet.Cells[errorValidationRow, 2];
                if (cell == null || string.IsNullOrWhiteSpace(cell.GetValue<string>()))
                {
                    counterBlank++;
                }
                else
                {
                    result.ErrorValidation.Add(cell.GetValue<string>());
                    counterBlank = 0;
                }
            }

            return result;
        }

        private UTValidationActivity GetUTActivity(ExcelWorksheet worksheet)
        {
            var result = new UTValidationActivity();
            var mode = worksheet.Cells[1, 1].GetValue<string>();
            if (mode != "mode")
            {
                return GetUTManualActivity(1, worksheet);
            }

            result.ExecutionType = worksheet.Cells[1, 2].GetValue<string>().ToLower() == "auto" ? ExecutionType.AUTO : ExecutionType.MANUAL;
            if (result.ExecutionType == ExecutionType.MANUAL)
            {
                return GetUTManualActivity(2, worksheet);
            }

            // TODO: Implement the logic to read the worksheet and validate it using automation process
            throw new Exception($"Invalid execution type on worksheet {worksheet.Name}.");

            // if (worksheet.Cells[2, 1].GetValue<string>().ToLower().Trim() != "title")
            // {
            //     throw new Exception($"Invalid title on worksheet {worksheet.Name}.");
            // }

            // if (worksheet.Cells[3, 1].GetValue<string>().ToLower().Trim() != "action")
            // {
            //     throw new Exception($"Invalid action on worksheet {worksheet.Name}.");
            // }

            // if (worksheet.Cells[4, 1].GetValue<string>().ToLower().Trim() != "parameters")
            // {
            //     throw new Exception($"Invalid parameters on worksheet {worksheet.Name}.");
            // }

            // result.Title = worksheet.Cells[2, 2].GetValue<string>();
            // result.ExpectedSheetName = worksheet.Cells[3, 2].GetValue<string>();
            // result.Action = worksheet.Cells[4, 2].GetValue<string>();

            // #region Collect Parameters
            // result.Parameters = new List<string>();
            // int row = 5;
            // while (true)
            // {
            //     if (row > 3)
            //     {
            //         if (worksheet.Cells[row, 1].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].GetValue<string>()))
            //         {
            //             break;
            //         }
            //     }

            //     ExcelRange cell = worksheet.Cells[row, 2];
            //     if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
            //     {
            //         throw new Exception($"Invalid parameters on worksheet {worksheet.Name} row {row} column 2, parameter can't be empty.");
            //     }

            //     result.Parameters.Add(cell.GetValue<string>());
            //     row++;
            // }

            // var startRowResultValidation = row;
            // result.ResultValidation = new List<string>();
            // while (true)
            // {
            //     if (row > startRowResultValidation)
            //     {
            //         if (worksheet.Cells[row, 1].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].GetValue<string>()))
            //         {
            //             break;
            //         }
            //     }

            //     ExcelRange cell = worksheet.Cells[row, 2];
            //     if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
            //     {
            //         break;
            //     }

            //     result.ResultValidation.Add(cell.GetValue<string>());
            //     row++;
            // }
            // #endregion

            // return result;
        }

        public async Task AutoRunAsync(DirectoryInfo directoryInfo, IServiceProvider mainProvider)
        {
            if (directoryInfo == null)
            {
                throw new ArgumentNullException(nameof(directoryInfo));
            }

            if (!directoryInfo.Exists)
            {
                throw new DirectoryNotFoundException($"Directory not found: {directoryInfo.FullName}");
            }

            foreach (var file in directoryInfo.GetFiles("*.xlsx"))
            {
                await AutoRunPerFile(file, mainProvider);
            }
        }

        private async Task AutoRunPerFile(FileInfo file, IServiceProvider mainProvider)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            if (!file.Exists)
            {
                throw new FileNotFoundException($"File not found: {file.FullName}");
            }

            using (var package = new ExcelPackage(file))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    await RunTestAsync(file, mainProvider, worksheet);
                }
            }
        }

        private async Task RunTestAsync(FileInfo file, IServiceProvider mainProvider, ExcelWorksheet worksheet)
        {
            var activity = GetUTActivity(worksheet);
            if (activity.ExecutionType != ExecutionType.AUTO)
            {
                return;
            }

            Console.WriteLine($"Start Testing: " + activity.Title);
            Console.WriteLine($"Processing file: {file.FullName}, worksheet: {worksheet.Name}");

            try
            {
                using (var scope = mainProvider.CreateScope())
                {
                    var provider = scope.ServiceProvider;
                    // TODO: Implement the logic to read the worksheet and validate it using automation process
                    var validator = new ExcelValidator(this, file.FullName, worksheet.Name, activity.ExpectedSheetName, this);

                    // TODO: Run Activity Here
                    validator.Validate();
                }
                Console.WriteLine($"Completed Testing: " + activity.Title);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed Processing worksheet: {file.FullName}, worksheet: {worksheet.Name}", ex);
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine($"Completed Testing With Error: " + activity.Title);
            }
        }

        protected async Task RunTestAsync(FileInfo excelFile, string initSheetName, string expectedSheetName, Func<Task> action)
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

            await action.Invoke();
            ResetConnection();

            validator.Validate();
            ResetConnection();
        }

        protected async Task RunTestAsync(FileInfo excelFile, string initSheetName, string expectedSheetName,
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
        
        protected async Task RunTestAsync<T>(FileInfo excelFile, string initSheetName, string expectedSheetName,
            Func<IServiceProvider> createServiceProvider,
            Func<IServiceProvider, UTContext<T>, Task> action)
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
            
            var utContext = new UTContext<T>();
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
            validator.Validate<T>(utAct, utContext);
            ResetConnection();
        }
    }
}