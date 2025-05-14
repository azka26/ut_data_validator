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

        private UTValidationActivity CollectActivity(ExcelWorksheet worksheet) 
        {
            var result = new UTValidationActivity();
            var mode = worksheet.Cells[1, 1].GetValue<string>();
            if (mode != "mode") {
                return result;
            }

            result.ExecutionType = worksheet.Cells[1, 2].GetValue<string>().ToLower() == "auto" ? ExecutionType.AUTO : ExecutionType.MANUAL;
            if (result.ExecutionType == ExecutionType.MANUAL)
            {
                return result;
            }

            if (worksheet.Cells[2, 1].GetValue<string>().ToLower().Trim() != "title")
            {
                throw new Exception($"Invalid title on worksheet {worksheet.Name}.");
            }

            if (worksheet.Cells[3, 1].GetValue<string>().ToLower().Trim() != "action")
            {
                throw new Exception($"Invalid action on worksheet {worksheet.Name}.");
            }

            if (worksheet.Cells[4, 1].GetValue<string>().ToLower().Trim() != "parameters")
            {
                throw new Exception($"Invalid parameters on worksheet {worksheet.Name}.");
            }

            result.Title = worksheet.Cells[2, 2].GetValue<string>();
            result.ExpectedSheetName = worksheet.Cells[3, 2].GetValue<string>();
            result.Action = worksheet.Cells[4, 2].GetValue<string>();

            #region Collect Parameters
            result.Parameters = new List<string>();
            int row = 5;
            while (true)
            {
                if (row > 3) 
                {
                    if (worksheet.Cells[row, 1].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].GetValue<string>()))
                    {
                        break;
                    }
                }

                ExcelRange cell = worksheet.Cells[row, 2];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    throw new Exception($"Invalid parameters on worksheet {worksheet.Name} row {row} column 2, parameter can't be empty.");
                }

                result.Parameters.Add(cell.GetValue<string>());
                row++;
            }

            var startRowResultValidation = row;
            result.ResultValidation = new List<string>();
            while (true) 
            {
                if (row > startRowResultValidation) 
                {
                    if (worksheet.Cells[row, 1].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].GetValue<string>()))
                    {
                        break;
                    }
                }

                ExcelRange cell = worksheet.Cells[row, 2];
                if (cell == null || string.IsNullOrEmpty(cell.GetValue<string>()))
                {
                    break;
                }

                result.ResultValidation.Add(cell.GetValue<string>());
                row++;
            }
            #endregion

            return result;
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
                    bool flowControl = await NewMethod(file, mainProvider, worksheet);
                    if (!flowControl)
                    {
                        continue;
                    }
                }
            }
        }

        private async Task NewMethod(FileInfo file, IServiceProvider mainProvider, ExcelWorksheet worksheet)
        {
            var activity = CollectActivity(worksheet);
            if (activity.ExecutionType != ExecutionType.AUTO)
            {
                return;
            }

            try {
                using (var scope = mainProvider.CreateScope())
                {
                    var provider = scope.ServiceProvider;
                    // TODO: Implement the logic to read the worksheet and validate it using automation process
                    var validator = new ExcelValidator(this, file.FullName, worksheet.Name, activity.ExpectedSheetName, this);

                    // TODO: Run Activity Here

                    validator.Validate();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed Processing worksheet: {0}", ex);
            }

            return;
        }
    }
}