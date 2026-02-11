using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace UTDataValidator
{
    public enum ValidationPhase
    {
        Initialization,
        DataReading,
        DataValidation
    }

    public class ErrorDescription
    {
        public string ErrorMessage { get; set; } = string.Empty;
        public string WorksheetName { get; set;} = string.Empty;
        public Exception? Exception { get; set;} = null;
        public int RowNumber { get; set; }
    }

    public interface IEventExcelValidator
    {
        DataTable ReadTable(ExcelDataDefinition excelDataDefinition);
        void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition);
        bool ConvertType(Type type, ExcelRange excelRange, out object outputValue);
        List<ErrorDescription> GetErrors();
        void AddError(ErrorDescription error);
        void ThrowIfErrors();
    }
}
