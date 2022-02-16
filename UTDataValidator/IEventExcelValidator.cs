using System.Collections.Generic;
using System.Data;

namespace UTDataValidator
{
    public interface IEventExcelValidator
    {
        DataTable ReadTable(ExcelDataDefinition excelDataDefinition);
        void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition);
        void ExecuteAction(TestAction testAction);
    }
}
