using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace UTDataValidator
{
    public interface IEventExcelValidator
    {
        DataTable ReadTable(ExcelDataDefinition excelDataDefinition);
        void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition);
        bool ConvertType(Type type, ExcelRange excelRange, out object outputValue);
    }
}
