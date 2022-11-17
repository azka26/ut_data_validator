using System.Collections.Generic;

namespace UTDataValidator
{
    public class ExcelTestDefinition
    {
        public IEnumerable<TestAction> TestActions { get; set; }
        public IEnumerable<ExcelDataDefinition> ExcelDataDefinitions { get; set; }
    }
}