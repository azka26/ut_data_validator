using System.Collections.Generic;

namespace UTDataValidator
{
    public class UTValidationActivity
    {
        public ExecutionType ExecutionType { get; set; } = ExecutionType.MANUAL;
        public string Title { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public List<string> Parameters { get; set; } = new List<string>();
        public string ExpectedSheetName { get; set; } = string.Empty;
        public List<string> ResultValidation { get; set; } = new List<string>();
    }
}
