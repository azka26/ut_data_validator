using System.Collections.Generic;

namespace UTDataValidator
{
    public class UTValidationActivity
    {
        public string TestCase { get; set; } = string.Empty;
        public string ExpectedSheet { get; set; } = string.Empty;
        public List<string> ErrorValidation { get; } = new List<string>();

        // public string Action { get; set; } = string.Empty;
        // public List<string> Parameters { get; set; } = new List<string>();
        // public ExecutionType ExecutionType { get; set; } = ExecutionType.MANUAL;
    }
}
