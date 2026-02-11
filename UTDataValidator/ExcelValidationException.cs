namespace UTDataValidator
{
    public class ExcelValidationException : Exception
    {
        public List<ErrorDescription> Errors { get; private set; } = new List<ErrorDescription>();

        public ExcelValidationException(string message, List<ErrorDescription> errors)
            : base(message)
        {
            Errors = errors;
        }
    }
}