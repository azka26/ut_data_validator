namespace UTDataValidator
{
    public class ExcelValidationException : Exception
    {
        public List<ErrorDescription> Errors { get; private set; } = new List<ErrorDescription>();

        public ExcelValidationException(string message, List<ErrorDescription> errors)
            : base(message + Environment.NewLine + string.Join(Environment.NewLine, errors.ConvertAll(e => e.ErrorMessage)))
        {
            Errors = errors;
        }
    }
}