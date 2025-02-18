using System.Data;

namespace UTDataValidator
{
    public interface IAssertion
    {
        void AreEqual<T>(T expected, T actual, string message);
        void IsTrue(bool condition, string message);
        
        /// <summary>
        /// Implement this method to add custom validation
        /// </summary>
        /// <param name="validationName">value is Validation Name, value UPPER CASE, example validate(empty_or_null) will be EMPTY_OR_NULL</param>
        /// <param name="expected"></param>
        /// <param name="actual"></param>
        /// <param name="tableName"></param>
        /// <param name="columnName"></param>
        /// <param name="excelRowNumber"></param>
        void CustomValidate(string validationName, DataRow expected, DataRow actual, string tableName, string columnName, int excelRowNumber);
    }
}