using NUnit.Framework;
using System;
using System.Data;
using UTDataValidator;

namespace UnitTestProject1
{
    public class Assertion : IAssertion
    {
        public void AreEqual<T>(T expected, T actual, string message)
        {
            Assert.AreEqual(expected, actual, message);
        }

        public void IsTrue(bool condition, string message)
        {
            Assert.IsTrue(condition, message);
        }

        public void CustomValidate(string validationName, DataRow expected, DataRow actual, string tableName, string columnName, int excelRowNumber)
        {
            if (validationName == "NULL_OR_EMPTY")
            {
                var expectedCondition = expected[columnName] == DBNull.Value || string.IsNullOrEmpty(expected[columnName].ToString());
                var actualCondition = actual[columnName] == DBNull.Value || string.IsNullOrEmpty(actual[columnName].ToString());
                Assert.AreEqual(expectedCondition, actualCondition, $"Different value on table {expected.Table} column {columnName} row {excelRowNumber}.");

                return;
            }

            throw new NotImplementedException();
        }
    }
}