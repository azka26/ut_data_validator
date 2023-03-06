using NUnit.Framework;
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
    }
}