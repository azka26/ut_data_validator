namespace UTDataValidator
{
    public interface IAssertion
    {
        void AreEqual<T>(T expected, T actual, string message);
        void IsTrue(bool condition, string message);
    }
}