using System;

namespace UnitTest.SampleService;

public interface ISampleService
{
    Task DoSomethingAsync(string input);
    List<string> GetErrors();
}
