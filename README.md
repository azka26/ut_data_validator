# UTDataValidator Documentation

`UTDataValidator` is a library designed to simplify unit testing by providing tools for validating data and error messages. This documentation explains how to set up and use `UTDataValidator` in your project.

## Prerequisites

Before using `UTDataValidator`, ensure you have the following:
- A base unit test class that inherits from `SqlServerUnitTestBase` and implements required methods.
- Unit test classes that extend your base unit test class and define test cases.

## Step 1: Create a Base Unit Test Class

The base unit test class should inherit from `SqlServerUnitTestBase` and implement methods for validation and service provider configuration. Below is an example implementation:

```csharp
// filepath: /home/andika/Documents/azka/ut_data_validator/UnitTest/SampleUnitTestBase.cs
public abstract class SampleUnitTestBase : SqlServerUnitTestBase
{
    #region VALIDATION
    public override void AreEqual<T>(T expected, T actual, string message)
    {
        var dateTimeFormat = "{0:yyyy-MM-dd HH:mm:ss}";
        if (expected != null && actual != null)
        {
            if (expected is DateTime expectedDate && actual is DateTime actualDate)
            {
                var expectedString = string.Format(dateTimeFormat, expectedDate);
                var actualString = string.Format(dateTimeFormat, actualDate);
                Assert.True(expectedString == actualString, message);
                return;
            }
        }

        Assert.True(expected.Equals(actual), message);
    }

    public override void IsTrue(bool condition, string message)
    {
        Assert.True(condition, message);
    }

    public override void CustomValidate(string validationName, DataRow expected, DataRow actual, string tableName, string columnName, int excelRowNumber)
    {
        throw new NotImplementedException();
    }
    #endregion

    protected override string GetSqlServerConnectionString()
    {
        return "Server=localhost,2022;Database=unit_test;User Id=sa;Password=AZDev@2022;TrustServerCertificate=True;Encrypt=False;Connection Timeout=30;";
    }

    private IServiceProvider _serviceProvider;
    protected IServiceProvider GetServiceProvider()
    {
        if (_serviceProvider != null)
        {
            return _serviceProvider;
        }

        var serviceCollection = new ServiceCollection();
        serviceCollection.AddTransient<SqlConnection>((provider) => new SqlConnection(GetSqlServerConnectionString()));
        serviceCollection.AddTransient<ISampleService, SampleService.SampleService>();
        _serviceProvider = serviceCollection.BuildServiceProvider();
        return _serviceProvider;
    }
}
```

### Key Points:
- Implement the `AreEqual`, `IsTrue`, and `CustomValidate` methods for custom validation logic.
- Provide a SQL Server connection string in `GetSqlServerConnectionString`.
- Configure dependency injection in `GetServiceProvider`.

## Step 2: Create Unit Test Classes

Unit test classes should extend your base unit test class and define test cases using the `RunTestAsync` method. Below is an example:

```csharp
// filepath: /home/andika/Documents/azka/ut_data_validator/UnitTest/SampleUnitTest.cs
public class SampleUnitTest : SampleUnitTestBase
{
    private FileInfo file = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sample", "Item", "item.xlsx"));

    [Fact]
    public async Task CompareDataOnly_WithProvider_Valid_Test()
    {
        await RunTestAsync(
            file,
            "InitSubmitValidData",
            "ExpSubmitValidData",
            () => GetServiceProvider(),
            async (IServiceProvider provider) =>
            {
                var sampleService = provider.GetRequiredService<ISampleService>();
                await sampleService.DoSomethingAsync("no error");
            }
        );
    }

    [Fact]
    public async Task CompareDataAndError_Invalid_Test()
    {
        await RunTestAsync(
            file,
            "InitSubmitInvalidData",
            "InitSubmitInvalidData",
            () => GetServiceProvider(),
            async (IServiceProvider provider, UTContext utContext) =>
            {
                var sampleService = provider.GetRequiredService<ISampleService>();
                await sampleService.DoSomethingAsync("error_2");
                utContext.ErrorMessages = sampleService.GetErrors();
            }
        );
    }
    
    [Fact]
    public async Task CompareDataOnly_WithProvider_Valid_AutoSelectExpected_Test()
    {
        await RunTestAsync(
            file,
            "InitSubmitValidData",
            () => GetServiceProvider(),
            async (IServiceProvider provider, UTContext context) =>
            {
                var sampleService = provider.GetRequiredService<ISampleService>();
                await sampleService.DoSomethingAsync("no error");
            }
        );
    }
    
    [Fact]
    public async Task CompareDataOnly_WithProvider_Invalid_AutoSelectExpected_Test()
    {
        await RunTestAsync(
            file,
            "InitSubmitInvalidData",
            () => GetServiceProvider(),
            async (IServiceProvider provider, UTContext context) =>
            {
                var sampleService = provider.GetRequiredService<ISampleService>();
                await sampleService.DoSomethingAsync("error_2");
                context.ErrorMessages = sampleService.GetErrors();
            }
        );
    }
}
```

### Key Points:
- Use the `RunTestAsync` method to define test cases.
- Pass the file, initialization data, expected data, and service provider to `RunTestAsync`.
- Use the `UTContext` object to capture error messages or additional context during testing.

## Step 3: Run Your Tests

Run your tests using your preferred test runner (e.g., `dotnet test` for .NET projects). Ensure your test data files (e.g., Excel files) are correctly placed in the expected directory.

## Conclusion

By following the steps above, you can effectively use `UTDataValidator` to validate data and error messages in your unit tests. Customize the base unit test class and test cases as needed to fit your project's requirements.