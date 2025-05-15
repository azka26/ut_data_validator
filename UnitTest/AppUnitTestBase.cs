using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.DependencyInjection;
using UnitTest.SampleService;
using UTDataValidator.SqlServer;

namespace UnitTest;

public abstract class AppUnitTestBase : SqlServerUnitTestBase
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
