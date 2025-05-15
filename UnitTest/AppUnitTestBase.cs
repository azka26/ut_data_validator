using System;
using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.DependencyInjection;
using UnitTest.SampleService;

namespace UnitTest;

public abstract class AppUnitTestBase : SqlServerUnitTestBase
{
    public override void CustomValidate(string validationName, DataRow expected, DataRow actual, string tableName, string columnName, int excelRowNumber)
    {
        throw new NotImplementedException();
    }

    protected override string GetSqlServerConnectionString()
    {
        return "Server=localhost,2022;Database=unit_test;User Id=sa;Password=AMNDev@2022;TrustServerCertificate=True;Encrypt=False;Connection Timeout=30;";
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
