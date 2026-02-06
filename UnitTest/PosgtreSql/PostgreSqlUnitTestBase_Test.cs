using System.Data;
using Npgsql;
using Microsoft.Extensions.DependencyInjection;
using UnitTest.SampleService;
using UTDataValidator.PostgreSql;
using UTDataValidator;

namespace UnitTest.PosgtreSql;

public abstract class PostgreSqlUnitTestBase_Test : PostgreSqlUnitTestBase
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

        Assert.True(expected?.Equals(actual) ?? false, message);
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

    protected override string GetConnectionString()
    {
        return "Server=localhost;Port=5432;Database=dbname;User ID=postgres;Password=test123;";
    }

    private IServiceProvider? _serviceProvider;
    protected IServiceProvider GetServiceProvider()
    {
        if (_serviceProvider != null)
        {
            return _serviceProvider;
        }

        var serviceCollection = new ServiceCollection();
        serviceCollection.AddTransient<NpgsqlConnection>((provider) => new NpgsqlConnection(GetConnectionString()));
        serviceCollection.AddTransient<ISampleService, SampleService.SampleService>();
        _serviceProvider = serviceCollection.BuildServiceProvider();
        return _serviceProvider;
    }
}

public class PostgreSqlSampleUnitTest : PostgreSqlUnitTestBase_Test
{
    private FileInfo file = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sample", "Item", "ProvinceCity.xlsx"));

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
                var sql = $@"
                    INSERT INTO citys 
                        (name, province_id, company_id, created_by, created_date, is_draft_record) 
                        VALUES ('Jakarta', '1', '1', 'system', '2025-01-01 00:00:00.000 +0700', '0')";
                var pgCon = new NpgsqlConnection(GetConnectionString());
                await pgCon.OpenAsync();
                var cmd = new NpgsqlCommand(sql, pgCon);
                await cmd.ExecuteNonQueryAsync();
            }
        );
    }

    // [Fact]
    // public async Task CompareDataAndError_Invalid_Test()
    // {
    //     await RunTestAsync(
    //         file,
    //         "InitSubmitInvalidData",
    //         "InitSubmitInvalidData",
    //         () => GetServiceProvider(),
    //         async (IServiceProvider provider, UTContext utContext) =>
    //         {
    //             // CALL SERVICE HERE
    //             var sampleService = provider.GetRequiredService<ISampleService>();
    //             await sampleService.DoSomethingAsync("error_2");

    //             // SET ERROR TO UT CONTEXT
    //             utContext.ErrorMessages = sampleService.GetErrors();
    //         }
    //     );
    // }

    // [Fact]
    // public async Task CompareDataAndError_Invalid_ErrorActNotExistOnExp_Test()
    // {
    //     await Assert.ThrowsAsync<Exception>(async () =>
    //     {
    //         await RunTestAsync(
    //             file,
    //             "InitSubmitInvalidData",
    //             "InitSubmitInvalidData",
    //             () => GetServiceProvider(),
    //             async (IServiceProvider provider, UTContext utContext) =>
    //             {
    //                 // CALL SERVICE HERE
    //                 var sampleService = provider.GetRequiredService<ISampleService>();
    //                 await sampleService.DoSomethingAsync("error_3");

    //                 // SET ERROR TO UT CONTEXT
    //                 utContext.ErrorMessages = sampleService.GetErrors();
    //             }
    //         );
    //     });
    // }

    // [Fact]
    // public async Task CompareDataAndError_Invalid_ErrorExpNotExistOnAct_Test()
    // {
    //     await Assert.ThrowsAsync<Exception>(async () =>
    //     {
    //         await RunTestAsync(
    //             file,
    //             "InitSubmitInvalidData",
    //             "InitSubmitInvalidData",
    //             () => GetServiceProvider(),
    //             async (IServiceProvider provider, UTContext utContext) =>
    //             {
    //                 // CALL SERVICE HERE
    //                 var sampleService = provider.GetRequiredService<ISampleService>();
    //                 await sampleService.DoSomethingAsync("error_1");

    //                 // SET ERROR TO UT CONTEXT
    //                 utContext.ErrorMessages = sampleService.GetErrors();
    //             }
    //         );
    //     });
    // }

    // [Fact]
    // public async Task CompareDataOnly_WithProvider_Valid_AutoSelectExpected_Test()
    // {
    //     await RunTestAsync(
    //         file,
    //         "InitSubmitValidData",
    //         () => GetServiceProvider(),
    //         async (IServiceProvider provider, UTContext context) =>
    //         {
    //             var sampleService = provider.GetRequiredService<ISampleService>();
    //             await sampleService.DoSomethingAsync("no error");
    //         }
    //     );
    // }

    // [Fact]
    // public async Task CompareDataOnly_WithProvider_Invalid_AutoSelectExpected_Test()
    // {
    //     await RunTestAsync(
    //         file,
    //         "InitSubmitInvalidData",
    //         () => GetServiceProvider(),
    //         async (IServiceProvider provider, UTContext context) =>
    //         {
    //             // CALL SERVICE HERE
    //             var sampleService = provider.GetRequiredService<ISampleService>();
    //             await sampleService.DoSomethingAsync("error_2");

    //             // SET ERROR TO UT CONTEXT
    //             context.ErrorMessages = sampleService.GetErrors();
    //         }
    //     );
    // }
}
