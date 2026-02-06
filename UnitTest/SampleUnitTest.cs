using System;
using Microsoft.Extensions.DependencyInjection;
using UnitTest.SampleService;
using UTDataValidator;

namespace UnitTest;

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
                // CALL SERVICE HERE
                var sampleService = provider.GetRequiredService<ISampleService>();
                await sampleService.DoSomethingAsync("error_2");

                // SET ERROR TO UT CONTEXT
                utContext.ErrorMessages = sampleService.GetErrors();
            }
        );
    }

    [Fact]
    public async Task CompareDataAndError_Invalid_ErrorActNotExistOnExp_Test()
    {
        await Assert.ThrowsAsync<Exception>(async () =>
        {
            await RunTestAsync(
                file,
                "InitSubmitInvalidData",
                "InitSubmitInvalidData",
                () => GetServiceProvider(),
                async (IServiceProvider provider, UTContext utContext) =>
                {
                    // CALL SERVICE HERE
                    var sampleService = provider.GetRequiredService<ISampleService>();
                    await sampleService.DoSomethingAsync("error_3");

                    // SET ERROR TO UT CONTEXT
                    utContext.ErrorMessages = sampleService.GetErrors();
                }
            );
        });
    }

    [Fact]
    public async Task CompareDataAndError_Invalid_ErrorExpNotExistOnAct_Test()
    {
        await Assert.ThrowsAsync<Exception>(async () =>
        {
            await RunTestAsync(
                file,
                "InitSubmitInvalidData",
                "InitSubmitInvalidData",
                () => GetServiceProvider(),
                async (IServiceProvider provider, UTContext utContext) =>
                {
                    // CALL SERVICE HERE
                    var sampleService = provider.GetRequiredService<ISampleService>();
                    await sampleService.DoSomethingAsync("error_1");

                    // SET ERROR TO UT CONTEXT
                    utContext.ErrorMessages = sampleService.GetErrors();
                }
            );
        });
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
                // CALL SERVICE HERE
                var sampleService = provider.GetRequiredService<ISampleService>();
                await sampleService.DoSomethingAsync("error_2");

                // SET ERROR TO UT CONTEXT
                context.ErrorMessages = sampleService.GetErrors();
            }
        );
    }

    protected override string GetConnectionString()
    {
        throw new NotImplementedException();
    }
}
