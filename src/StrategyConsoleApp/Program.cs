using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using StrategyConsoleApp.Configuration.Aufb;
using StrategyConsoleApp.Jobs;
using StrategyConsoleApp.Strategies;
using StrategyConsoleApp.Strategies.Aufb;
using StrategyConsoleApp.Strategies.Interfaces;
using StrategyConsoleApp.Strategies.Off;
using StrategyConsoleApp.Strategies.QS;


Console.WriteLine("Program starts...");

IHost host = new HostBuilder()
.ConfigureAppConfiguration((hostContext, config) =>
{
    config.SetBasePath(Directory.GetCurrentDirectory());
    config.AddJsonFile("appsettings.json", optional: true);
})
.ConfigureServices((hostContext, services) =>
{
    IConfiguration config = hostContext.Configuration;
    services.AddSingleton(config);

    services.AddSingleton(config.GetSection("FileProcessorConfig").Get<FileProcessorConfig>()!);
    services.AddSingleton<IFileProcessor, FileProcessor>();

    services.AddSingleton(config.GetSection("QsStrategyConfig").Get<QsStrategyConfig>()!);
    services.AddTransient<QSTxtFileReader>();
    services.AddTransient<QSToExcelProcessor>();
    services.AddTransient<QsStrategy>();

    services.AddSingleton(config.GetSection("AufbStrategyConfig").Get<AufbStrategyConfig>()!);
    services.AddTransient<AufbStrategy>();

    services.AddSingleton(config.GetSection("OffStrategyConfig").Get<OffStrategyConfig>()!);
    services.AddTransient<OffTxtFileReader>();
    services.AddTransient< OffToExcelProcessor>();
    services.AddTransient<OffStrategy>();
})
.Build();

IServiceProvider serviceProvider = host.Services;
IFileProcessor fileProcessor = serviceProvider.GetRequiredService<IFileProcessor>();

await fileProcessor.ProcessFilesAsync();