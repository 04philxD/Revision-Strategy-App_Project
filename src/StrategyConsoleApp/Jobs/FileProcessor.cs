using Microsoft.Extensions.DependencyInjection;
using StrategyConsoleApp.Strategies.QS;
using StrategyConsoleApp.Strategies.Off;
using StrategyConsoleApp.Strategies.Aufb;
using System.Diagnostics;

namespace StrategyConsoleApp.Jobs;

public class FileProcessor : IFileProcessor
{
    private readonly FileProcessorConfig _configuration;
    private readonly IServiceProvider _serviceProvider;
    private readonly Stopwatch _stopwatch = new();

    public FileProcessor(FileProcessorConfig configuration, IServiceProvider serviceProvider)
    {
        _configuration = configuration;
        _serviceProvider = serviceProvider;
    }

    public async Task ProcessFilesAsync()
    {
        _stopwatch.Start();
        string[] files = Directory.GetFiles(_configuration.WorkingPath, "*.txt");

        foreach(string strategy in _configuration.Strategies)
        {
            (IStrategy StrgProc, IEnumerable<string> StrgFiles) strg = GetStrategy(strategy, files);
             await strg.StrgProc.ProcessAsync(strg.StrgFiles);
        }

        _stopwatch.Stop();
        Console.WriteLine($"Processing time: {_stopwatch.Elapsed}");
    }

    private (IStrategy StrgProc, IEnumerable<string> StringFiles) GetStrategy(string strategy, IEnumerable<string> files)
    {
        switch (strategy)
        {
            case "Aufb":
                return (_serviceProvider.GetRequiredService<AufbStrategy>(), files.Where(f => f.Contains(strategy, StringComparison.CurrentCultureIgnoreCase)).ToList());
            case "QS":
                var test = files.Where(f => f.Contains(strategy, StringComparison.CurrentCultureIgnoreCase) && !f.Contains("Aufb", StringComparison.CurrentCultureIgnoreCase) && !f.Contains("Off", StringComparison.CurrentCultureIgnoreCase)).ToList();
                return (_serviceProvider.GetRequiredService<QsStrategy>(), test);
            case "Off":
                return (_serviceProvider.GetRequiredService<OffStrategy>(), files.Where(f => f.Contains(strategy, StringComparison.CurrentCultureIgnoreCase)).ToList());
            default:
                break;
        }

        return (null!, null!);
    }
}
