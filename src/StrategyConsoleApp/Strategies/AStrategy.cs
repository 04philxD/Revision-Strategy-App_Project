using System.Collections.Concurrent;
using System.Data;
using StrategyConsoleApp.Jobs;
using StrategyConsoleApp.Strategies.Interfaces;

namespace StrategyConsoleApp.Strategies;

public abstract class AStrategy : IStrategy
{
    protected FileProcessorConfig _generalConfig;
    protected readonly ITxtFileReader _txtFileReader;
    protected readonly ConcurrentDictionary<string, List<DataRow>> _departmentsData = [];
    protected readonly IProcessToExcel _toExcelProcessor;
    protected ConcurrentDictionary<string, string> _filesRecipient = [];

    public AStrategy(FileProcessorConfig generalConfig, ITxtFileReader txtFileReader, IProcessToExcel processToExcel)
    {
        _generalConfig = generalConfig;
        _txtFileReader = txtFileReader;
        _toExcelProcessor = processToExcel;
    }

    public abstract Task ProcessAsync(IEnumerable<string> files);
}
