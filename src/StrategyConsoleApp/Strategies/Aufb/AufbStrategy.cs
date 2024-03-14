using StrategyConsoleApp.Configuration.Aufb;
using StrategyConsoleApp.Jobs;
using StrategyConsoleApp.Strategies.QS;

namespace StrategyConsoleApp.Strategies.Aufb;

public class AufbStrategy : AStrategy
{
    private readonly AufbStrategyConfig _config;

    public AufbStrategy(FileProcessorConfig generalConfig, QSTxtFileReader reader, AufbStrategyConfig config, QSToExcelProcessor toExcelProcessor) : base(generalConfig, reader, toExcelProcessor)
    {
        _config = config;
    }

    public override async Task ProcessAsync(IEnumerable<string> files)
    {
        if(!Path.Exists(_config.WrongTextTargetPath))
        {
            Directory.CreateDirectory(_config.WrongTextTargetPath);
        }

        foreach (string file in files)
        {
            await _txtFileReader.MoveTxtToProcessed(file, _config.WrongTextTargetPath);
        }
    }   
}
