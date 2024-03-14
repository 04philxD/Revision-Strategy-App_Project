using StrategyConsoleApp.Jobs;
using StrategyConsoleApp.Strategies.Interfaces;

namespace StrategyConsoleApp.Strategies.Off;

public class OffStrategy : AStrategy
{
    private readonly OffStrategyConfig _config;

    public OffStrategy(FileProcessorConfig generalConfig, OffStrategyConfig config, OffTxtFileReader txtFileReader, OffToExcelProcessor processToExcel) : base(generalConfig, txtFileReader, processToExcel)
    {
        _config = config;
    }

    public override async Task ProcessAsync(IEnumerable<string> files)
    {
        if(!files.Any()) return;

        foreach (string file in files)
        {
            await Task.Run(() => _txtFileReader.ReadFile(file, _config.OffTxtTargetpath, _departmentsData));
        }
        
         _filesRecipient = _toExcelProcessor.ProcessToExcel(_generalConfig.WorkingPath, _departmentsData);

        foreach(KeyValuePair<string, string> pair in _filesRecipient)
        {
            bool containsTeam = _config.FileRecipients.Any(r => r.Team == pair.Key);

            if (containsTeam)
            {
                string recipient = _config.FileRecipients.FirstOrDefault(r => r.Team == pair.Key).TargetPath;
                await Task.Run(() => _toExcelProcessor.SaveFileToRecipientPath(pair.Value, recipient));
            }
        }

        foreach (string file in files.ToList())
        {
            await _txtFileReader.MoveTxtToProcessed(file, _config.OffTxtTargetpath);
        }
    }
}
