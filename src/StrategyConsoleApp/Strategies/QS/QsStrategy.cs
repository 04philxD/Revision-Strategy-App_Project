using System.Collections.Concurrent;
using StrategyConsoleApp.Jobs;
using StrategyConsoleApp.Strategies.Interfaces;

namespace StrategyConsoleApp.Strategies.QS;

public class QsStrategy : AStrategy
{
    private readonly QsStrategyConfig _config;

    public QsStrategy(FileProcessorConfig generalConfig, QsStrategyConfig configuration, QSTxtFileReader txtFileReader, QSToExcelProcessor processToExcel) : base(generalConfig, txtFileReader, processToExcel)
    {
        _config = configuration;
    }

    public override async Task ProcessAsync(IEnumerable<string> files)
    {
        if (!files.Any())
            return;

        foreach (string file in files.ToList())
        {
            await Task.Run(() => _txtFileReader.ReadFile(file, _config.TxtTargetPath, _departmentsData));
        }

        _filesRecipient = _toExcelProcessor.ProcessToExcel(_generalConfig.WorkingPath, _departmentsData);
        

        foreach (KeyValuePair<string, string> team in _filesRecipient)
        {
            bool containsTeam = _config.FileRecipients.Any(r => r.Team == team.Key);

            if (containsTeam)
            {
                string? recipient = _config.FileRecipients.FirstOrDefault(r => r.Team == team.Key).TargetPath;
                await Task.Run(() => _toExcelProcessor.SaveFileToRecipientPath(team.Value, recipient)) ;
            }
        }

        foreach (string file in files.ToList())
        {
            await _txtFileReader.MoveTxtToProcessed(file, _config.TxtTargetPath);
        }
    }
}
