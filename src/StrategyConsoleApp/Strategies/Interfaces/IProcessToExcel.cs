using System.Collections.Concurrent;
using System.Data;
using NPOI.SS.UserModel;

namespace StrategyConsoleApp.Strategies.Interfaces
{
    public interface IProcessToExcel
    {
        ConcurrentDictionary<string, string> ProcessToExcel(string workingPath, ConcurrentDictionary<string, List<DataRow>> departmentsData);
        Task SaveFileToRecipientPath(string path, string recipient);
    }
}