using System.Collections.Concurrent;
using System.Data;

public interface ITxtFileReader
{
    void ReadFile(string file, string targetPath, ConcurrentDictionary<string, List<DataRow>> departments);
    Task MoveTxtToProcessed(string file, string targetPath);
}
