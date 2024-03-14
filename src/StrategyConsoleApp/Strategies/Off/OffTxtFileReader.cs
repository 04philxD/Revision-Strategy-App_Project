using System.Collections.Concurrent;
using System.Data;
using Microsoft.Extensions.Configuration;
using StrategyConsoleApp.Strategies;

public class OffTxtFileReader : ATxtFileReader
{
    private readonly OffStrategyConfig _config;

    public OffTxtFileReader(OffStrategyConfig config)
    {
        _config = config;
    }

    public override void ReadFile(string file, string targetPath, ConcurrentDictionary<string, List<DataRow>> offDepartments)
    {
        try
        {
            Console.WriteLine($"{file} is now being read");
            try
            {
                Directory.CreateDirectory(targetPath);
                string[] lines = File.ReadAllLines(file, System.Text.Encoding.Latin1);

                AddColumsToDataTable(_config.OffTxtFields);
                LoadData(lines);
                FillDepartments(offDepartments);
            }

            catch (FileLoadException ex)
            {
                Console.WriteLine($"Error processing the file: {file}\n{ex.Message}");
            }
        }
        catch (DirectoryNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void AddColumsToDataTable(OffTxtFields columns)
    {
        DataTable columnNames = new();
        foreach (OffField column in columns)
        {
            columnNames.Columns.Add(new DataColumn(column.Field));
        }

        _rawDataTable = columnNames.Clone();
    }

    private void LoadData(string[] lines)
    {
        foreach (string? line in lines)
        {
            List<string> values = new();
            string trimmedLine = line.Split(' ').Aggregate((a, b) => $"{a}{b}");

            int start = 0;

            foreach (OffField field in _config.OffTxtFields)
            {
                int end = start + field.Length;
                values.Add(trimmedLine.Substring(start, field.Length));
                start = end;
            }

            _rawDataTable.LoadDataRow(values.ToArray(), true);
        }
    }
}

