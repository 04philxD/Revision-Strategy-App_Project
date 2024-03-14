using System.Collections.Concurrent;
using System.Data;

namespace StrategyConsoleApp.Strategies.QS
{
    public class QSTxtFileReader : ATxtFileReader
    {
        private readonly DataTable _columnNamesTable = new();
        private bool _isFirstFile = true;

        public override void ReadFile(string file, string targetPath, ConcurrentDictionary<string, List<DataRow>> departments)
        {
            try
            {
                Console.WriteLine($"Es wird nun die Datei {file} verarbeitet.");
                try
                {
                    Directory.CreateDirectory(targetPath);
                    string[] lines = File.ReadAllLines(file, System.Text.Encoding.Latin1);

                    if (!lines.Any())
                        throw new Exception();

                    if (_isFirstFile)
                    {
                        _isFirstFile = CreateColumnNames(lines[0]);
                        departments.TryAdd("columnNames", _columnNamesTable.AsEnumerable().ToList());
                    }

                    LoadData(lines, file);
                    Console.WriteLine($"File {Path.GetFileName(file)} succesfully processed.");

                    FillDepartments(departments);

                    //TestDictionary(departments)
                }
                catch
                {
                    Console.WriteLine($"Error processing file '{file}'");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void LoadData(string[] lines, string fileName)
        {
            if (!_rawDataTable.Columns.Contains("SourceFile"))
                _rawDataTable.Columns.Add("SourceFile");

            int i = 0;
            foreach (string? line in lines.Skip(1))
            {
                i++;
                string[] values = line.Split(";").Select(val => val.Trim()).ToArray();
                values = values.Concat(new[] { fileName }).ToArray();

                _rawDataTable.LoadDataRow(values, true);
                //Console.WriteLine($"{i} Zeile wurde gelesen.");
            }
        }

        private bool CreateColumnNames(string line)
        {
            bool namesCreated = false;
            IEnumerable<string> columnNames = line.Split(';').Select(col => col.Trim());

            foreach (string columnName in columnNames)
            {
                _columnNamesTable.Columns.Add(new DataColumn(columnName));
            }

            _rawDataTable = _columnNamesTable.Clone();
            _columnNamesTable.LoadDataRow(columnNames.ToArray(), true);

            return namesCreated;
        }
    }
}
