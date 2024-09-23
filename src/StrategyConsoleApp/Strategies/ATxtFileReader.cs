using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StrategyConsoleApp.Strategies
{
    public abstract class ATxtFileReader : ITxtFileReader
    {
        protected DataTable _rawDataTable = new();

        public abstract void ReadFile(string file, string targetPath, ConcurrentDictionary<string, List<DataRow>> departments);

        public async Task MoveTxtToProcessed(string file, string targetPath)
        {
            var sourceFileInfo = new FileInfo(file);
            string destinationFilePath = Path.Combine(targetPath, sourceFileInfo.Name);

            await Task.Run(() => File.Move(file, destinationFilePath));
        }

        public void FillDepartments(ConcurrentDictionary<string, List<DataRow>> departments, List<string> specialNames = null)
        {
            EnumerableRowCollection<DataRow> dt = _rawDataTable.AsEnumerable();

            IEnumerable<object> dept = dt.Select(rw => rw[0]).Distinct();
            foreach (string depp in dept.Cast<string>())
            {
                EnumerableRowCollection<DataRow> deptData = dt.Where(d => d[0].ToString() == depp);
                IEnumerable<object> teams = deptData.Select(rw => rw[1]).Distinct();

                foreach (string team in teams)
                {
                    IEnumerable<DataRow> teamData = deptData.Where(t => t[1].ToString() == team).ToList();
                    string teamName = $"{depp}{team}";

                    // Überprüfen, ob spezielle Namen vorhanden sind
                    var specialRows = teamData.Where(row => specialNames != null && specialNames.Contains(row["name"].ToString())).ToList();

                    // Füge die speziellen Namen unter ihren eigenen Schlüsseln hinzu
                    foreach (DataRow row in specialRows)
                    {
                        string name = row[2].ToString(); // Ersetze "name" durch den tatsächlichen Spaltennamen
                        string specialKey = $"Special_{name}";

                        departments.AddOrUpdate(
                            specialKey, key => new List<DataRow> { row },
                            (key, oldValue) => {
                                if (!oldValue.Any(existingRow => existingRow.ItemArray.SequenceEqual(row.ItemArray)))
                                {
                                    oldValue.Add(row);
                                }
                                return oldValue;
                            }
                        );
                    }

                    // Füge die Team-Daten zum Dictionary hinzu, ohne die speziellen Namen
                    if (departments.TryGetValue(teamName, out var existingRows))
                    {
                        // Filtern der neuen Daten, um Duplikate zu vermeiden
                        var newRows = teamData.Where(newRow => !existingRows.Any(existingRow => existingRow.ItemArray.SequenceEqual(newRow.ItemArray))).ToList();
                        existingRows.AddRange(newRows);
                    }
                    else
                    {
                        // Wenn der Schlüssel nicht existiert, einfach hinzufügen
                        departments.TryAdd(teamName, teamData.Except(specialRows).ToList());
                    }
                }
            }
        }

    }
}
