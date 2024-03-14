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

        public void FillDepartments(ConcurrentDictionary<string, List<DataRow>> departments)
        {
            EnumerableRowCollection<DataRow> dt = _rawDataTable.AsEnumerable();

            IEnumerable<object> dept = dt.Select(rw => rw[0]).Distinct();
            foreach (string depp in dept.Cast<string>())
            {
                EnumerableRowCollection<DataRow> deptData = dt.Where(d => d[0].ToString() == depp);

                IEnumerable<object> teams = deptData.Select(rw => rw[1]).Distinct();

                foreach (string team in teams)
                {
                    IEnumerable<DataRow> teamData = deptData.Where(t => t[1].ToString() == team);
                    string teamName = $"{depp}{team}";

                    departments.AddOrUpdate(
                        teamName, key => teamData.ToList(),
                        (key, oldValue) => { oldValue.AddRange(teamData); return oldValue; }
                    );
                }
            }
        }
    }
}
