using System.Collections.Concurrent;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace StrategyConsoleApp.Strategies.QS
{
    public class QSToExcelProcessor : AToExcelProcessor
    {
        private readonly QsStrategyConfig _config;

        public QSToExcelProcessor(QsStrategyConfig config)
        {
            _config = config;
        }

        public override ConcurrentDictionary<string, string> ProcessToExcel(string path, ConcurrentDictionary<string, List<DataRow>> departments)
        {
            string? dirRoot = Directory.CreateDirectory(Path.Combine(path, DateTime.Now.ToString("yyyy"))).ToString();
            string? dir = Directory.CreateDirectory(Path.Combine(dirRoot, date, "QS")).ToString();

            try
            {
                Console.WriteLine("Processing to excel starts...");
                Directory.CreateDirectory(dir);

                DataRow columnNames = departments.Where(p => p.Key == _config.HeaderRowStyleName).SelectMany(x => x.Value).First();

                foreach (KeyValuePair<string, List<DataRow>> pair in departments.Where(a => a.Key != _config.HeaderRowStyleName))
                {
                    ProcessDepartment(pair, dir, columnNames);
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return _result;
        }

        public override void ProcessDepartment(KeyValuePair<string, List<DataRow>> pair, string dir, DataRow columnNames)
        {
            IWorkbook workbook = new XSSFWorkbook();
            Console.WriteLine(workbook);
            List<string> fileNames = pair.Value
                .Where(row => row.Table.Columns.Contains("SourceFile"))
                .Select(row => row["SourceFile"].ToString())
                .Distinct()
                .ToList()!;

            foreach (string fileName in fileNames)
            {
                string[] splitt = fileName.Split('_');
                string? fileNameSplit = splitt[1];

                ISheet sheet = workbook.CreateSheet(fileNameSplit);

                DataRow[] rowsForThisFile = pair.Value
                    .Where(row => row.Table.Columns.Contains("SourceFile") && row["SourceFile"].ToString() == fileName)
                    .ToArray();

                WriteHeaderRow(sheet, columnNames, workbook);
                WriteDataRows(sheet, rowsForThisFile, workbook, fileNameSplit);

                int numberOfColumns = sheet.GetRow(0).Cells.Count - 1;
                for (int i = 0; i <= numberOfColumns; i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                AddDataValidation(sheet, workbook, rowsForThisFile.Length, sheet.GetRow(0).Cells.Count - 1, _config.DataValidation);
                HideColumnsByName(sheet, fileNameSplit, _config.HideColumnsByName, _config.GeneralColumnsToHide);
            }

            destinationPath = Path.Combine(dir, $"QS_{pair.Key}_{date}.xlsx");
            SaveWorkbook(workbook, destinationPath);

            _result.TryAdd(pair.Key, destinationPath);
        }

        private void WriteHeaderRow(ISheet sheet, DataRow columnNamesRow, IWorkbook workbook)
        {
            XSSFCellStyle cellStyle = GetCellStyle(workbook, _config.HeaderRowStyleName);

            CreateRow(sheet, columnNamesRow, 0, cellStyle);
            CreateRevisionCell(sheet, columnNamesRow, cellStyle);
        }
        
        private void CreateRevisionCell(ISheet sheet, DataRow columnNamesRow, ICellStyle cellStyle)
        {
            int revisionColumn = sheet.GetRow(0).Cells.Count;

            sheet.GetRow(0).CreateCell(revisionColumn).SetCellValue("-REVISION-");
            sheet.GetRow(0).GetCell(revisionColumn).CellStyle = cellStyle;
            sheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, columnNamesRow.ItemArray.Length - 1));
        }

        private void WriteDataRows(ISheet sheet, IEnumerable<DataRow> dataRows, IWorkbook workbook, string key)
        {
            int rowNumber = 1;
            var colorStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            foreach (DataRow dataRow in dataRows)
            {
                colorStyle = GetCellStyle(workbook, dataRow["Sourcefile"].ToString());
                CreateRow(sheet, dataRow, rowNumber++, colorStyle);
            }

            int signColumnIndex = sheet.GetRow(0).Cells.Count -2;
                string? signatureField = sheet.GetRow(0).GetCell(signColumnIndex).StringCellValue;

            int timeStampColumnIndex = sheet.GetRow(0).Cells.Count -1;
                string? timeStampField = sheet.GetRow(0).GetCell(timeStampColumnIndex).StringCellValue;

            AddRevisionValidationFields(workbook, rowNumber, key, signatureField, timeStampField, signColumnIndex, timeStampColumnIndex);
        }

        private void CreateRow(ISheet sheet, DataRow dataRow, int index, ICellStyle? style = null)
        {
            IWorkbook wb = new XSSFWorkbook();
            ICreationHelper factory = wb.GetCreationHelper();

            IRow row = sheet.CreateRow(index);
            for (int i = 0; i < dataRow.ItemArray.Length - 1; i++)
            {
                IRichTextString str = factory.CreateRichTextString(dataRow[i].ToString());
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(str);
                if (style != null)
                    cell.CellStyle = style;
            }
        }

        private void HideColumnsByName(ISheet sheet, string fileName, HideColumns addColumnsToHide, List<string> columnsToHide)
        {
            List<string> generalColumnsToHide = columnsToHide;

            var columnNames = addColumnsToHide.Where(r => r.SourceFileType == fileName).SelectMany(x => x.ColumnsToHide).ToList();

            generalColumnsToHide.AddRange(columnNames);

            IRow headerRow = sheet.GetRow(0);

            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                ICell cell = headerRow.GetCell(i);
                if (cell != null && generalColumnsToHide.Contains(cell.StringCellValue))
                    sheet.SetColumnHidden(i, true);
            }
        }
    }
}
