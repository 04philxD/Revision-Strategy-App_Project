using System.Collections.Concurrent;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using StrategyConsoleApp.Strategies;

public class OffToExcelProcessor: AToExcelProcessor
{
    private readonly OffStrategyConfig _config;
    private Dictionary<string, OffField> _mappedHeaderRow = new();

    public OffToExcelProcessor(OffStrategyConfig config)
    {
        _config = config;
    }

    public override ConcurrentDictionary<string, string> ProcessToExcel(string workingPath, ConcurrentDictionary<string, List<DataRow>> departsData)
    {
        string dirRoot = Path.Combine(workingPath, DateTime.Now.ToString("yyyy"));
        string dir = Path.Combine(dirRoot, date, "Offene-Referenzen");

        try
        {
            Console.WriteLine("Processing to excel starts...");
            Directory.CreateDirectory(dir);

            foreach(KeyValuePair<string, List<DataRow>> pair in departsData)
            {
                ProcessDepartment(pair, dir);
            }
        }

        catch (KeyNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
        }

        catch(Exception ex)
        {
            Console.WriteLine(ex.Message + Environment.NewLine);
            throw;
        }

        return _result;
    }

    public override void ProcessDepartment(KeyValuePair<string, List<DataRow>> pair, string dir, DataRow columnNames = null!)
    {
        IWorkbook wb = new XSSFWorkbook();
        string file = pair.Key.Contains("Rs", StringComparison.CurrentCultureIgnoreCase)
            ? _config.OffRsExcelToFill : _config.OffExcelToFill;

        destinationPath = Path.Combine(dir, $"{pair.Key}_offene-Referenzen.xlsx");

        CopySheetToWorkbook(wb, pair.Key, file);
        OffExcelSheetFill(pair, wb.GetSheet(pair.Key), wb);

        SaveWorkbook(wb, destinationPath);

        _result.TryAdd(pair.Key, destinationPath);
    }

    private void CopySheetToWorkbook(IWorkbook wb, string pairKey, string filePath)
    {
        ISheet sheet = wb.CreateSheet(pairKey);
        ISheet existingSheet = GetSheetFromExcel(filePath);

        for (int rowIndex = 0; rowIndex <= existingSheet.LastRowNum; rowIndex++)
        {
            IRow existingRow = existingSheet.GetRow(rowIndex);
            IRow newRow = sheet.CreateRow(rowIndex);

            if (existingRow != null)
            {
                for (int cellIndex = 0; cellIndex < existingRow.LastCellNum; cellIndex++)
                {
                    CopyCellToWorkbookRow(existingRow, cellIndex, wb, sheet, newRow, existingSheet.GetColumnWidth(cellIndex));
                }
            }
        }
    }

    private void CopyCellToWorkbookRow(IRow existingRow, int cellIndex, IWorkbook wb, ISheet sheet, IRow newRow, int columnWidth)
    {
        ICell existingCell = existingRow.GetCell(cellIndex);
        ICell newCell = newRow.CreateCell(cellIndex);

        if (existingCell != null)
        {
            newCell.SetCellValue(existingCell.StringCellValue);

            ICellStyle existingCellStyle = existingCell.CellStyle;
            ICellStyle newCellStyle = wb.CreateCellStyle();
            newCellStyle.CloneStyleFrom(existingCellStyle);
            newCell.CellStyle = newCellStyle;

            sheet.SetColumnWidth(cellIndex, columnWidth);
        }
    }

    private void OffExcelSheetFill(KeyValuePair<string, List<DataRow>> pair, ISheet sheet, IWorkbook workbook)
    {
        _mappedHeaderRow = MapHeaderRowToConfig(sheet);
        ICellStyle style = GetCellStyleFromExcel(workbook);
        string? columnNameForIndex = GetColumnNameForValidation(_mappedHeaderRow);

        int dataRowIndex = 1;
        foreach(DataRow row in pair.Value)
        {
            IRow dataRowInExcel = sheet.GetRow(dataRowIndex) ?? sheet.CreateRow(dataRowIndex);
            WriteDataToRow(dataRowInExcel, row, _mappedHeaderRow, style);
            dataRowIndex++;
        }

        if (dataRowIndex < 17) //default number of created rows (source: _config.ExcelFileToFill)
            dataRowIndex = 17;

        string? signatureField = _mappedHeaderRow.Where(s => s.Value.Signature).Select(f => f.Value.Field).FirstOrDefault();
        string? timeStampField = _mappedHeaderRow.Where(s => s.Value.TimeStamp).Select(f => f.Value.Field).FirstOrDefault();

        AddDataValidation(sheet
            ,workbook
            ,dataRowIndex
            ,GetColumnIndex(columnNameForIndex, _mappedHeaderRow)
            ,_config.DataValidation);
        AddRevisionValidationFields(workbook
            ,dataRowIndex
            ,pair.Key
            ,signatureField
            ,timeStampField
            ,GetColumnIndex(signatureField, _mappedHeaderRow)
            ,GetColumnIndex(timeStampField, _mappedHeaderRow));
    }

    private void WriteDataToRow(IRow dataRowInExcel, DataRow row, Dictionary<string, OffField> mappedHeaderRow, ICellStyle style)
    {
        foreach (KeyValuePair<string, OffField> mappedHeaderField in mappedHeaderRow)
        {
            WriteDataToCell(dataRowInExcel, mappedHeaderField, row, style);
        }
    }

    private void WriteDataToCell(IRow dataRowInExcel, KeyValuePair<string, OffField> mappedHeaderField, DataRow row, ICellStyle style)
    {
        ICell cell = dataRowInExcel.GetCell(GetColumnIndex(mappedHeaderField.Key, _mappedHeaderRow))
            ?? dataRowInExcel.CreateCell(GetColumnIndex(mappedHeaderField.Key, _mappedHeaderRow));

        cell.SetCellValue(row[mappedHeaderField.Key].ToString());
        cell.CellStyle = style;
    }

    private ISheet GetSheetFromExcel(string pathToExcel)
    {
        try
        {
            using (FileStream file = new FileStream(pathToExcel, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(file);

                return workbook.GetSheet(_config.OffExistingExcelSheetName);
            }
        }

        catch (IOException ex)
        {
            Console.WriteLine($"Error accessing file: {ex.Message}");
            throw;
        }
    }

    private Dictionary<string, OffField> MapHeaderRowToConfig(ISheet sheet)
    {
        Dictionary<string, OffField> headerRowMap = new();

        var fieldsToUse = _config.OffTxtFields.Where(f => f.ToUse).ToDictionary(f => f.Field);
        IRow headerRow = sheet.GetRow(0);

        if(headerRow != null)
        {
            foreach(ICell cell in headerRow.Cells)
            {
                string cellValue = cell.StringCellValue.Trim().ToString();
                if(cellValue != null && fieldsToUse.TryGetValue(cellValue, out OffField field))
                {
                    headerRowMap[cellValue] = field;
                }
            }
        }

        return headerRowMap;
    }

    private int GetColumnIndex(string? columnName, Dictionary<string, OffField> headerRowDic)
    {
        try
        {
            int columnIndex = headerRowDic.ToList().FindIndex(f => f.Key == columnName);

            return columnIndex;
        }

        catch (NullReferenceException ex)
        {
            Console.WriteLine($"Column name can't be null: {ex.Message}");
            throw;
        }
    }

    private ICellStyle GetCellStyleFromExcel(IWorkbook workbook)
    {
        ISheet existingSheet = GetSheetFromExcel(_config.OffExcelToFill);
        IRow existingRow = existingSheet.GetRow(1);
        ICell cell = existingRow.GetCell(1);
        ICellStyle existingCellStyle = cell.CellStyle;

        ICellStyle newCellStyle = workbook.CreateCellStyle();
        newCellStyle.CloneStyleFrom(existingCellStyle);

        return newCellStyle;
    }

    private string? GetColumnNameForValidation(Dictionary<string, OffField> mappedHeaderRow)
    {
        return mappedHeaderRow.Values.FirstOrDefault(x => x.Validation)?.Field;
    }
}
