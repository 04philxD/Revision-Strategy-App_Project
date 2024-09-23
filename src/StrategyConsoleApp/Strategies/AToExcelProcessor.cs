using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using StrategyConsoleApp.Strategies.Interfaces;

namespace StrategyConsoleApp.Strategies
{
    public abstract class AToExcelProcessor: IProcessToExcel
    {
        protected readonly ConcurrentDictionary<string, string> _result = [];
        protected readonly string date = DateTime.Now.ToString("MMyy");
        //protected readonly string date = "0624";      //nur für die Nacherstellung von Monaten
        protected string destinationPath = string.Empty;

        public abstract ConcurrentDictionary<string, string> ProcessToExcel(string workingPath, ConcurrentDictionary<string, List<DataRow>> departmentsData);

        public abstract void ProcessDepartment(KeyValuePair<string, List<DataRow>> pair, string dir, DataRow columnNames = null!);

        public void AddDataValidation(ISheet sheet, IWorkbook workbook, int lastRow, int dataValidationColumn, List<string> dataValidation)
        {
            int maxLength = dataValidation.OrderByDescending(s => s.Length).First().Length;

            IDataValidationHelper dvHelper = sheet.GetDataValidationHelper();
            IDataValidationConstraint dvConstraint = dvHelper.CreateExplicitListConstraint(dataValidation.ToArray());

            CellRangeAddressList addressList = new(1, lastRow, dataValidationColumn, dataValidationColumn);
            IDataValidation validation = dvHelper.CreateValidation(dvConstraint, addressList);

            sheet.AddValidationData(validation);
            sheet.SetColumnWidth(dataValidationColumn, maxLength * 256);
        }

        public void AddRevisionValidationFields(IWorkbook workbook, int dataRowIndex, string key, string? signatureField, string? timeStampField, int signColumnIndex, int timeStampColumnIndex)
        {
            IRow signatureRow = workbook.GetSheet(key).CreateRow(dataRowIndex + 3);

            AddSignatureCell(workbook, signatureRow, signColumnIndex, signatureField);
            AddTimeStampCell(workbook, signatureRow, timeStampField, timeStampColumnIndex);
        }

        private void AddSignatureCell(IWorkbook workbook, IRow signatureRow, int signColumnIndex, string? signatureField)
        {
            AddCell(workbook, signatureRow, signColumnIndex, signatureField);
        }

        private void AddTimeStampCell(IWorkbook workbook, IRow signatureRow, string? timeStampField, int timeStampColumnIndex)
        {
            AddCell(workbook, signatureRow, timeStampColumnIndex, timeStampField, cell => cell.SetCellValue("=HEUTE()"));
        }

        private void AddCell(IWorkbook workbook, IRow row, int columnIndex, string? field, Action<ICell> additionalSetup = null!)
        {
            ICell cell = row.CreateCell(columnIndex);
            XSSFCellStyle cellStyle = GetCellStyle(workbook, field);
            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle;
            }

            additionalSetup?.Invoke(cell);
        }

        public XSSFColor GetColorForFile(string parameterName)
        {
            string[] split = parameterName.Split('_');
            if (split.Length < 2)
            {
                switch (parameterName)
                {
                    case "Prüfergebnis\n(Bitte auswählen)":
                        return new XSSFColor(new byte[] { 192, 192, 192 });
                    default:
                        return new XSSFColor(new byte[] {250, 250, 250});
                }
            }

            string fileType = split[1];

            return fileType switch
            {
                "PROVISION" => new XSSFColor(new byte[] { 255, 242, 204 }),
                "BUCHUNG" => new XSSFColor(new byte[] { 221, 235, 247 }),
                "ZAHL" => new XSSFColor(new byte[] { 226, 239, 218 }),
                _ => new XSSFColor(new byte[] { 255, 255, 255 }),
            };
        }

        public XSSFCellStyle GetCellStyle(IWorkbook workbook, string parameterName)
        {
            var style = (XSSFCellStyle)workbook.CreateCellStyle();

            switch (parameterName)
            {
                case $"Prüfergebnis\n(Bitte auswählen)":
                    style.SetFillForegroundColor(GetColorForFile(parameterName));
                    style.FillPattern = FillPattern.SolidForeground;
                    return style;
                case $"Kommentar \nnur bei Prüfergebnis \"03\"":
                    style.BorderBottom = BorderStyle.Thick;
                    style.BorderLeft = BorderStyle.Thick;
                    style.BorderRight = BorderStyle.Thick;
                    style.BorderTop = BorderStyle.Thick;
                    style.Alignment = HorizontalAlignment.Right;
                    return style;
                case "columnNames":
                    style.BorderBottom = BorderStyle.Thick;
                    style.Alignment = HorizontalAlignment.Center;
                    return style;
                default:
                    style.SetFillForegroundColor(GetColorForFile(parameterName));
                    style.FillPattern = FillPattern.SolidForeground;
                    style.BorderBottom = BorderStyle.Thin;
                    style.BorderLeft = BorderStyle.Thin;
                    style.BorderRight = BorderStyle.Thin;
                    style.BorderTop = BorderStyle.Thin;
                    return style;
            }
        }

        public void SaveWorkbook(IWorkbook wb, string destinationPath)
        {
            try
            {
                using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(fileStream);
                    Console.WriteLine($"Excel was saved to: {destinationPath}");
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error saving Excel file: {ex.Message}");
                throw;
            }
        }

        public async Task SaveFileToRecipientPath(string sourceFilePath, string recipientPath)
        {
            string fileName = Path.GetFileName(sourceFilePath);
            string targetPath = $"{recipientPath}\\{fileName}";
            try
            {
                if (System.IO.File.Exists(sourceFilePath))
                {
                    await Task.Run(() =>
                    {
                        Directory.CreateDirectory(recipientPath);
                        File.Copy(sourceFilePath, targetPath);

                        Console.WriteLine($"File Copied to {targetPath}");
                    });
                }
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine($"There is no matching recipient path: {ex.Message}");
            }
        }
    }
}
