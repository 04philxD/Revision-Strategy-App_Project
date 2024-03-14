public class OffStrategyConfig
{
    public string OffTxtTargetpath { get; set; } = null!;
    public OffTxtFields OffTxtFields { get; set; } = null!;
    public string OffExcelToFill { get; set; } = null!;
    public string OffExistingExcelSheetName { get; set; } = null!;
    public FileRecipients FileRecipients { get; set; } = null!;
    public string OffRsExcelToFill { get; set; } = null!;
    public List<string> DataValidation { get; set; } = null!;
}
