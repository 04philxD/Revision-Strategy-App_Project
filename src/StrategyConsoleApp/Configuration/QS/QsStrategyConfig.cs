using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class QsStrategyConfig
{
    public string TxtTargetPath { get; set; } = null!;
    public HideColumns HideColumnsByName { get; set; } = null!;
    public FileRecipients FileRecipients { get; set; } = null!;
    public List<string> GeneralColumnsToHide { get; set; } = null!;
    public string HeaderRowStyleName { get; set; } = null!;
    public List<string> DataValidation { get; set; } = null!;
    public List<string> SpecialNames {  get; set; } = null!;

}
