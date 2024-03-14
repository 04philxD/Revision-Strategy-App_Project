using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StrategyConsoleApp.Jobs;

public class FileProcessorConfig
{
    public string WorkingPath { get; set; } = string.Empty;

    public List<string> Strategies { get; set; } = [];
}
