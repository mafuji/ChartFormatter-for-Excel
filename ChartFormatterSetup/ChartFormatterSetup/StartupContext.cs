using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChartFormatterSetup
{
    public class StartupContext
    {
        public bool IsBomb { get; }
        public string? TargetPath { get; }

        public StartupContext(string[] args)
        {
            if (args.Length >= 2 && args[0].Equals("bomb", StringComparison.OrdinalIgnoreCase))
            {
                IsBomb = true;
                TargetPath = args[1];
            }
            else
            {
                IsBomb = false;
                TargetPath = null;
            }
        }

        public bool IsTargetMyFolder()
        {
            if (!IsBomb || TargetPath == null)
                return false;

            string normArg = Path.GetFullPath(TargetPath).TrimEnd('\\');
            string normStartup = Path.GetFullPath(Application.StartupPath).TrimEnd('\\');

            return string.Equals(normArg, normStartup, StringComparison.OrdinalIgnoreCase);
        }
    }
}
