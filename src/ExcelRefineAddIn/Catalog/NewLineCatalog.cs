using ExcelRefineAddIn.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRefineAddIn.Catalog
{
    public static class NewLineCatalog
    {
        public static readonly IReadOnlyList<NewLineOption> NewLineOptions = new List<NewLineOption>
        {
            new NewLineOption { DisplayName = "\\r\\n (Windows)", NewLine = "\r\n" },
            new NewLineOption { DisplayName = "\\n (Linux)", NewLine = "\n" },
        };
    }
}
