using ExcelRefineAddIn.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRefineAddIn.Catalog
{
    public static class CharsetCatalog
    {
        public static readonly IReadOnlyList<CharsetOption> CharsetOptions = new List<CharsetOption>
        {
            new CharsetOption { DisplayName = "Shift_JIS", Encoding = Encoding.GetEncoding("shift_jis") },
            new CharsetOption { DisplayName = "UTF-8", Encoding = Encoding.UTF8 },
        };

        static CharsetCatalog()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
    }
}
