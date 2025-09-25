using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRefineAddIn.Service
{
    public class CsvService
    {
        private static readonly Lazy<CsvService> _instance = new Lazy<CsvService>(() => new CsvService());

        public static CsvService Instance => _instance.Value;

        private CsvService() {}

        public void save(List<List<object>> rows, String fullpath, String delimiter, String newLine, Encoding encoding)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
                NewLine = newLine,
                ShouldQuote = args =>
                    args.Field.Contains('\n') || args.Field.Contains('\r') || args.Field.Contains(delimiter) || args.Field.Contains('"')
            };

            using (var writer = new StreamWriter(fullpath, false, encoding))
            using (var csv = new CsvWriter(writer, config))
            {
                foreach (var row in rows)
                {
                    foreach (var cell in row)
                    {
                        csv.WriteField(cell);
                    }
                    csv.NextRecord();
                }
            }
        }
    }
}
