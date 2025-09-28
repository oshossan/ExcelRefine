using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRefineAddIn.Service
{
    public class SerilogService
    {
        private static readonly Lazy<SerilogService> _instance = new Lazy<SerilogService>(() => new SerilogService());
        public static SerilogService Instance => _instance.Value;

        private SerilogService()
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelRefine.log");

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(logPath)
                .CreateLogger();
        }

        public void Error(string message, Exception ex = null)
        {
            if (ex == null)
            {
                Log.Error(message);
                return;
            }
            
            Log.Error(ex, message);
        }

        public void Debug(string message)
        {
            Log.Debug(message);
        }
    }
}
