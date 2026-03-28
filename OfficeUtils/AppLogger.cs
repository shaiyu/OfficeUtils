using System;
using System.IO;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using Serilog.Formatting.Display;

namespace OfficeUtils
{
    // Simple application logger wrapper using Serilog
    public static class AppLogger
    {
        public static ILogger? Logger { get; private set; }

        public static void Init()
        {
            var logDir = Path.Combine(AppContext.BaseDirectory, "logs");
            try
            {
                Directory.CreateDirectory(logDir);
            }
            catch { }

            Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Debug()
                .WriteTo.File(Path.Combine(logDir, "officeutils-.log"), rollingInterval: RollingInterval.Day)
                .CreateLogger();
        }

        // Attach a UI sink that writes formatted messages to an action (e.g., append to a control)
        public static void AttachUi(Action<string> append)
        {
            if (append == null) return;
            var formatter = new MessageTemplateTextFormatter("{Timestamp:HH:mm:ss} [{Level}] {Message}{NewLine}{Exception}", null);
            var uiSink = new DelegatingSink(append, formatter);

            Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Debug()
                .WriteTo.File(Path.Combine(AppContext.BaseDirectory, "logs", "officeutils-.log"), rollingInterval: RollingInterval.Day)
                .WriteTo.Sink(uiSink)
                .CreateLogger();
        }

        private class DelegatingSink : ILogEventSink
        {
            private readonly Action<string> _append;
            private readonly MessageTemplateTextFormatter _formatter;

            public DelegatingSink(Action<string> append, MessageTemplateTextFormatter formatter)
            {
                _append = append;
                _formatter = formatter;
            }

            public void Emit(LogEvent logEvent)
            {
                try
                {
                    using (var sw = new StringWriter())
                    {
                        _formatter.Format(logEvent, sw);
                        _append(sw.ToString());
                    }
                }
                catch
                {
                    // swallow UI sink errors
                }
            }
        }
    }
}
