using System;
using System.IO;
using System.Windows.Forms;
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
            catch
            {
                // ignore
            }

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

        // Attach RichTextBox sink
        public static void AttachRichTextBox(RichTextBox rtb)
        {
            if (rtb == null) return;
            var formatter = new MessageTemplateTextFormatter("{Timestamp:HH:mm:ss} [{Level}] {Message}{NewLine}{Exception}", null);
            var richSink = new RichTextBoxSink(rtb, formatter);

            Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Debug()
                .WriteTo.File(Path.Combine(AppContext.BaseDirectory, "logs", "officeutils-.log"), rollingInterval: RollingInterval.Day)
                .WriteTo.Sink(richSink)
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

        // Lightweight RichTextBox sink implementation
        private class RichTextBoxSink : ILogEventSink
        {
            private readonly RichTextBox _rtb;
            private readonly MessageTemplateTextFormatter _formatter;

            public RichTextBoxSink(RichTextBox rtb, MessageTemplateTextFormatter formatter)
            {
                _rtb = rtb;
                _formatter = formatter;
            }

            public void Emit(LogEvent logEvent)
            {
                try
                {
                    using var sw = new StringWriter();
                    _formatter.Format(logEvent, sw);
                    var text = sw.ToString();
                    System.Drawing.Color c = System.Drawing.Color.Black;
                    switch (logEvent.Level)
                    {
                        case LogEventLevel.Debug:
                            c = System.Drawing.Color.Gray; break;
                        case LogEventLevel.Information:
                            c = System.Drawing.Color.Black; break;
                        case LogEventLevel.Warning:
                            c = System.Drawing.Color.Orange; break;
                        case LogEventLevel.Error:
                            c = System.Drawing.Color.Red; break;
                        case LogEventLevel.Fatal:
                            c = System.Drawing.Color.DarkRed; break;
                    }
                    if (_rtb.IsDisposed) return;
                    if (_rtb.InvokeRequired)
                    {
                        _rtb.Invoke(new Action(() => AppendTextColored(_rtb, text, c)));
                    }
                    else
                    {
                        AppendTextColored(_rtb, text, c);
                    }
                }
                catch
                {
                    // ignore
                }
            }

            private static void AppendTextColored(RichTextBox rtb, string text, System.Drawing.Color color)
            {
                rtb.SelectionStart = rtb.TextLength;
                rtb.SelectionColor = color;
                rtb.AppendText(text);
                rtb.SelectionColor = rtb.ForeColor;
                rtb.SelectionStart = rtb.TextLength;
                rtb.ScrollToCaret();
            }
        }
    }
}
