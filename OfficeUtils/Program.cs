using Serilog;

namespace OfficeUtils
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Log.Information("=== 程序启动 ===");

                // To customize application configuration such as set high DPI settings or default font,
                // see https://aka.ms/applicationconfiguration.
                // initialize logger
                AppLogger.Init();
                ApplicationConfiguration.Initialize();
                Application.Run(new MainForm());

                Log.Information("=== 程序正常退出 ===");
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "程序崩溃");
            }
            finally
            {
                // 必须写，否则日志可能丢
                Log.CloseAndFlush();
            }
        }
    }
}