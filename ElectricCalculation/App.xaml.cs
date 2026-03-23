using System.Configuration;
using System.Data;
using System.Windows;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using ElectricCalculation.Services;

namespace ElectricCalculation
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

            base.OnStartup(e);
        }

        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            LogException(e.Exception, "DispatcherUnhandledException");
            ShowCrashMessage(e.Exception);
            e.Handled = true;
        }

        private void OnDomainUnhandledException(object? sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception ?? new Exception("Unknown unhandled exception.");
            LogException(exception, "AppDomain.UnhandledException");

            try
            {
                Dispatcher.Invoke(() => ShowCrashMessage(exception));
            }
            catch
            {
                // ignored
            }
        }

        private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            LogException(e.Exception, "TaskScheduler.UnobservedTaskException");

            try
            {
                Dispatcher.Invoke(() => ShowCrashMessage(e.Exception));
            }
            catch
            {
                // ignored
            }

            e.SetObserved();
        }

        private static void ShowCrashMessage(Exception exception)
        {
            MessageBox.Show(
                $"Đã xảy ra lỗi, ứng dụng đã chặn để tránh bị crash.\n\n{exception.GetType().Name}: {exception.Message}\n\nChi tiết đã được ghi vào thư mục Logs trong Documents\\ElectricCalculation.",
                "Electric Calculation - Lỗi",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private static void LogException(Exception exception, string source)
        {
            try
            {
                var saveRoot = SaveGameService.GetSaveRootDirectory(); // Documents\\ElectricCalculation\\Saves
                var appRoot = Path.GetFullPath(Path.Combine(saveRoot, ".."));
                var logDir = Path.Combine(appRoot, "Logs");
                Directory.CreateDirectory(logDir);

                var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var path = Path.Combine(logDir, $"crash_{stamp}.txt");

                var sb = new StringBuilder(2048);
                sb.AppendLine($"Time: {DateTime.Now:O}");
                sb.AppendLine($"Source: {source}");
                sb.AppendLine($"Message: {exception.Message}");
                sb.AppendLine();
                sb.AppendLine(exception.ToString());

                File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // Best-effort logging only.
            }
        }
    }

}
