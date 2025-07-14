using System;
using System.Windows.Forms;
using ZebraPrinterMonitor.Forms;
using ZebraPrinterMonitor.Services;
using ZebraPrinterMonitor.Utils;

namespace ZebraPrinterMonitor
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                // 设置应用程序样式
                Application.SetHighDpiMode(HighDpiMode.SystemAware);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 初始化日志记录
                Logger.Initialize();
                Logger.Info("应用程序启动中...");

                // 检查单实例运行
                using (var mutex = new System.Threading.Mutex(false, "ZebraPrinterMonitor_SingleInstance"))
                {
                    if (!mutex.WaitOne(0, false))
                    {
                        MessageBox.Show("程序已经在运行中！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 初始化配置
                    ConfigurationManager.Initialize();

                    // 设置全局异常处理
                    Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                    Application.ThreadException += Application_ThreadException;
                    AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

                    // 启动主窗体
                    var mainForm = new MainForm();
                    Application.Run(mainForm);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"程序启动失败: {ex.Message}", ex);
                MessageBox.Show($"程序启动失败:\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Logger.Info("应用程序退出");
                Logger.Shutdown();
            }
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            Logger.Error($"UI线程异常: {e.Exception.Message}", e.Exception);
            MessageBox.Show($"发生未处理的异常:\n{e.Exception.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            Logger.Error($"应用程序域异常: {ex?.Message ?? "Unknown"}", ex);
            MessageBox.Show($"发生严重错误，程序将退出:\n{ex?.Message ?? "Unknown error"}", "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
} 