using System;
using System.Windows;
using HRToolkit.Wpf.Services;
using HRToolkit.Wpf.ViewModels;

namespace HRToolkit.Wpf
{
    public partial class App : Application
    {
        private PythonProcessManager? _processManager;

        private async void Application_Startup(object sender, StartupEventArgs e)
        {
            _processManager = new PythonProcessManager();
            try
            {
                _processManager.StartEngine();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动 Python 业务引擎失败：{ex.Message}\n请确认已安装 Python 并在项目目录就绪。", "HR Toolkit 启动错误", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
                return;
            }

            var mainVm = new MainViewModel(_processManager);
            var mainWindow = new MainWindow
            {
                DataContext = mainVm
            };
            mainWindow.Show();

            await mainVm.InitializeAsync();
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            _processManager?.Dispose();
        }
    }
}
