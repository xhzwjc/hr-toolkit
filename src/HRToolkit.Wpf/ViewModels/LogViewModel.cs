using System.Collections.ObjectModel;
using System.Windows;
using HRToolkit.Wpf.Models;

namespace HRToolkit.Wpf.ViewModels
{
    public class LogViewModel : ViewModelBase
    {
        public ObservableCollection<LogEntry> Entries { get; } = new();

        private double _progressRatio = 0.0;
        public double ProgressRatio
        {
            get => _progressRatio;
            set => SetProperty(ref _progressRatio, value);
        }

        private string _statusMessage = "就绪";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private bool _isBusy = false;
        public bool IsBusy
        {
            get => _isBusy;
            set => SetProperty(ref _isBusy, value);
        }

        public void AppendLog(string message, string level = "info")
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                Entries.Add(new LogEntry { Message = message, Level = level });
                // Keep last 1000 items to maintain high performance
                if (Entries.Count > 1000)
                {
                    Entries.RemoveAt(0);
                }
            });
        }

        public void Clear()
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                Entries.Clear();
                ProgressRatio = 0.0;
                StatusMessage = "就绪";
            });
        }
    }
}
