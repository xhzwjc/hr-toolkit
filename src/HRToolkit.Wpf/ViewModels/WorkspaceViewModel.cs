using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HRToolkit.Wpf.Models;
using HRToolkit.Wpf.Services;

namespace HRToolkit.Wpf.ViewModels
{
    public class WorkspaceViewModel : ViewModelBase
    {
        private readonly JsonRpcClient _client;
        public ObservableCollection<WorkspaceFileItem> Files { get; } = new();
        public ObservableCollection<TrashItem> TrashItems { get; } = new();

        private bool _isDrawerOpen = false;
        public bool IsDrawerOpen
        {
            get => _isDrawerOpen;
            set => SetProperty(ref _isDrawerOpen, value);
        }

        public ICommand RefreshCommand { get; }
        public ICommand ToggleDrawerCommand { get; }

        public WorkspaceViewModel(JsonRpcClient client)
        {
            _client = client;
            RefreshCommand = new RelayCommand(async () => await LoadWorkspaceFilesAsync());
            ToggleDrawerCommand = new RelayCommand(() => IsDrawerOpen = !IsDrawerOpen);
        }

        public async Task LoadWorkspaceFilesAsync()
        {
            try
            {
                var files = await _client.SendRequestAsync<WorkspaceFileItem[]>("list_workspace_files");
                Application.Current?.Dispatcher.Invoke(() =>
                {
                    Files.Clear();
                    if (files != null)
                    {
                        foreach (var f in files) Files.Add(f);
                    }
                });
            }
            catch { }
        }

        public async Task ImportFilesAsync(string[] paths)
        {
            if (paths.Length == 0) return;
            try
            {
                await _client.SendRequestAsync<object>("import_workspace_files", new { sources = paths });
                await LoadWorkspaceFilesAsync();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"导入失败: {ex.Message}", "文件导入", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
