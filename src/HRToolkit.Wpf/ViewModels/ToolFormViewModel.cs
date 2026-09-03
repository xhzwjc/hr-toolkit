using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HRToolkit.Wpf.Models;
using HRToolkit.Wpf.Services;

namespace HRToolkit.Wpf.ViewModels
{
    public class ToolFormViewModel : ViewModelBase
    {
        private readonly JsonRpcClient _client;
        private readonly LogViewModel _logVm;

        private ToolItem? _selectedTool;
        public ToolItem? SelectedTool
        {
            get => _selectedTool;
            set
            {
                if (SetProperty(ref _selectedTool, value))
                {
                    ResetForm();
                }
            }
        }

        public ObservableCollection<string> InputFiles { get; } = new();

        private string? _supportFile;
        public string? SupportFile
        {
            get => _supportFile;
            set => SetProperty(ref _supportFile, value);
        }

        // Common Tool Options
        private bool _dryRun;
        public bool DryRun
        {
            get => _dryRun;
            set => SetProperty(ref _dryRun, value);
        }

        private string _libraryMode = "person_folder";
        public string LibraryMode
        {
            get => _libraryMode;
            set => SetProperty(ref _libraryMode, value);
        }

        private bool _useOcrCache = true;
        public bool UseOcrCache
        {
            get => _useOcrCache;
            set => SetProperty(ref _useOcrCache, value);
        }

        public ICommand RunCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand AddFilesCommand { get; }
        public ICommand ClearFilesCommand { get; }

        public ToolFormViewModel(JsonRpcClient client, LogViewModel logVm)
        {
            _client = client;
            _logVm = logVm;

            RunCommand = new RelayCommand(async () => await RunCurrentToolAsync(), () => !_logVm.IsBusy && SelectedTool != null);
            CancelCommand = new RelayCommand(async () => await CancelCurrentToolAsync(), () => _logVm.IsBusy);
            ClearFilesCommand = new RelayCommand(() => InputFiles.Clear());
            AddFilesCommand = new RelayCommand(OnAddFiles);
        }

        private void OnAddFiles()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = SelectedTool?.MultiInput ?? true,
                Filter = "常用文件 (*.xlsx;*.xls;*.zip;*.rar;*.7z)|*.xlsx;*.xls;*.zip;*.rar;*.7z|所有文件 (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                foreach (var file in dlg.FileNames)
                {
                    if (!InputFiles.Contains(file))
                    {
                        InputFiles.Add(file);
                    }
                }
            }
        }

        public async Task RunCurrentToolAsync()
        {
            if (SelectedTool == null) return;
            if (InputFiles.Count == 0 && SelectedTool.Id != "data_statistics")
            {
                MessageBox.Show("请先添加需要处理的输入文件或文件夹。", "缺少输入文件", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _logVm.IsBusy = true;
            _logVm.StatusMessage = $"正在处理 {SelectedTool.Name}…";
            _logVm.AppendLog($"开始任务：{SelectedTool.Name}", "info");

            var payload = new
            {
                tool_id = SelectedTool.Id,
                inputs = InputFiles,
                support_file = SupportFile,
                options = new
                {
                    dry_run = DryRun,
                    library_mode = LibraryMode,
                    use_ocr_cache = UseOcrCache
                }
            };

            try
            {
                await _client.SendRequestAsync<object>("run_tool", payload);
            }
            catch (Exception ex)
            {
                _logVm.IsBusy = false;
                _logVm.AppendLog($"启动失败: {ex.Message}", "error");
                MessageBox.Show(ex.Message, "运行失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public async Task CancelCurrentToolAsync()
        {
            try
            {
                await _client.SendRequestAsync<object>("cancel_tool");
                _logVm.AppendLog("已发送取消请求，正在安全终止…", "warning");
            }
            catch { }
        }

        public void ResetForm()
        {
            InputFiles.Clear();
            SupportFile = null;
            DryRun = false;
            LibraryMode = "person_folder";
            UseOcrCache = true;
        }
    }
}
