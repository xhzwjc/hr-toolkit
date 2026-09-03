using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using HRToolkit.Wpf.Models;
using HRToolkit.Wpf.Services;

namespace HRToolkit.Wpf.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private readonly PythonProcessManager _processManager;
        private JsonRpcClient Client => _processManager.Client ?? throw new InvalidOperationException("IPC 未就绪");

        public ObservableCollection<NavigationGroup> NavGroups { get; } = new();
        public ObservableCollection<ToolItem> AllTools { get; } = new();

        public LogViewModel Log { get; } = new();
        public WorkspaceViewModel Workspace { get; private set; } = null!;
        public ToolFormViewModel Form { get; private set; } = null!;

        private ProjectStatus _project = new();
        public ProjectStatus Project
        {
            get => _project;
            set => SetProperty(ref _project, value);
        }

        private ToolItem? _currentTool;
        public ToolItem? CurrentTool
        {
            get => _currentTool;
            set
            {
                if (SetProperty(ref _currentTool, value) && value != null)
                {
                    Form.SelectedTool = value;
                }
            }
        }

        public ICommand OpenProjectCommand { get; }
        public ICommand CreateProjectCommand { get; }
        public ICommand SelectToolCommand { get; }

        public MainViewModel(PythonProcessManager processManager)
        {
            _processManager = processManager;
            Workspace = new WorkspaceViewModel(Client);
            Form = new ToolFormViewModel(Client, Log);

            OpenProjectCommand = new RelayCommand(async () => await OnOpenProjectAsync());
            CreateProjectCommand = new RelayCommand(async () => await OnCreateProjectAsync());
            SelectToolCommand = new RelayCommand(t => CurrentTool = t as ToolItem);

            // Wire up IPC events to ViewModels
            Client.OnLogReceived += (level, msg) => Log.AppendLog(msg, level);
            Client.OnProgressReceived += (ratio, msg) =>
            {
                Log.ProgressRatio = ratio;
                if (!string.IsNullOrEmpty(msg)) Log.StatusMessage = msg;
            };
            Client.OnProjectChanged += elem =>
            {
                var status = JsonSerializer.Deserialize<ProjectStatus>(elem.GetRawText());
                if (status != null) Project = status;
            };
            Client.OnWorkspaceChanged += async () => await Workspace.LoadWorkspaceFilesAsync();
            Client.OnToolFinished += (success, msg) =>
            {
                Log.IsBusy = false;
                Log.StatusMessage = success ? "处理完成" : "处理未完成";
                Log.AppendLog(msg, success ? "success" : "error");
            };
        }

        public async Task InitializeAsync()
        {
            try
            {
                var meta = await Client.SendRequestAsync<MetadataResult>("get_metadata");
                if (meta != null)
                {
                    AllTools.Clear();
                    foreach (var t in meta.Tools) AllTools.Add(t);

                    NavGroups.Clear();
                    foreach (var g in meta.NavGroups) NavGroups.Add(g);

                    if (AllTools.Count > 0)
                    {
                        CurrentTool = AllTools[0];
                    }
                }

                var status = await Client.SendRequestAsync<ProjectStatus>("get_project_status");
                if (status != null) Project = status;

                Log.AppendLog("HR Toolkit 核心引擎与 IPC 通信已连接。", "success");
            }
            catch (Exception ex)
            {
                Log.AppendLog($"初始化失败: {ex.Message}", "error");
            }
        }

        private async Task OnOpenProjectAsync()
        {
            var dlg = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "打开工作项目文件夹"
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var status = await Client.SendRequestAsync<ProjectStatus>("open_project", new { path = dlg.FolderName });
                    if (status != null) Project = status;
                    await Workspace.LoadWorkspaceFilesAsync();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "打开项目失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private async Task OnCreateProjectAsync()
        {
            var dlg = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "选择新建项目的父级保存目录"
            };
            if (dlg.ShowDialog() == true)
            {
                string parentDir = dlg.FolderName;
                string projectName = $"人事工作项目_{DateTime.Now:yyyyMMdd}";

                try
                {
                    var status = await Client.SendRequestAsync<ProjectStatus>("create_project", new { name = projectName, parent = parentDir });
                    if (status != null) Project = status;
                    await Workspace.LoadWorkspaceFilesAsync();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "创建项目失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
    }
}
