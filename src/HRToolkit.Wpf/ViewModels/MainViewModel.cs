using System;
using System.Collections.Generic;
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

            // Populate fallback default tools immediately so UI is never blank
            LoadDefaultTools();

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

        private void LoadDefaultTools()
        {
            var defaultTools = new List<ToolItem>
            {
                new() { Id = "social_security", Name = "社保明细与汇总", Group = "社保与保险", Description = "需求1：生成社保明细表和社保汇总表", MultiInput = true },
                new() { Id = "insurance_ledger", Name = "保险台账与预警", Group = "社保与保险", Description = "需求3：生成保险台账和人员增减预警", MultiInput = true },
                new() { Id = "data_statistics", Name = "考勤与周月报", Group = "考勤与统计", Description = "需求2：生成考勤和周月报统计表", MultiInput = true },
                new() { Id = "salary_split", Name = "工资表拆分", Group = "薪酬管理", Description = "需求4：将工资表按入职公司拆分为多个工作簿", MultiInput = false },
                new() { Id = "salary_merge", Name = "多月工资合并", Group = "薪酬管理", Description = "需求5：合并多个月工资表，生成个人应发工资汇总", MultiInput = true },
                new() { Id = "personnel_change_merge", Name = "异动汇总", Group = "人员与档案", Description = "需求6：汇总多个项目异动表", MultiInput = true },
                new() { Id = "archive_import", Name = "档案入库", Group = "人员与档案", Description = "需求7：将项目档案移交表写入公司档案汇总表", MultiInput = true },
                new() { Id = "material_collector", Name = "员工资料打包", Group = "人员与档案", Description = "需求9：员工资料自动打包与信息提取", MultiInput = false },
                new() { Id = "folder_rename", Name = "资料文件夹改名", Group = "人员与档案", Description = "需求8：人员资料文件夹批量改名", MultiInput = false }
            };

            AllTools.Clear();
            foreach (var t in defaultTools) AllTools.Add(t);
            CurrentTool = AllTools[0];
        }

        public async Task InitializeAsync()
        {
            Log.AppendLog("正在连接 Python 业务引擎…", "info");
            try
            {
                var meta = await Client.SendRequestAsync<MetadataResult>("get_metadata");
                if (meta != null && meta.Tools.Count > 0)
                {
                    AllTools.Clear();
                    foreach (var t in meta.Tools) AllTools.Add(t);

                    NavGroups.Clear();
                    foreach (var g in meta.NavGroups) NavGroups.Add(g);

                    CurrentTool = AllTools[0];
                }

                var status = await Client.SendRequestAsync<ProjectStatus>("get_project_status");
                if (status != null) Project = status;

                Log.AppendLog("HR Toolkit 核心引擎已连接就绪。", "success");
            }
            catch (Exception ex)
            {
                Log.AppendLog($"连接 Python 引擎失败: {ex.Message}", "error");
                if (!string.IsNullOrEmpty(_processManager.LastStderr))
                {
                    Log.AppendLog($"[Python 报错信息] {_processManager.LastStderr}", "error");
                }
                Log.AppendLog($"Python 路径: {_processManager.ResolvedPythonPath}", "warning");
                Log.AppendLog($"运行工作目录: {_processManager.ResolvedWorkingDir}", "warning");

                MessageBox.Show(
                    $"连接 Python 后端失败：{ex.Message}\n\n" +
                    $"检测到的 Python: {_processManager.ResolvedPythonPath}\n" +
                    $"运行目录: {_processManager.ResolvedWorkingDir}\n\n" +
                    (!string.IsNullOrEmpty(_processManager.LastStderr) ? $"Python 报错:\n{_processManager.LastStderr}\n\n" : "") +
                    "请确认：\n1. 是否已安装所需依赖（pip install -e . 或 pip install -r requirements.txt）\n2. 可在终端手动测试：python -m hr_toolkit --ipc",
                    "Python 核心未就绪",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
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
