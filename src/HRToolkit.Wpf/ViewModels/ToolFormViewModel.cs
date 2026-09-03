using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
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
                    OnToolChanged();
                }
            }
        }

        // 1. Primary Inputs
        public ObservableCollection<string> InputFiles { get; } = new();

        private string _inputFilesLabel = "待处理输入文件";
        public string InputFilesLabel
        {
            get => _inputFilesLabel;
            set => SetProperty(ref _inputFilesLabel, value);
        }

        // 2. Secondary / Support File
        private string? _supportFile;
        public string? SupportFile
        {
            get => _supportFile;
            set => SetProperty(ref _supportFile, value);
        }

        private bool _hasSupportFile;
        public bool HasSupportFile
        {
            get => _hasSupportFile;
            set => SetProperty(ref _hasSupportFile, value);
        }

        private string _supportFileTitle = "辅助表 / 花名册文件";
        public string SupportFileTitle
        {
            get => _supportFileTitle;
            set => SetProperty(ref _supportFileTitle, value);
        }

        private string _supportFileHelp = "";
        public string SupportFileHelp
        {
            get => _supportFileHelp;
            set => SetProperty(ref _supportFileHelp, value);
        }

        // 3. Common Tool Options
        private bool _dryRun;
        public bool DryRun
        {
            get => _dryRun;
            set => SetProperty(ref _dryRun, value);
        }

        private bool _useOcrCache = true;
        public bool UseOcrCache
        {
            get => _useOcrCache;
            set => SetProperty(ref _useOcrCache, value);
        }

        // 4. Material Collector Specific Options
        public bool IsMaterialCollector => SelectedTool?.Id == "material_collector";

        private string _targetEmployee = "";
        public string TargetEmployee
        {
            get => _targetEmployee;
            set => SetProperty(ref _targetEmployee, value);
        }

        private bool _collectAllMaterials = true;
        public bool CollectAllMaterials
        {
            get => _collectAllMaterials;
            set => SetProperty(ref _collectAllMaterials, value);
        }

        private bool _createZip = false;
        public bool CreateZip
        {
            get => _createZip;
            set => SetProperty(ref _createZip, value);
        }

        private string _libraryMode = "person_folder";
        public string LibraryMode
        {
            get => _libraryMode;
            set => SetProperty(ref _libraryMode, value);
        }

        public ObservableCollection<MaterialTypeOption> MaterialOptions { get; } = new();

        // 5. Folder Rename Specific Options
        public bool IsFolderRename => SelectedTool?.Id == "folder_rename";

        private string _renameMode = "append"; // append, remove, replace, excel
        public string RenameMode
        {
            get => _renameMode;
            set
            {
                if (SetProperty(ref _renameMode, value))
                {
                    OnPropertyChanged(nameof(IsAppendOrRemoveMode));
                    OnPropertyChanged(nameof(IsReplaceMode));
                    OnPropertyChanged(nameof(IsExcelRenameMode));
                }
            }
        }

        public bool IsAppendOrRemoveMode => RenameMode == "append" || RenameMode == "remove";
        public bool IsReplaceMode => RenameMode == "replace";
        public bool IsExcelRenameMode => RenameMode == "excel";

        private string _renameText = "";
        public string RenameText
        {
            get => _renameText;
            set => SetProperty(ref _renameText, value);
        }

        private string _targetName = "";
        public string TargetName
        {
            get => _targetName;
            set => SetProperty(ref _targetName, value);
        }

        private string _replacementName = "";
        public string ReplacementName
        {
            get => _replacementName;
            set => SetProperty(ref _replacementName, value);
        }

        private string _renameFileType = "folder";
        public string RenameFileType
        {
            get => _renameFileType;
            set => SetProperty(ref _renameFileType, value);
        }

        // 6. Data Statistics Specific Options
        public bool IsDataStatistics => SelectedTool?.Id == "data_statistics";

        private string _weekStart = "";
        public string WeekStart
        {
            get => _weekStart;
            set => SetProperty(ref _weekStart, value);
        }

        private string _weekEnd = "";
        public string WeekEnd
        {
            get => _weekEnd;
            set => SetProperty(ref _weekEnd, value);
        }

        private string _monthStart = "";
        public string MonthStart
        {
            get => _monthStart;
            set => SetProperty(ref _monthStart, value);
        }

        private string _monthEnd = "";
        public string MonthEnd
        {
            get => _monthEnd;
            set => SetProperty(ref _monthEnd, value);
        }

        private bool _includeBusinessTrip = false;
        public bool IncludeBusinessTrip
        {
            get => _includeBusinessTrip;
            set => SetProperty(ref _includeBusinessTrip, value);
        }

        // 7. Salary Tools Specific Options
        public bool IsSalarySplit => SelectedTool?.Id == "salary_split";
        public bool IsSalaryMerge => SelectedTool?.Id == "salary_merge";

        private string _salaryYear = DateTime.Now.Year.ToString();
        public string SalaryYear
        {
            get => _salaryYear;
            set => SetProperty(ref _salaryYear, value);
        }

        // Commands
        public ICommand RunCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand AddFilesCommand { get; }
        public ICommand AddFolderCommand { get; }
        public ICommand ClearFilesCommand { get; }
        public ICommand SelectSupportFileCommand { get; }
        public ICommand ClearSupportFileCommand { get; }

        public ToolFormViewModel(JsonRpcClient client, LogViewModel logVm)
        {
            _client = client;
            _logVm = logVm;

            RunCommand = new RelayCommand(async () => await RunCurrentToolAsync(), () => !_logVm.IsBusy && SelectedTool != null);
            CancelCommand = new RelayCommand(async () => await CancelCurrentToolAsync(), () => _logVm.IsBusy);
            ClearFilesCommand = new RelayCommand(() => InputFiles.Clear());
            AddFilesCommand = new RelayCommand(OnAddFiles);
            AddFolderCommand = new RelayCommand(OnAddFolder);
            SelectSupportFileCommand = new RelayCommand(OnSelectSupportFile);
            ClearSupportFileCommand = new RelayCommand(() => SupportFile = null);

            InitDefaultMaterials();
        }

        private void InitDefaultMaterials()
        {
            string[] defaultTypes = { "身份证", "学历证", "银行卡", "劳动合同", "驾驶证", "特种证书", "体检表", "离职证明" };
            MaterialOptions.Clear();
            foreach (var t in defaultTypes)
            {
                MaterialOptions.Add(new MaterialTypeOption { Name = t, IsSelected = false });
            }
        }

        private void OnToolChanged()
        {
            InputFiles.Clear();
            SupportFile = null;
            DryRun = false;

            OnPropertyChanged(nameof(IsMaterialCollector));
            OnPropertyChanged(nameof(IsFolderRename));
            OnPropertyChanged(nameof(IsDataStatistics));
            OnPropertyChanged(nameof(IsSalarySplit));
            OnPropertyChanged(nameof(IsSalaryMerge));

            if (SelectedTool == null) return;

            switch (SelectedTool.Id)
            {
                case "social_security":
                    InputFilesLabel = "待处理社保费用申报表 (Excel / 压缩包 / 文件夹)";
                    HasSupportFile = true;
                    SupportFileTitle = "参保人员花名册 (Excel)";
                    SupportFileHelp = "选择公司现行员工花名册，用于校对社保缴纳人员名单与部门归属。";
                    break;

                case "insurance_ledger":
                    InputFilesLabel = "待处理保单 / 保险账单 (Excel / 压缩包)";
                    HasSupportFile = true;
                    SupportFileTitle = "人力资源分析表 / 员工花名册 (Excel)";
                    SupportFileHelp = "选择现行花名册，用于核对参保增减人员并生成预警。";
                    break;

                case "data_statistics":
                    InputFilesLabel = "考勤打卡原始记录表 (Excel / 压缩包)";
                    HasSupportFile = true;
                    SupportFileTitle = "应汇报人员名单 (Excel)";
                    SupportFileHelp = "选择需统计周报、月报的应出勤/汇报人员名单。";
                    break;

                case "salary_split":
                    InputFilesLabel = "待拆分综合工资表 (Excel)";
                    HasSupportFile = false;
                    SupportFileTitle = "";
                    SupportFileHelp = "";
                    break;

                case "salary_merge":
                    InputFilesLabel = "各月度待合并工资表 (Excel / 压缩包)";
                    HasSupportFile = true;
                    SupportFileTitle = "已有工资汇总表 (Excel，可选追加)";
                    SupportFileHelp = "如果已有往期汇总表，在此选择可将新月份数据自动增量追加合并。";
                    break;

                case "personnel_change_merge":
                    InputFilesLabel = "各项目部人事异动表 (Excel / 压缩包)";
                    HasSupportFile = true;
                    SupportFileTitle = "异动汇总模板表 (Excel，可选)";
                    SupportFileHelp = "可选指定专用的汇总模板；未选择时自动采用系统标准模板。";
                    break;

                case "archive_import":
                    InputFilesLabel = "项目档案移交清单 (Excel / 压缩包)";
                    HasSupportFile = true;
                    SupportFileTitle = "公司档案汇总表 (Excel)";
                    SupportFileHelp = "选择公司各公司档案总表，系统将按所属公司将移交记录自动写入对应工作表。";
                    break;

                case "material_collector":
                    InputFilesLabel = "待提取员工资料库目录 (文件夹 / 扫描件)";
                    HasSupportFile = true;
                    SupportFileTitle = "员工名单花名册 (Excel，可选)";
                    SupportFileHelp = "可选择包含姓名/身份证的 Excel 名单，或直接在下方文本框输入单个员工姓名进行过滤提取。";
                    break;

                case "folder_rename":
                    InputFilesLabel = "待处理目标文件夹 / 文件所在目录";
                    HasSupportFile = true;
                    SupportFileTitle = "人员重命名名单 (Excel)";
                    SupportFileHelp = "当改名模式选择“按 Excel 名单顺序批量重命名”时，需指定此 Excel 表格。";
                    break;

                default:
                    InputFilesLabel = "待处理输入文件";
                    HasSupportFile = false;
                    SupportFileTitle = "";
                    SupportFileHelp = "";
                    break;
            }
        }

        private void OnAddFiles()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = SelectedTool?.MultiInput ?? true,
                Filter = "常用文件 (*.xlsx;*.xls;*.zip;*.rar;*.7z;*.pdf)|*.xlsx;*.xls;*.zip;*.rar;*.7z;*.pdf|所有文件 (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                foreach (var file in dlg.FileNames)
                {
                    if (!InputFiles.Contains(file)) InputFiles.Add(file);
                }
            }
        }

        private void OnAddFolder()
        {
            var dlg = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "选择待处理的文件夹"
            };
            if (dlg.ShowDialog() == true)
            {
                if (!InputFiles.Contains(dlg.FolderName))
                {
                    InputFiles.Add(dlg.FolderName);
                }
            }
        }

        private void OnSelectSupportFile()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = false,
                Filter = "Excel 表格 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                Title = $"选择{SupportFileTitle}"
            };
            if (dlg.ShowDialog() == true)
            {
                SupportFile = dlg.FileName;
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

            // Build options dictionary
            var options = new Dictionary<string, object>
            {
                ["dry_run"] = DryRun,
                ["use_ocr_cache"] = UseOcrCache,
                ["library_mode"] = LibraryMode
            };

            if (IsMaterialCollector)
            {
                options["target_input"] = TargetEmployee.Trim();
                options["collect_all"] = CollectAllMaterials;
                options["create_zip"] = CreateZip;
                options["material_types"] = CollectAllMaterials
                    ? new List<string>()
                    : MaterialOptions.Where(m => m.IsSelected).Select(m => m.Name).ToList();
            }
            else if (IsFolderRename)
            {
                options["mode"] = RenameMode;
                options["file_type"] = RenameFileType;
                options["rename_text"] = RenameText.Trim();
                options["target_name"] = TargetName.Trim();
                options["replacement_name"] = ReplacementName.Trim();
                options["preview"] = DryRun;
            }
            else if (IsDataStatistics)
            {
                if (!string.IsNullOrWhiteSpace(WeekStart) && !string.IsNullOrWhiteSpace(WeekEnd))
                {
                    options["week_start"] = WeekStart.Trim();
                    options["week_end"] = WeekEnd.Trim();
                }
                if (!string.IsNullOrWhiteSpace(MonthStart))
                {
                    options["month_start"] = MonthStart.Trim();
                    options["month_end"] = !string.IsNullOrWhiteSpace(MonthEnd) ? MonthEnd.Trim() : MonthStart.Trim();
                }
                options["include_business_trip"] = IncludeBusinessTrip;
            }

            var payload = new
            {
                tool_id = SelectedTool.Id,
                inputs = InputFiles.ToList(),
                support_file = SupportFile,
                options = options
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
    }
}
