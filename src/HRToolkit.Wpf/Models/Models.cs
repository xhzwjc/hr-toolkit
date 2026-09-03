using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace HRToolkit.Wpf.Models
{
    public class NavigationGroup
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("tool_ids")]
        public List<string> ToolIds { get; set; } = new();
    }

    public class ToolItem
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;

        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("group")]
        public string Group { get; set; } = string.Empty;

        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        [JsonPropertyName("cli_command")]
        public string CliCommand { get; set; } = string.Empty;

        [JsonPropertyName("multi_input")]
        public bool MultiInput { get; set; }
    }

    public class MetadataResult
    {
        [JsonPropertyName("nav_groups")]
        public List<NavigationGroup> NavGroups { get; set; } = new();

        [JsonPropertyName("tools")]
        public List<ToolItem> Tools { get; set; } = new();

        [JsonPropertyName("builtin_materials")]
        public List<string> BuiltinMaterials { get; set; } = new();

        [JsonPropertyName("material_presets")]
        public Dictionary<string, List<string>> MaterialPresets { get; set; } = new();

        [JsonPropertyName("default_project_name")]
        public string DefaultProjectName { get; set; } = string.Empty;

        [JsonPropertyName("default_project_parent")]
        public string DefaultProjectParent { get; set; } = string.Empty;
    }

    public class ProjectStatus
    {
        [JsonPropertyName("has_project")]
        public bool HasProject { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("path")]
        public string Path { get; set; } = string.Empty;

        [JsonPropertyName("writable")]
        public bool Writable { get; set; }

        [JsonPropertyName("read_only_reason")]
        public string ReadOnlyReason { get; set; } = string.Empty;
    }

    public class WorkspaceFileItem
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        [JsonPropertyName("relative_path")]
        public string RelativePath { get; set; } = string.Empty;

        [JsonPropertyName("category")]
        public string Category { get; set; } = string.Empty;

        [JsonPropertyName("batch_id")]
        public string? BatchId { get; set; }

        [JsonPropertyName("size")]
        public long Size { get; set; }

        [JsonPropertyName("mtime")]
        public double Mtime { get; set; }

        public string FormattedSize
        {
            get
            {
                if (Size >= 1024 * 1024 * 1024) return $"{Size / (1024.0 * 1024 * 1024):F1} GB";
                if (Size >= 1024 * 1024) return $"{Size / (1024.0 * 1024):F1} MB";
                if (Size >= 1024) return $"{Size / 1024.0:F1} KB";
                return $"{Size} B";
            }
        }
    }

    public class TrashItem
    {
        [JsonPropertyName("batch_id")]
        public string BatchId { get; set; } = string.Empty;

        [JsonPropertyName("tool_name")]
        public string ToolName { get; set; } = string.Empty;

        [JsonPropertyName("business_description")]
        public string BusinessDescription { get; set; } = string.Empty;

        [JsonPropertyName("deleted_at")]
        public string DeletedAt { get; set; } = string.Empty;

        [JsonPropertyName("total_size")]
        public long TotalSize { get; set; }

        [JsonPropertyName("file_count")]
        public int FileCount { get; set; }
    }

    public class LogEntry
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string Level { get; set; } = "info"; // info, success, warning, error
        public string Message { get; set; } = string.Empty;

        public string FormattedTime => Timestamp.ToString("HH:mm:ss");
    }
}
