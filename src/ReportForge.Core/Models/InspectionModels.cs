using System.Collections.Generic;

namespace ReportForge.Core.Models
{
    /// <summary>巡检问题</summary>
    public class Issue
    {
        public string Id { get; set; } = "";
        public IssueSeverity Severity { get; set; } = IssueSeverity.Warning;
        public string Category { get; set; } = "";
        public LocalizedString Message { get; set; } = new();
        /// <summary>问题段落在文档中的索引</summary>
        public int ParagraphIndex { get; set; }
        /// <summary>段落文本预览（前80字符）</summary>
        public string TextPreview { get; set; } = "";
        /// <summary>是否可自动修复</summary>
        public bool AutoFixable { get; set; }
    }

    public enum IssueSeverity { Info, Warning, Error }

    /// <summary>巡检报告</summary>
    public class InspectionReport
    {
        public List<Issue> Issues { get; set; } = new();
        public int TotalParagraphs { get; set; }
        public int TotalTables { get; set; }
        public int TotalSections { get; set; }
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public int InfoCount { get; set; }
    }

    /// <summary>修复报告</summary>
    public class RepairReport
    {
        public int DirectFormattingCleared { get; set; }
        public int StylesReapplied { get; set; }
        public int NumberingsRebuilt { get; set; }
        public int CaptionsRebuilt { get; set; }
        public int TablesNormalized { get; set; }
        public int FieldsUpdated { get; set; }
        public List<string> Warnings { get; set; } = new();
    }

    /// <summary>修复选项</summary>
    public class RepairOptions
    {
        public bool ClearDirectFormatting { get; set; } = true;
        public bool ReapplyStyles { get; set; } = true;
        public bool RebuildNumbering { get; set; } = true;
        public bool RebuildCaptions { get; set; } = true;
        public bool UpdateTOC { get; set; } = true;
        public bool UpdateCrossReferences { get; set; } = true;
        public bool NormalizeTables { get; set; } = true;
        public bool UpdatePageFields { get; set; } = true;
    }
}
