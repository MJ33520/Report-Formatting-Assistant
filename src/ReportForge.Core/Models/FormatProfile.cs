using System.Collections.Generic;

namespace ReportForge.Core.Models
{
    /// <summary>
    /// 格式配置 Profile — 完整描述一份报告的格式规范。
    /// 从规范文档解析或用户手动配置生成，以 JSON 持久化。
    /// </summary>
    public class FormatProfile
    {
        /// <summary>Profile ID，如 "gov-report-2020"</summary>
        public string Id { get; set; } = "";
        public string Version { get; set; } = "1.0";
        public string Locale { get; set; } = "zh-CN";
        public LocalizedString DisplayName { get; set; } = new();

        // --- 页面 ---
        public PageSetup PageSetup { get; set; } = new();

        // --- 标题 ---
        public List<HeadingLevel> HeadingLevels { get; set; } = new();

        // --- 编号 ---
        public NumberingScheme NumberingScheme { get; set; } = new();

        // --- 正文 ---
        public StyleDefinition BodyStyle { get; set; } = new();

        // --- 题注 ---
        public CaptionConfiguration CaptionConfig { get; set; } = new();

        // --- 表格 ---
        public TableStyleConfig TableStyle { get; set; } = new();

        // --- 分节 ---
        public List<SectionTemplate> SectionTemplates { get; set; } = new();

        // --- 目录 ---
        public TocConfig TocConfig { get; set; } = new();

        // --- 打印 ---
        public PrintConfig PrintConfig { get; set; } = new();

        // --- 附加样式（页眉页脚、附录标题、图注表注等） ---
        public List<StyleDefinition> AdditionalStyles { get; set; } = new();
    }
}
