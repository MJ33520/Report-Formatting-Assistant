namespace ReportForge.Core.Models
{
    /// <summary>题注配置（图题/表题）</summary>
    public class CaptionConfiguration
    {
        public CaptionTypeConfig Figure { get; set; } = new()
        {
            Label = new LocalizedString { ["zh-CN"] = "图", ["en-US"] = "Figure" },
            ChapterNumbering = true,
            Separator = "-",
            Position = CaptionPosition.Below,
            StyleId = "RF_FigureCaption",
            Font = new FontConfig { Zh = "黑体" },
            FontSizePt = 10.5,
            Alignment = TextAlignment.Center
        };

        public CaptionTypeConfig Table { get; set; } = new()
        {
            Label = new LocalizedString { ["zh-CN"] = "表", ["en-US"] = "Table" },
            ChapterNumbering = true,
            Separator = "-",
            Position = CaptionPosition.Above,
            StyleId = "RF_TableCaption",
            Font = new FontConfig { Zh = "黑体" },
            FontSizePt = 10.5,
            Alignment = TextAlignment.Center
        };
    }

    public class CaptionTypeConfig
    {
        /// <summary>标签文字（多语言），如"图"/"Figure"</summary>
        public LocalizedString Label { get; set; } = new();
        /// <summary>是否按章编号</summary>
        public bool ChapterNumbering { get; set; } = true;
        /// <summary>章节号与序号之间的分隔符</summary>
        public string Separator { get; set; } = "-";
        /// <summary>题注位置</summary>
        public CaptionPosition Position { get; set; } = CaptionPosition.Below;
        /// <summary>题注段落样式ID</summary>
        public string StyleId { get; set; } = "";
        public FontConfig Font { get; set; } = new();
        public double FontSizePt { get; set; } = 10.5;
        public TextAlignment Alignment { get; set; } = TextAlignment.Center;
    }

    public enum CaptionPosition
    {
        Above,
        Below
    }
}
