using System.Collections.Generic;

namespace ReportForge.Core.Models
{
    /// <summary>分节模板配置</summary>
    public class SectionTemplate
    {
        /// <summary>节模板ID，如 "cover", "toc", "main_body"</summary>
        public string Id { get; set; } = "";
        public LocalizedString DisplayName { get; set; } = new();
        public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;
        public PageNumberConfig? PageNumbering { get; set; }
        public PrintMode PrintMode { get; set; } = PrintMode.DoubleSide;
        /// <summary>页眉页脚是否链接到前一节</summary>
        public bool LinkToPrevious { get; set; } = true;
        /// <summary>是否启用奇偶页差异</summary>
        public bool DifferentOddEvenPages { get; set; }
        /// <summary>自定义页边距（可选，为 null 则继承全局设置）</summary>
        public MarginConfig? CustomMargins { get; set; }

        // --- 页眉 ---
        /// <summary>是否显示页眉</summary>
        public bool ShowHeader { get; set; } = true;
        /// <summary>页眉文字</summary>
        public string HeaderText { get; set; } = "";
        /// <summary>页眉字体</summary>
        public FontConfig HeaderFont { get; set; } = new() { Zh = "仿宋_GB2312", Latin = "Times New Roman" };
        /// <summary>页眉字号 pt</summary>
        public double HeaderFontSizePt { get; set; } = 12; // 小四
        /// <summary>页眉下方是否有边线</summary>
        public bool HeaderBorderBottom { get; set; } = true;

        // --- 页脚 ---
        /// <summary>是否显示页脚</summary>
        public bool ShowFooter { get; set; } = true;

        // --- 页眉/页脚距离 ---
        /// <summary>页眉距页面顶端(cm)，null则不设</summary>
        public double? HeaderDistanceCm { get; set; }
        /// <summary>页脚距页面底端(cm)，null则不设</summary>
        public double? FooterDistanceCm { get; set; }
    }

    public class PageNumberConfig
    {
        /// <summary>页码格式</summary>
        public PageNumberFormat Format { get; set; } = PageNumberFormat.Arabic;
        /// <summary>起始页码（null 表示续前节）</summary>
        public int? Start { get; set; }
        /// <summary>是否续前节页码</summary>
        public bool ContinueFromPrevious { get; set; }
        /// <summary>页码位置</summary>
        public PageNumberPosition Position { get; set; } = PageNumberPosition.BottomCenter;
        public FontConfig Font { get; set; } = new() { Latin = "Times New Roman" };
        public double FontSizePt { get; set; } = 10.5;
    }

    public enum PageOrientation { Portrait, Landscape }
    public enum PrintMode { SingleSide, DoubleSide }
    public enum PageNumberFormat { None, Arabic, RomanUpper, RomanLower, LetterUpper, LetterLower }
    public enum PageNumberPosition { BottomCenter, BottomRight, BottomOutside, TopCenter, TopRight, TopOutside }

    /// <summary>目录配置</summary>
    public class TocConfig
    {
        public int MaxLevel { get; set; } = 3;
        public bool ShowPageNumbers { get; set; } = true;
        public bool RightAlignPageNumbers { get; set; } = true;
        public bool UseHyperlinks { get; set; } = true;
        public string LeaderType { get; set; } = "dots";
    }

    /// <summary>打印配置</summary>
    public class PrintConfig
    {
        public bool CoverSingleSide { get; set; } = true;
        public bool TocSingleSide { get; set; } = true;
        public string BodySections { get; set; } = "double_side";
        public bool MirrorMargins { get; set; }
    }
}
