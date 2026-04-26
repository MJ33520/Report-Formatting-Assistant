using System.Collections.Generic;

namespace ReportForge.Core.Models
{
    /// <summary>
    /// 标题层级定义——将标题建模为文档结构层级，而非单纯字体样式。
    /// 包含业务层级、样式层级、大纲级别、编号级别的四维关系。
    /// </summary>
    public class HeadingLevel
    {
        /// <summary>稳定ID，如 "heading.level1"</summary>
        public string Id { get; set; } = "";

        /// <summary>多语言显示名，如 zh-CN:"一级标题", en-US:"Heading 1"</summary>
        public LocalizedString DisplayName { get; set; } = new();

        /// <summary>是否属于正文结构（vs. 附录标题等）</summary>
        public bool IsBodyStructure { get; set; } = true;

        /// <summary>对应的段落样式ID</summary>
        public string StyleId { get; set; } = "";

        /// <summary>大纲级别（1-9），用于导航窗格和目录</summary>
        public int OutlineLevel { get; set; } = 1;

        /// <summary>多级列表中的编号级别（0-based）</summary>
        public int NumberingLevel { get; set; }

        /// <summary>字体配置</summary>
        public FontConfig Font { get; set; } = new();

        /// <summary>字号（磅）</summary>
        public double FontSizePt { get; set; } = 16;

        public bool Bold { get; set; } = true;
        public bool Italic { get; set; }

        public TextAlignment Alignment { get; set; } = TextAlignment.Left;
        public IndentConfig? FirstLineIndent { get; set; }
        public double LeftIndentCm { get; set; }
        public double RightIndentCm { get; set; }
        public LineSpacingConfig LineSpacing { get; set; } = new();
        public double SpaceBeforePt { get; set; }
        public double SpaceAfterPt { get; set; }

        /// <summary>段前分页</summary>
        public bool PageBreakBefore { get; set; }

        /// <summary>是否进入目录</summary>
        public bool IncludeInTOC { get; set; } = true;

        /// <summary>是否参与导航窗格</summary>
        public bool IncludeInNavPane { get; set; } = true;

        /// <summary>是否作为图表按章编号的依据</summary>
        public bool IsCaptionChapterBasis { get; set; }
    }
}
