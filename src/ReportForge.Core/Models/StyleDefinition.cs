namespace ReportForge.Core.Models
{
    /// <summary>
    /// 样式定义——段落样式或字符样式的完整描述。
    /// 与 Word Style 对象一一映射。
    /// </summary>
    public class StyleDefinition
    {
        /// <summary>内部稳定ID，如 "RF_Heading1"、"RF_BodyText"</summary>
        public string StyleId { get; set; } = "";

        /// <summary>多语言显示名</summary>
        public LocalizedString DisplayName { get; set; } = new();

        /// <summary>样式类型：Paragraph 或 Character</summary>
        public StyleType Type { get; set; } = StyleType.Paragraph;

        /// <summary>基于的父样式ID（可选）</summary>
        public string? BasedOnStyleId { get; set; }

        /// <summary>后续段落样式ID（可选）</summary>
        public string? NextParagraphStyleId { get; set; }

        // --- 字体属性 ---
        public FontConfig Font { get; set; } = new();
        public double FontSizePt { get; set; } = 14; // 四号 = 14pt
        public bool Bold { get; set; }
        public bool Italic { get; set; }

        // --- 段落属性（仅 Paragraph 样式有效）---
        public TextAlignment Alignment { get; set; } = TextAlignment.Justify;
        public IndentConfig? FirstLineIndent { get; set; }
        public double LeftIndentCm { get; set; }
        public double RightIndentCm { get; set; }
        public LineSpacingConfig LineSpacing { get; set; } = new();
        public double SpaceBeforePt { get; set; }
        public double SpaceAfterPt { get; set; }

        /// <summary>大纲级别（1-9），0 表示正文</summary>
        public int OutlineLevel { get; set; }

        /// <summary>段前分页</summary>
        public bool PageBreakBefore { get; set; }
    }

    public enum StyleType
    {
        Paragraph,
        Character
    }
}
