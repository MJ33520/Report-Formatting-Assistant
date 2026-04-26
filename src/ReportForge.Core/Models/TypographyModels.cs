namespace ReportForge.Core.Models
{
    /// <summary>字体配置，支持中文/西文分别指定</summary>
    public class FontConfig
    {
        /// <summary>中文字体族，如"黑体""楷体""仿宋_GB2312"</summary>
        public string Zh { get; set; } = "仿宋_GB2312";

        /// <summary>西文/数字字体族，如"Times New Roman"</summary>
        public string Latin { get; set; } = "Times New Roman";
    }

    /// <summary>行距配置</summary>
    public class LineSpacingConfig
    {
        /// <summary>行距规则：Exact（固定值）、Multiple（倍数）、AtLeast（最小值）</summary>
        public LineSpacingRule Rule { get; set; } = LineSpacingRule.Exact;

        /// <summary>行距数值（磅或倍数，取决于Rule）</summary>
        public double Value { get; set; } = 30;
    }

    public enum LineSpacingRule
    {
        Exact,      // 固定值
        Multiple,   // 倍数
        AtLeast     // 最小值
    }

    /// <summary>缩进配置</summary>
    public class IndentConfig
    {
        /// <summary>缩进类型：None, Chars, Cm, Pt</summary>
        public IndentUnit Unit { get; set; } = IndentUnit.Chars;
        public double Value { get; set; } = 2;
    }

    public enum IndentUnit
    {
        None,
        Chars,
        Cm,
        Pt
    }

    /// <summary>对齐方式</summary>
    public enum TextAlignment
    {
        Left,
        Center,
        Right,
        Justify     // 两端对齐
    }
}
