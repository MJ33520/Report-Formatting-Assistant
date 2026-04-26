namespace ReportForge.Core.Models
{
    /// <summary>表格样式配置</summary>
    public class TableStyleConfig
    {
        public string Id { get; set; } = "RF_StandardTable";
        public TextAlignment TableAlignment { get; set; } = TextAlignment.Center;
        public TableWidthStrategy WidthStrategy { get; set; } = TableWidthStrategy.AutoFitWindow;
        public TableBorderConfig Borders { get; set; } = new();
        public TableCellStyleConfig HeaderRow { get; set; } = new()
        {
            Font = new FontConfig { Zh = "黑体" },
            FontSizePt = 12,
            Bold = true,
            Alignment = TextAlignment.Center,
            RepeatOnNewPage = true
        };
        public TableCellStyleConfig BodyCell { get; set; } = new()
        {
            Font = new FontConfig { Zh = "仿宋_GB2312" },
            FontSizePt = 12,
            LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 23 }
        };
        public CellPaddingConfig CellPadding { get; set; } = new();
        /// <summary>是否允许跨页断行</summary>
        public bool AllowBreakAcrossPages { get; set; }
    }

    public class TableBorderConfig
    {
        public BorderDef Outside { get; set; } = new() { Style = BorderStyle.Single, WidthPt = 1.5 };
        public BorderDef Inside { get; set; } = new() { Style = BorderStyle.Single, WidthPt = 0.5 };
    }

    public class BorderDef
    {
        public BorderStyle Style { get; set; } = BorderStyle.Single;
        public double WidthPt { get; set; } = 0.5;
    }

    public enum BorderStyle { None, Single, Double, Dashed }

    public class TableCellStyleConfig
    {
        public FontConfig Font { get; set; } = new();
        public double FontSizePt { get; set; } = 12;
        public bool Bold { get; set; }
        public TextAlignment Alignment { get; set; } = TextAlignment.Left;
        public LineSpacingConfig? LineSpacing { get; set; }
        public bool RepeatOnNewPage { get; set; }
    }

    public class CellPaddingConfig
    {
        public double TopMm { get; set; }
        public double BottomMm { get; set; }
        public double LeftMm { get; set; } = 1.27;
        public double RightMm { get; set; } = 1.27;
    }

    public enum TableWidthStrategy
    {
        AutoFitWindow,
        AutoFitContent,
        FixedWidth
    }
}
