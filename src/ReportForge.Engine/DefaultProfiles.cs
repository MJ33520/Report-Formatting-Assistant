using System.Collections.Generic;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>内置默认配置工厂——基于真实格式要求生成标准 Profile</summary>
    public static class DefaultProfiles
    {
        /// <summary>
        /// 政府公文报告格式（来源：用户提供的真实格式要求截图）
        /// 页面: A4, 上下3cm, 左右2.6cm, 页眉1.5cm, 页脚2cm
        /// 正文: 仿宋_GB2312 四号(14pt), 固定值30磅, 两端对齐, 首行缩进2字符
        /// 一级标题: 黑体 三号(16pt); 二级: 楷体 三号(16pt); 三四级: 仿宋_GB2312 四号(14pt) 加粗
        /// 编号: 一→（一）→1.→1)→①→a)→(1)→i)→A)
        /// 图表: 按一级标题-序号编号, 如图3-3, 表2-2
        /// 表格: 外框1.5磅, 表头黑体小四(12pt), 表内仿宋小四(12pt), 行距23磅
        /// 页码: 四号(14pt) Times New Roman, 居中
        /// </summary>
        public static FormatProfile CreateGovReport(string locale = "zh-CN")
        {
            return new FormatProfile
            {
                Id = "gov-report-default",
                Version = "2.0",
                Locale = locale,
                DisplayName = new LocalizedString
                {
                    ["zh-CN"] = "政府公文报告格式",
                    ["en-US"] = "Government Report Format"
                },

                PageSetup = new PageSetup
                {
                    PaperSize = "A4",
                    Margins = new MarginConfig { TopCm = 3.0, BottomCm = 3.0, LeftCm = 2.6, RightCm = 2.6, GutterCm = 0 },
                    HeaderDistanceCm = 1.5,
                    FooterDistanceCm = 2.0
                },

                // ==================== 9级标题 ====================
                HeadingLevels = new List<HeadingLevel>
                {
                    // 一级: 黑体 三号(16pt) 加粗
                    new()
                    {
                        Id = "heading.level1",
                        DisplayName = new LocalizedString { ["zh-CN"] = "一级标题", ["en-US"] = "Heading 1" },
                        StyleId = "RF_Heading1", OutlineLevel = 1, NumberingLevel = 0,
                        Font = new FontConfig { Zh = "黑体", Latin = "Times New Roman" },
                        FontSizePt = 16, // 三号
                        Bold = true, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 0 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = true, IncludeInNavPane = true, IsCaptionChapterBasis = true
                    },
                    // 二级: 楷体 三号(16pt)
                    new()
                    {
                        Id = "heading.level2",
                        DisplayName = new LocalizedString { ["zh-CN"] = "二级标题", ["en-US"] = "Heading 2" },
                        StyleId = "RF_Heading2", OutlineLevel = 2, NumberingLevel = 1,
                        Font = new FontConfig { Zh = "楷体", Latin = "Times New Roman" },
                        FontSizePt = 16, // 三号
                        Bold = false, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = true, IncludeInNavPane = true
                    },
                    // 三级: 仿宋_GB2312 四号(14pt) 加粗
                    new()
                    {
                        Id = "heading.level3",
                        DisplayName = new LocalizedString { ["zh-CN"] = "三级标题", ["en-US"] = "Heading 3" },
                        StyleId = "RF_Heading3", OutlineLevel = 3, NumberingLevel = 2,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, // 四号
                        Bold = true, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = true
                    },
                    // 四级: 仿宋_GB2312 四号(14pt) 加粗
                    new()
                    {
                        Id = "heading.level4",
                        DisplayName = new LocalizedString { ["zh-CN"] = "四级标题", ["en-US"] = "Heading 4" },
                        StyleId = "RF_Heading4", OutlineLevel = 4, NumberingLevel = 3,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, Bold = true, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = false
                    },
                    // 五级: 仿宋_GB2312 四号(14pt)
                    new()
                    {
                        Id = "heading.level5",
                        DisplayName = new LocalizedString { ["zh-CN"] = "五级标题", ["en-US"] = "Heading 5" },
                        StyleId = "RF_Heading5", OutlineLevel = 5, NumberingLevel = 4,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, Bold = false, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = false
                    },
                    // 六级: 仿宋_GB2312 四号(14pt)
                    new()
                    {
                        Id = "heading.level6",
                        DisplayName = new LocalizedString { ["zh-CN"] = "六级标题", ["en-US"] = "Heading 6" },
                        StyleId = "RF_Heading6", OutlineLevel = 6, NumberingLevel = 5,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, Bold = false, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = false
                    },
                    // 七级: 仿宋_GB2312 四号(14pt)
                    new()
                    {
                        Id = "heading.level7",
                        DisplayName = new LocalizedString { ["zh-CN"] = "七级标题", ["en-US"] = "Heading 7" },
                        StyleId = "RF_Heading7", OutlineLevel = 7, NumberingLevel = 6,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, Bold = false, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = false
                    },
                    // 八级
                    new()
                    {
                        Id = "heading.level8",
                        DisplayName = new LocalizedString { ["zh-CN"] = "八级标题", ["en-US"] = "Heading 8" },
                        StyleId = "RF_Heading8", OutlineLevel = 8, NumberingLevel = 7,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, Bold = false, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = false
                    },
                    // 九级
                    new()
                    {
                        Id = "heading.level9",
                        DisplayName = new LocalizedString { ["zh-CN"] = "九级标题", ["en-US"] = "Heading 9" },
                        StyleId = "RF_Heading9", OutlineLevel = 9, NumberingLevel = 8,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14, Bold = false, Alignment = TextAlignment.Left,
                        FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 },
                        IncludeInTOC = false
                    }
                },

                // ==================== 编号方案 ====================
                // 图片要求: 一→（一）→1→1)→⑴→①
                NumberingScheme = new NumberingScheme
                {
                    Id = "gov-cn-mixed",
                    DisplayName = new LocalizedString { ["zh-CN"] = "政府公文混合编号", ["en-US"] = "Gov Mixed Numbering" },
                    Levels = new List<NumberingLevelDef>
                    {
                        new() { Level = 0, Format = NumberFormat.ChineseCounting, Suffix = "、", Example = "一、二、三、" },
                        new() { Level = 1, Format = NumberFormat.ChineseCounting, Prefix = "（", Suffix = "）", Example = "（一）（二）" },
                        new() { Level = 2, Format = NumberFormat.Arabic, Suffix = ".", Example = "1. 2. 3." },
                        new() { Level = 3, Format = NumberFormat.Arabic, Suffix = ")", Example = "1) 2) 3)" },
                        new() { Level = 4, Format = NumberFormat.CircledNumber, Example = "①②③" },
                        new() { Level = 5, Format = NumberFormat.LetterLower, Suffix = ")", Example = "a) b) c)" },
                        new() { Level = 6, Format = NumberFormat.Arabic, Prefix = "(", Suffix = ")", Example = "(1)(2)(3)" },
                        new() { Level = 7, Format = NumberFormat.RomanLower, Suffix = ")", Example = "i) ii) iii)" },
                        new() { Level = 8, Format = NumberFormat.LetterUpper, Suffix = ")", Example = "A) B) C)" }
                    }
                },

                // ==================== 正文 ====================
                BodyStyle = new StyleDefinition
                {
                    StyleId = "RF_BodyText",
                    DisplayName = new LocalizedString { ["zh-CN"] = "RF正文", ["en-US"] = "RF Body Text" },
                    Type = StyleType.Paragraph,
                    Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                    FontSizePt = 14, // 四号
                    Alignment = TextAlignment.Justify,
                    FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars, Value = 2 },
                    LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 30 }
                },

                CaptionConfig = new CaptionConfiguration(),

                // ==================== 表格 ====================
                // 外框1.5磅, 表头黑体小四(12pt), 表内仿宋_GB2312小四(12pt), 行距23磅
                TableStyle = new TableStyleConfig
                {
                    Borders = new TableBorderConfig
                    {
                        Outside = new BorderDef { Style = BorderStyle.Single, WidthPt = 1.5 },
                        Inside = new BorderDef { Style = BorderStyle.Single, WidthPt = 0.5 }
                    },
                    HeaderRow = new TableCellStyleConfig
                    {
                        Font = new FontConfig { Zh = "黑体", Latin = "Times New Roman" },
                        FontSizePt = 12, // 小四
                        Bold = true,
                        Alignment = TextAlignment.Center,
                        RepeatOnNewPage = true
                    },
                    BodyCell = new TableCellStyleConfig
                    {
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 12, // 小四
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 23 }
                    },
                },

                // ==================== 分节 ====================
                SectionTemplates = new List<SectionTemplate>
                {
                    new()
                    {
                        Id = "cover", DisplayName = "封面",
                        ShowHeader = false, ShowFooter = false,
                        PageNumbering = new PageNumberConfig { Format = PageNumberFormat.None },
                        PrintMode = PrintMode.SingleSide, LinkToPrevious = false,
                        CustomMargins = new MarginConfig { TopCm = 3.0, BottomCm = 3.0, LeftCm = 2.6, RightCm = 2.6, GutterCm = 0 },
                        HeaderDistanceCm = 1.5, FooterDistanceCm = 2.0
                    },
                    new()
                    {
                        Id = "toc", DisplayName = "目录",
                        ShowHeader = false, ShowFooter = true,
                        PageNumbering = new PageNumberConfig
                        {
                            Format = PageNumberFormat.RomanUpper, Start = 1,
                            Position = PageNumberPosition.BottomCenter,
                            Font = new FontConfig { Latin = "Times New Roman" },
                            FontSizePt = 14 // 四号
                        },
                        PrintMode = PrintMode.SingleSide, LinkToPrevious = false,
                        CustomMargins = new MarginConfig { TopCm = 3.0, BottomCm = 3.0, LeftCm = 2.6, RightCm = 2.6, GutterCm = 0 },
                        HeaderDistanceCm = 1.5, FooterDistanceCm = 2.0
                    },
                    new()
                    {
                        Id = "main_body", DisplayName = "正文",
                        ShowHeader = true, ShowFooter = true,
                        HeaderText = "", // 由用户设置具体页眉文字
                        HeaderFont = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        HeaderFontSizePt = 12, // 小四
                        HeaderBorderBottom = true,
                        PageNumbering = new PageNumberConfig
                        {
                            Format = PageNumberFormat.Arabic, Start = 1,
                            Position = PageNumberPosition.BottomCenter,
                            Font = new FontConfig { Latin = "Times New Roman" },
                            FontSizePt = 14 // 四号
                        },
                        PrintMode = PrintMode.DoubleSide, LinkToPrevious = false,
                        CustomMargins = new MarginConfig { TopCm = 3.0, BottomCm = 3.0, LeftCm = 2.6, RightCm = 2.6, GutterCm = 0 },
                        HeaderDistanceCm = 1.5, FooterDistanceCm = 2.0
                    },
                    new()
                    {
                        Id = "landscape_page", DisplayName = "横向页面",
                        ShowHeader = true, ShowFooter = true,
                        Orientation = PageOrientation.Landscape,
                        PageNumbering = new PageNumberConfig { ContinueFromPrevious = true },
                        CustomMargins = new MarginConfig { TopCm = 3.0, BottomCm = 3.0, LeftCm = 2.6, RightCm = 2.6, GutterCm = 0 },
                        HeaderDistanceCm = 1.5, FooterDistanceCm = 2.0
                    },
                    new()
                    {
                        Id = "appendix", DisplayName = "附录",
                        ShowHeader = true, ShowFooter = true,
                        PageNumbering = new PageNumberConfig { ContinueFromPrevious = true },
                        CustomMargins = new MarginConfig { TopCm = 3.0, BottomCm = 3.0, LeftCm = 2.6, RightCm = 2.6, GutterCm = 0 },
                        HeaderDistanceCm = 1.5, FooterDistanceCm = 2.0
                    }
                },

                TocConfig = new TocConfig { MaxLevel = 3 },
                PrintConfig = new PrintConfig { CoverSingleSide = true, TocSingleSide = true },

                // ==================== 附加样式 ====================
                AdditionalStyles = new List<StyleDefinition>
                {
                    new()
                    {
                        StyleId = "RF_Figure",
                        DisplayName = new LocalizedString { ["zh-CN"] = "图片段落", ["en-US"] = "Figure Paragraph" },
                        Type = StyleType.Paragraph,
                        Font = new FontConfig { Zh = "仿宋_GB2312", Latin = "Times New Roman" },
                        FontSizePt = 14,
                        Alignment = TextAlignment.Center,
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Multiple, Value = 1.0 },
                        SpaceBeforePt = 6, SpaceAfterPt = 0,
                    },
                    new()
                    {
                        StyleId = "RF_FigureCaption",
                        DisplayName = new LocalizedString { ["zh-CN"] = "图题样式", ["en-US"] = "Figure Caption" },
                        Type = StyleType.Paragraph,
                        Font = new FontConfig { Zh = "黑体", Latin = "Times New Roman" },
                        FontSizePt = 10.5, // 五号
                        Alignment = TextAlignment.Center,
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 22 },
                        SpaceBeforePt = 0, SpaceAfterPt = 6,
                    },
                    new()
                    {
                        StyleId = "RF_TableCaption",
                        DisplayName = new LocalizedString { ["zh-CN"] = "表题样式", ["en-US"] = "Table Caption" },
                        Type = StyleType.Paragraph,
                        Font = new FontConfig { Zh = "黑体", Latin = "Times New Roman" },
                        FontSizePt = 10.5,
                        Alignment = TextAlignment.Center,
                        LineSpacing = new LineSpacingConfig { Rule = LineSpacingRule.Exact, Value = 22 },
                        SpaceBeforePt = 6, SpaceAfterPt = 0,
                    }
                }
            };
        }
    }
}
