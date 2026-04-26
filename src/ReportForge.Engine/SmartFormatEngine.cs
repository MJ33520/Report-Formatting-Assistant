using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>
    /// 智能格式化引擎——扫描全文段落，根据编号模式自动识别标题层级，
    /// 支持多种编号体系（中文编号、数字点号、章节关键词），
    /// 批量应用目标样式。
    /// </summary>
    public class SmartFormatEngine
    {
        /// <summary>识别结果：段落索引 → 检测到的层级（0=正文, 1-9=标题级别）</summary>
        public class DetectionResult
        {
            public int ParagraphIndex { get; set; }
            public int DetectedLevel { get; set; } // 0 = 正文, 1-9 = 标题
            public string PreviewText { get; set; } = "";
            public string MatchedPattern { get; set; } = "";
        }

        /// <summary>分析统计</summary>
        public class AnalysisReport
        {
            public List<DetectionResult> Results { get; set; } = new();
            public int[] LevelCounts { get; set; } = new int[10]; // [0]=正文, [1]-[9]=标题
            public string DetectedSystem { get; set; } = ""; // 检测到的编号体系
        }

        // ==================== 正则模式 ====================

        // 数字点号层级: 1 / 1.1 / 1.1.1 / ...
        private static readonly Regex RxDotLevel1 = new(@"^(\d+)\s+\S", RegexOptions.Compiled);
        private static readonly Regex RxDotLevel2 = new(@"^(\d+\.\d+)\s+\S", RegexOptions.Compiled);
        private static readonly Regex RxDotLevel3 = new(@"^(\d+\.\d+\.\d+)\s+\S", RegexOptions.Compiled);
        private static readonly Regex RxDotLevel4 = new(@"^(\d+\.\d+\.\d+\.\d+)\s+\S", RegexOptions.Compiled);
        private static readonly Regex RxDotLevel5 = new(@"^(\d+\.\d+\.\d+\.\d+\.\d+)\s+\S", RegexOptions.Compiled);
        private static readonly Regex RxDotLevel6 = new(@"^(\d+\.\d+\.\d+\.\d+\.\d+\.\d+)\s+\S", RegexOptions.Compiled);

        // 中文编号体系
        private static readonly Regex RxChinese1 = new(@"^[一二三四五六七八九十百]+、", RegexOptions.Compiled);
        private static readonly Regex RxChinese2 = new(@"^[（\(][一二三四五六七八九十百]+[）\)]", RegexOptions.Compiled);
        private static readonly Regex RxArabicDot = new(@"^\d+[.．]\s*\S", RegexOptions.Compiled);
        private static readonly Regex RxArabicParen = new(@"^\d+[)）]\s*\S", RegexOptions.Compiled);
        private static readonly Regex RxCircled = new(@"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]", RegexOptions.Compiled);
        private static readonly Regex RxLetterLower = new(@"^[a-z][)）]\s*\S", RegexOptions.Compiled);
        private static readonly Regex RxParenArabic = new(@"^[（\(]\d+[）\)]\s*\S", RegexOptions.Compiled);

        // 章节关键词
        private static readonly Regex RxChapter = new(@"^第[一二三四五六七八九十百\d]+章", RegexOptions.Compiled);
        private static readonly Regex RxSection = new(@"^第[一二三四五六七八九十百\d]+节", RegexOptions.Compiled);

        /// <summary>
        /// 扫描文档所有段落，识别标题层级
        /// </summary>
        public AnalysisReport Analyze(IWordDocumentAdapter doc)
        {
            var report = new AnalysisReport();
            var paragraphs = doc.GetAllParagraphs().ToList();
            bool hasDotSystem = false;
            bool hasChineseSystem = false;

            // 第一遍：检测主要编号体系
            foreach (var para in paragraphs)
            {
                var text = para.Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(text)) continue;
                if (RxDotLevel2.IsMatch(text) || RxDotLevel3.IsMatch(text)) hasDotSystem = true;
                if (RxChinese1.IsMatch(text) || RxChinese2.IsMatch(text)) hasChineseSystem = true;
            }

            // 确定主编号体系
            if (hasDotSystem && !hasChineseSystem) report.DetectedSystem = "数字点号体系";
            else if (hasChineseSystem && !hasDotSystem) report.DetectedSystem = "中文编号体系";
            else if (hasDotSystem && hasChineseSystem) report.DetectedSystem = "混合编号体系";
            else report.DetectedSystem = "无明确体系";

            // 第二遍：逐段识别
            int index = 0;
            foreach (var para in paragraphs)
            {
                var text = para.Text?.Trim() ?? "";
                int level = 0;
                string pattern = "";

                if (!string.IsNullOrEmpty(text) && !IsSkippableParagraph(text) && !para.IsInsideTable)
                {
                    // === 优先级 0：已有 Word 列表级别（自动编号）===
                    if (para.ListLevel > 0)
                    {
                        level = Math.Min(para.ListLevel, 9);
                        var listStr = para.ListString;
                        pattern = $"列表级别{para.ListLevel}" + (string.IsNullOrEmpty(listStr) ? "" : $" ({listStr})");
                    }
                    // === 优先级 1：已有大纲级别 ===
                    else if (para.OutlineLevel >= 1 && para.OutlineLevel <= 9)
                    {
                        level = para.OutlineLevel;
                        pattern = $"大纲级别{para.OutlineLevel}";
                    }
                    // === 优先级 2：内置标题样式名 ===
                    else if (IsBuiltinHeadingStyle(para.StyleName, out int headingLevel))
                    {
                        level = headingLevel;
                        pattern = $"内置样式 {para.StyleName}";
                    }
                    // === 优先级 3：文本正则匹配（含编号文字拼接）===
                    else
                    {
                        // 把编号文字拼到段落文本前面再匹配
                        var fullText = para.ListString + text;
                        (level, pattern) = DetectLevel(fullText, hasDotSystem, hasChineseSystem);
                    }

                    // 辅助判断：超过80字或以句号结尾的大概率不是标题
                    if (level > 0 && (text.Length > 80 || text.EndsWith("。") || text.EndsWith(".")))
                    {
                        level = 0;
                        pattern = "长段落/句号结尾排除";
                    }
                }

                report.Results.Add(new DetectionResult
                {
                    ParagraphIndex = index,
                    DetectedLevel = level,
                    PreviewText = text.Length > 40 ? text.Substring(0, 40) + "..." : text,
                    MatchedPattern = pattern
                });
                report.LevelCounts[level]++;
                index++;
            }

            return report;
        }

        /// <summary>应用结果统计</summary>
        public class ApplyResult
        {
            public int ImageCount { get; set; }
            public int TableCount { get; set; }
            public int UncaptionedImages { get; set; }
            public int UncaptionedTables { get; set; }
        }

        /// <summary>
        /// 根据分析结果批量应用样式，同时处理图片和表格
        /// </summary>
        /// <param name="levelOffset">级别偏移量。如设为2，则检测到的1级→目标3级，2级→4级...</param>
        public ApplyResult Apply(IWordDocumentAdapter doc, FormatProfile profile, AnalysisReport report, int levelOffset = 0)
        {
            var paragraphs = doc.GetAllParagraphs().ToList();
            var result = new ApplyResult();

            // 题注样式名（用于检测已有题注）
            var captionStyles = new HashSet<string> { "RF_FigureCaption", "RF_TableCaption", "题注", "Caption" };

            for (int i = 0; i < report.Results.Count && i < paragraphs.Count; i++)
            {
                var para = paragraphs[i];
                var detection = report.Results[i];

                // 跳过表格内段落（表格有专门的格式化逻辑）
                if (para.IsInsideTable) continue;

                // === 图片段落：优先用 RF_Figure ===
                if (para.HasInlineImage)
                {
                    result.ImageCount++;
                    try
                    {
                        doc.Styles.ApplyStyle(para.Range, "RF_Figure");
                    }
                    catch { }

                    // 检查下一段是否有题注
                    bool hasCaptionBelow = (i + 1 < paragraphs.Count) &&
                        captionStyles.Contains(paragraphs[i + 1].StyleName);
                    if (!hasCaptionBelow) result.UncaptionedImages++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(detection.PreviewText)) continue;

                string targetStyleId;
                if (detection.DetectedLevel == 0)
                {
                    targetStyleId = profile.BodyStyle.StyleId;
                }
                else
                {
                    int targetLevel = detection.DetectedLevel + levelOffset;
                    if (targetLevel >= 1 && targetLevel <= profile.HeadingLevels.Count)
                    {
                        targetStyleId = profile.HeadingLevels[targetLevel - 1].StyleId;
                    }
                    else
                    {
                        targetStyleId = profile.BodyStyle.StyleId;
                    }
                }

                // 先清直接格式，再应用样式
                try
                {
                    doc.Styles.ClearDirectFormatting(para.Range);
                    doc.Styles.ApplyStyle(para.Range, targetStyleId);
                }
                catch { }
            }

            // === 表格格式化 ===
            foreach (var table in doc.Tables.GetAllTables())
            {
                result.TableCount++;
                try
                {
                    doc.Tables.ApplyStandardFormat(table, profile.TableStyle);
                    doc.Tables.SetHeaderRowRepeat(table, profile.TableStyle.HeaderRow.RepeatOnNewPage);
                    doc.Tables.SetAllowBreakAcrossPages(table, profile.TableStyle.AllowBreakAcrossPages);
                }
                catch { }
            }
            // 简单估算：无题注表格 = 总表格数（精确检测需要更复杂的逻辑）
            result.UncaptionedTables = result.TableCount;

            return result;
        }

        // ==================== 内部方法 ====================

        private static (int level, string pattern) DetectLevel(string text, bool hasDotSystem, bool hasChineseSystem)
        {
            // 章节关键词（最高优先级）
            if (RxChapter.IsMatch(text)) return (1, "第X章");
            if (RxSection.IsMatch(text)) return (2, "第X节");

            // 如果检测到数字点号体系，优先用点号层级
            if (hasDotSystem)
            {
                if (RxDotLevel6.IsMatch(text)) return (6, "N.N.N.N.N.N");
                if (RxDotLevel5.IsMatch(text)) return (5, "N.N.N.N.N");
                if (RxDotLevel4.IsMatch(text)) return (4, "N.N.N.N");
                if (RxDotLevel3.IsMatch(text)) return (3, "N.N.N");
                if (RxDotLevel2.IsMatch(text)) return (2, "N.N");
                if (RxDotLevel1.IsMatch(text)) return (1, "N");
            }

            // 中文编号体系
            if (RxChinese1.IsMatch(text)) return (1, "一、");
            if (RxChinese2.IsMatch(text)) return (2, "（一）");
            if (RxArabicDot.IsMatch(text)) return (3, "1.");
            if (RxArabicParen.IsMatch(text)) return (4, "1)");
            if (RxCircled.IsMatch(text)) return (5, "①");
            if (RxLetterLower.IsMatch(text)) return (6, "a)");
            if (RxParenArabic.IsMatch(text)) return (7, "(1)");

            // 纯数字开头（如果没有点号体系，单独数字可能是一级标题）
            if (!hasDotSystem && RxDotLevel1.IsMatch(text)) return (1, "N");

            return (0, "正文");
        }

        private static bool IsSkippableParagraph(string text)
        {
            // 跳过空段、目录标记、图表等
            if (text.StartsWith("图") && text.Contains("-") && text.Length < 30) return true;
            if (text.StartsWith("表") && text.Contains("-") && text.Length < 30) return true;
            if (text == "目录" || text == "目  录") return true;
            return false;
        }

        private static bool IsBuiltinHeadingStyle(string styleName, out int level)
        {
            level = 0;
            if (string.IsNullOrEmpty(styleName)) return false;

            // Word 中文内置标题样式
            var zhMap = new Dictionary<string, int>
            {
                ["标题 1"] = 1, ["标题 2"] = 2, ["标题 3"] = 3,
                ["标题 4"] = 4, ["标题 5"] = 5, ["标题 6"] = 6,
                ["标题 7"] = 7, ["标题 8"] = 8, ["标题 9"] = 9,
            };
            if (zhMap.TryGetValue(styleName, out level)) return true;

            // Word 英文内置标题样式
            if (styleName.StartsWith("Heading ", StringComparison.OrdinalIgnoreCase))
            {
                var numPart = styleName.Substring(8).Trim();
                if (int.TryParse(numPart, out level) && level >= 1 && level <= 9) return true;
            }

            return false;
        }
    }
}
