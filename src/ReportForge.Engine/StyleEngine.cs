using System.Collections.Generic;
using System.Linq;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>
    /// 样式引擎——负责在 Word 文档中创建/更新目标样式体系，清理历史样式。
    /// 核心原则：段落样式管结构，字符样式管局部，直接格式视为脏数据。
    /// </summary>
    public class StyleEngine : IStyleEngine
    {
        /// <summary>
        /// 根据 FormatProfile 在文档中创建/更新全部目标样式
        /// </summary>
        public void InitializeTargetStyles(FormatProfile profile, IWordDocumentAdapter doc)
        {
            // 1. 创建正文样式
            doc.Styles.CreateOrUpdateParagraphStyle(profile.BodyStyle);

            // 2. 创建标题样式
            foreach (var heading in profile.HeadingLevels)
            {
                var styleDef = HeadingToStyleDef(heading);
                doc.Styles.CreateOrUpdateParagraphStyle(styleDef);
            }

            // 3. 创建题注样式
            CreateCaptionStyle(doc, profile.CaptionConfig.Figure);
            CreateCaptionStyle(doc, profile.CaptionConfig.Table);

            // 4. 创建附加样式（页眉页脚、图注表注等）
            foreach (var style in profile.AdditionalStyles)
            {
                if (style.Type == StyleType.Paragraph)
                    doc.Styles.CreateOrUpdateParagraphStyle(style);
                else
                    doc.Styles.CreateOrUpdateCharacterStyle(style);
            }
        }

        /// <summary>清理模板中的历史遗留/无用自定义样式</summary>
        public void CleanupOrphanStyles(IWordDocumentAdapter doc, CleanupOptions options)
        {
            var unusedStyles = doc.Styles.GetUnusedCustomStyles().ToList();

            foreach (var styleName in unusedStyles)
            {
                if (options.WhitelistedStyleIds.Contains(styleName))
                    continue;

                if (!options.DryRun && options.DeleteUnused)
                {
                    try { doc.Styles.DeleteCustomStyle(styleName); }
                    catch { /* 某些内建样式无法删除，静默跳过 */ }
                }
            }
        }

        /// <summary>对当前选区应用指定样式</summary>
        public void ApplyStyleToSelection(IWordDocumentAdapter doc, string styleId)
        {
            var selection = doc.GetSelection();
            // 先清除直接格式，再应用样式
            doc.Styles.ClearDirectFormatting(selection);
            doc.Styles.ApplyStyle(selection, styleId);
        }

        /// <summary>清除指定范围的直接格式</summary>
        public void ClearDirectFormatting(IWordDocumentAdapter doc, IRangeAdapter range)
        {
            doc.Styles.ClearDirectFormatting(range);
        }

        private static StyleDefinition HeadingToStyleDef(HeadingLevel heading)
        {
            return new StyleDefinition
            {
                StyleId = heading.StyleId,
                DisplayName = heading.DisplayName,
                Type = StyleType.Paragraph,
                Font = heading.Font,
                FontSizePt = heading.FontSizePt,
                Bold = heading.Bold,
                Italic = heading.Italic,
                Alignment = heading.Alignment,
                FirstLineIndent = heading.FirstLineIndent,
                LineSpacing = heading.LineSpacing,
                SpaceBeforePt = heading.SpaceBeforePt,
                SpaceAfterPt = heading.SpaceAfterPt,
                OutlineLevel = heading.OutlineLevel,
                PageBreakBefore = heading.PageBreakBefore
            };
        }

        private static void CreateCaptionStyle(IWordDocumentAdapter doc, CaptionTypeConfig config)
        {
            var def = new StyleDefinition
            {
                StyleId = config.StyleId,
                DisplayName = config.Label,
                Type = StyleType.Paragraph,
                Font = config.Font,
                FontSizePt = config.FontSizePt,
                Alignment = config.Alignment
            };
            doc.Styles.CreateOrUpdateParagraphStyle(def);
        }
    }
}
