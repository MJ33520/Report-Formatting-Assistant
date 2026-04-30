using System;
using System.Collections.Generic;
using System.Linq;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ReportForge.WordAdapter
{
    /// <summary>
    /// Word 样式适配器——创建/更新/删除/应用段落样式和字符样式。
    /// 核心原则：通过样式控制格式，直接格式视为脏数据需清理。
    /// </summary>
    public class WordStyleAdapter : IStyleAdapter
    {
        private readonly Word.Document _doc;

        public WordStyleAdapter(Word.Document doc) { _doc = doc; }

        public bool StyleExists(string styleId)
        {
            try { var _ = _doc.Styles[styleId]; return true; }
            catch { return false; }
        }

        public void CreateOrUpdateParagraphStyle(StyleDefinition def)
        {
            Word.Style style;
            if (StyleExists(def.StyleId))
            {
                style = _doc.Styles[def.StyleId];
            }
            else
            {
                style = _doc.Styles.Add(def.StyleId, Word.WdStyleType.wdStyleTypeParagraph);
            }

            try
            {
                var desiredName = def.DisplayName.Get();
                if (!string.IsNullOrEmpty(desiredName) && style.NameLocal != desiredName && !style.NameLocal.StartsWith(desiredName + ","))
                {
                    // 使用别名机制，保留 StyleId 作为别名以便后续可以通过 StyleId 获取
                    style.NameLocal = desiredName + "," + def.StyleId;
                }
            }
            catch { /* 显示名冲突时保留原有名字 */ }

            // 字体
            style.Font.NameFarEast = def.Font.Zh;
            style.Font.Name = def.Font.Latin;
            style.Font.Size = (float)def.FontSizePt;
            style.Font.Bold = def.Bold ? 1 : 0;
            style.Font.Italic = def.Italic ? 1 : 0;

            // 段落格式
            var pf = style.ParagraphFormat;
            pf.Alignment = ToWordAlignment(def.Alignment);

            // 行距
            if (def.LineSpacing != null)
            {
                switch (def.LineSpacing.Rule)
                {
                    case LineSpacingRule.Exact:
                        pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                        pf.LineSpacing = (float)def.LineSpacing.Value;
                        break;
                    case LineSpacingRule.Multiple:
                        pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                        pf.LineSpacing = (float)(def.LineSpacing.Value * 12);
                        break;
                    case LineSpacingRule.AtLeast:
                        pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;
                        pf.LineSpacing = (float)def.LineSpacing.Value;
                        break;
                }
            }

            // 段前段后
            pf.SpaceBefore = (float)def.SpaceBeforePt;
            pf.SpaceAfter = (float)def.SpaceAfterPt;

            // 首行缩进
            if (def.FirstLineIndent != null)
            {
                switch (def.FirstLineIndent.Unit)
                {
                    case IndentUnit.Chars:
                        pf.CharacterUnitFirstLineIndent = (float)def.FirstLineIndent.Value;
                        break;
                    case IndentUnit.Cm:
                        pf.FirstLineIndent = CmToPoints(def.FirstLineIndent.Value);
                        break;
                    case IndentUnit.Pt:
                        pf.FirstLineIndent = (float)def.FirstLineIndent.Value;
                        break;
                }
            }

            // 大纲级别
            if (def.OutlineLevel >= 1 && def.OutlineLevel <= 9)
            {
                pf.OutlineLevel = (Word.WdOutlineLevel)def.OutlineLevel;
            }
            else
            {
                pf.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            }

            // 段前分页
            pf.PageBreakBefore = def.PageBreakBefore ? 1 : 0;
        }

        public void CreateOrUpdateCharacterStyle(StyleDefinition def)
        {
            Word.Style style;
            if (StyleExists(def.StyleId))
            {
                style = _doc.Styles[def.StyleId];
            }
            else
            {
                style = _doc.Styles.Add(def.StyleId, Word.WdStyleType.wdStyleTypeCharacter);
            }

            try
            {
                var desiredName = def.DisplayName.Get();
                if (!string.IsNullOrEmpty(desiredName) && style.NameLocal != desiredName)
                    style.NameLocal = desiredName;
            }
            catch { }

            style.Font.NameFarEast = def.Font.Zh;
            style.Font.Name = def.Font.Latin;
            style.Font.Size = (float)def.FontSizePt;
            style.Font.Bold = def.Bold ? 1 : 0;
            style.Font.Italic = def.Italic ? 1 : 0;
        }

        public void ApplyStyle(IRangeAdapter range, string styleId)
        {
            var wordRange = ((WordRangeAdapter)range)._range;
            // 如果样式不存在，忽略（应先调用 InitializeStyles）
            if (!StyleExists(styleId)) return;
            object styleObj = _doc.Styles[styleId];
            wordRange.set_Style(ref styleObj);
        }

        public void ClearDirectFormatting(IRangeAdapter range)
        {
            var wordRange = ((WordRangeAdapter)range)._range;

            // 跳过目录区域——检查选区是否在TOC字段内
            foreach (Word.Field field in _doc.Fields)
            {
                if (field.Type == Word.WdFieldType.wdFieldTOC ||
                    field.Type == Word.WdFieldType.wdFieldTOCEntry)
                {
                    // 如果目标 range 与 TOC 字段重叠，收缩范围跳过
                    if (wordRange.Start >= field.Result.Start && wordRange.Start <= field.Result.End)
                        return; // 完全在 TOC 内，直接跳过
                    if (wordRange.End >= field.Result.Start && wordRange.End <= field.Result.End)
                        return; // 尾部在 TOC 内，跳过
                }
            }

            wordRange.Font.Reset();
            wordRange.ParagraphFormat.Reset();
        }

        public void DeleteCustomStyle(string styleId)
        {
            try
            {
                var style = _doc.Styles[styleId];
                if (style.BuiltIn == false)
                {
                    style.Delete();
                }
            }
            catch { /* 静默跳过无法删除的样式 */ }
        }

        public IEnumerable<string> GetAllCustomStyleNames()
        {
            var names = new List<string>();
            foreach (Word.Style style in _doc.Styles)
            {
                if (!style.BuiltIn)
                    names.Add(style.NameLocal);
            }
            return names;
        }

        public IEnumerable<string> GetUnusedCustomStyles()
        {
            // 收集文档中实际使用的样式
            var usedStyles = new HashSet<string>();
            foreach (Word.Paragraph para in _doc.Paragraphs)
            {
                usedStyles.Add(((Word.Style)para.get_Style()).NameLocal);
            }

            // 找出未使用的自定义样式
            return GetAllCustomStyleNames().Where(name => !usedStyles.Contains(name));
        }

        private static Word.WdParagraphAlignment ToWordAlignment(TextAlignment alignment)
        {
            switch (alignment)
            {
                case TextAlignment.Left: return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                case TextAlignment.Center: return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case TextAlignment.Right: return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case TextAlignment.Justify: return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                default: return Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }
        }

        private static float CmToPoints(double cm) => (float)(cm * 28.3465);
    }
}
