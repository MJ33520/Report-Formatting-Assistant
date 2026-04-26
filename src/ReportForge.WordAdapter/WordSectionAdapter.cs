using System;
using System.Collections.Generic;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ReportForge.WordAdapter
{
    /// <summary>分节适配器——管理Word文档的Section对象</summary>
    public class WordSectionAdapter : ISectionAdapter
    {
        private readonly Word.Document _doc;
        public WordSectionAdapter(Word.Document doc) { _doc = doc; }

        public int SectionCount => _doc.Sections.Count;

        public void InsertSectionBreak(IRangeAdapter position, SectionBreakType type)
        {
            var range = ((WordRangeAdapter)position)._range;
            var wdBreak = ToWordBreakType(type);
            range.InsertBreak(wdBreak);
        }

        public void ConfigureSection(int sectionIndex, SectionTemplate template)
        {
            var section = _doc.Sections[sectionIndex];
            var ps = section.PageSetup;
            ps.Orientation = template.Orientation == PageOrientation.Landscape
                ? Word.WdOrientation.wdOrientLandscape
                : Word.WdOrientation.wdOrientPortrait;
        }

        public void SetPageNumbering(int sectionIndex, PageNumberConfig config)
        {
            var section = _doc.Sections[sectionIndex];
            var footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

            if (config.Format == PageNumberFormat.None)
            {
                // 无页码——清空页脚
                footer.Range.Text = "";
                return;
            }

            // 取消链接前一节
            if (sectionIndex > 1)
                footer.LinkToPrevious = false;

            // 设置页码格式
            footer.PageNumbers.NumberStyle = ToWordPageNumberStyle(config.Format);
            if (config.Start.HasValue)
            {
                footer.PageNumbers.RestartNumberingAtSection = true;
                footer.PageNumbers.StartingNumber = config.Start.Value;
            }
            else if (config.ContinueFromPrevious)
            {
                footer.PageNumbers.RestartNumberingAtSection = false;
            }

            // 添加页码字段
            var alignment = ToWordPageNumAlignment(config.Position);
            footer.PageNumbers.Add(alignment, FirstPage: true);

            // 设置页码字体
            footer.Range.Font.Name = config.Font.Latin;
            footer.Range.Font.Size = (float)config.FontSizePt;
        }

        public void SetOrientation(int sectionIndex, PageOrientation orientation)
        {
            var section = _doc.Sections[sectionIndex];
            section.PageSetup.Orientation = orientation == PageOrientation.Landscape
                ? Word.WdOrientation.wdOrientLandscape
                : Word.WdOrientation.wdOrientPortrait;
        }

        public void SetMargins(int sectionIndex, MarginConfig margins)
        {
            var ps = _doc.Sections[sectionIndex].PageSetup;
            ps.TopMargin = CmToPoints(margins.TopCm);
            ps.BottomMargin = CmToPoints(margins.BottomCm);
            ps.LeftMargin = CmToPoints(margins.LeftCm);
            ps.RightMargin = CmToPoints(margins.RightCm);
            ps.Gutter = CmToPoints(margins.GutterCm);
        }

        public void UnlinkHeaderFooter(int sectionIndex)
        {
            var section = _doc.Sections[sectionIndex];
            foreach (Word.HeaderFooter hf in section.Headers)
                hf.LinkToPrevious = false;
            foreach (Word.HeaderFooter hf in section.Footers)
                hf.LinkToPrevious = false;
        }

        /// <summary>根据SectionTemplate配置页眉页脚</summary>
        public void SetupHeaderFooter(int sectionIndex, SectionTemplate template)
        {
            var section = _doc.Sections[sectionIndex];

            // 取消与前一节的链接
            if (sectionIndex > 1)
            {
                section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            }

            // === 页眉 ===
            var header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            if (!template.ShowHeader)
            {
                // 清除页眉内容和边框
                header.Range.Text = "";
                header.Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            }
            else
            {
                // 设置页眉文字
                if (!string.IsNullOrEmpty(template.HeaderText))
                    header.Range.Text = template.HeaderText;

                // 设置字体
                header.Range.Font.NameFarEast = template.HeaderFont.Zh;
                header.Range.Font.Name = template.HeaderFont.Latin;
                header.Range.Font.Size = (float)template.HeaderFontSizePt;
                header.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // 设置下边线
                if (template.HeaderBorderBottom)
                {
                    var border = header.Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom];
                    border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    border.LineWidth = Word.WdLineWidth.wdLineWidth100pt; // 1磅
                }
                else
                {
                    header.Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                }
            }

            // === 页脚 ===
            var footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            if (!template.ShowFooter)
            {
                footer.Range.Text = "";
                footer.PageNumbers.RestartNumberingAtSection = true;
            }
            // 页脚页码由 SetPageNumbering 单独处理
        }

        private static float CmToPoints(double cm) => (float)(cm * 28.3465);

        private static Word.WdBreakType ToWordBreakType(SectionBreakType type)
        {
            switch (type)
            {
                case SectionBreakType.NextPage: return Word.WdBreakType.wdSectionBreakNextPage;
                case SectionBreakType.Continuous: return Word.WdBreakType.wdSectionBreakContinuous;
                case SectionBreakType.EvenPage: return Word.WdBreakType.wdSectionBreakEvenPage;
                case SectionBreakType.OddPage: return Word.WdBreakType.wdSectionBreakOddPage;
                default: return Word.WdBreakType.wdSectionBreakNextPage;
            }
        }

        private static Word.WdPageNumberStyle ToWordPageNumberStyle(PageNumberFormat fmt)
        {
            switch (fmt)
            {
                case PageNumberFormat.Arabic: return Word.WdPageNumberStyle.wdPageNumberStyleArabic;
                case PageNumberFormat.RomanUpper: return Word.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
                case PageNumberFormat.RomanLower: return Word.WdPageNumberStyle.wdPageNumberStyleLowercaseRoman;
                case PageNumberFormat.LetterUpper: return Word.WdPageNumberStyle.wdPageNumberStyleUppercaseLetter;
                case PageNumberFormat.LetterLower: return Word.WdPageNumberStyle.wdPageNumberStyleLowercaseLetter;
                default: return Word.WdPageNumberStyle.wdPageNumberStyleArabic;
            }
        }

        private static Word.WdPageNumberAlignment ToWordPageNumAlignment(PageNumberPosition pos)
        {
            switch (pos)
            {
                case PageNumberPosition.BottomCenter:
                case PageNumberPosition.TopCenter:
                    return Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
                case PageNumberPosition.BottomRight:
                case PageNumberPosition.TopRight:
                    return Word.WdPageNumberAlignment.wdAlignPageNumberRight;
                case PageNumberPosition.BottomOutside:
                case PageNumberPosition.TopOutside:
                    return Word.WdPageNumberAlignment.wdAlignPageNumberOutside;
                default:
                    return Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
            }
        }

        public void SetHeaderFooterDistance(int sectionIndex, double? headerCm, double? footerCm)
        {
            if (sectionIndex < 1 || sectionIndex > _doc.Sections.Count) return;
            var ps = _doc.Sections[sectionIndex].PageSetup;
            if (headerCm.HasValue) ps.HeaderDistance = CmToPoints(headerCm.Value);
            if (footerCm.HasValue) ps.FooterDistance = CmToPoints(footerCm.Value);
        }

        public void SetPaperSize(int sectionIndex, string paperSize)
        {
            if (sectionIndex < 1 || sectionIndex > _doc.Sections.Count) return;
            var ps = _doc.Sections[sectionIndex].PageSetup;
            switch (paperSize?.ToUpper())
            {
                case "A4": ps.PaperSize = Word.WdPaperSize.wdPaperA4; break;
                case "A3": ps.PaperSize = Word.WdPaperSize.wdPaperA3; break;
                case "LETTER": ps.PaperSize = Word.WdPaperSize.wdPaperLetter; break;
                default: ps.PaperSize = Word.WdPaperSize.wdPaperA4; break;
            }
        }
    }
}
