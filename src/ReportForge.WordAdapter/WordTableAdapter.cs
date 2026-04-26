using System.Collections.Generic;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ReportForge.WordAdapter
{
    /// <summary>表格适配器——应用标准表格格式</summary>
    public class WordTableAdapter : ITableAdapter
    {
        private readonly Word.Document _doc;
        public WordTableAdapter(Word.Document doc) { _doc = doc; }

        public IEnumerable<ITableProxy> GetAllTables()
        {
            for (int i = 1; i <= _doc.Tables.Count; i++)
            {
                yield return new WordTableProxy(_doc.Tables[i], i - 1);
            }
        }

        public void ApplyStandardFormat(ITableProxy table, TableStyleConfig config)
        {
            var wordTable = ((WordTableProxy)table).Raw;

            // 表格对齐
            wordTable.Rows.Alignment = config.TableAlignment == TextAlignment.Center
                ? Word.WdRowAlignment.wdAlignRowCenter
                : Word.WdRowAlignment.wdAlignRowLeft;

            // 宽度策略
            if (config.WidthStrategy == TableWidthStrategy.AutoFitWindow)
                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            else if (config.WidthStrategy == TableWidthStrategy.AutoFitContent)
                wordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

            // 边框
            SetBorder(wordTable.Borders[Word.WdBorderType.wdBorderTop], config.Borders.Outside);
            SetBorder(wordTable.Borders[Word.WdBorderType.wdBorderBottom], config.Borders.Outside);
            SetBorder(wordTable.Borders[Word.WdBorderType.wdBorderLeft], config.Borders.Outside);
            SetBorder(wordTable.Borders[Word.WdBorderType.wdBorderRight], config.Borders.Outside);
            SetBorder(wordTable.Borders[Word.WdBorderType.wdBorderHorizontal], config.Borders.Inside);
            SetBorder(wordTable.Borders[Word.WdBorderType.wdBorderVertical], config.Borders.Inside);

            // 表头行格式
            if (wordTable.Rows.Count > 0)
            {
                var headerRow = wordTable.Rows[1];
                headerRow.HeadingFormat = -1; // true
                headerRow.Range.Font.NameAscii = config.HeaderRow.Font.Latin;
                headerRow.Range.Font.NameOther = config.HeaderRow.Font.Latin;
                headerRow.Range.Font.NameFarEast = config.HeaderRow.Font.Zh;
                headerRow.Range.Font.Size = (float)config.HeaderRow.FontSizePt;
                headerRow.Range.Font.Bold = config.HeaderRow.Bold ? 1 : 0;
                headerRow.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRow.Range.ParagraphFormat.SpaceBefore = 0;
                headerRow.Range.ParagraphFormat.SpaceAfter = 0;
                // 表头行距也设为23磅固定
                if (config.BodyCell.LineSpacing != null)
                {
                    headerRow.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    headerRow.Range.ParagraphFormat.LineSpacing = (float)config.BodyCell.LineSpacing.Value;
                }
            }

            // 表内文字格式（第2行起）
            for (int r = 2; r <= wordTable.Rows.Count; r++)
            {
                var row = wordTable.Rows[r];
                row.Range.Font.NameAscii = config.BodyCell.Font.Latin;
                row.Range.Font.NameOther = config.BodyCell.Font.Latin;
                row.Range.Font.NameFarEast = config.BodyCell.Font.Zh;
                row.Range.Font.Size = (float)config.BodyCell.FontSizePt;
                row.Range.Font.Bold = 0;

                row.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                row.Range.ParagraphFormat.SpaceBefore = 0;
                row.Range.ParagraphFormat.SpaceAfter = 0;

                if (config.BodyCell.LineSpacing != null)
                {
                    row.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    row.Range.ParagraphFormat.LineSpacing = (float)config.BodyCell.LineSpacing.Value;
                }
            }

            // 跨页断行
            wordTable.Rows.AllowBreakAcrossPages = config.AllowBreakAcrossPages ? -1 : 0;
        }

        public void SetHeaderRowRepeat(ITableProxy table, bool repeat)
        {
            var wordTable = ((WordTableProxy)table).Raw;
            if (wordTable.Rows.Count > 0)
                wordTable.Rows[1].HeadingFormat = repeat ? -1 : 0;
        }

        public void SetAllowBreakAcrossPages(ITableProxy table, bool allow)
        {
            var wordTable = ((WordTableProxy)table).Raw;
            wordTable.Rows.AllowBreakAcrossPages = allow ? -1 : 0;
        }

        private void SetBorder(Word.Border border, BorderDef def)
        {
            border.LineStyle = def.Style == BorderStyle.Single
                ? Word.WdLineStyle.wdLineStyleSingle
                : def.Style == BorderStyle.Double
                    ? Word.WdLineStyle.wdLineStyleDouble
                    : Word.WdLineStyle.wdLineStyleNone;
            border.LineWidth = PointsToLineWidth(def.WidthPt);
        }

        private Word.WdLineWidth PointsToLineWidth(double pt)
        {
            // WdLineWidth 枚举值：
            // wdLineWidth025pt = 2,  wdLineWidth050pt = 4,  wdLineWidth075pt = 6
            // wdLineWidth100pt = 8,  wdLineWidth150pt = 12, wdLineWidth225pt = 18
            // wdLineWidth300pt = 24, wdLineWidth450pt = 36, wdLineWidth600pt = 48
            if (pt <= 0.25) return Word.WdLineWidth.wdLineWidth025pt;
            if (pt <= 0.5)  return Word.WdLineWidth.wdLineWidth050pt;
            if (pt <= 0.75) return Word.WdLineWidth.wdLineWidth075pt;
            if (pt <= 1.0)  return Word.WdLineWidth.wdLineWidth100pt;
            if (pt <= 1.5)  return Word.WdLineWidth.wdLineWidth150pt;
            if (pt <= 2.25) return Word.WdLineWidth.wdLineWidth225pt;
            if (pt <= 3.0)  return Word.WdLineWidth.wdLineWidth300pt;
            if (pt <= 4.5)  return Word.WdLineWidth.wdLineWidth450pt;
            return Word.WdLineWidth.wdLineWidth600pt;
        }
    }

    /// <summary>Word表格代理</summary>
    public class WordTableProxy : ITableProxy
    {
        internal readonly Word.Table Raw;
        public WordTableProxy(Word.Table table, int index) { Raw = table; Index = index; }
        public int Index { get; }
        public int RowCount => Raw.Rows.Count;
        public int ColumnCount => Raw.Columns.Count;
        public IRangeAdapter Range => new WordRangeAdapter(Raw.Range);
    }
}
