using System;
using System.Collections.Generic;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ReportForge.WordAdapter
{
    /// <summary>字段/题注适配器——插入原生Caption、目录、交叉引用</summary>
    public class WordFieldAdapter : IFieldAdapter
    {
        private readonly Word.Application _app;
        private readonly Word.Document _doc;

        public WordFieldAdapter(Word.Application app, Word.Document doc)
        {
            _app = app;
            _doc = doc;
        }

        public void InsertCaption(IRangeAdapter position, CaptionTypeConfig config, string captionText)
        {
            var range = ((WordRangeAdapter)position)._range;
            range.Select();

            // 确保标签存在
            var label = config.Label.Get();
            EnsureCaptionLabel(label);

            // 插入题注
            var insertPos = config.Position == CaptionPosition.Below
                ? Word.WdCaptionPosition.wdCaptionPositionBelow
                : Word.WdCaptionPosition.wdCaptionPositionAbove;

            _app.Selection.InsertCaption(
                Label: label,
                Title: " " + captionText,
                Position: insertPos,
                ExcludeLabel: 0);

            // 应用题注样式
            if (!string.IsNullOrEmpty(config.StyleId))
            {
                try {
                    object styleObj = _doc.Styles[config.StyleId];
                    _app.Selection.Paragraphs[1].set_Style(ref styleObj);
                }
                catch { /* 样式不存在时静默跳过 */ }
            }
        }

        public void InsertTableOfContents(IRangeAdapter position, TocConfig config)
        {
            var range = ((WordRangeAdapter)position)._range;
            var toc = _doc.TablesOfContents.Add(
                Range: range,
                UseHeadingStyles: true,
                UpperHeadingLevel: 1,
                LowerHeadingLevel: config.MaxLevel,
                UseHyperlinks: config.UseHyperlinks,
                RightAlignPageNumbers: config.RightAlignPageNumbers);

            // 设置目录字体和行距
            try
            {
                var tocRange = toc.Range;
                tocRange.Font.Name = "Times New Roman";
                tocRange.Font.NameFarEast = "仿宋_GB2312";
                tocRange.Font.Size = 14; // 四号
                tocRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                tocRange.ParagraphFormat.LineSpacing = 30;
            }
            catch { }
        }

        public void InsertFigureTableOfContents(IRangeAdapter position, string label)
        {
            var range = ((WordRangeAdapter)position)._range;
            _doc.TablesOfFigures.Add(
                Range: range,
                Caption: label);
        }

        public void InsertCrossReference(IRangeAdapter position, string referenceType, string refItemText)
        {
            var range = ((WordRangeAdapter)position)._range;
            range.Select();
            // Word 交叉引用需要通过 Selection 操作
            // TODO: 实现具体的交叉引用插入（需要遍历图表书签）
        }

        public void UpdateAllFields()
        {
            _doc.Fields.Update();
            // 更新目录
            foreach (Word.TableOfContents toc in _doc.TablesOfContents)
            {
                toc.Update();
            }
        }

        private void EnsureCaptionLabel(string label)
        {
            try
            {
                var _ = _app.CaptionLabels[label];
            }
            catch
            {
                // 标签不存在，创建新标签
                _app.CaptionLabels.Add(label);
            }
        }
    }
}
