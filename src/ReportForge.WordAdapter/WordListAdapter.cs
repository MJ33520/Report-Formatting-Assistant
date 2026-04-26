using System;
using System.Collections.Generic;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ReportForge.WordAdapter
{
    /// <summary>多级列表适配器——创建 ListTemplate 并绑定到样式</summary>
    public class WordListAdapter : IListAdapter
    {
        private readonly Word.Document _doc;
        private Word.ListTemplate? _currentTemplate;

        public WordListAdapter(Word.Document doc) { _doc = doc; }

        public void CreateMultiLevelList(NumberingScheme scheme)
        {
            Word.ListTemplate template;

            // 检查是否已存在同名列表模板——幂等
            try
            {
                template = _doc.ListTemplates.Add(
                    OutlineNumbered: true,
                    Name: scheme.Id);
            }
            catch
            {
                // 已存在，遍历找到它
                foreach (Word.ListTemplate lt in _doc.ListTemplates)
                {
                    if (lt.Name == scheme.Id)
                    {
                        _currentTemplate = lt;
                        return;
                    }
                }
                // 找不到就用唯一名重新创建
                template = _doc.ListTemplates.Add(
                    OutlineNumbered: true,
                    Name: scheme.Id + "_" + System.DateTime.Now.Ticks);
            }

            for (int i = 0; i < scheme.Levels.Count && i < 9; i++)
            {
                var levelDef = scheme.Levels[i];
                var level = template.ListLevels[i + 1]; // 1-based in COM

                level.NumberFormat = BuildNumberFormat(levelDef, i);
                level.NumberStyle = ToWordNumberStyle(levelDef.Format);
                level.StartAt = levelDef.StartAt;

                if (levelDef.IndentPt > 0)
                {
                    level.NumberPosition = (float)levelDef.IndentPt;
                }

                level.ResetOnHigher = i;
            }

            _currentTemplate = template;
        }

        public void LinkStyleToListLevel(string styleId, int level)
        {
            if (_currentTemplate == null)
                throw new InvalidOperationException("请先调用 CreateMultiLevelList 创建编号方案");

            try
            {
                var style = _doc.Styles[styleId];
                // 将样式链接到多级列表的指定级别
                style.LinkToListTemplate(_currentTemplate, level + 1); // COM 是 1-based
            }
            catch (Exception ex)
            {
                // 写到日志文件以便调试
                try
                {
                    var logPath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ReportForge_log.txt");
                    System.IO.File.AppendAllText(logPath,
                        $"[{DateTime.Now}] LinkStyleToListLevel({styleId}, {level}) ERROR: {ex.Message}\r\n");
                }
                catch { }
            }
        }

        public void RestartNumbering(IRangeAdapter range)
        {
            var wordRange = ((WordRangeAdapter)range)._range;
            wordRange.ListFormat.ApplyListTemplateWithLevel(
                _currentTemplate,
                ContinuePreviousList: false);
        }

        public void RemoveNumbering(IRangeAdapter range)
        {
            var wordRange = ((WordRangeAdapter)range)._range;
            wordRange.ListFormat.RemoveNumbers();
        }

        private string BuildNumberFormat(NumberingLevelDef def, int levelIndex)
        {
            // Word NumberFormat 使用 %1, %2 等占位符表示各级编号
            string placeholder = $"%{levelIndex + 1}";
            return $"{def.Prefix}{placeholder}{def.Suffix}";
        }

        private Word.WdListNumberStyle ToWordNumberStyle(NumberFormat format)
        {
            switch (format)
            {
                case NumberFormat.Arabic: return Word.WdListNumberStyle.wdListNumberStyleArabic;
                case NumberFormat.ChineseCounting: return Word.WdListNumberStyle.wdListNumberStyleSimpChinNum1;
                case NumberFormat.CircledNumber: return Word.WdListNumberStyle.wdListNumberStyleNumberInCircle;
                case NumberFormat.RomanUpper: return Word.WdListNumberStyle.wdListNumberStyleUppercaseRoman;
                case NumberFormat.RomanLower: return Word.WdListNumberStyle.wdListNumberStyleLowercaseRoman;
                case NumberFormat.LetterUpper: return Word.WdListNumberStyle.wdListNumberStyleUppercaseLetter;
                case NumberFormat.LetterLower: return Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter;
                default: return Word.WdListNumberStyle.wdListNumberStyleArabic;
            }
        }
    }
}
