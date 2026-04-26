using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using Word = Microsoft.Office.Interop.Word;

namespace ReportForge.WordAdapter
{
    /// <summary>
    /// Word COM 适配器——封装 Microsoft.Office.Interop.Word 调用，
    /// 向上层提供稳定接口，隔离 COM 复杂性。
    /// </summary>
    public class WordDocumentAdapter : IWordDocumentAdapter
    {
        private readonly Word.Application _app;
        private readonly Word.Document _doc;

        public WordDocumentAdapter(Word.Application app, Word.Document doc)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            Styles = new WordStyleAdapter(_doc);
            Lists = new WordListAdapter(_doc);
            Sections = new WordSectionAdapter(_doc);
            Fields = new WordFieldAdapter(_app, _doc);
            Tables = new WordTableAdapter(_doc);
        }

        public IStyleAdapter Styles { get; }
        public IListAdapter Lists { get; }
        public ISectionAdapter Sections { get; }
        public IFieldAdapter Fields { get; }
        public ITableAdapter Tables { get; }

        public string WordVersion => _app.Version;
        public string FileName => _doc.Name;

        public IRangeAdapter GetSelection()
        {
            return new WordRangeAdapter(_app.Selection.Range);
        }

        public IRangeAdapter GetContent()
        {
            return new WordRangeAdapter(_doc.Content);
        }

        public IEnumerable<IParagraphAdapter> GetAllParagraphs()
        {
            for (int i = 1; i <= _doc.Paragraphs.Count; i++)
            {
                yield return new WordParagraphAdapter(_doc.Paragraphs[i], i - 1);
            }
        }

        public void UpdateAllFields()
        {
            _doc.Fields.Update();
            // 也更新页眉页脚中的字段
            foreach (Word.Section section in _doc.Sections)
            {
                foreach (Word.HeaderFooter hf in section.Headers)
                {
                    hf.Range.Fields.Update();
                }
                foreach (Word.HeaderFooter hf in section.Footers)
                {
                    hf.Range.Fields.Update();
                }
            }
        }

        public void Dispose()
        {
            // 不负责关闭文档或退出 Word，只释放 COM 引用
        }
    }

    /// <summary>Word Range 适配器</summary>
    public class WordRangeAdapter : IRangeAdapter
    {
        internal readonly Word.Range _range;

        public WordRangeAdapter(Word.Range range) { _range = range; }
        public string Text => _range.Text ?? "";
        public int Start => _range.Start;
        public int End => _range.End;
        public void Select() => _range.Select();

        internal Word.Range Raw => _range;
    }

    /// <summary>Word 段落适配器</summary>
    public class WordParagraphAdapter : IParagraphAdapter
    {
        private readonly Word.Paragraph _para;

        public WordParagraphAdapter(Word.Paragraph para, int index)
        {
            _para = para;
            Index = index;
        }

        public int Index { get; }
        public string Text => _para.Range.Text ?? "";
        public string StyleName => ((Word.Style)_para.get_Style()).NameLocal;
        public int OutlineLevel => (int)_para.OutlineLevel;
        public IRangeAdapter Range => new WordRangeAdapter(_para.Range);

        public int ListLevel
        {
            get
            {
                try
                {
                    var lf = _para.Range.ListFormat;
                    if (lf != null && lf.ListType != Word.WdListType.wdListNoNumbering)
                        return lf.ListLevelNumber; // 1-based
                }
                catch { }
                return 0;
            }
        }

        public string ListString
        {
            get
            {
                try
                {
                    var lf = _para.Range.ListFormat;
                    if (lf != null && lf.ListType != Word.WdListType.wdListNoNumbering)
                        return lf.ListString ?? "";
                }
                catch { }
                return "";
            }
        }

        public bool HasDirectFormatting
        {
            get
            {
                // 检查是否有直接格式覆盖
                try
                {
                    var font = _para.Range.Font;
                    // 如果字体名称不是 null（不继承自样式），可能有直接格式
                    return font.Name != null && font.Bold != 9999999; // wdUndefined
                }
                catch { return false; }
            }
        }

        public bool HasInlineImage
        {
            get
            {
                try { return _para.Range.InlineShapes.Count > 0; }
                catch { return false; }
            }
        }

        public bool IsInsideTable
        {
            get
            {
                try { return (bool)_para.Range.Information[Word.WdInformation.wdWithInTable]; }
                catch { return false; }
            }
        }
    }
}
