using System;
using System.Collections.Generic;
using ReportForge.Core.Models;

namespace ReportForge.Core.Interfaces
{
    /// <summary>Range抽象——对Word Range/Selection的封装</summary>
    public interface IRangeAdapter
    {
        string Text { get; }
        int Start { get; }
        int End { get; }
        void Select();
    }

    /// <summary>段落抽象</summary>
    public interface IParagraphAdapter
    {
        int Index { get; }
        string Text { get; }
        string StyleName { get; }
        int OutlineLevel { get; }
        IRangeAdapter Range { get; }
        bool HasDirectFormatting { get; }
        /// <summary>段落的列表级别 (1-based)，无列表时返回 0</summary>
        int ListLevel { get; }
        /// <summary>段落的编号文字（如 "一、" "1.1"），无编号时返回空</summary>
        string ListString { get; }
        /// <summary>段落是否包含嵌入图片</summary>
        bool HasInlineImage { get; }
        /// <summary>段落是否位于表格内</summary>
        bool IsInsideTable { get; }
    }

    /// <summary>表格抽象</summary>
    public interface ITableProxy
    {
        int Index { get; }
        int RowCount { get; }
        int ColumnCount { get; }
        IRangeAdapter Range { get; }
    }

    /// <summary>Word文档适配器——顶层入口</summary>
    public interface IWordDocumentAdapter : IDisposable
    {
        IStyleAdapter Styles { get; }
        IListAdapter Lists { get; }
        ISectionAdapter Sections { get; }
        IFieldAdapter Fields { get; }
        ITableAdapter Tables { get; }

        IRangeAdapter GetSelection();
        IRangeAdapter GetContent();
        IEnumerable<IParagraphAdapter> GetAllParagraphs();
        void UpdateAllFields();
        string WordVersion { get; }
        string FileName { get; }
    }

    /// <summary>样式适配器</summary>
    public interface IStyleAdapter
    {
        bool StyleExists(string styleId);
        void CreateOrUpdateParagraphStyle(StyleDefinition def);
        void CreateOrUpdateCharacterStyle(StyleDefinition def);
        void ApplyStyle(IRangeAdapter range, string styleId);
        void ClearDirectFormatting(IRangeAdapter range);
        void DeleteCustomStyle(string styleId);
        IEnumerable<string> GetAllCustomStyleNames();
        IEnumerable<string> GetUnusedCustomStyles();
    }

    /// <summary>多级列表适配器</summary>
    public interface IListAdapter
    {
        bool CreateMultiLevelList(NumberingScheme scheme);
        void LinkStyleToListLevel(string styleId, int level);
        void RestartNumbering(IRangeAdapter range);
        void RemoveNumbering(IRangeAdapter range);
    }

    /// <summary>分节适配器</summary>
    public interface ISectionAdapter
    {
        int SectionCount { get; }
        void InsertSectionBreak(IRangeAdapter position, SectionBreakType type);
        void ConfigureSection(int sectionIndex, SectionTemplate template);
        void SetPageNumbering(int sectionIndex, PageNumberConfig config);
        void SetOrientation(int sectionIndex, PageOrientation orientation);
        void SetMargins(int sectionIndex, MarginConfig margins);
        void UnlinkHeaderFooter(int sectionIndex);
        void SetupHeaderFooter(int sectionIndex, SectionTemplate template);
        void SetHeaderFooterDistance(int sectionIndex, double? headerCm, double? footerCm);
        void SetPaperSize(int sectionIndex, string paperSize);
    }

    public enum SectionBreakType { NextPage, Continuous, EvenPage, OddPage }

    /// <summary>字段/题注适配器</summary>
    public interface IFieldAdapter
    {
        void InsertCaption(IRangeAdapter position, CaptionTypeConfig config, string captionText);
        void InsertTableOfContents(IRangeAdapter position, TocConfig config);
        void InsertFigureTableOfContents(IRangeAdapter position, string label);
        void UpdateAllFields();
        void InsertCrossReference(IRangeAdapter position, string referenceType, string refItemText);
    }

    /// <summary>表格适配器</summary>
    public interface ITableAdapter
    {
        IEnumerable<ITableProxy> GetAllTables();
        void ApplyStandardFormat(ITableProxy table, TableStyleConfig config);
        void SetHeaderRowRepeat(ITableProxy table, bool repeat);
        void SetAllowBreakAcrossPages(ITableProxy table, bool allow);
    }
}
