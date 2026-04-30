using System.Collections.Generic;
using ReportForge.Core.Models;

namespace ReportForge.Core.Interfaces
{
    /// <summary>Profile 管理服务</summary>
    public interface IProfileService
    {
        FormatProfile LoadProfile(string profileId);
        void SaveProfile(FormatProfile profile);
        IEnumerable<ProfileSummary> ListProfiles();
        FormatProfile CreateDefault(string locale = "zh-CN");
        FormatProfile ParseFromText(string rawText);
    }

    public class ProfileSummary
    {
        public string Id { get; set; } = "";
        public LocalizedString DisplayName { get; set; } = new();
        public string Locale { get; set; } = "zh-CN";
    }

    /// <summary>样式引擎</summary>
    public interface IStyleEngine
    {
        void InitializeTargetStyles(FormatProfile profile, IWordDocumentAdapter doc);
        void CleanupOrphanStyles(IWordDocumentAdapter doc, CleanupOptions options);
        void ApplyStyleToSelection(IWordDocumentAdapter doc, string styleId);
        void ClearDirectFormatting(IWordDocumentAdapter doc, IRangeAdapter range);
    }

    public class CleanupOptions
    {
        /// <summary>要保留的样式ID白名单</summary>
        public HashSet<string> WhitelistedStyleIds { get; set; } = new();
        /// <summary>是否删除未使用的自定义样式</summary>
        public bool DeleteUnused { get; set; } = true;
        /// <summary>试运行模式（只报告不删除）</summary>
        public bool DryRun { get; set; }
    }

    /// <summary>编号引擎</summary>
    public interface INumberingEngine
    {
        bool CreateNumberingScheme(IWordDocumentAdapter doc, NumberingScheme scheme);
        void BindStylesToNumbering(IWordDocumentAdapter doc, FormatProfile profile);
        void RebuildAllNumbering(IWordDocumentAdapter doc);
    }

    /// <summary>题注引擎</summary>
    public interface ICaptionEngine
    {
        void InsertFigureCaption(IWordDocumentAdapter doc, string text);
        void InsertTableCaption(IWordDocumentAdapter doc, string text);
        void RebuildAllCaptions(IWordDocumentAdapter doc, FormatProfile profile);
        void GenerateFigureTOC(IWordDocumentAdapter doc, IRangeAdapter position);
        void GenerateTableTOC(IWordDocumentAdapter doc, IRangeAdapter position);
    }

    /// <summary>表格引擎</summary>
    public interface ITableEngine
    {
        void ApplyStandardTable(IWordDocumentAdapter doc, ITableProxy table, TableStyleConfig config);
        void NormalizeAllTables(IWordDocumentAdapter doc, FormatProfile profile);
    }

    /// <summary>分节引擎</summary>
    public interface ISectionEngine
    {
        void SetupDocumentSections(IWordDocumentAdapter doc, IList<string> sectionTemplateIds, FormatProfile profile);
        void ApplySectionTemplate(IWordDocumentAdapter doc, int sectionIndex, SectionTemplate template);
        void FixPageNumberContinuity(IWordDocumentAdapter doc);
    }

    /// <summary>全文修复引擎</summary>
    public interface IRepairEngine
    {
        RepairReport RunFullRepair(IWordDocumentAdapter doc, FormatProfile profile, RepairOptions options);
    }

    /// <summary>巡检引擎</summary>
    public interface IInspectEngine
    {
        InspectionReport RunFullInspection(IWordDocumentAdapter doc, FormatProfile profile);
        IEnumerable<Issue> CheckManualNumbering(IWordDocumentAdapter doc);
        IEnumerable<Issue> CheckStyleCompliance(IWordDocumentAdapter doc, FormatProfile profile);
        IEnumerable<Issue> CheckHeadingLevelSkips(IWordDocumentAdapter doc);
        IEnumerable<Issue> CheckCaptionIntegrity(IWordDocumentAdapter doc);
        IEnumerable<Issue> CheckTableCompliance(IWordDocumentAdapter doc, FormatProfile profile);
        void NavigateToIssue(IWordDocumentAdapter doc, Issue issue);
    }
}
