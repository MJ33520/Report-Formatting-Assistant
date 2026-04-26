using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>全文修复引擎——按顺序执行清理直接格式→重应用样式→重建编号→更新字段</summary>
    public class RepairEngine : IRepairEngine
    {
        private readonly IStyleEngine _styleEngine;
        private readonly INumberingEngine _numberingEngine;
        private readonly ICaptionEngine _captionEngine;
        private readonly ITableEngine _tableEngine;

        public RepairEngine(IStyleEngine styleEngine, INumberingEngine numberingEngine,
                            ICaptionEngine captionEngine, ITableEngine tableEngine)
        {
            _styleEngine = styleEngine;
            _numberingEngine = numberingEngine;
            _captionEngine = captionEngine;
            _tableEngine = tableEngine;
        }

        public RepairReport RunFullRepair(IWordDocumentAdapter doc, FormatProfile profile, RepairOptions options)
        {
            var report = new RepairReport();

            // Step 1: 清除直接格式
            if (options.ClearDirectFormatting)
            {
                _styleEngine.ClearDirectFormatting(doc, doc.GetContent());
                report.DirectFormattingCleared++;
            }

            // Step 2: 重新初始化目标样式
            if (options.ReapplyStyles)
            {
                _styleEngine.InitializeTargetStyles(profile, doc);
                report.StylesReapplied++;
            }

            // Step 3: 重建编号
            if (options.RebuildNumbering)
            {
                _numberingEngine.RebuildAllNumbering(doc);
                report.NumberingsRebuilt++;
            }

            // Step 4: 重建题注
            if (options.RebuildCaptions)
            {
                _captionEngine.RebuildAllCaptions(doc, profile);
                report.CaptionsRebuilt++;
            }

            // Step 5: 规范化表格
            if (options.NormalizeTables)
            {
                _tableEngine.NormalizeAllTables(doc, profile);
                report.TablesNormalized++;
            }

            // Step 6: 更新所有字段（目录、交叉引用、页码）
            if (options.UpdateTOC || options.UpdateCrossReferences || options.UpdatePageFields)
            {
                doc.UpdateAllFields();
                report.FieldsUpdated++;
            }

            return report;
        }
    }
}
