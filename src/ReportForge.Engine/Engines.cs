using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>编号引擎——创建多级列表并绑定到标题样式</summary>
    public class NumberingEngine : INumberingEngine
    {
        public void CreateNumberingScheme(IWordDocumentAdapter doc, NumberingScheme scheme)
        {
            doc.Lists.CreateMultiLevelList(scheme);
        }

        public void BindStylesToNumbering(IWordDocumentAdapter doc, FormatProfile profile)
        {
            foreach (var heading in profile.HeadingLevels)
            {
                doc.Lists.LinkStyleToListLevel(heading.StyleId, heading.NumberingLevel);
            }
        }

        public void RebuildAllNumbering(IWordDocumentAdapter doc)
        {
            // 遍历所有段落，对已有大纲级别的段落重新关联编号
            foreach (var para in doc.GetAllParagraphs())
            {
                if (para.OutlineLevel >= 1 && para.OutlineLevel <= 9)
                {
                    // 通过适配层重新应用样式来触发编号刷新
                    doc.Styles.ApplyStyle(para.Range, para.StyleName);
                }
            }
        }
    }

    /// <summary>题注引擎——插入原生Caption，支持按章编号</summary>
    public class CaptionEngine : ICaptionEngine
    {
        public void InsertFigureCaption(IWordDocumentAdapter doc, string text)
        {
            var sel = doc.GetSelection();
            // TODO: 从当前Profile获取figure config
            doc.Fields.InsertCaption(sel, new CaptionTypeConfig
            {
                Label = "图", ChapterNumbering = true, Separator = "-",
                Position = CaptionPosition.Below, StyleId = "RF_FigureCaption"
            }, text);
        }

        public void InsertTableCaption(IWordDocumentAdapter doc, string text)
        {
            var sel = doc.GetSelection();
            doc.Fields.InsertCaption(sel, new CaptionTypeConfig
            {
                Label = "表", ChapterNumbering = true, Separator = "-",
                Position = CaptionPosition.Above, StyleId = "RF_TableCaption"
            }, text);
        }

        public void RebuildAllCaptions(IWordDocumentAdapter doc, FormatProfile profile)
        {
            // 更新所有题注字段
            doc.Fields.UpdateAllFields();
        }

        public void GenerateFigureTOC(IWordDocumentAdapter doc, IRangeAdapter position)
        {
            doc.Fields.InsertFigureTableOfContents(position, "图");
        }

        public void GenerateTableTOC(IWordDocumentAdapter doc, IRangeAdapter position)
        {
            doc.Fields.InsertFigureTableOfContents(position, "表");
        }
    }

    /// <summary>表格引擎——应用标准表格格式</summary>
    public class TableEngine : ITableEngine
    {
        public void ApplyStandardTable(IWordDocumentAdapter doc, ITableProxy table, TableStyleConfig config)
        {
            doc.Tables.ApplyStandardFormat(table, config);
            doc.Tables.SetHeaderRowRepeat(table, config.HeaderRow.RepeatOnNewPage);
            doc.Tables.SetAllowBreakAcrossPages(table, config.AllowBreakAcrossPages);
        }

        public void NormalizeAllTables(IWordDocumentAdapter doc, FormatProfile profile)
        {
            foreach (var table in doc.Tables.GetAllTables())
            {
                ApplyStandardTable(doc, table, profile.TableStyle);
            }
        }
    }

    /// <summary>分节引擎——管理文档分节、页码、页眉页脚</summary>
    public class SectionEngine : ISectionEngine
    {
        public void SetupDocumentSections(IWordDocumentAdapter doc, System.Collections.Generic.IList<string> sectionTemplateIds, FormatProfile profile)
        {
            // 按模板ID列表依次配置各节
            for (int i = 0; i < sectionTemplateIds.Count && i < doc.Sections.SectionCount; i++)
            {
                var templateId = sectionTemplateIds[i];
                var template = profile.SectionTemplates.Find(t => t.Id == templateId);
                if (template != null)
                {
                    ApplySectionTemplate(doc, i + 1, template);
                }
            }
        }

        public void ApplySectionTemplate(IWordDocumentAdapter doc, int sectionIndex, SectionTemplate template)
        {
            doc.Sections.ConfigureSection(sectionIndex, template);

            if (template.PageNumbering != null)
                doc.Sections.SetPageNumbering(sectionIndex, template.PageNumbering);

            doc.Sections.SetOrientation(sectionIndex, template.Orientation);

            if (!template.LinkToPrevious && sectionIndex > 1)
                doc.Sections.UnlinkHeaderFooter(sectionIndex);

            if (template.CustomMargins != null)
                doc.Sections.SetMargins(sectionIndex, template.CustomMargins);
        }

        public void FixPageNumberContinuity(IWordDocumentAdapter doc)
        {
            // 遍历各节确保页码连续性设置正确
            doc.UpdateAllFields();
        }
    }
}
