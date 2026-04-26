using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    public class InspectEngine : IInspectEngine
    {
        private static readonly Regex[] ManualNumPatterns = new[]
        {
            new Regex(@"^[一二三四五六七八九十]+[、.．]", RegexOptions.Compiled),
            new Regex(@"^（[一二三四五六七八九十]+）", RegexOptions.Compiled),
            new Regex(@"^\d+[、.．]\s", RegexOptions.Compiled),
            new Regex(@"^\d+\.\d+[\s\.．]", RegexOptions.Compiled),
            new Regex(@"^[（\(]\d+[）\)]", RegexOptions.Compiled),
            new Regex(@"^[①②③④⑤⑥⑦⑧⑨⑩]", RegexOptions.Compiled),
        };

        public InspectionReport RunFullInspection(IWordDocumentAdapter doc, FormatProfile profile)
        {
            var all = new List<Issue>();
            all.AddRange(CheckManualNumbering(doc));
            all.AddRange(CheckStyleCompliance(doc, profile));
            all.AddRange(CheckHeadingLevelSkips(doc));
            all.AddRange(CheckCaptionIntegrity(doc));
            all.AddRange(CheckTableCompliance(doc, profile));
            return new InspectionReport
            {
                Issues = all,
                ErrorCount = all.Count(i => i.Severity == IssueSeverity.Error),
                WarningCount = all.Count(i => i.Severity == IssueSeverity.Warning),
                InfoCount = all.Count(i => i.Severity == IssueSeverity.Info)
            };
        }

        public IEnumerable<Issue> CheckManualNumbering(IWordDocumentAdapter doc)
        {
            foreach (var p in doc.GetAllParagraphs())
            {
                var t = p.Text.TrimStart();
                if (string.IsNullOrWhiteSpace(t) || t.Length > 200 || (p.OutlineLevel >= 1 && p.OutlineLevel <= 9)) continue;
                foreach (var rx in ManualNumPatterns)
                {
                    if (rx.IsMatch(t))
                    {
                        yield return new Issue { Id = $"mn-{p.Index}", Severity = IssueSeverity.Warning, Category = "ManualNumbering",
                            Message = $"疑似手动编号：{(t.Length > 50 ? t.Substring(0, 50) + "…" : t)}", ParagraphIndex = p.Index, TextPreview = t.Length > 80 ? t.Substring(0, 80) : t };
                        break;
                    }
                }
            }
        }

        public IEnumerable<Issue> CheckStyleCompliance(IWordDocumentAdapter doc, FormatProfile profile)
        {
            var targets = new HashSet<string> { profile.BodyStyle.StyleId };
            foreach (var h in profile.HeadingLevels) targets.Add(h.StyleId);
            targets.Add(profile.CaptionConfig.Figure.StyleId);
            targets.Add(profile.CaptionConfig.Table.StyleId);

            foreach (var p in doc.GetAllParagraphs())
            {
                if (string.IsNullOrWhiteSpace(p.Text)) continue;
                if (!targets.Contains(p.StyleName) && !p.StyleName.StartsWith("RF_"))
                    yield return new Issue { Id = $"sc-{p.Index}", Severity = IssueSeverity.Info, Category = "StyleMismatch",
                        Message = $"未使用目标样式（当前：{p.StyleName}）", ParagraphIndex = p.Index };
            }
        }

        public IEnumerable<Issue> CheckHeadingLevelSkips(IWordDocumentAdapter doc)
        {
            int last = 0;
            foreach (var p in doc.GetAllParagraphs())
            {
                if (p.OutlineLevel >= 1 && p.OutlineLevel <= 9)
                {
                    if (last > 0 && p.OutlineLevel > last + 1)
                        yield return new Issue { Id = $"hs-{p.Index}", Severity = IssueSeverity.Warning, Category = "HeadingSkip",
                            Message = $"标题层级跳级：{last}→{p.OutlineLevel}", ParagraphIndex = p.Index };
                    last = p.OutlineLevel;
                }
            }
        }

        public IEnumerable<Issue> CheckCaptionIntegrity(IWordDocumentAdapter doc) => Enumerable.Empty<Issue>(); // TODO Phase2
        public IEnumerable<Issue> CheckTableCompliance(IWordDocumentAdapter doc, FormatProfile profile) => Enumerable.Empty<Issue>(); // TODO Phase2

        public void NavigateToIssue(IWordDocumentAdapter doc, Issue issue)
        {
            var paras = doc.GetAllParagraphs().ToList();
            if (issue.ParagraphIndex >= 0 && issue.ParagraphIndex < paras.Count)
                paras[issue.ParagraphIndex].Range.Select();
        }
    }
}
