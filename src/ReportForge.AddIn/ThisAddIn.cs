using System;
using System.Runtime.InteropServices;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;
using ReportForge.Engine;
using ReportForge.WordAdapter;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Extensibility;

namespace ReportForge.AddIn
{
    /// <summary>
    /// ReportForge Word Add-in 入口——实现 COM Add-in 接口，
    /// 在 Word 启动时加载 Ribbon 和 TaskPane。
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("ReportForge.AddIn")]
    public class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        private static readonly string LogPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ReportForge_log.txt");

        private static readonly string AddInDir = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "";

        static ThisAddIn()
        {
            // COM Add-in 在 Word 进程中运行，需要手动解析依赖程序集
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                var name = new System.Reflection.AssemblyName(args.Name).Name + ".dll";
                var path = System.IO.Path.Combine(AddInDir, name);
                if (System.IO.File.Exists(path))
                    return System.Reflection.Assembly.LoadFrom(path);
                return null;
            };
        }

        private Word.Application? _app;
        private Office.ICTPFactory? _ctpFactory;
        private Office._CustomTaskPane? _mainTaskPane;

        // 核心服务
        private IProfileService? _profileService;
        private IStyleEngine? _styleEngine;
        private INumberingEngine? _numberingEngine;
        private ICaptionEngine? _captionEngine;
        private ITableEngine? _tableEngine;
        private ISectionEngine? _sectionEngine;
        private IRepairEngine? _repairEngine;
        private IInspectEngine? _inspectEngine;

        #region IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            try
            {
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] OnConnection called\r\n");
                _app = (Word.Application)application;
                InitializeServices();
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] OnConnection OK\r\n");
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] OnConnection ERROR: {ex}\r\n");
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                var xml = RibbonXml.GetRibbonXml();
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] GetCustomUI OK, length={xml.Length}\r\n");
                return xml;
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] GetCustomUI ERROR: {ex}\r\n");
                return "";
            }
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory ctpFactory)
        {
            _ctpFactory = ctpFactory;
        }

        #endregion

        #region 服务初始化

        private SmartFormatEngine? _smartFormatEngine;
        private ProfileManager? _profileManager;

        private void InitializeServices()
        {
            _profileManager = new ProfileManager();
            _profileManager.EnsureDefaultExists();
            _profileService = new ProfileService();
            _styleEngine = new StyleEngine();
            _numberingEngine = new NumberingEngine();
            _captionEngine = new CaptionEngine();
            _tableEngine = new TableEngine();
            _sectionEngine = new SectionEngine();
            _repairEngine = new RepairEngine(_styleEngine, _numberingEngine, _captionEngine, _tableEngine);
            _inspectEngine = new InspectEngine();
            _smartFormatEngine = new SmartFormatEngine();
        }

        /// <summary>获取当前配置（优先 JSON 文件，回退内置默认）</summary>
        private FormatProfile GetCurrentProfile()
        {
            return _profileManager?.LoadProfile() ?? DefaultProfiles.CreateGovReport();
        }

        private IWordDocumentAdapter GetActiveDocAdapter()
        {
            if (_app?.ActiveDocument == null)
                throw new InvalidOperationException("没有活动文档");
            return new WordDocumentAdapter(_app, _app.ActiveDocument);
        }

        #endregion

        #region Ribbon 回调方法（供 Ribbon XML 调用）

        private void SafeRun(string action, Action work)
        {
            try
            {
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] {action} START\r\n");
                work();
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] {action} OK\r\n");
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(LogPath, $"[{DateTime.Now}] {action} ERROR: {ex}\r\n");
                System.Windows.Forms.MessageBox.Show($"操作失败：{ex.Message}", "ReportForge 错误",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnApplyBodyStyle(Office.IRibbonControl control) => SafeRun("ApplyBodyStyle", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            _styleEngine!.ApplyStyleToSelection(doc, profile.BodyStyle.StyleId);
        });

        public void OnApplyHeadingStyle(Office.IRibbonControl control) => SafeRun($"ApplyHeading[{control.Tag}]", () =>
        {
            var tag = control.Tag;
            if (!int.TryParse(tag, out int level)) return;
            var profile = GetCurrentProfile();
            if (level < 1 || level > profile.HeadingLevels.Count) return;
            var heading = profile.HeadingLevels[level - 1];
            var doc = GetActiveDocAdapter();
            
            EnsureStylesAndNumbering(profile, doc);
            
            _styleEngine!.ApplyStyleToSelection(doc, heading.StyleId);
        });

        public void OnInsertFigureCaption(Office.IRibbonControl control) => SafeRun("InsertFigureCaption", () =>
        {
            var doc = GetActiveDocAdapter();
            _captionEngine!.InsertFigureCaption(doc, "");
        });

        public void OnInsertTableCaption(Office.IRibbonControl control) => SafeRun("InsertTableCaption", () =>
        {
            var doc = GetActiveDocAdapter();
            _captionEngine!.InsertTableCaption(doc, "");
        });

        public void OnApplyStandardTable(Office.IRibbonControl control) => SafeRun("ApplyStandardTable", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            foreach (var table in doc.Tables.GetAllTables())
            {
                _tableEngine!.ApplyStandardTable(doc, table, profile.TableStyle);
                break;
            }
        });

        public void OnApplyFigureStyle(Office.IRibbonControl control) => SafeRun("ApplyFigureStyle", () =>
        {
            var doc = GetActiveDocAdapter();
            _styleEngine!.ApplyStyleToSelection(doc, "RF_Figure");
        });

        public void OnInitializeStyles(Office.IRibbonControl control) => SafeRun("InitializeStyles", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            EnsureStylesAndNumbering(profile, doc);
            System.Windows.Forms.MessageBox.Show("目标样式体系已创建完成！", "ReportForge",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });

        public void OnFullRepair(Office.IRibbonControl control) => SafeRun("FullRepair", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            var report = _repairEngine!.RunFullRepair(doc, profile, new Core.Models.RepairOptions());
            System.Windows.Forms.MessageBox.Show(
                $"修复完成！\n清理直接格式: {report.DirectFormattingCleared}\n样式重应用: {report.StylesReapplied}\n编号重建: {report.NumberingsRebuilt}\n字段更新: {report.FieldsUpdated}",
                "ReportForge - 全文修复", System.Windows.Forms.MessageBoxButtons.OK);
        });

        public void OnRunInspection(Office.IRibbonControl control) => SafeRun("RunInspection", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            var report = _inspectEngine!.RunFullInspection(doc, profile);

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"巡检完成！ 错误:{report.ErrorCount}  警告:{report.WarningCount}  提示:{report.InfoCount}");
            sb.AppendLine(new string('─', 40));

            int shown = 0;
            foreach (var issue in report.Issues)
            {
                if (shown >= 20) { sb.AppendLine($"... 还有 {report.Issues.Count - 20} 个问题"); break; }
                var icon = issue.Severity == Core.Models.IssueSeverity.Error ? "❌"
                         : issue.Severity == Core.Models.IssueSeverity.Warning ? "⚠️" : "ℹ️";
                sb.AppendLine($"{icon} [第{issue.ParagraphIndex + 1}段] {issue.Message}");
                shown++;
            }

            if (report.Issues.Count == 0)
                sb.AppendLine("✅ 未发现问题，文档格式良好！");

            System.Windows.Forms.MessageBox.Show(sb.ToString(), "ReportForge - 格式巡检",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });

        public void OnClearDirectFormatting(Office.IRibbonControl control) => SafeRun("ClearDirectFormatting", () =>
        {
            var doc = GetActiveDocAdapter();
            var sel = doc.GetSelection();
            // 如果选区长度 <= 1（只是光标），则处理全文
            var target = (sel.End - sel.Start) > 1 ? sel : doc.GetContent();
            _styleEngine!.ClearDirectFormatting(doc, target);
            System.Windows.Forms.MessageBox.Show(
                (sel.End - sel.Start) > 1 ? "已清理选区的直接格式" : "已清理全文的直接格式（目录区域已自动跳过）",
                "ReportForge", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });

        public void OnUpdateTOC(Office.IRibbonControl control) => SafeRun("UpdateTOC", () =>
        {
            var doc = GetActiveDocAdapter();
            doc.UpdateAllFields();
        });

        public void OnSmartFormat(Office.IRibbonControl control) => SafeRun("SmartFormat", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();

            // 先确保样式和编号体系已创建
            EnsureStylesAndNumbering(profile, doc);

            // 分析文档
            var report = _smartFormatEngine!.Analyze(doc);

            // 找出检测到的最高级别
            int minDetected = 0;
            for (int lv = 1; lv <= 9; lv++)
                if (report.LevelCounts[lv] > 0) { minDetected = lv; break; }

            // 构建预览 + 级别选择对话框
            using (var form = new System.Windows.Forms.Form())
            {
                form.Text = "ReportForge - 智能格式化";
                form.Width = 520;
                form.Height = 420;
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;

                // 识别结果
                var txtInfo = new System.Windows.Forms.TextBox
                {
                    Multiline = true, ReadOnly = true, ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                    Location = new System.Drawing.Point(15, 15), Size = new System.Drawing.Size(475, 200),
                    Font = new System.Drawing.Font("Consolas", 9)
                };
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"检测到编号体系：{report.DetectedSystem}");
                sb.AppendLine(new string('─', 45));
                for (int lv = 1; lv <= 9; lv++)
                {
                    if (report.LevelCounts[lv] > 0)
                    {
                        var example = "";
                        foreach (var r in report.Results)
                            if (r.DetectedLevel == lv) { example = r.PreviewText; break; }
                        sb.AppendLine($"  来源 {lv} 级 × {report.LevelCounts[lv]}  (如：\"{example}\")");
                    }
                }
                sb.AppendLine($"  正文段落 × {report.LevelCounts[0]}");
                txtInfo.Text = sb.ToString();
                form.Controls.Add(txtInfo);

                // 级别映射选择
                var lblMap = new System.Windows.Forms.Label
                {
                    Text = "来源的最高级别对应到报告的：",
                    Location = new System.Drawing.Point(15, 230), AutoSize = true,
                    Font = new System.Drawing.Font("Microsoft YaHei UI", 10)
                };
                form.Controls.Add(lblMap);

                var combo = new System.Windows.Forms.ComboBox
                {
                    DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                    Location = new System.Drawing.Point(260, 227), Width = 180,
                    Font = new System.Drawing.Font("Microsoft YaHei UI", 10)
                };
                string[] levelNames = { "一级标题 (一、)", "二级标题 (（一）)", "三级标题 (1.)",
                    "四级标题 (1))", "五级 (①)", "六级 (a))", "七级 ((1))", "八级 (i))", "九级 (A))" };
                for (int i = 0; i < 9; i++) combo.Items.Add(levelNames[i]);
                combo.SelectedIndex = (minDetected > 0) ? minDetected - 1 : 0; // 默认=检测到的最高级别
                form.Controls.Add(combo);

                // 预览映射
                var lblPreview = new System.Windows.Forms.Label
                {
                    Location = new System.Drawing.Point(15, 265), Size = new System.Drawing.Size(475, 60),
                    Font = new System.Drawing.Font("Microsoft YaHei UI", 9)
                };
                void UpdatePreview()
                {
                    int offset = combo.SelectedIndex - (minDetected > 0 ? minDetected - 1 : 0);
                    var prev = new System.Text.StringBuilder();
                    for (int lv = 1; lv <= 9; lv++)
                    {
                        if (report.LevelCounts[lv] > 0)
                        {
                            int target = lv + offset;
                            string targetName = (target >= 1 && target <= 9) ? levelNames[target - 1] : "→正文";
                            prev.AppendLine($"  来源 {lv} 级 → {targetName}");
                        }
                    }
                    lblPreview.Text = prev.ToString();
                }
                combo.SelectedIndexChanged += (s, e) => UpdatePreview();
                UpdatePreview();
                form.Controls.Add(lblPreview);

                // 按钮
                var btnOK = new System.Windows.Forms.Button
                {
                    Text = "应用", DialogResult = System.Windows.Forms.DialogResult.OK,
                    Location = new System.Drawing.Point(290, 340), Size = new System.Drawing.Size(90, 32)
                };
                var btnCancel = new System.Windows.Forms.Button
                {
                    Text = "取消", DialogResult = System.Windows.Forms.DialogResult.Cancel,
                    Location = new System.Drawing.Point(395, 340), Size = new System.Drawing.Size(90, 32)
                };
                form.Controls.Add(btnOK);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOK;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    int offset = combo.SelectedIndex - (minDetected > 0 ? minDetected - 1 : 0);
                    var applyResult = _smartFormatEngine.Apply(doc, profile, report, offset);

                    // 构建完成报告
                    var doneSb = new System.Text.StringBuilder();
                    doneSb.AppendLine("✅ 智能格式化完成！");
                    doneSb.AppendLine();
                    if (applyResult.ImageCount > 0)
                        doneSb.AppendLine($"📷 图片：{applyResult.ImageCount} 个（已应用居中格式）");
                    if (applyResult.TableCount > 0)
                        doneSb.AppendLine($"📊 表格：{applyResult.TableCount} 个（已应用标准格式）");
                    doneSb.AppendLine();

                    // 待标注提示
                    if (applyResult.UncaptionedImages > 0 || applyResult.UncaptionedTables > 0)
                    {
                        doneSb.AppendLine("⚠️ 以下内容尚无题注，请手动添加：");
                        if (applyResult.UncaptionedImages > 0)
                            doneSb.AppendLine($"  · {applyResult.UncaptionedImages} 个图片无图题 → 点击图片后使用「图题」按钮");
                        if (applyResult.UncaptionedTables > 0)
                            doneSb.AppendLine($"  · {applyResult.UncaptionedTables} 个表格无表题 → 点击表格后使用「表题」按钮");
                    }

                    System.Windows.Forms.MessageBox.Show(doneSb.ToString(), "ReportForge",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        });

        public void OnPageSetup(Office.IRibbonControl control) => SafeRun("PageSetup", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();

            // 对所有节应用页面设置
            int sectionCount = doc.Sections.SectionCount;
            for (int i = 1; i <= sectionCount; i++)
            {
                // 纸张
                doc.Sections.SetPaperSize(i, profile.PageSetup.PaperSize);

                // 页边距：优先用 SectionTemplate 的 CustomMargins，否则用全局
                MarginConfig margins = profile.PageSetup.Margins;
                if (i - 1 < profile.SectionTemplates.Count && profile.SectionTemplates[i - 1].CustomMargins != null)
                    margins = profile.SectionTemplates[i - 1].CustomMargins;
                doc.Sections.SetMargins(i, margins);

                // 页眉/页脚距离
                double? headerDist = profile.PageSetup.HeaderDistanceCm;
                double? footerDist = profile.PageSetup.FooterDistanceCm;
                if (i - 1 < profile.SectionTemplates.Count)
                {
                    var st = profile.SectionTemplates[i - 1];
                    if (st.HeaderDistanceCm.HasValue) headerDist = st.HeaderDistanceCm;
                    if (st.FooterDistanceCm.HasValue) footerDist = st.FooterDistanceCm;
                }
                doc.Sections.SetHeaderFooterDistance(i, headerDist, footerDist);
            }

            System.Windows.Forms.MessageBox.Show(
                $"页面设置完成！已应用到 {sectionCount} 个节：\n\n" +
                $"  纸张：A4\n" +
                $"  页边距：上下 {profile.PageSetup.Margins.TopCm}cm，左右 {profile.PageSetup.Margins.LeftCm}cm\n" +
                $"  装订线：{profile.PageSetup.Margins.GutterCm}cm\n" +
                $"  页眉距顶：{profile.PageSetup.HeaderDistanceCm}cm\n" +
                $"  页脚距底：{profile.PageSetup.FooterDistanceCm}cm",
                "ReportForge", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });


        public void OnInsertTOC(Office.IRibbonControl control) => SafeRun("InsertTOC", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            var sel = doc.GetSelection();
            doc.Fields.InsertTableOfContents(sel, profile.TocConfig);
            System.Windows.Forms.MessageBox.Show("规范目录已插入！", "ReportForge",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });

        public void OnSetupHeaderFooter(Office.IRibbonControl control) => SafeRun("SetupHeaderFooter", () =>
        {
            var profile = GetCurrentProfile();
            var doc = GetActiveDocAdapter();
            int sectionCount = doc.Sections.SectionCount;

            // 对每个节应用对应的模板（如果有的话）
            for (int i = 1; i <= sectionCount && i <= profile.SectionTemplates.Count; i++)
            {
                var template = profile.SectionTemplates[i - 1];
                doc.Sections.SetupHeaderFooter(i, template);
                if (template.PageNumbering != null)
                    doc.Sections.SetPageNumbering(i, template.PageNumbering);
            }

            System.Windows.Forms.MessageBox.Show(
                $"已配置 {System.Math.Min(sectionCount, profile.SectionTemplates.Count)} 个节的页眉页脚",
                "ReportForge", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });

        #endregion

        #region 配置与帮助

        public void OnOpenConfig(Office.IRibbonControl control) => SafeRun("OpenConfig", () =>
        {
            _profileManager!.OpenConfigInEditor();
        });

        public void OnReloadConfig(Office.IRibbonControl control) => SafeRun("ReloadConfig", () =>
        {
            var profile = _profileManager!.ReloadProfile();
            System.Windows.Forms.MessageBox.Show(
                $"配置已重新加载！\n\n" +
                $"当前配置：{profile.DisplayName.Get()}\n" +
                $"正文字体：{profile.BodyStyle.Font.Zh} {profile.BodyStyle.FontSizePt}pt\n" +
                $"行距：{profile.BodyStyle.LineSpacing?.Value}磅\n" +
                $"\n下次点击「初始化样式」时将使用新配置。",
                "ReportForge", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        });

        public void OnAbout(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show(
                "ReportForge — 政府公文报告格式化工具\n\n" +
                "版本：v7.0\n" +
                "授权：CC BY-NC-SA 4.0（永久免费，禁止商用）\n\n" +
                "江老师借天才程序员手搓\n\n" +
                "如需帮助，请联系作者。",
                "关于 ReportForge",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        #endregion

        private void EnsureStylesAndNumbering(Core.Models.FormatProfile profile, Core.Interfaces.IWordDocumentAdapter doc)
        {
            _styleEngine!.InitializeTargetStyles(profile, doc);
            bool isNew = _numberingEngine!.CreateNumberingScheme(doc, profile.NumberingScheme);
            if (isNew)
            {
                _numberingEngine.BindStylesToNumbering(doc, profile);
            }
        }
    }
}
