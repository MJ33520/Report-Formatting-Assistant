using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ReportForge.Core.Models;
using System.Collections.Generic;

namespace ReportForge.Engine
{
    /// <summary>
    /// 配置序列化器——简化JSON ↔ FormatProfile 互转。
    /// JSON 格式面向普通用户，只暴露常改的字段（字体/字号/行距/编号）。
    /// 内部自动映射为完整的 FormatProfile。
    /// </summary>
    public static class ProfileSerializer
    {
        /// <summary>将完整 FormatProfile 导出为简化 JSON 字符串</summary>
        public static string ExportToJson(FormatProfile profile)
        {
            var obj = new JObject
            {
                ["_comment"] = "ReportForge 格式配置文件 | 永久免费，禁止商用 | CC BY-NC-SA 4.0",
                ["_author"] = "江老师借天才程序员手搓",
                ["displayName"] = profile.DisplayName.Get(),
                ["pageSetup"] = new JObject
                {
                    ["paperSize"] = profile.PageSetup.PaperSize,
                    ["margins"] = new JObject
                    {
                        ["topCm"] = profile.PageSetup.Margins.TopCm,
                        ["bottomCm"] = profile.PageSetup.Margins.BottomCm,
                        ["leftCm"] = profile.PageSetup.Margins.LeftCm,
                        ["rightCm"] = profile.PageSetup.Margins.RightCm,
                        ["gutterCm"] = profile.PageSetup.Margins.GutterCm
                    },
                    ["headerDistanceCm"] = profile.PageSetup.HeaderDistanceCm,
                    ["footerDistanceCm"] = profile.PageSetup.FooterDistanceCm
                },
                ["bodyStyle"] = new JObject
                {
                    ["fontZh"] = profile.BodyStyle.Font.Zh,
                    ["fontLatin"] = profile.BodyStyle.Font.Latin,
                    ["fontSizePt"] = profile.BodyStyle.FontSizePt,
                    ["lineSpacingRule"] = profile.BodyStyle.LineSpacing?.Rule.ToString() ?? "Exact",
                    ["lineSpacingValue"] = profile.BodyStyle.LineSpacing?.Value ?? 30,
                    ["firstLineIndentChars"] = profile.BodyStyle.FirstLineIndent?.Value ?? 2
                }
            };

            // 标题
            var headingsArr = new JArray();
            foreach (var h in profile.HeadingLevels)
            {
                headingsArr.Add(new JObject
                {
                    ["level"] = h.OutlineLevel,
                    ["fontZh"] = h.Font.Zh,
                    ["fontLatin"] = h.Font.Latin,
                    ["fontSizePt"] = h.FontSizePt,
                    ["bold"] = h.Bold,
                    ["firstLineIndentChars"] = h.FirstLineIndent?.Value ?? 0
                });
            }
            obj["headings"] = headingsArr;

            // 表格
            obj["table"] = new JObject
            {
                ["outsideBorderPt"] = profile.TableStyle.Borders.Outside.WidthPt,
                ["insideBorderPt"] = profile.TableStyle.Borders.Inside.WidthPt,
                ["headerFontZh"] = profile.TableStyle.HeaderRow.Font.Zh,
                ["headerFontSizePt"] = profile.TableStyle.HeaderRow.FontSizePt,
                ["bodyFontZh"] = profile.TableStyle.BodyCell.Font.Zh,
                ["bodyFontLatin"] = profile.TableStyle.BodyCell.Font.Latin,
                ["bodyFontSizePt"] = profile.TableStyle.BodyCell.FontSizePt,
                ["lineSpacingPt"] = profile.TableStyle.BodyCell.LineSpacing?.Value ?? 23
            };

            // 编号
            var numArr = new JArray();
            foreach (var lv in profile.NumberingScheme.Levels)
            {
                numArr.Add($"{lv.Prefix}{lv.Example}");
            }
            obj["numbering"] = numArr;

            return obj.ToString(Formatting.Indented);
        }

        /// <summary>从简化 JSON 加载配置，覆盖到默认 Profile 上</summary>
        public static FormatProfile ImportFromJson(string json)
        {
            // 先创建默认 Profile 作为基础
            var profile = DefaultProfiles.CreateGovReport();
            var obj = JObject.Parse(json);

            // 页面设置
            var ps = obj["pageSetup"];
            if (ps != null)
            {
                if (ps["paperSize"] != null) profile.PageSetup.PaperSize = ps["paperSize"].ToString();
                var m = ps["margins"];
                if (m != null)
                {
                    if (m["topCm"] != null) profile.PageSetup.Margins.TopCm = (double)m["topCm"];
                    if (m["bottomCm"] != null) profile.PageSetup.Margins.BottomCm = (double)m["bottomCm"];
                    if (m["leftCm"] != null) profile.PageSetup.Margins.LeftCm = (double)m["leftCm"];
                    if (m["rightCm"] != null) profile.PageSetup.Margins.RightCm = (double)m["rightCm"];
                    if (m["gutterCm"] != null) profile.PageSetup.Margins.GutterCm = (double)m["gutterCm"];
                }
                if (ps["headerDistanceCm"] != null) profile.PageSetup.HeaderDistanceCm = (double)ps["headerDistanceCm"];
                if (ps["footerDistanceCm"] != null) profile.PageSetup.FooterDistanceCm = (double)ps["footerDistanceCm"];
            }

            // 正文
            var bs = obj["bodyStyle"];
            if (bs != null)
            {
                if (bs["fontZh"] != null) profile.BodyStyle.Font.Zh = bs["fontZh"].ToString();
                if (bs["fontLatin"] != null) profile.BodyStyle.Font.Latin = bs["fontLatin"].ToString();
                if (bs["fontSizePt"] != null) profile.BodyStyle.FontSizePt = (double)bs["fontSizePt"];
                if (bs["lineSpacingValue"] != null && profile.BodyStyle.LineSpacing != null)
                    profile.BodyStyle.LineSpacing.Value = (double)bs["lineSpacingValue"];
                if (bs["firstLineIndentChars"] != null && profile.BodyStyle.FirstLineIndent != null)
                    profile.BodyStyle.FirstLineIndent.Value = (double)bs["firstLineIndentChars"];
            }

            // 标题
            var headings = obj["headings"] as JArray;
            if (headings != null)
            {
                for (int i = 0; i < headings.Count && i < profile.HeadingLevels.Count; i++)
                {
                    var h = headings[i];
                    var target = profile.HeadingLevels[i];
                    if (h["fontZh"] != null) target.Font.Zh = h["fontZh"].ToString();
                    if (h["fontLatin"] != null) target.Font.Latin = h["fontLatin"].ToString();
                    if (h["fontSizePt"] != null) target.FontSizePt = (double)h["fontSizePt"];
                    if (h["bold"] != null) target.Bold = (bool)h["bold"];
                    if (h["firstLineIndentChars"] != null)
                    {
                        if (target.FirstLineIndent == null) target.FirstLineIndent = new IndentConfig { Unit = IndentUnit.Chars };
                        target.FirstLineIndent.Value = (double)h["firstLineIndentChars"];
                    }
                }
            }

            // 表格
            var tbl = obj["table"];
            if (tbl != null)
            {
                if (tbl["outsideBorderPt"] != null) profile.TableStyle.Borders.Outside.WidthPt = (double)tbl["outsideBorderPt"];
                if (tbl["insideBorderPt"] != null) profile.TableStyle.Borders.Inside.WidthPt = (double)tbl["insideBorderPt"];
                if (tbl["headerFontZh"] != null) profile.TableStyle.HeaderRow.Font.Zh = tbl["headerFontZh"].ToString();
                if (tbl["headerFontSizePt"] != null) profile.TableStyle.HeaderRow.FontSizePt = (double)tbl["headerFontSizePt"];
                if (tbl["bodyFontZh"] != null) profile.TableStyle.BodyCell.Font.Zh = tbl["bodyFontZh"].ToString();
                if (tbl["bodyFontLatin"] != null) profile.TableStyle.BodyCell.Font.Latin = tbl["bodyFontLatin"].ToString();
                if (tbl["bodyFontSizePt"] != null) profile.TableStyle.BodyCell.FontSizePt = (double)tbl["bodyFontSizePt"];
                if (tbl["lineSpacingPt"] != null && profile.TableStyle.BodyCell.LineSpacing != null)
                    profile.TableStyle.BodyCell.LineSpacing.Value = (double)tbl["lineSpacingPt"];
            }

            return profile;
        }

        /// <summary>将 JSON 保存到文件</summary>
        public static void SaveToFile(FormatProfile profile, string filePath)
        {
            var json = ExportToJson(profile);
            File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
        }

        /// <summary>从文件加载配置</summary>
        public static FormatProfile LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                return DefaultProfiles.CreateGovReport();
            var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
            return ImportFromJson(json);
        }
    }
}
