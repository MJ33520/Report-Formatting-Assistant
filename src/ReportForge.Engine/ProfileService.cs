using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using ReportForge.Core.Interfaces;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>
    /// Profile 管理服务实现——负责配置的读写、列表、默认值生成。
    /// Profile 以 JSON 文件持久化在用户 AppData 目录下。
    /// </summary>
    public class ProfileService : IProfileService
    {
        private readonly string _profileDir;
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
        };

        public ProfileService(string? profileDirectory = null)
        {
            _profileDir = profileDirectory
                ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                "ReportForge", "Profiles");
            Directory.CreateDirectory(_profileDir);
        }

        public FormatProfile LoadProfile(string profileId)
        {
            var path = GetProfilePath(profileId);
            if (!File.Exists(path))
                throw new FileNotFoundException($"Profile not found: {profileId}", path);

            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<FormatProfile>(json, JsonOptions)
                   ?? throw new InvalidOperationException("Failed to deserialize profile");
        }

        public void SaveProfile(FormatProfile profile)
        {
            var path = GetProfilePath(profile.Id);
            var json = JsonSerializer.Serialize(profile, JsonOptions);
            File.WriteAllText(path, json);
        }

        public IEnumerable<ProfileSummary> ListProfiles()
        {
            var results = new List<ProfileSummary>();
            if (!Directory.Exists(_profileDir)) return results;

            foreach (var file in Directory.GetFiles(_profileDir, "*.json"))
            {
                try
                {
                    var json = File.ReadAllText(file);
                    var profile = JsonSerializer.Deserialize<FormatProfile>(json, JsonOptions);
                    if (profile != null)
                    {
                        results.Add(new ProfileSummary
                        {
                            Id = profile.Id,
                            DisplayName = profile.DisplayName,
                            Locale = profile.Locale
                        });
                    }
                }
                catch { /* 跳过损坏的配置文件 */ }
            }
            return results;
        }

        /// <summary>创建默认配置（政府公文 A4 格式）</summary>
        public FormatProfile CreateDefault(string locale = "zh-CN")
        {
            return DefaultProfiles.CreateGovReport(locale);
        }

        /// <summary>从粘贴的规范文本中解析配置草案（Phase 2 增强）</summary>
        public FormatProfile ParseFromText(string rawText)
        {
            // MVP: 创建默认 profile 作为基础，后续增强文本解析
            var profile = CreateDefault();
            profile.Id = $"parsed-{DateTime.Now:yyyyMMdd-HHmmss}";
            // TODO: Phase 2 - 使用正则/NLP解析文本中的格式规则
            return profile;
        }

        private string GetProfilePath(string profileId) =>
            Path.Combine(_profileDir, $"{profileId}.json");
    }
}
