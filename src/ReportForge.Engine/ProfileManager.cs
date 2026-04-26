using System;
using System.IO;
using ReportForge.Core.Models;

namespace ReportForge.Engine
{
    /// <summary>
    /// 配置文件管理器——管理 profiles/ 目录下的配置文件
    /// </summary>
    public class ProfileManager
    {
        private readonly string _profileDir;
        private readonly string _defaultPath;
        private FormatProfile _cachedProfile;

        public ProfileManager(string installDir = null)
        {
            if (string.IsNullOrEmpty(installDir))
                installDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ReportForge");

            _profileDir = Path.Combine(installDir, "profiles");
            _defaultPath = Path.Combine(_profileDir, "default.json");
        }

        /// <summary>配置文件完整路径</summary>
        public string ConfigFilePath => _defaultPath;

        /// <summary>profiles 目录路径</summary>
        public string ProfileDirectory => _profileDir;

        /// <summary>
        /// 加载配置。优先用 default.json，不存在则自动生成。
        /// </summary>
        public FormatProfile LoadProfile()
        {
            if (_cachedProfile != null) return _cachedProfile;

            if (File.Exists(_defaultPath))
            {
                try
                {
                    _cachedProfile = ProfileSerializer.LoadFromFile(_defaultPath);
                    return _cachedProfile;
                }
                catch
                {
                    // JSON 损坏时回退到内置默认
                }
            }

            // 不存在或解析失败，生成默认配置文件
            _cachedProfile = DefaultProfiles.CreateGovReport();
            EnsureDefaultExists();
            return _cachedProfile;
        }

        /// <summary>重新加载（用户改了 JSON 后调用）</summary>
        public FormatProfile ReloadProfile()
        {
            _cachedProfile = null;
            return LoadProfile();
        }

        /// <summary>确保 default.json 存在</summary>
        public void EnsureDefaultExists()
        {
            if (File.Exists(_defaultPath)) return;

            Directory.CreateDirectory(_profileDir);
            var profile = DefaultProfiles.CreateGovReport();
            ProfileSerializer.SaveToFile(profile, _defaultPath);
        }

        /// <summary>用记事本打开配置文件</summary>
        public void OpenConfigInEditor()
        {
            EnsureDefaultExists();
            System.Diagnostics.Process.Start("notepad.exe", _defaultPath);
        }
    }
}
