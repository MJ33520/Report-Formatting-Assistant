using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportForge.Core.Models
{
    /// <summary>
    /// 多语言字符串，键为locale（如 zh-CN, en-US），值为本地化文本
    /// </summary>
    public class LocalizedString : Dictionary<string, string>
    {
        public LocalizedString() : base(StringComparer.OrdinalIgnoreCase) { }

        public LocalizedString(string defaultValue) : this()
        {
            this["zh-CN"] = defaultValue;
        }

        /// <summary>按优先级获取：指定locale → zh-CN → 第一个可用值</summary>
        public string Get(string locale = "zh-CN")
        {
            if (TryGetValue(locale, out var v)) return v;
            if (TryGetValue("zh-CN", out var zh)) return zh;
            return Values.FirstOrDefault() ?? "";
        }

        public static implicit operator LocalizedString(string value) => new LocalizedString(value);
    }
}
