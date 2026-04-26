using System.Collections.Generic;

namespace ReportForge.Core.Models
{
    /// <summary>
    /// 多级编号方案——可配置，支持中文公文编号、阿拉伯数字层级编号、英文报告编号等。
    /// </summary>
    public class NumberingScheme
    {
        /// <summary>方案ID，如 "gov-cn-mixed"、"tech-arabic"</summary>
        public string Id { get; set; } = "";

        /// <summary>多语言显示名</summary>
        public LocalizedString DisplayName { get; set; } = new();

        /// <summary>各级编号定义</summary>
        public List<NumberingLevelDef> Levels { get; set; } = new();
    }

    /// <summary>单级编号定义</summary>
    public class NumberingLevelDef
    {
        /// <summary>编号级别（0-based）</summary>
        public int Level { get; set; }

        /// <summary>编号格式类型</summary>
        public NumberFormat Format { get; set; } = NumberFormat.Arabic;

        /// <summary>编号前缀，如 "（"、"第"</summary>
        public string Prefix { get; set; } = "";

        /// <summary>编号后缀，如 "、"、"）"、"."</summary>
        public string Suffix { get; set; } = "";

        /// <summary>组合模板（用于多级拼接），如 "{L1}.{L2}"</summary>
        public string? Template { get; set; }

        /// <summary>缩进位置（磅）</summary>
        public double IndentPt { get; set; }

        /// <summary>编号示例（供UI展示）</summary>
        public string Example { get; set; } = "";

        /// <summary>起始值</summary>
        public int StartAt { get; set; } = 1;
    }

    /// <summary>编号格式枚举</summary>
    public enum NumberFormat
    {
        /// <summary>阿拉伯数字 1,2,3</summary>
        Arabic,
        /// <summary>中文数字 一,二,三</summary>
        ChineseCounting,
        /// <summary>带圈数字 ①②③</summary>
        CircledNumber,
        /// <summary>大写罗马 Ⅰ Ⅱ Ⅲ</summary>
        RomanUpper,
        /// <summary>小写罗马 i ii iii</summary>
        RomanLower,
        /// <summary>大写字母 A B C</summary>
        LetterUpper,
        /// <summary>小写字母 a b c</summary>
        LetterLower,
        /// <summary>多级拼接（使用Template字段）</summary>
        Composite
    }
}
