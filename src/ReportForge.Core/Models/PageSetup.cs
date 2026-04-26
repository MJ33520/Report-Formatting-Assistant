namespace ReportForge.Core.Models
{
    /// <summary>页面设置</summary>
    public class PageSetup
    {
        public string PaperSize { get; set; } = "A4";
        public MarginConfig Margins { get; set; } = new();
        public double HeaderDistanceCm { get; set; } = 1.5;
        public double FooterDistanceCm { get; set; } = 2.0;
    }

    public class MarginConfig
    {
        public double TopCm { get; set; } = 3.0;
        public double BottomCm { get; set; } = 3.0;
        public double LeftCm { get; set; } = 2.6;
        public double RightCm { get; set; } = 2.6;
        public double GutterCm { get; set; } = 0;
    }
}
