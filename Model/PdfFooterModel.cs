namespace Proj.Utils
{
    public class PdfFooterModel
    {
        public string LeftText { get; set; } = string.Empty;
        public string RightText { get; set; } = string.Empty;
        public bool ShowPageNumbers { get; set; } = true;
    }
}
