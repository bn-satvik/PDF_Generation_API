namespace Proj.Utils
{
    public class PdfHeaderModel
    {
        // Path to the logo image file (e.g., "Utils/Logo.png")
        public string LogoPath { get; set; } = "Utils\\Logo.png";

        // Title for the report (e.g., "Top Attacked Users")
        public string Title { get; set; } = string.Empty;

        // Date when the report is generated
        public string GeneratedDate { get; set; } = DateTime.Now.ToString("MMM dd, yyyy");

        // Company name (e.g., "Barracuda Networks")
        public string CompanyName { get; set; } = string.Empty;

        // Name of the data inspector (e.g., "Data Inspector")
        public string InspectorName { get; set; } = string.Empty;

        // The date range for the report (e.g., "Data from Apr 26 - Oct 23, 2024")
        public string DateRange { get; set; } = string.Empty;
    }
}
