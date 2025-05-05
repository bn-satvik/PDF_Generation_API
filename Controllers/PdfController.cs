using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Proj.Utils;
using System.Text.Json;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Proj.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PdfController : ControllerBase
    {
        private readonly ILogger<PdfController> _logger;

        public PdfController(ILogger<PdfController> logger)
        {
            _logger = logger;
        }

        // API endpoint: POST /api/pdf/generate
        [HttpPost("generate")]
        public async Task<IActionResult> GeneratePdf()
        {
            try
            {
                // Read form data from request
                var form = await Request.ReadFormAsync();

                // Get image file, CSV file, and metadata from form
                var image = form.Files["image"];
                var csvFile = form.Files["tableData"];
                var metadataRaw = form["metadata"]; // JSON string: [Title, Inspector, Date Range]

                // Check if any required input is missing
                if (image == null || csvFile == null || string.IsNullOrWhiteSpace(metadataRaw))
                {
                    return BadRequest("Missing image, CSV file, or metadata.");
                }

                // Parse CSV into table data (list of rows)
                List<List<string>> tableData = new();
                using (var reader = new StreamReader(csvFile.OpenReadStream(), Encoding.UTF8))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = await reader.ReadLineAsync();
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            var row = ParseCsvRow(line); // Handle quoted values
                            tableData.Add(row);
                        }
                    }
                }

                // Ensure at least header + one data row
                if (tableData.Count < 2)
                {
                    return BadRequest("CSV must contain a header row and at least one data row.");
                }

                // Parse metadata JSON into list of strings
                List<string> metadata;
                try
                {
                    metadata = JsonSerializer.Deserialize<List<string>>(metadataRaw);
                }
                catch (JsonException ex)
                {
                    return BadRequest($"Invalid JSON format for metadata: {ex.Message}");
                }

                // Get image stream for PDF
                using var imageStream = image.OpenReadStream();

                // Prepare PDF header data
                var headerModel = new PdfHeaderModel
                {
                    LogoPath = "Utils/Logo.png",
                    Title = metadata.Count > 0 ? metadata[0] : "N/A",
                    GeneratedDate = DateTime.Now.ToString("MMM dd, yyyy"),
                    CompanyName = "Barracuda Networks",
                    InspectorName = metadata.Count > 1 ? metadata[1] : "N/A",
                    DateRange = metadata.Count > 2 ? metadata[2] : "N/A"
                };

                // Set footer options
                var footerModel = new PdfFooterModel
                {
                    ShowPageNumbers = true
                };

                // Generate the PDF using helper class
                byte[] pdfBytes = PdfGenerator.Generate(imageStream, tableData, headerModel, footerModel);

                // Create filename using metadata title and date
                string fileName = $"{headerModel.Title}_{headerModel.GeneratedDate}.pdf";

                // Return PDF file as HTTP response
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                _logger.LogError(ex, "Error generating PDF");
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        // Parses one line of CSV handling quoted strings and commas
        private List<string> ParseCsvRow(string line)
        {
            var row = new List<string>();
            var regex = new Regex("\"([^\"]*)\"|([^\",]+)", RegexOptions.Compiled);
            var matches = regex.Matches(line);

            foreach (Match match in matches)
            {
                var value = match.Groups[1].Value;
                row.Add(string.IsNullOrEmpty(value) ? match.Groups[2].Value : value);
            }

            return row;
        }
    }
}
