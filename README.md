
# 📄 PDF Generation API

This API generates a PDF document with an image on the first page and a tabular report on the second page. It supports uploading an image, CSV data, and metadata for dynamic report creation.

## 🔧 Technologies Used

- ASP.NET Core Web API
- MigraDoc & PdfSharpCore for PDF generation
- CsvHelper (optional future use for structured CSV parsing)

## 📦 NuGet Packages

```xml
<PackageReference Include="CsvHelper" Version="33.0.1" />
<PackageReference Include="PdfSharp" Version="6.1.1" />
<PackageReference Include="PdfSharp-MigraDoc" Version="6.1.1" />
<PackageReference Include="PdfSharpCore" Version="1.3.67" />
````

## 🚀 Endpoint

### `POST /api/pdf/generate`

Generates a PDF file with the uploaded image and table content.

#### 📥 Input (Multipart Form-Data)

| Key         | Type        | Description                                          |
| ----------- | ----------- | ---------------------------------------------------- |
| `image`     | File        | The image file to include on the first page          |
| `tableData` | File (.csv) | The CSV file containing the table data               |
| `metadata`  | Text (JSON) | JSON array of: `[Title, Inspector Name, Date Range]` |

**Example Metadata:**

```json
["Feeds", "Data Inspector", "01-02-2024 to 02-02-2024"]
```

#### 📤 Output

Returns a `PDF` file download with:

* Page 1: The uploaded image
* Page 2+: A clean, styled table rendered from the CSV content

---

## 📌 Example Request (with `curl`)

```bash
curl -X POST https://yourapiurl/api/pdf/generate \
  -F "image=@/path/to/image.jpg" \
  -F "tableData=@/path/to/data.csv" \
  -F "metadata=[\"Feeds\", \"Data Inspector\", \"01-02-2024 to 02-02-2024\"]"
```

---

## 🛠 Running the Project

### 1. Clone the repository

```bash
git clone https://github.com/bn-satvik/PDF_Generation_API.git
cd PDF-Generation-API-4
```

### 2. Restore dependencies

```bash
dotnet restore
```

### 3. Run the API

```bash
dotnet run
```

---

## 📝 Notes

* CSV must contain a **header row** and at least one **data row**.
* The PDF layout includes space adjustments for readability.
* The header includes title, inspector name, and date range dynamically.

---

## 📂 Sample CSV Format

```csv
Row,Recipient Email,Received Date,Clicked on Links,Sender Email,Subject
1,email@example.com,2024-04-01,Yes,sender@example.com,Sample Subject
```

---
