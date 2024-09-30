using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Text;

namespace ExcelController.Controllers
{
    public class ExcelController : Controller
    {
        private const string UploadedFileKey = "UploadedFile";

        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (!IsValidFile(file, out string message))
            {
                ViewBag.Message = message;
                return View();
            }

            // Set the EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Store the file in session
            StoreFileInSession(file);

            using (var package = new ExcelPackage(new MemoryStream(HttpContext.Session.Get(UploadedFileKey))))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Store the worksheet name in ViewBag
                ViewBag.WorksheetName = worksheet.Name;

                // Add a new column "Total Value before Taxing" and calculate its values
                int newColumnIndex = AddTotalValueColumn(worksheet);

                // Calculate sum of "Total Value After Taxing" and add it to a new row at the end
                AddTotalSumRow(worksheet, newColumnIndex);

                // Convert the Excel sheet to HTML for display
                ViewBag.ExcelTable = ConvertWorksheetToHtml(worksheet);
            }

            ViewBag.FileName = file.FileName;
            return View();
        }

        private bool IsValidFile(IFormFile file, out string message)
        {
            if (file == null || file.Length <= 0)
            {
                message = "File not selected.";
                return false;
            }

            message = string.Empty;
            return true;
        }

        private void StoreFileInSession(IFormFile file)
        {
            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                HttpContext.Session.Set(UploadedFileKey, stream.ToArray());
                HttpContext.Session.SetString("FileName", file.FileName);
            }
        }

        private int AddTotalValueColumn(ExcelWorksheet worksheet)
        {
            int newColumnIndex = worksheet.Dimension.End.Column + 1;
            worksheet.Cells[1, newColumnIndex].Value = "Total Value before Taxing";

            CalculateTotalValueBeforeTaxing(worksheet, newColumnIndex);
            return newColumnIndex;
        }

        private void CalculateTotalValueBeforeTaxing(ExcelWorksheet worksheet, int newColumnIndex)
        {
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                decimal totalValueAfterTaxing = GetCellValueAsDecimal(worksheet.Cells[row, 7]);
                decimal taxingValue = GetCellValueAsDecimal(worksheet.Cells[row, 8]);

                worksheet.Cells[row, newColumnIndex].Value = totalValueAfterTaxing + taxingValue;
            }
        }

        private decimal GetCellValueAsDecimal(ExcelRange cell)
        {
            return decimal.TryParse(cell.Value?.ToString(), out decimal value) ? value : 0;
        }

        private void AddTotalSumRow(ExcelWorksheet worksheet, int newColumnIndex)
        {
            int totalRowIndex = worksheet.Dimension.End.Row + 1;
            decimal totalSum = CalculateTotalSum(worksheet);

            ViewBag.TotalSum = totalSum;

            // Add the total sum to a new row and make it bold
            worksheet.Cells[totalRowIndex, 7].Value = totalSum;
            worksheet.Cells[totalRowIndex, 7].Style.Font.Bold = true;
            worksheet.Cells[totalRowIndex, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[totalRowIndex, 7].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            // Optionally, label the "Total" row in column 1
            worksheet.Cells[totalRowIndex, 1].Value = "Total";
            worksheet.Cells[totalRowIndex, 1].Style.Font.Bold = true;
        }

        private decimal CalculateTotalSum(ExcelWorksheet worksheet)
        {
            decimal totalSum = 0;

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                totalSum += GetCellValueAsDecimal(worksheet.Cells[row, 7]);
            }

            return totalSum;
        }

        private string ConvertWorksheetToHtml(ExcelWorksheet worksheet)
        {
            var sb = new StringBuilder();
            sb.Append("<form method='post' action='/Excel/Save'>");
            sb.Append("<table class='table table-bordered'>");

            // Add headers
            sb.Append("<thead><tr>");
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                sb.AppendFormat("<th>{0}</th>", worksheet.Cells[1, col].Value?.ToString() ?? string.Empty);
            }
            sb.Append("</tr></thead>");

            // Add rows
            sb.Append("<tbody>");
            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                sb.Append("<tr>");
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;

                    if (col == worksheet.Dimension.End.Column) // Assuming the "Total Value before Taxing" column is the last one
                    {
                        sb.AppendFormat("<td><input type='text' value='{0}' class='form-control' readonly /></td>", cellValue);
                    }
                    else
                    {
                        sb.AppendFormat("<td><input type='text' name='cell-{0}-{1}' value='{2}' class='form-control' /></td>", row, col, cellValue);
                    }
                }
                sb.Append("</tr>");
            }
            sb.Append("</tbody>");
            sb.Append("</table>");

            // Hidden field to store the original file name
            sb.Append($"<input type='hidden' name='fileName' value='{HttpContext.Session.GetString("FileName")}' />");
            sb.Append("<button type='submit' class='btn btn-primary'>Download</button>");
            sb.Append("</form>");

            return sb.ToString();
        }

        [HttpPost]
        public IActionResult Save(IFormCollection formData)
        {
            var fileBytes = HttpContext.Session.Get(UploadedFileKey);
            var fileName = formData["fileName"].ToString();

            if (fileBytes == null || fileBytes.Length == 0)
            {
                ViewBag.Message = "No file uploaded.";
                return View("Upload");
            }

            using (var stream = new MemoryStream(fileBytes))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    // Update the worksheet with form data
                    UpdateWorksheetWithFormData(worksheet, formData);

                    // Calculate and update total value before taxing in the new column
                    int newColumnIndex = worksheet.Dimension.End.Column + 1;
                    worksheet.Cells[1, newColumnIndex].Value = "Total Value before Taxing";

                    // This part remains unchanged
                    for (int row = 2; row <= worksheet.Dimension.End.Row - 1; row++)
                    {
                        decimal totalValueAfterTaxing = decimal.TryParse(worksheet.Cells[row, 7].Value?.ToString(), out decimal taxValue) ? taxValue : 0; // Column 7
                        decimal taxingValue = decimal.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out decimal taxing) ? taxing : 0; // Column 8

                        decimal totalValueBeforeTaxing = totalValueAfterTaxing + taxingValue;
                        worksheet.Cells[row, newColumnIndex].Value = totalValueBeforeTaxing;
                        worksheet.Cells[row, newColumnIndex].Style.Numberformat.Format = "0.00"; // Ensure correct formatting
                        worksheet.Cells[row, newColumnIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, newColumnIndex].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
                    }

                    // Create the HTML table structure to return to the user
                    var htmlTable = ConvertWorksheetToHtml(worksheet);
                    ViewBag.UpdatedTable = htmlTable;

                    // Save the modified Excel file
                    var modifiedStream = new MemoryStream();
                    package.SaveAs(modifiedStream);
                    modifiedStream.Position = 0;

                    // Return the updated Excel file as a downloadable file
                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    return File(modifiedStream, contentType, fileName);
                }
            }
        }

        private void UpdateWorksheetWithFormData(ExcelWorksheet worksheet, IFormCollection formData)
        {
            foreach (var key in formData.Keys)
            {
                if (key.StartsWith("cell-"))
                {
                    var parts = key.Split('-');
                    int row = int.Parse(parts[1]);
                    int col = int.Parse(parts[2]);
                    string newValue = formData[key];

                    worksheet.Cells[row, col].Value = newValue;
                }
            }
        }
    }
}
