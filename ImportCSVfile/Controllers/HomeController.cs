using CsvHelper;
using ImportCSVfile.Models;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using iText.Layout;
using PdfDocument = iText.Kernel.Pdf.PdfDocument;
using CsvHelper.Configuration;
using PdfSharpCore.Drawing;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Layout.Properties;

namespace ImportCSVfile.Controllers
{
    public class HomeController : Controller
    {
        [HttpPost]
        public ActionResult ImportCSV(IFormFile file)
        {
            List<Product> importedProducts = new List<Product>();

            if (file != null && file.Length > 0)
            {
                // Read the CSV file using CsvHelper
                using (var reader = new StreamReader(file.OpenReadStream()))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    importedProducts = csv.GetRecords<Product>().ToList();
                }

                // Calculate the sum of all products
                decimal totalSum = importedProducts.Sum(product => product.Quantity * product.Price);

                #region Create PDF Document
                // Create a new PDF document
                string pdfFilePath = "C:\\Users\\ACER\\Desktop\\My Visual Studio Ex\\EmailTestApp\\EmailTestApp\\ProductList.pdf";
                using (var pdfDoc = new PdfDocument(new PdfWriter(pdfFilePath)))
                using (var document = new Document(pdfDoc))
                {
                    // Add title for the table
                    PdfFont titleFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    Paragraph titleParagraph = new Paragraph("Product List").SetFont(titleFont).SetFontSize(16).SetTextAlignment(TextAlignment.CENTER);
                    document.Add(titleParagraph);

                    // Set up the table
                    Table table = new Table(new float[] { 1, 3, 1, 2, 2 });
                    table.SetWidth(UnitValue.CreatePercentValue(100));

                    // Set table headers' style
                    PdfFont headerFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    Cell headerCell = new Cell().SetFont(headerFont).SetFontSize(12).SetBackgroundColor(ColorConstants.LIGHT_GRAY).SetTextAlignment(TextAlignment.CENTER);
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Product ID")).SetFont(headerFont).SetFontSize(12).SetBackgroundColor(ColorConstants.LIGHT_GRAY).SetTextAlignment(TextAlignment.CENTER));
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Product Name")).SetFont(headerFont).SetFontSize(12).SetBackgroundColor(ColorConstants.LIGHT_GRAY).SetTextAlignment(TextAlignment.CENTER));
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Quantity")).SetFont(headerFont).SetFontSize(12).SetBackgroundColor(ColorConstants.LIGHT_GRAY).SetTextAlignment(TextAlignment.CENTER));
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Price")).SetFont(headerFont).SetFontSize(12).SetBackgroundColor(ColorConstants.LIGHT_GRAY).SetTextAlignment(TextAlignment.CENTER));
                    table.AddHeaderCell(new Cell().Add(new Paragraph("Total Price")).SetFont(headerFont).SetFontSize(12).SetBackgroundColor(ColorConstants.LIGHT_GRAY).SetTextAlignment(TextAlignment.CENTER));

                    // Add data rows
                    PdfFont cellFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                    foreach (var product in importedProducts)
                    {
                        table.AddCell(new Cell().Add(new Paragraph(product.Product_ID.ToString())).SetFont(cellFont).SetTextAlignment(TextAlignment.CENTER));
                        table.AddCell(new Cell().Add(new Paragraph(product.Product_Name)).SetFont(cellFont).SetTextAlignment(TextAlignment.LEFT));
                        table.AddCell(new Cell().Add(new Paragraph(product.Quantity.ToString())).SetFont(cellFont).SetTextAlignment(TextAlignment.CENTER));
                        table.AddCell(new Cell().Add(new Paragraph(product.Price.ToString("0.00"))).SetFont(cellFont).SetTextAlignment(TextAlignment.RIGHT));
                        table.AddCell(new Cell().Add(new Paragraph((product.Quantity * product.Price).ToString("0.00"))).SetFont(cellFont).SetTextAlignment(TextAlignment.RIGHT));
                    }

                    // Add the total sum row
                    Cell totalCell = new Cell(1, 4).Add(new Paragraph("Total")).SetFont(headerFont).SetTextAlignment(TextAlignment.RIGHT);
                    table.AddCell(totalCell);
                    table.AddCell(new Cell().Add(new Paragraph(totalSum.ToString("0.00"))).SetFont(cellFont).SetTextAlignment(TextAlignment.RIGHT));

                    // Add the table to the document
                    document.Add(table);

                    document.Close();
                }
                #endregion
            }


            return View("Index", importedProducts);
        }

        public ActionResult Index()
        {
            return View();
        }
    }
}

