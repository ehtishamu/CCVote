using Aspose.BarCode.Generation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using DinkToPdf;
using DinkToPdf.Contracts;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class QRController : ControllerBase
    {
        private readonly IConverter _converter;

        public QRController(IConverter converter)
        {
            _converter = converter;
        }

        [HttpPost]
        public IActionResult Post(Dictionary<string, string>[] KeyValuePairs)
        {
            for (int i = 0; i < KeyValuePairs.Length; i++)
            {
                var generator = new BarcodeGenerator(EncodeTypes.QR);
                // Specify code text to encode
               
                    KeyValuePairs.ToList()[i].TryGetValue("URL",out string colValue);
                generator.CodeText = colValue;
                // Specify the size of the image
                generator.Parameters.Barcode.XDimension.Pixels = 8;
                generator.Parameters.Resolution = 500;
                // Save the generated QR code
                generator.Save("C:\\Ehtisham\\CCLIVE\\CCVote\\CCVote\\wwwroot\\img"+i+".jpg");
            }
            return Ok(new
            {
                message = "QR Generated Successfully!!!",
            });
        }
        private string TryGetValueFromDist(Dictionary<string, string> rowData, string key)
        {
            rowData.TryGetValue(key, out string colValue);
            return colValue;
        }

        [HttpPost("convert")]
        public IActionResult ConvertHtmlToPdf(HTMLModel model)
        {
            string htmlContent = model.htmlContent;
            try
            {
                var doc = new HtmlToPdfDocument()
                {
                    GlobalSettings = new GlobalSettings
                    {
                        ColorMode = ColorMode.Color,
                        Orientation = Orientation.Portrait,
                        PaperSize = PaperKind.A4,
                    },
                    Objects = {
                        new ObjectSettings
                        {
                            HtmlContent = htmlContent,
                            WebSettings = { DefaultEncoding = "utf-8" },
                        }
                    }
                };

                byte[] pdf = _converter.Convert(doc);

                return File(pdf, "application/pdf", "converted.pdf");
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "An error occurred while converting HTML to PDF.", error = ex.Message });
            }
        }
        [HttpPost("combine")]
        public IActionResult CombinePDFs(List<IFormFile> files)
        {
            try
            {
                using (MemoryStream outputStream = new MemoryStream())
                {
                    using (PdfDocument outputDocument = new PdfDocument())
                    {
                        foreach (var file in files)
                        {
                            if (file.Length > 0)
                            {
                                using (var ms = new MemoryStream())
                                {
                                    file.CopyTo(ms);
                                    ms.Position = 0;
                                    PdfDocument inputDocument = PdfReader.Open(ms, PdfDocumentOpenMode.Import);

                                    // Iterate through pages of each document
                                    for (int i = 0; i < inputDocument.PageCount; i++)
                                    {
                                        PdfPage page = inputDocument.Pages[i];
                                        outputDocument.AddPage(page);
                                    }
                                }
                            }
                        }

                        outputDocument.Save(outputStream);
                    }

                    return File(outputStream.ToArray(), "application/pdf", "combined.pdf");
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "An error occurred while combining PDFs.", error = ex.Message });
            }
        }
    }
    public class HTMLModel
    {
        public string htmlContent { get; set; }
    }
    public class QRModel { 
    
    }
}
