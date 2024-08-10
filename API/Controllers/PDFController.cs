using DinkToPdf.Contracts;
using DinkToPdf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;

namespace PdfCombinerAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PdfController : ControllerBase
    {
        private readonly IConverter _converter;

        public PdfController(IConverter converter)
        {
            _converter = converter;
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


        [HttpPost("convert")]
        public IActionResult ConvertHtmlToPdf([FromBody] string htmlContent)
        {
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
    }

}
