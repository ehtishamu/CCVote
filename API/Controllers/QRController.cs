﻿using Aspose.BarCode.Generation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Drawing;
using System.Drawing.Imaging;
using QRCoder;
using static QRCoder.PayloadGenerator;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Globalization;

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

        [HttpPost("test")]
        public IActionResult TestQR(string url) {

            QRCodeGenerator QrGenerator = new QRCodeGenerator();
            QRCodeData QrCodeInfo = QrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.Q);
            QRCode QrCode = new QRCode(QrCodeInfo);
            Bitmap QrBitmap = QrCode.GetGraphic(60);
            byte[] BitmapArray = QrBitmap.BitmapToByteArray();
            string QrUri = string.Format("data:image/png;base64,{0}", Convert.ToBase64String(BitmapArray));
            return Ok(new
            {
                QrUri = QrUri,
                message = "QR Generated Successfully!!!",
            });
        }
        [HttpPost("import")]
        public async Task<IActionResult> ImportExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded or file is empty.");
            }

            try
            {
                using (var ms = new MemoryStream())
                {
                    await file.CopyToAsync(ms);
                    ms.Position = 0;

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(ms))
                    {
                        var ws = package.Workbook.Worksheets.FirstOrDefault();
                        if (ws == null)
                        {
                            return BadRequest("No worksheet found in the uploaded file.");
                        }

                        int rows = ws.Dimension.End.Row;
                        int cols = ws.Dimension.End.Column;

                        // Extract headers
                        var headers = new string[cols];
                        for (int col = 1; col <= cols; col++)
                        {
                            headers[col - 1] = ws.Cells[1, col].Text;
                        }

                        // Extract rows
                        var keyValuePairs = new Dictionary<string, string>[rows - 1];
                        for (int row = 2; row <= rows; row++)
                        {
                            var kvp = new Dictionary<string, string>();
                            for (int col = 1; col <= cols; col++)
                            {
                                var key = headers[col - 1];
                                var value = ws.Cells[row, col].Text;
                                kvp.Add(key, value);
                            }
                            keyValuePairs[row - 2] = kvp;
                        }

                        string htmlContent = GenerateHtmlAll(keyValuePairs);

                        var doc = new HtmlToPdfDocument()
                        {
                            GlobalSettings = new GlobalSettings
                            {
                                ColorMode = DinkToPdf.ColorMode.Color,
                                Orientation = Orientation.Portrait,
                                PaperSize = PaperKind.Letter,
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
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
        [HttpPost("GeneratePDFfromExcel")]
        public IActionResult GeneratePDFfromExcel(string url)
        {

            string QrUri = generateQRBase64(url);
            return Ok(new
            {
                QrUri = QrUri,
                message = "QR Generated Successfully!!!",
            });
        }
        private string generateQRBase64(string url) {
            QRCodeGenerator QrGenerator = new QRCodeGenerator();
            QRCodeData QrCodeInfo = QrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.Q);
            QRCode QrCode = new QRCode(QrCodeInfo);
            Bitmap QrBitmap = QrCode.GetGraphic(60);
            byte[] BitmapArray = QrBitmap.BitmapToByteArray();
            string QrUri = string.Format("data:image/png;base64,{0}", Convert.ToBase64String(BitmapArray));
            return QrUri;
        }
        private string GenerateHtmlAll(Dictionary<string, string>[] KeyValuePairs)
        {
            var allhtmls = @"
<!DOCTYPE html>
<html lang=""en"">

<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Document</title>
    <style>



        html {{ -webkit-print-color-adjust: exact; }}
        body {{
            margin: 0;
            padding: 0;
           line-height: 18px;
            font-size: 14px;
            text-align: justify;
        }}
       .dinktopdf-only {{    display: none; }}
       .wordletter {{
            width: 1027px;
            margin: 0 auto;
        }}
        #page-break {{
        page-break-before: always; /* or 'page-break-after: always;' depending on the requirement */
    }}
        table {{
            width: 100%;
            font-size: 14px;
            text-align: justify;
        }}

        .text-center{{
            text-align: center;
        }}
        ol li{{
            margin-bottom: 20px;
        }}
        ol{{
            background-color: #ccc;
        }}
        u{{
            font-weight: bold;
        }}
        ul li{{
            margin-bottom: 10px;
        }}
        p{{
            margin-bottom: 5px;
        }}

        @media print {
                html {{ -webkit-print-color-adjust: exact; }}
        body {{
            margin: 0;
            padding: 0;
            line-height: 18px;
            font-size: 14px;
            text-align: justify;
            font-family: ""Times New Roman"", Times, serif;
        }}
       .wordletter {{
            width: 1027px;
            margin: 0 auto;
        }}
        #page-break {{
        page-break-before: always; /* or 'page-break-after: always;' depending on the requirement */
    }}
        table {{
            width: 100%;
            font-size: 20px;
            text-align: justify;
        }}
        .text-center{{
            text-align: center;
        }}
        ol li{{
            margin-bottom: 20px;
        }}
        ol{{
            background-color: #ccc;
        }}
        u{{
            font-weight: bold;
        }}
        ul li{{
            margin-bottom: 10px;
        }}
        p{{
            margin-bottom: 5px;
        }}
        }
    </style>
</head>

<body>";
            for (int i = 0; i < KeyValuePairs.Length; i++)
            {
                var dist = KeyValuePairs[i];
                // Define the name variable
                string Name = TryGetValueFromDist(dist, "First") + " " + TryGetValueFromDist(dist, "Last");
                string Address1 = TryGetValueFromDist(dist, "Street");
                string Address2 = TryGetValueFromDist(dist, "Address 2");
                string City = TryGetValueFromDist(dist, "City");
                string State = TryGetValueFromDist(dist, "State");
                string Zip = TryGetValueFromDist(dist, "Zip");
                string ShareNumber = TryGetValueFromDist(dist, "Shares");
                string URL = TryGetValueFromDist(dist, "Link");
                string emailAddress = "ir@carecloud.com";
                string phone = "732-873-1351";
                string QrUri = generateQRBase64(URL);
                string htmlContent = $@"

      <div class=""wordletter"" >
<table align=""center"" class=""text-center"">
<tr>
<td align=""center"">
<img style=""margin: 0 auto;"" width=""250"" src=""https://localhost:7018/carecloud.png"">
</td>
</tr>
</table>
<table style=""margin-top:15px;"">
  <tr>
                    <td width=""650"">
                            <table>

                                <tr>

                                    <td colspan=""3"" style=""line-height: 22px;font-size: 20px; text-transform:titlecase;"">{ToTitleCase(Name.ToLower())}</td>
                                </tr>
                                <tr>

                                    <td style=""line-height: 22px;font-size: 20px; text-transform:titlecase;"" colspan=""3"">{ToTitleCase(Address1.ToLower())}</td>
                                </tr>
                                <tr>

                                    <td style=""line-height: 22px;font-size: 20px; "" colspan=""3"">{ToTitleCase(City.ToLower())},&nbsp;{State} {Zip}</td>

                                </tr>
                                <tr style=""vertical-align: top;"">
                                    <td style=""line-height: 22px;font-size: 20px; text-align:left;"" width=""50"" align=""left"" style=""vertical-align: top;""><strong style=""margin-top:27px; vertical-align: top; display:block"">Re:</strong></td>
                                    <td style=""line-height: 22px;font-size: 20px;"">

                                        <table style=""margin-top:25px; vertical-align: top;"">
                                            <tr>
                                                <td style=""line-height: 22px;font-size: 20px;"">CareCloud Series A Preferred Special Proxy Vote</td>
                                            </tr>
                                            <tr>
                                                <td style=""line-height: 22px;font-size: 20px; text-transform:titlecase;"">Shareholder: {ToTitleCase(Name.ToLower())}</td>
                                            </tr>
                                            <tr>
                                                <td style=""line-height: 22px;font-size: 20px;"">Number of shares entitled to vote: {ShareNumber}</td>
                                            </tr>
                                        </table>                         
                            </td>
                                </tr>
                            </table>
                    </td>
                    <td style=""line-height: 15px;font-size: 18px;"" width=""200"" align=""right"" style=""vertical-align: bottom; text-align: center;"">
                        <table>
                            <tr>
                                <td align=""center"" style=""line-height: 10px;font-size: 20px; text-align:center"">Vote Now</td>
                            </tr>
                            <tr>
                                <td align=""center"">  <img width=""150"" src=""{QrUri}""></td>
                            </tr>
                            <tr>
                                <td align=""center"" style=""line-height: 10px;font-size: 20px; text-align:center"" ><strong style=""line-height: 10px;font-size: 20px;"">SCAN HERE</strong></td>
                            </tr>
                        </table>

                    </td>
                </tr>

            </table>
        <table >
            <tr>
                <td style=""line-height: 22px;font-size: 20px; text-align: justify;""  colspan=""2""></br>
                    <p style=""text-transform:titlecase;"">Dear {ToTitleCase(Name.ToLower())},</p>
                    <p style = ""text-align: justify; line-height: 22px;font-size: 20px;"">We are pleased to share with you that as of today <strong><i>87%</i></strong> of your fellow Series A Preferred Shareholders
                        have submitted proxy votes in favor of both proposals being considered in the special proxy vote.
                        While there has been tremendous support, a passing vote will require a minimum quorum, which has not
                        yet been met – <i>we are close but your vote is critical.</i></p>
                    <p  style=""line-height: 22px;font-size: 20px;"">As you may have seen:</p>
                    <ul style=""style=""line-height: 22px;font-size: 20px;"""">
                        <li style=""line-height: 22px;font-size: 20px;""><i>Glass Lewis</i>, a leading proxy vote advisory firm, recommends a vote <strong>“FOR”</strong> both proposals.</li>
                        <li style=""line-height: 22px;font-size: 20px;""><i>87% of Series A Shareholders</i> indicated a vote <strong>“FOR”</strong> both proposals, as of August 8, 2024.</li>
                        <li style=""line-height: 22px;font-size: 20px;"">For your vote to count, you’ll need to vote <strong>“FOR”</strong> both proposals by <strong><i><u>August 21, 2024.</u></i></strong></li>
                    </ul>
                    <p style=""line-height: 18px;font-size: 23px; margin-bottom:4px;""><strong><u>How to Cast Your Vote:</u></strong></p>
                    <p style=""line-height: 22px;font-size: 20px; margin-bottom: 5px; margin-top:0px;"">To ensure your vote is counted you have several options:</p>
                </td>
            </tr>
        </table>

        <table style=""background-color: #d1f1fe; border:1px solid #009bde;"">
            <tr>
                <td style=""style=""line-height: 22px;font-size: 20px;"""" style=""background-color: #d1f1fe;"">
                    <ol>
                        <li style=""margin-bottom: 10px; line-height: 26px;font-size: 20px;""><strong><u style=""font-size: 23px; text-transform:uppercase; "">Vote Securely Online:</u></strong> Scan the above QR Code or visit:<br> <a style=""text-decoration:none"" href=""{URL}"">{URL}</a>.</li>
                        <li style=""margin-bottom: 10px; line-height: 26px;font-size: 20px;""><strong><u style=""font-size: 23px; text-transform:uppercase;"">Call to Vote:</u> </strong>You can vote by phone now or reach out with questions regarding the voting process at <strong>844-874-6164.</strong></li>
                        <li style=""line-height: 22px;font-size: 20px;""><strong><u style=""font-size: 23px; text-transform:uppercase;"">Send an Email:</u></strong> Send an email today to <a style=""text-decoration:none"" href=""mailto:carecloud@allianceadvisors.com"">carecloud@allianceadvisors.com</a> indicating that you would like to vote and then you will receive voting instructions.</li>
                    </ol>
                </td>
            </tr>
        </table>

        <table>
            <tr>
                <td style=""line-height: 22px;font-size: 20px;""colspan=""2"">
                    <p style = ""text-align: justify; margin-bottom: 0px; margin-top: 5px; line-height: 22px;font-size: 20px;"">To learn more about the special proxy, it is important that you review the Series A Preferred special
                        proxy filings carefully, which are available on the SEC’s website and at </p> 
                    <a style=""margin-bottom: 10px; display: block; text-decoration:none""  href=""https://ir.carecloud.com/series-a-special-proxy"">https://ir.carecloud.com/series-a-special-proxy</a>

                    <p style = ""line-height: 22px;font-size: 20px; text-align: justify; margin-bottom:25px !important; "">Please don’t hesitate to contact me via email <a style=""text-decoration:none"" href=""mailto:{emailAddress}"">{emailAddress}</a> or on my cell {phone} if I can be of any assistance. Thank you for your continued support of CareCloud.</p>
                    
                <table style=""margin-top:10px;"">
                    <tr>
                        <td>
                         <p style=""line-height: 15px;font-size: 20px; margin-bottom: 10px;"">Sincerely,</p>
                    <p style=""margin-bottom: 5px ; margin-top:5px;""><img style=""margin: 0 auto; width:auto; height: 40px; text-decoration:none""  src=""https://localhost:7018/signature.jpg""></p>
                    <p style=""line-height: 15px;font-size: 20px; margin-bottom: 10px; margin-top:0px;"">Stephen A. Snyder </p>
                    <p style=""line-height: 15px;font-size: 20px; margin-top:5px !important;"">President</p>
                        </td>    
                    </tr>
                </table>
                   
                </td>
            </tr>
        </table>
    </div><div style=""display:block; clear:both; page-break-after: always;""></div>";

                allhtmls += htmlContent;

            }
            allhtmls += @"</body></html>";
            return allhtmls;


        }
        private string ToTitleCase(string input)
        {
            TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
            return textInfo.ToTitleCase(input);
        }
        [HttpPost]
        public IActionResult Post(Dictionary<string, string>[] KeyValuePairs)
        {
            for (int i = 0; i < KeyValuePairs.Length; i++)
            {
                var generator = new BarcodeGenerator(EncodeTypes.QR);
                // Specify code text to encode
               
                    KeyValuePairs.ToList()[i].TryGetValue("Link",out string colValue);
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
                        ColorMode = DinkToPdf.ColorMode.Color,
                        Orientation = Orientation.Portrait,
                        PaperSize = PaperKind.Letter,
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
    public static class BitmapExtension
    {
        public static byte[] BitmapToByteArray(this Bitmap bitmap)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }
    }
}
