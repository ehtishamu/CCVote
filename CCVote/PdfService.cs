using Aspose.Pdf;
using Aspose.Pdf.Text;
using System.IO;

public class PdfService
{
    public byte[] ConvertHtmlToPdf(string htmlContent)
    {
        // Initialize a new Document
        Document pdfDocument = new Document();

        // Add a page to the PDF document
        Page page = pdfDocument.Pages.Add();

        // Create an HtmlFragment object from the HTML string
        HtmlFragment htmlFragment = new HtmlFragment(htmlContent);

        // Add the HTML content to the page
        page.Paragraphs.Add(htmlFragment);

        // Save the document to a MemoryStream
        using (MemoryStream ms = new MemoryStream())
        {
            pdfDocument.Save(ms);
            return ms.ToArray();
        }
    }
}
