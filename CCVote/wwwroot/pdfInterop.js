window.htmlToPdf = function (htmlContent, fileName) {
    const { jsPDF } = window.jspdf;

    // Define A4 page dimensions (in points)
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 10; // Set a margin (optional)
    const availableWidth = pageWidth - margin * 2;
    const availableHeight = pageHeight - margin * 2;

    // Sanitize HTML content using DOMPurify
    const cleanHtml = DOMPurify.sanitize(htmlContent);

    pdf.html(cleanHtml, {
        callback: function (doc) {
            doc.save(fileName);
        },
        x: margin,
        y: margin,
        width: availableWidth, // Set width to fit content within page margins
        windowWidth: availableWidth * 2.83465, // Scale content to fit width
        html2canvas: {
            scale: 0.3 // Adjust this scale factor as needed to fit content on the page
        }
    });
};
