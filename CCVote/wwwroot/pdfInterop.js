window.htmlToPdf = function (htmlContent, fileName) {
    const { jsPDF } = window.jspdf;

    // Define Letter page dimensions (in mm)
    const pdf = new jsPDF('p', 'mm', 'letter');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 10; // Set a margin (optional)
    const availableWidth = pageWidth - margin * 2;

    // Sanitize HTML content using DOMPurify
    const cleanHtml = DOMPurify.sanitize(htmlContent);

    pdf.html(cleanHtml, {
        callback: function (doc) {
            doc.save(fileName);
        },
        x: margin,
        y: margin,
        width: availableWidth, // Set width to fit content within page margins
        windowWidth: availableWidth * 4, // Scale content to fit width
        html2canvas: {
            scale: 0.25, // Reduce the scale to fit more content on the page
            useCORS: true // Enable cross-origin resource sharing if needed
        }
    });
};





window.htmlToPdfAll = function (htmlContent, fileName) {
    const { jsPDF } = window.jspdf;

    // Define Letter page dimensions (in mm)
    const pdf = new jsPDF('p', 'mm', 'letter');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 10; // Set a margin (optional)
    const availableWidth = pageWidth - margin * 2;

    // Sanitize HTML content using DOMPurify
    const cleanHtml = DOMPurify.sanitize(htmlContent);

    pdf.html(cleanHtml, {
        callback: function (doc) {
            doc.save(fileName);
        },
        x: margin,
        y: margin,
        width: availableWidth, // Set width to fit content within page margins
        windowWidth: availableWidth * 4, // Scale content to fit width
        htmlToPdfAll: {
            scale: 0.25, // Reduce the scale to fit more content on the page
            useCORS: true // Enable cross-origin resource sharing if needed
        },
        autoPaging: true,
    });
};

function openPdfInNewTab(pdfData) {
    const blob = new Blob([pdfData], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    window.open(url, '_blank');
}

