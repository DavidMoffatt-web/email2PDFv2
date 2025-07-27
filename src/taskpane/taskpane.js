console.log('taskpane.js loaded');
// Load pdf-lib from CDN for browser compatibility
let PDFDocument;
console.log('Loading PDFLib script...');
const pdfLibScript = document.createElement('script');
pdfLibScript.src = 'https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js';
pdfLibScript.onload = () => {
    console.log('PDFLib script loaded.');
    PDFDocument = window.PDFLib.PDFDocument;
};
document.head.appendChild(pdfLibScript);



Office.onReady(() => {
    console.log('Office.js is ready');
    const btn = document.getElementById('convertBtn');
    console.log('convertBtn found:', btn);
    btn.onclick = async function() {
        console.log('Convert button clicked.');
        this.disabled = true;
        this.textContent = 'Converting...';
        try {
            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const htmlBody = result.value;
                    const attachments = Office.context.mailbox.item.attachments;
                    const imgSources = [];
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(htmlBody, 'text/html');
                    const imgs = doc.querySelectorAll('img');
                    imgs.forEach(img => {
                        if (img.src) imgSources.push(img.src);
                    });

                    // Wait for pdf-lib to load
                    if (!PDFDocument) {
                        console.error('PDF library not loaded yet.');
                        alert('PDF library not loaded yet. Please try again in a moment.');
                        this.disabled = false;
                        this.textContent = 'Convert Email to PDF';
                        return;
                    }

                    console.log('Calling createPdf...');
                    try {
                        await createPdf(htmlBody, attachments, imgSources);
                    } catch (err) {
                        console.error('Error in createPdf:', err);
                        alert('PDF generation failed: ' + err.message);
                    }
                    this.disabled = false;
                    this.textContent = 'Convert Email to PDF';
                } else {
                    console.error('Failed to get email body:', result.error);
                    alert('Failed to get email body: ' + result.error);
                    this.disabled = false;
                    this.textContent = 'Convert Email to PDF';
                }
            });
        } catch (err) {
            console.error('Unexpected error in button handler:', err);
            alert('Unexpected error: ' + err.message);
            this.disabled = false;
            this.textContent = 'Convert Email to PDF';
        }
    };
});

// ...existing Office.js code to get htmlBody, attachments, imgSources...


async function createPdf(htmlBody, attachments, imgSources) {
    console.log('createPdf called');
    console.log('Starting PDF creation...');
    // Create a container for the email HTML
    const emailContainer = document.createElement('div');
    // For debugging: use a simple static string instead of the email HTML
    emailContainer.innerHTML = '<h1>Hello PDF!</h1><p>This is a test of html2pdf in Office add-in.</p>';
    emailContainer.style.background = '#fff';
    emailContainer.style.padding = '24px';
    emailContainer.style.fontFamily = 'Arial, sans-serif';
    emailContainer.style.width = '800px';
    emailContainer.style.maxWidth = '100%';
    emailContainer.style.minHeight = '400px';
    emailContainer.style.height = 'auto';
    emailContainer.style.overflow = 'visible';
    emailContainer.style.position = 'fixed';
    emailContainer.style.left = '0';
    emailContainer.style.top = '0';
    emailContainer.style.zIndex = '9999';
    emailContainer.style.visibility = 'visible'; // Make visible for testing
    emailContainer.style.border = '2px solid red'; // Debug border
    document.body.appendChild(emailContainer);
    // Log computed style and bounding rect for diagnostics
    setTimeout(() => {
        const rect = emailContainer.getBoundingClientRect();
        const style = window.getComputedStyle(emailContainer);
        console.log('EMAIL container bounding rect:', rect);
        console.log('EMAIL container computed style:', style);
        console.log('EMAIL container offsetHeight:', emailContainer.offsetHeight);
        console.log('EMAIL container innerHTML:', emailContainer.innerHTML);
    }, 100);
    // Wait for images to load before generating PDF
    const emailImages = Array.from(emailContainer.querySelectorAll('img'));
    let loaded = 0;
    function generatePdf() {
        html2pdf().set({
            margin: 10,
            filename: 'email.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
            pagebreak: { mode: ['avoid-all', 'css'] }
        }).from(emailContainer).save().then(() => {
            document.body.removeChild(emailContainer);
            console.log('Email body PDF generated and download triggered.');
        }).catch(err => {
            document.body.removeChild(emailContainer);
            console.error('html2pdf.js failed:', err);
        });
    }
    if (emailImages.length > 0) {
        emailImages.forEach(img => {
            if (img.complete) {
                loaded++;
            } else {
                img.onload = img.onerror = () => {
                    loaded++;
                    if (loaded === emailImages.length) generatePdf();
                };
            }
        });
        if (loaded === emailImages.length) generatePdf();
    } else {
        generatePdf();
    }
}
