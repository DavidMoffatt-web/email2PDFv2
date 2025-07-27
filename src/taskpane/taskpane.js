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
    // Use the visible emailToPdfDiv for PDF generation
    const emailDiv = document.getElementById('emailToPdfDiv');
    if (!emailDiv) {
        alert('Email PDF div not found!');
        return;
    }
    // Sanitize email HTML: remove <meta>, <script>, <style>, <link>
    let sanitizedHtml = htmlBody.replace(/<\/?(meta|script|style|link)[^>]*>/gi, '');
    emailDiv.innerHTML = sanitizedHtml;
    emailDiv.style.display = 'block';
    // Wait for images to load before generating PDF
    const emailImages = Array.from(emailDiv.querySelectorAll('img'));
    let loaded = 0;
    function generatePdf() {
        html2pdf().set({
            margin: 10,
            filename: 'email.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
            pagebreak: { mode: ['avoid-all', 'css'] }
        }).from(emailDiv).save().then(() => {
            emailDiv.style.display = 'none';
            console.log('Email div PDF generated and download triggered.');
        }).catch(err => {
            emailDiv.style.display = 'none';
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
