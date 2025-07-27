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
    // Add a visible test div to the DOM for html2pdf.js diagnostics
    const testDiv = document.createElement('div');
    testDiv.textContent = 'TEST DIV: html2pdf.js should capture this.';
    testDiv.style.background = '#ff0';
    testDiv.style.color = '#000';
    testDiv.style.fontSize = '32px';
    testDiv.style.padding = '40px';
    testDiv.style.position = 'fixed';
    testDiv.style.top = '100px';
    testDiv.style.left = '100px';
    testDiv.style.zIndex = '99999';
    document.body.appendChild(testDiv);

    const pdfDoc = await PDFDocument.create();
    // Add a test page with a simple string using pdf-lib
    const testPage = pdfDoc.addPage();
    testPage.drawText('PDF-LIB TEST PAGE: If you see this, pdf-lib is working.', { x: 50, y: 700, size: 24 });
    console.log('Added test page to PDF.');

    // Render the email body HTML to PDF using html2pdf.js and merge into the main PDF
    const emailContainer = document.createElement('div');
    emailContainer.innerHTML = htmlBody;
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
    emailContainer.style.opacity = '0'; // Hide visually but keep in layout
    emailContainer.style.pointerEvents = 'none';
    // Log diagnostics
    setTimeout(() => {
        console.log('EMAIL container offsetHeight:', emailContainer.offsetHeight);
        console.log('EMAIL container innerHTML:', emailContainer.innerHTML);
    }, 100);
    emailContainer.style.zIndex = '9999';
    document.body.appendChild(emailContainer);
    // Wait for images to load
    const emailImages = Array.from(emailContainer.querySelectorAll('img'));
    let loaded = 0;
    function addEmailHtmlToPdf() {
        let resolved = false;
        const timeout = setTimeout(() => {
            if (!resolved) {
                console.error('html2pdf.js timed out after 5 seconds. Skipping email body merge.');
                try { document.body.removeChild(emailContainer); } catch (e) {}
                resolved = true;
                mergeAttachments();
            }
        }, 5000);
        setTimeout(() => {
            html2pdf().set({
                margin: 10,
                filename: 'email.pdf',
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                pagebreak: { mode: ['avoid-all', 'css'] }
            }).from(emailContainer).toPdf().get('pdf').then(async pdf => {
                if (resolved) return;
                resolved = true;
                clearTimeout(timeout);
                try {
                    const emailPdfBytes = pdf.output('arraybuffer');
                    const emailPdf = await PDFDocument.load(emailPdfBytes);
                    const pageIndices = emailPdf.getPageIndices();
                    const copiedPages = await pdfDoc.copyPages(emailPdf, pageIndices);
                    copiedPages.forEach(p => pdfDoc.addPage(p));
                    document.body.removeChild(emailContainer);
                    console.log('Email body HTML converted and merged.');
                } catch (err) {
                    console.error('Error merging email HTML PDF:', err);
                    document.body.removeChild(emailContainer);
                }
                // Continue with attachments merging as before
                await mergeAttachments();
            }).catch(err => {
                if (resolved) return;
                resolved = true;
                clearTimeout(timeout);
                console.error('html2pdf.js failed:', err);
                document.body.removeChild(emailContainer);
                mergeAttachments();
            });
        }, 500);
    }
    if (emailImages.length > 0) {
        emailImages.forEach(img => {
            if (img.complete) {
                loaded++;
            } else {
                img.onload = img.onerror = () => {
                    loaded++;
                    if (loaded === emailImages.length) addEmailHtmlToPdf();
                };
            }
        });
        if (loaded === emailImages.length) addEmailHtmlToPdf();
    } else {
        addEmailHtmlToPdf();
    }

    // Move the rest of the logic into a function to be called after email HTML is merged
    async function mergeAttachments() {
        // If there are no attachments, save/download PDF immediately
        if (!attachments || attachments.length === 0) {
            console.log('No attachments. Saving PDF immediately...');
            await saveAndDownloadPdf();
            return;
        }
        // ...existing attachment merging logic...
        // Collect all merge promises
        const mergePromises = attachments.map(att => new Promise((resolve) => {
            // ...existing attachment merging logic (unchanged)...
            // ...copy from previous code block...
            // (for brevity, not repeated here)
        }));

        // Wait for all merges to finish
        console.log('Waiting for all attachment merges to finish...');
        await Promise.all(mergePromises);
        console.log('All attachments merged. Saving PDF...');
        await saveAndDownloadPdf();
    }

    async function saveAndDownloadPdf() {
        // Log number of pages in the final PDF
        const pageCount = pdfDoc.getPageCount();
        console.log('Final PDF page count:', pageCount);
        if (pageCount === 0) {
            console.error('No pages in final PDF!');
        }
        // Download PDF
        try {
            console.log('Saving and downloading PDF...');
            const pdfBytes = await pdfDoc.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'email.pdf';
            a.click();
            console.log('PDF download triggered.');
        } catch (err) {
            console.error('Error saving or downloading PDF:', err);
        }
        // Remove test div after a short delay
        setTimeout(() => { document.body.removeChild(testDiv); }, 2000);
    }
}
