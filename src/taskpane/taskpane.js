
// Load pdf-lib from CDN for browser compatibility
let PDFDocument;
const pdfLibScript = document.createElement('script');
pdfLibScript.src = 'https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js';
pdfLibScript.onload = () => {
    PDFDocument = window.PDFLib.PDFDocument;
};
document.head.appendChild(pdfLibScript);

Office.onReady(() => {
    document.getElementById('convertBtn').onclick = async function() {
        // Show loading feedback
        this.disabled = true;
        this.textContent = 'Converting...';

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
                    alert('PDF library not loaded yet. Please try again in a moment.');
                    this.disabled = false;
                    this.textContent = 'Convert Email to PDF';
                    return;
                }

                await createPdf(htmlBody, attachments, imgSources);
                this.disabled = false;
                this.textContent = 'Convert Email to PDF';
            } else {
                alert('Failed to get email body: ' + result.error);
                this.disabled = false;
                this.textContent = 'Convert Email to PDF';
            }
        });
    };
});

// ...existing Office.js code to get htmlBody, attachments, imgSources...


async function createPdf(htmlBody, attachments, imgSources) {
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage();
    page.drawText(htmlBody.replace(/<[^>]+>/g, '').slice(0, 4000)); // Strips HTML tags, limits length

    // Add embedded images
    for (const src of imgSources) {
        try {
            const imgBytes = await fetch(src).then(res => res.arrayBuffer());
            const img = await pdfDoc.embedPng(imgBytes);
            const { width, height } = img.scale(0.5);
            page.drawImage(img, { x: 50, y: 400, width, height });
        } catch (e) {
            console.error('Image error:', src, e);
        }
    }

    // Add image attachments (PNG/JPG only)
    for (const att of attachments) {
        if (att.contentType.startsWith('image/')) {
            await new Promise((resolve) => {
                Office.context.mailbox.item.getAttachmentContentAsync(att.id, { asyncContext: att }, async (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const { content, contentType } = result.value;
                        let img;
                        if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                            const byteArr = Uint8Array.from(atob(content), c => c.charCodeAt(0));
                            if (contentType === 'image/png') {
                                img = await pdfDoc.embedPng(byteArr);
                            } else if (contentType === 'image/jpeg') {
                                img = await pdfDoc.embedJpg(byteArr);
                            }
                            if (img) {
                                const attPage = pdfDoc.addPage();
                                const { width, height } = img.scale(0.5);
                                attPage.drawImage(img, { x: 50, y: 400, width, height });
                            }
                        }
                    } else {
                        console.error('Failed to get attachment content:', result.error);
                    }
                    resolve();
                });
            });
        }
    }

    // Download PDF
    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'email.pdf';
    a.click();
}
