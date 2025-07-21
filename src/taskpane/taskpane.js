document.addEventListener("DOMContentLoaded", function () {
    // Initialize the task pane
    document.getElementById("myButton").onclick = function () {
        // Handle button click
        Office.context.mailbox.item.body.getAsync("text", function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("output").innerText = result.value;
            } else {
                document.getElementById("output").innerText = "Error: " + result.error.message;
            }
        });
    };
});

// Wait for Office to initialize
Office.onReady(() => {
    document.getElementById('convertBtn').onclick = async function() {
        // Get the email body (HTML)
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const htmlBody = result.value;

                // Get attachments
                const attachments = Office.context.mailbox.item.attachments; // Array of attachment objects

                // Parse embedded images from HTML body
                const imgSources = [];
                const parser = new DOMParser();
                const doc = parser.parseFromString(htmlBody, 'text/html');
                const imgs = doc.querySelectorAll('img');
                imgs.forEach(img => {
                    if (img.src) imgSources.push(img.src);
                });

                // You now have:
                // - htmlBody: the email body as HTML
                // - attachments: array of attachments
                // - imgSources: array of embedded image URLs

                // Next: Pass these to your PDF generation logic
                console.log({ htmlBody, attachments, imgSources });
            } else {
                console.error('Failed to get email body:', result.error);
            }
        });
    };
});

import { PDFDocument } from 'pdf-lib';

// ...existing Office.js code to get htmlBody, attachments, imgSources...

async function createPdf(htmlBody, attachments, imgSources) {
    const pdfDoc = await PDFDocument.create();

    // Add email body as a page (simple text for demo)
    const page = pdfDoc.addPage();
    page.drawText(htmlBody.replace(/<[^>]+>/g, '')); // Strips HTML tags for demo

    // Add embedded images
    for (const src of imgSources) {
        try {
            const imgBytes = await fetch(src).then(res => res.arrayBuffer());
            const img = await pdfDoc.embedPng(imgBytes); // or embedJpg for JPGs
            const { width, height } = img.scale(0.5);
            page.drawImage(img, { x: 50, y: 400, width, height });
        } catch (e) {
            console.error('Image error:', src, e);
        }
    }

for (const att of attachments) {
    // Only process image attachments for demo (e.g., PNG, JPG)
    if (att.contentType.startsWith('image/')) {
        Office.context.mailbox.item.getAttachmentContentAsync(att.id, { asyncContext: att }, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const { content, contentType } = result.value;
                let img;
                if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                    if (contentType === 'image/png') {
                        img = await pdfDoc.embedPng(Uint8Array.from(atob(content), c => c.charCodeAt(0)));
                    } else if (contentType === 'image/jpeg') {
                        img = await pdfDoc.embedJpg(Uint8Array.from(atob(content), c => c.charCodeAt(0)));
                    }
                    if (img) {
                        const page = pdfDoc.addPage();
                        const { width, height } = img.scale(0.5);
                        page.drawImage(img, { x: 50, y: 400, width, height });
                    }
                }
            } else {
                console.error('Failed to get attachment content:', result.error);
            }
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
