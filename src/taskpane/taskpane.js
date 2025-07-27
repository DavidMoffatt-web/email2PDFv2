console.log('taskpane.js loaded');
// Load pdf-lib from CDN for browser compatibility
let PDFDocument, StandardFonts;
console.log('Loading PDFLib script...');
const pdfLibScript = document.createElement('script');
pdfLibScript.src = 'https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js';
pdfLibScript.onload = () => {
    console.log('PDFLib script loaded.');
    PDFDocument = window.PDFLib.PDFDocument;
    StandardFonts = window.PDFLib.StandardFonts;
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
            // Gather metadata first
            const item = Office.context.mailbox.item;
            const getMetadata = () => {
                return {
                    from: (item.from && item.from.displayName ? item.from.displayName + ' <' + item.from.emailAddress + '>' : ''),
                    to: (item.to && item.to.length ? item.to.map(e => e.displayName + ' <' + e.emailAddress + '>').join('; ') : ''),
                    cc: (item.cc && item.cc.length ? item.cc.map(e => e.displayName + ' <' + e.emailAddress + '>').join('; ') : ''),
                    subject: item.subject || '',
                    date: (item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : '')
                };
            };
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
                        // Use UI message div instead of alert
                        let msgDiv = document.getElementById('pdf2email-message');
                        if (!msgDiv) {
                            msgDiv = document.createElement('div');
                            msgDiv.id = 'pdf2email-message';
                            msgDiv.style.color = '#b00';
                            msgDiv.style.fontSize = '1em';
                            msgDiv.style.margin = '12px 0';
                            msgDiv.style.display = 'block';
                            const appDiv = document.getElementById('app');
                            if (appDiv) appDiv.parentNode.insertBefore(msgDiv, appDiv.nextSibling);
                            else document.body.insertBefore(msgDiv, document.body.firstChild);
                        }
                        msgDiv.textContent = 'PDF library not loaded yet. Please try again in a moment.';
                        this.disabled = false;
                        this.textContent = 'Convert Email to PDF';
                        return;
                    }

                    // Get selected mode
                    const modeSel = document.getElementById('pdfMode');
                    const mode = modeSel ? modeSel.value : 'full';
                    console.log('PDF Generation Mode:', mode);

                    // Compose metadata header HTML
                    const meta = getMetadata();
                    const metaHtml = `<div style="background:#f5f5f5;padding:12px 18px 12px 18px;margin-bottom:18px;border-radius:6px;font-size:1em;line-height:1.5;">
                        <div><b>From:</b> ${meta.from}</div>
                        <div><b>To:</b> ${meta.to}</div>
                        ${meta.cc ? `<div><b>CC:</b> ${meta.cc}</div>` : ''}
                        <div><b>Subject:</b> ${meta.subject}</div>
                        <div><b>Sent:</b> ${meta.date}</div>
                    </div>`;

                    // Pass metadata HTML to PDF logic
                    try {
                        if (mode === 'full') {
                            await createPdf(metaHtml + htmlBody, attachments, imgSources);
                        } else if (mode === 'individual') {
                            await createIndividualPdfs(metaHtml + htmlBody, attachments, imgSources, false, metaHtml);
                        } else if (mode === 'images') {
                            await createIndividualPdfs(metaHtml + htmlBody, attachments, imgSources, true, metaHtml);
                        } else {
                            await createPdf(metaHtml + htmlBody, attachments, imgSources);
                        }
                    } catch (err) {
                        console.error('Error in PDF generation:', err);
                        let msgDiv = document.getElementById('pdf2email-message');
                        if (!msgDiv) {
                            msgDiv = document.createElement('div');
                            msgDiv.id = 'pdf2email-message';
                            msgDiv.style.color = '#b00';
                            msgDiv.style.fontSize = '1em';
                            msgDiv.style.margin = '12px 0';
                            msgDiv.style.display = 'block';
                            const appDiv = document.getElementById('app');
                            if (appDiv) appDiv.parentNode.insertBefore(msgDiv, appDiv.nextSibling);
                            else document.body.insertBefore(msgDiv, document.body.firstChild);
                        }
                        msgDiv.textContent = 'PDF generation failed: ' + err.message;
                    }
                    this.disabled = false;
                    this.textContent = 'Convert Email to PDF';
                } else {
                    console.error('Failed to get email body:', result.error);
                    let msgDiv = document.getElementById('pdf2email-message');
                    if (!msgDiv) {
                        msgDiv = document.createElement('div');
                        msgDiv.id = 'pdf2email-message';
                        msgDiv.style.color = '#b00';
                        msgDiv.style.fontSize = '1em';
                        msgDiv.style.margin = '12px 0';
                        msgDiv.style.display = 'block';
                        const appDiv = document.getElementById('app');
                        if (appDiv) appDiv.parentNode.insertBefore(msgDiv, appDiv.nextSibling);
                        else document.body.insertBefore(msgDiv, document.body.firstChild);
                    }
                    msgDiv.textContent = 'Failed to get email body: ' + result.error;
                    this.disabled = false;
                    this.textContent = 'Convert Email to PDF';
                }
            });
        } catch (err) {
            console.error('Unexpected error in button handler:', err);
            let msgDiv = document.getElementById('pdf2email-message');
            if (!msgDiv) {
                msgDiv = document.createElement('div');
                msgDiv.id = 'pdf2email-message';
                msgDiv.style.color = '#b00';
                msgDiv.style.fontSize = '1em';
                msgDiv.style.margin = '12px 0';
                msgDiv.style.display = 'block';
                const appDiv = document.getElementById('app');
                if (appDiv) appDiv.parentNode.insertBefore(msgDiv, appDiv.nextSibling);
                else document.body.insertBefore(msgDiv, document.body.firstChild);
            }
            msgDiv.textContent = 'Unexpected error: ' + err.message;
            this.disabled = false;
            this.textContent = 'Convert Email to PDF';
        }
    };
// Helper for individual PDFs mode
// metaHtml is optional, for prepending to the email body only
async function createIndividualPdfs(htmlBody, attachments, imgSources, includeEmailImages, metaHtml) {
    // UI message div
    let msgDiv = document.getElementById('pdf2email-message');
    function showMessage(msg) {
        msgDiv.textContent = msg;
        msgDiv.style.display = 'block';
    }
    function hideMessage() {
        msgDiv.textContent = '';
        msgDiv.style.display = 'none';
    }
    hideMessage();

    // Helper: Download attachment as Blob
    async function getAttachmentContent(attachment) {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    let content = result.value.content;
                    let type = result.value.format;
                    let name = attachment.name;
                    let mimeType = attachment.contentType || '';
                    if (type === 'base64') {
                        const byteChars = atob(content);
                        const byteNumbers = new Array(byteChars.length);
                        for (let i = 0; i < byteChars.length; i++) {
                            byteNumbers[i] = byteChars.charCodeAt(i);
                        }
                        const byteArray = new Uint8Array(byteNumbers);
                        resolve(new Blob([byteArray], { type: mimeType }));
                    } else if (type === 'file') {
                        fetch(content).then(r => r.blob()).then(resolve).catch(reject);
                    } else {
                        reject(new Error('Unknown attachment format: ' + type));
                    }
                } else {
                    reject(result.error);
                }
            });
        });
    }

    // Download all attachments as blobs and detect type
    let attachmentBlobs = [];
    if (attachments && attachments.length > 0) {
        for (let att of attachments) {
            try {
                const blob = await getAttachmentContent(att);
                attachmentBlobs.push({
                    name: att.name,
                    type: att.contentType || blob.type,
                    blob
                });
                console.log('Downloaded attachment:', att.name, att.contentType || blob.type);
            } catch (err) {
                console.error('Failed to download attachment', att.name, err);
            }
        }
    }

    // Helper: Convert HTML element to PDF and trigger download
    async function htmlToPdfAndDownload(element, filename) {
        return new Promise((resolve, reject) => {
            html2pdf().set({
                margin: 0,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                pagebreak: { mode: ['avoid-all', 'css'] }
            }).from(element).toPdf().get('pdf').then(pdf => {
                const blob = pdf.output('blob');
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                a.click();
                resolve();
            }).catch(reject);
        });
    }

    // 1. Email body as PDF (with metadata header if provided)
    const emailDiv = document.getElementById('emailToPdfDiv');
    let sanitizedHtml = htmlBody.replace(/<\/?(meta|script|style|link)[^>]*>/gi, '');
    // If metaHtml is provided, ensure it's prepended (avoid double-prepending)
    if (metaHtml && !sanitizedHtml.startsWith(metaHtml)) {
        sanitizedHtml = metaHtml + sanitizedHtml;
    }
    emailDiv.innerHTML = sanitizedHtml;
    emailDiv.style.display = 'block';
    await htmlToPdfAndDownload(emailDiv, 'email.pdf');
    emailDiv.style.display = 'none';

    // 2. Attachments as PDFs
    for (let att of attachmentBlobs) {
        if (att.type.startsWith('image/')) {
            const img = document.createElement('img');
            img.src = URL.createObjectURL(att.blob);
            img.style.maxWidth = '100%';
            img.style.display = 'block';
            const imgDiv = document.createElement('div');
            imgDiv.style.background = '#fff';
            imgDiv.style.width = '100%';
            imgDiv.style.maxWidth = '794px';
            imgDiv.style.minHeight = '400px';
            imgDiv.style.margin = '0 auto';
            imgDiv.style.boxSizing = 'border-box';
            imgDiv.style.padding = '24px';
            imgDiv.appendChild(img);
            document.body.appendChild(imgDiv);
            await new Promise(res => {
                if (img.complete) res();
                else img.onload = img.onerror = res;
            });
            try {
                await htmlToPdfAndDownload(imgDiv, att.name.replace(/\.[^/.]+$/, '') + '.pdf');
            } catch (err) {
                console.error('Failed to convert image attachment to PDF:', att.name, err);
            }
            document.body.removeChild(imgDiv);
        } else if (att.name.toLowerCase().endsWith('.docx')) {
            try {
                const arrayBuffer = await att.blob.arrayBuffer();
                const mammothResult = await window.mammoth.convertToHtml({ arrayBuffer });
                const docxDiv = document.createElement('div');
                docxDiv.innerHTML = mammothResult.value;
                docxDiv.style.background = '#fff';
                docxDiv.style.width = '100%';
                docxDiv.style.maxWidth = '794px';
                docxDiv.style.minHeight = '400px';
                docxDiv.style.margin = '0 auto';
                docxDiv.style.boxSizing = 'border-box';
                docxDiv.style.padding = '24px';
                document.body.appendChild(docxDiv);
                await htmlToPdfAndDownload(docxDiv, att.name.replace(/\.[^/.]+$/, '') + '.pdf');
                document.body.removeChild(docxDiv);
            } catch (err) {
                console.error('Failed to convert DOCX attachment to PDF:', att.name, err);
            }
        } else if (att.type === 'application/pdf' || att.name.toLowerCase().endsWith('.pdf')) {
            try {
                // Download as-is
                const arrayBuffer = await att.blob.arrayBuffer();
                const blob = new Blob([arrayBuffer], { type: 'application/pdf' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = att.name;
                a.click();
            } catch (err) {
                console.error('Failed to download PDF attachment:', att.name, err);
            }
        } else {
            console.warn('Skipping unsupported attachment type:', att.name, att.type);
        }
    }

    // 3. Embedded images in email (if requested)
    if (includeEmailImages) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlBody, 'text/html');
        const imgs = doc.querySelectorAll('img');
        let imgCount = 0;
        for (let img of imgs) {
            if (img.src) {
                try {
                    const imgEl = document.createElement('img');
                    imgEl.src = img.src;
                    imgEl.style.maxWidth = '100%';
                    imgEl.style.display = 'block';
                    const imgDiv = document.createElement('div');
                    imgDiv.style.background = '#fff';
                    imgDiv.style.width = '100%';
                    imgDiv.style.maxWidth = '794px';
                    imgDiv.style.minHeight = '400px';
                    imgDiv.style.margin = '0 auto';
                    imgDiv.style.boxSizing = 'border-box';
                    imgDiv.style.padding = '24px';
                    imgDiv.appendChild(imgEl);
                    document.body.appendChild(imgDiv);
                    await new Promise(res => {
                        if (imgEl.complete) res();
                        else imgEl.onload = imgEl.onerror = res;
                    });
                    await htmlToPdfAndDownload(imgDiv, `email_image_${++imgCount}.pdf`);
                    document.body.removeChild(imgDiv);
                } catch (err) {
                    console.error('Failed to convert embedded email image to PDF:', img.src, err);
                }
            }
        }
    }

    showMessage('All PDFs have been generated and downloaded.');
    setTimeout(hideMessage, 4000);
}
});

// ...existing Office.js code to get htmlBody, attachments, imgSources...


async function createPdf(htmlBody, attachments, imgSources) {
    // Add or get a message div for UI feedback
    let msgDiv = document.getElementById('pdf2email-message');
    if (!msgDiv) {
        msgDiv = document.createElement('div');
        msgDiv.id = 'pdf2email-message';
        msgDiv.style.color = '#b00';
        msgDiv.style.fontSize = '1em';
        msgDiv.style.margin = '12px 0';
        msgDiv.style.display = 'none';
        const appDiv = document.getElementById('app');
        if (appDiv) appDiv.parentNode.insertBefore(msgDiv, appDiv.nextSibling);
        else document.body.insertBefore(msgDiv, document.body.firstChild);
    }
    function showMessage(msg) {
        msgDiv.textContent = msg;
        msgDiv.style.display = 'block';
    }
    function hideMessage() {
        msgDiv.textContent = '';
        msgDiv.style.display = 'none';
    }
    console.log('createPdf called');
    console.log('Starting PDF creation...');
    // Use the visible emailToPdfDiv for PDF generation
    const emailDiv = document.getElementById('emailToPdfDiv');
    if (!emailDiv) {
        showMessage('Email PDF div not found!');
        return;
    }
    // Sanitize email HTML: remove <meta>, <script>, <style>, <link>
    let sanitizedHtml = htmlBody.replace(/<\/?(meta|script|style|link)[^>]*>/gi, '');
    emailDiv.innerHTML = sanitizedHtml;
    emailDiv.style.display = 'block';
    hideMessage();
    // Wait for images to load before generating PDF
    const emailImages = Array.from(emailDiv.querySelectorAll('img'));
    let loaded = 0;

    // Helper: Download attachment as Blob
    async function getAttachmentContent(attachment) {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    let content = result.value.content;
                    let type = result.value.format;
                    let name = attachment.name;
                    let mimeType = attachment.contentType || '';
                    // Convert base64 to Blob if needed
                    if (type === 'base64') {
                        const byteChars = atob(content);
                        const byteNumbers = new Array(byteChars.length);
                        for (let i = 0; i < byteChars.length; i++) {
                            byteNumbers[i] = byteChars.charCodeAt(i);
                        }
                        const byteArray = new Uint8Array(byteNumbers);
                        resolve(new Blob([byteArray], { type: mimeType }));
                    } else if (type === 'file') {
                        // File is a URL (not common in Outlook web)
                        fetch(content).then(r => r.blob()).then(resolve).catch(reject);
                    } else {
                        reject(new Error('Unknown attachment format: ' + type));
                    }
                } else {
                    reject(result.error);
                }
            });
        });
    }

    // Step 1: Download all attachments as blobs and detect type
    let attachmentBlobs = [];
    if (attachments && attachments.length > 0) {
        for (let att of attachments) {
            try {
                const blob = await getAttachmentContent(att);
                attachmentBlobs.push({
                    name: att.name,
                    type: att.contentType || blob.type,
                    blob
                });
                console.log('Downloaded attachment:', att.name, att.contentType || blob.type);
            } catch (err) {
                console.error('Failed to download attachment', att.name, err);
            }
        }
    }

    // Step 2: Convert email HTML to PDF (as ArrayBuffer) using html2pdf
    async function htmlToPdfBuffer(element) {
        return new Promise((resolve, reject) => {
            html2pdf().set({
                margin: 0,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                pagebreak: { mode: ['avoid-all', 'css'] }
            }).from(element).toPdf().get('pdf').then(pdf => {
                resolve(pdf.output('arraybuffer'));
            }).catch(reject);
        });
    }

    // Step 3: Convert/merge attachments
    let pdfBuffers = [];
    let skippedFiles = [];
    // 1. Email body as first PDF buffer
    async function processEmailAndAttachments() {
        // Email PDF
        const emailPdfBuffer = await htmlToPdfBuffer(emailDiv);
        pdfBuffers.push(emailPdfBuffer);
        emailDiv.style.display = 'none';

        // 2. Attachments
        for (let att of attachmentBlobs) {
            // Images
            if (att.type.startsWith('image/')) {
                // Create an <img> element for the image
                const img = document.createElement('img');
                img.src = URL.createObjectURL(att.blob);
                img.style.maxWidth = '100%';
                img.style.display = 'block';
                // Create a container div for the image
                const imgDiv = document.createElement('div');
                imgDiv.style.background = '#fff';
                imgDiv.style.width = '100%';
                imgDiv.style.maxWidth = '794px';
                imgDiv.style.minHeight = '400px';
                imgDiv.style.margin = '0 auto';
                imgDiv.style.boxSizing = 'border-box';
                imgDiv.style.padding = '24px';
                imgDiv.appendChild(img);
                document.body.appendChild(imgDiv);
                // Wait for image to load
                await new Promise(res => {
                    if (img.complete) res();
                    else img.onload = img.onerror = res;
                });
                // Convert to PDF buffer
                try {
                    const imgPdfBuffer = await htmlToPdfBuffer(imgDiv);
                    pdfBuffers.push(imgPdfBuffer);
                } catch (err) {
                    console.error('Failed to convert image attachment to PDF:', att.name, err);
                    skippedFiles.push(att.name + ' (image conversion failed)');
                }
                document.body.removeChild(imgDiv);
            }
            // DOCX
            else if (att.name.toLowerCase().endsWith('.docx')) {
                try {
                    const arrayBuffer = await att.blob.arrayBuffer();
                    const mammothResult = await window.mammoth.convertToHtml({ arrayBuffer });
                    // Render HTML to a div
                    const docxDiv = document.createElement('div');
                    docxDiv.innerHTML = mammothResult.value;
                    docxDiv.style.background = '#fff';
                    docxDiv.style.width = '100%';
                    docxDiv.style.maxWidth = '794px';
                    docxDiv.style.minHeight = '400px';
                    docxDiv.style.margin = '0 auto';
                    docxDiv.style.boxSizing = 'border-box';
                    docxDiv.style.padding = '24px';
                    document.body.appendChild(docxDiv);
                    // Convert to PDF buffer
                    const docxPdfBuffer = await htmlToPdfBuffer(docxDiv);
                    pdfBuffers.push(docxPdfBuffer);
                    document.body.removeChild(docxDiv);
                } catch (err) {
                    console.error('Failed to convert DOCX attachment to PDF:', att.name, err);
                    skippedFiles.push(att.name + ' (DOCX conversion failed)');
                }
            }
            // PDF
            else if (att.type === 'application/pdf' || att.name.toLowerCase().endsWith('.pdf')) {
                try {
                    const pdfArrayBuffer = await att.blob.arrayBuffer();
                    pdfBuffers.push(pdfArrayBuffer);
                } catch (err) {
                    console.error('Failed to add PDF attachment:', att.name, err);
                    skippedFiles.push(att.name + ' (PDF could not be loaded)');
                }
            }
            // Other types: skip for now, could add more logic here
            else {
                console.warn('Skipping unsupported attachment type:', att.name, att.type);
                skippedFiles.push(att.name + ' (unsupported type: ' + att.type + ")");
            }
        }

        // Step 4: Merge all PDF buffers using pdf-lib
        let failedMerges = [];
        try {
            const mergedPdf = await PDFDocument.create();
            for (let i = 0; i < pdfBuffers.length; i++) {
                let buf = pdfBuffers[i];
                try {
                    const srcPdf = await PDFDocument.load(buf, { ignoreEncryption: true });
                    const pages = await mergedPdf.copyPages(srcPdf, srcPdf.getPageIndices());
                    pages.forEach(page => mergedPdf.addPage(page));
                } catch (err) {
                    console.error('Failed to load/merge a PDF (possibly encrypted or corrupt):', err);
                    if (i > 0) skippedFiles.push((attachmentBlobs[i-1]?.name || 'Attachment') + ' (PDF merge failed)');
                    failedMerges.push(i);
                }
            }
            // Optionally, add a summary page if any files were skipped
            if (skippedFiles.length > 0) {
                const summaryPage = mergedPdf.addPage();
                const { width, height } = summaryPage.getSize();
                const font = await mergedPdf.embedFont(StandardFonts.Helvetica);
                const text = 'The following attachments could not be merged into the PDF:\n' + skippedFiles.map(f => '- ' + f).join('\n');
                summaryPage.drawText(text, {
                    x: 50,
                    y: height - 50,
                    size: 12,
                    font: font,
                    color: undefined,
                    maxWidth: width - 100,
                    lineHeight: 16
                });
            }
            const mergedBytes = await mergedPdf.save();
            // Download merged PDF
            const blob = new Blob([mergedBytes], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'email_and_attachments.pdf';
            a.click();
            console.log('Merged PDF download triggered.');
            // Show warning in UI if any files were skipped
            if (skippedFiles.length > 0) {
                showMessage('Some attachments could not be merged into the PDF: ' + skippedFiles.join('; '));
            } else {
                hideMessage();
            }
        } catch (err) {
            console.error('Failed to merge PDFs:', err);
            if (skippedFiles.length > 0) {
                showMessage('Some attachments could not be merged into the PDF: ' + skippedFiles.join('; '));
            } else {
                showMessage('Failed to merge PDFs: ' + err.message);
            }
        }
    }

    // Wait for images in email, then process everything
    function onEmailImagesLoaded() {
        processEmailAndAttachments();
    }
    if (emailImages.length > 0) {
        emailImages.forEach(img => {
            if (img.complete) {
                loaded++;
            } else {
                img.onload = img.onerror = () => {
                    loaded++;
                    if (loaded === emailImages.length) onEmailImagesLoaded();
                };
            }
        });
        if (loaded === emailImages.length) onEmailImagesLoaded();
    } else {
        onEmailImagesLoaded();
    }
}
