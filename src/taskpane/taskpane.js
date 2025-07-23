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
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage();
    // Remove unsupported Unicode characters for WinAnsi font
    const safeText = htmlBody.replace(/<[^>]+>/g, '').replace(/[^\0-\x7f]/g, ' ').slice(0, 4000);
    page.drawText(safeText); // Strips HTML tags, replaces non-ASCII chars, limits length
    console.log('Added email body to PDF.');

    // Add embedded images
    const proxyBase = 'https://email2-pd-fv2.vercel.app/api/proxy?url=';
    for (const src of imgSources) {
        try {
            console.log('Processing embedded image:', src);
            // Always use proxy for external images
            const proxiedSrc = src.startsWith('http') ? proxyBase + encodeURIComponent(src) : src;
            const response = await fetch(proxiedSrc);
            if (!response.ok) throw new Error('Failed to fetch');
            const contentType = response.headers.get('content-type');
            const imgBytes = await response.arrayBuffer();
            let img;
            if (contentType && contentType.includes('png')) {
                img = await pdfDoc.embedPng(imgBytes);
            } else if (contentType && contentType.includes('jpeg')) {
                img = await pdfDoc.embedJpg(imgBytes);
            } else if (contentType && contentType.includes('jpg')) {
                img = await pdfDoc.embedJpg(imgBytes);
            } else {
                throw new Error('Unsupported image format: ' + contentType);
            }
            const { width, height } = img.scale(0.5);
            page.drawImage(img, { x: 50, y: 400, width, height });
            console.log('Embedded image added to PDF.');
        } catch (e) {
            console.error('Image error:', src, e);
        }
    }
    // Collect all merge promises
    const mergePromises = attachments.map(att => new Promise((resolve) => {
        console.log('Processing attachment:', att.name, att.contentType || '');
        Office.context.mailbox.item.getAttachmentContentAsync(att.id, { asyncContext: att }, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const { content, contentType } = result.value;
                if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                    const byteArr = Uint8Array.from(atob(content), c => c.charCodeAt(0));
                    if (contentType === 'image/png') {
                        try {
                            console.log('Merging PNG image attachment:', att.name);
                            const img = await pdfDoc.embedPng(byteArr);
                            const attPage = pdfDoc.addPage();
                            const { width, height } = img.scale(0.5);
                            attPage.drawImage(img, { x: 50, y: 400, width, height });
                            console.log('PNG image merged.');
                        } catch (e) {
                            console.error('Attachment image error:', att.name, e);
                        }
                        resolve();
                    } else if (contentType === 'image/jpeg') {
                        try {
                            console.log('Merging JPEG image attachment:', att.name);
                            const img = await pdfDoc.embedJpg(byteArr);
                            const attPage = pdfDoc.addPage();
                            const { width, height } = img.scale(0.5);
                            attPage.drawImage(img, { x: 50, y: 400, width, height });
                            console.log('JPEG image merged.');
                        } catch (e) {
                            console.error('Attachment image error:', att.name, e);
                        }
                        resolve();
                    } else if (contentType === 'application/pdf' || att.name.endsWith('.pdf')) {
                        try {
                            console.log('Merging PDF attachment:', att.name);
                            const pdfToMerge = await PDFDocument.load(byteArr);
                            const pageIndices = pdfToMerge.getPageIndices();
                            const copiedPages = await pdfDoc.copyPages(pdfToMerge, pageIndices);
                            copiedPages.forEach(page => pdfDoc.addPage(page));
                            console.log('PDF merged.');
                        } catch (e) {
                            console.error('PDF merge error:', att.name, e);
                        }
                        resolve();
                    } else if (contentType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || att.name.endsWith('.docx')) {
                        try {
                            console.log('Converting DOCX attachment to PDF:', att.name);
                            const arrayBuffer = byteArr.buffer;
                            // Check different ways mammoth might be available
                            const mammothLib = window.mammoth || mammoth;
                            if (!mammothLib) {
                                console.error('Mammoth library not available');
                                resolve();
                                return;
                            }
                            const result = await mammothLib.convertToHtml({arrayBuffer});
                            const html = result.value;
                            const docxContainer = document.createElement('div');
                            docxContainer.innerHTML = html;
                            docxContainer.style.background = '#fff';
                            docxContainer.style.padding = '24px';
                            docxContainer.style.fontFamily = 'Arial, sans-serif';
                            docxContainer.style.width = '800px';
                            docxContainer.style.maxWidth = '100%';
                            docxContainer.style.height = 'auto';
                            docxContainer.style.overflow = 'visible';
                            document.body.appendChild(docxContainer);
                            const images = Array.from(docxContainer.querySelectorAll('img'));
                            let loaded = 0;
                            function addDocxToPdf() {
                                html2pdf().set({
                                    margin: 10,
                                    filename: 'email.pdf',
                                    image: { type: 'jpeg', quality: 0.98 },
                                    html2canvas: { scale: 2, useCORS: true },
                                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                                    pagebreak: { mode: ['avoid-all', 'css'] }
                                }).from(docxContainer).toPdf().get('pdf').then(pdf => {
                                    const docxPdfBytes = pdf.output('arraybuffer');
                                    PDFDocument.load(docxPdfBytes).then(docxPdf => {
                                        docxPdf.getPageIndices().forEach(idx => {
                                            docxPdf.copyPages(docxPdf, [idx]).then(copiedPages => {
                                                copiedPages.forEach(p => pdfDoc.addPage(p));
                                            });
                                        });
                                        document.body.removeChild(docxContainer);
                                        console.log('DOCX converted and merged.');
                                        resolve();
                                    });
                                });
                            }
                            if (images.length > 0) {
                                images.forEach(img => {
                                    if (img.complete) {
                                        loaded++;
                                    } else {
                                        img.onload = img.onerror = () => {
                                            loaded++;
                                            if (loaded === images.length) addDocxToPdf();
                                        };
                                    }
                                });
                                if (loaded === images.length) addDocxToPdf();
                            } else {
                                addDocxToPdf();
                            }
                        } catch (e) {
                            console.error('DOCX conversion error:', att.name, e);
                            resolve();
                        }
                    } else if (contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || att.name.endsWith('.xlsx')) {
                        try {
                            console.log('Converting XLSX attachment to PDF:', att.name);
                            const arrayBuffer = byteArr.buffer;
                            // Check different ways SheetJS might be available
                            const xlsxLib = window.XLSX || XLSX;
                            if (!xlsxLib) {
                                console.error('SheetJS library not available');
                                resolve();
                                return;
                            }
                            const workbook = xlsxLib.read(arrayBuffer, {type: 'array'});
                            let html = '';
                            workbook.SheetNames.forEach(sheetName => {
                                const worksheet = workbook.Sheets[sheetName];
                                html += xlsxLib.utils.sheet_to_html(worksheet, {id: sheetName});
                            });
                            const xlsxContainer = document.createElement('div');
                            xlsxContainer.innerHTML = html;
                            xlsxContainer.style.background = '#fff';
                            xlsxContainer.style.padding = '24px';
                            xlsxContainer.style.fontFamily = 'Arial, sans-serif';
                            xlsxContainer.style.width = '800px';
                            xlsxContainer.style.maxWidth = '100%';
                            xlsxContainer.style.height = 'auto';
                            xlsxContainer.style.overflow = 'visible';
                            document.body.appendChild(xlsxContainer);
                            const images = Array.from(xlsxContainer.querySelectorAll('img'));
                            let loaded = 0;
                            function addXlsxToPdf() {
                                html2pdf().set({
                                    margin: 10,
                                    filename: 'email.pdf',
                                    image: { type: 'jpeg', quality: 0.98 },
                                    html2canvas: { scale: 2, useCORS: true },
                                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                                    pagebreak: { mode: ['avoid-all', 'css'] }
                                }).from(xlsxContainer).toPdf().get('pdf').then(pdf => {
                                    const xlsxPdfBytes = pdf.output('arraybuffer');
                                    PDFDocument.load(xlsxPdfBytes).then(xlsxPdf => {
                                        xlsxPdf.getPageIndices().forEach(idx => {
                                            xlsxPdf.copyPages(xlsxPdf, [idx]).then(copiedPages => {
                                                copiedPages.forEach(p => pdfDoc.addPage(p));
                                            });
                                        });
                                        document.body.removeChild(xlsxContainer);
                                        console.log('XLSX converted and merged.');
                                        resolve();
                                    });
                                });
                            }
                            if (images.length > 0) {
                                images.forEach(img => {
                                    if (img.complete) {
                                        loaded++;
                                    } else {
                                        img.onload = img.onerror = () => {
                                            loaded++;
                                            if (loaded === images.length) addXlsxToPdf();
                                        };
                                    }
                                });
                                if (loaded === images.length) addXlsxToPdf();
                            } else {
                                addXlsxToPdf();
                            }
                        } catch (e) {
                            console.error('XLSX conversion error:', att.name, e);
                            resolve();
                        }
                    } else {
                        console.warn('Unsupported attachment type:', att.name, contentType);
                        resolve();
                    }
                }
            } else {
                console.error('Failed to get attachment content:', result.error);
                resolve();
            }
        });
    }));

    // Wait for all merges to finish
    console.log('Waiting for all attachment merges to finish...');
    await Promise.all(mergePromises);
    console.log('All attachments merged. Saving PDF...');

    // Download PDF
    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'email.pdf';
    a.click();
    console.log('PDF download triggered.');
}
