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





// Attach handler after both Office.js and DOM are ready
// Global error handler for debugging
window.addEventListener('error', function(e) {
    console.error('Global JS error:', e.message, e);
});
window.addEventListener('unhandledrejection', function(e) {
    console.error('Unhandled promise rejection:', e.reason);
});

Office.onReady(() => {
    console.log('Office.js is ready');
    document.addEventListener('DOMContentLoaded', function() {
        console.log('Attempting to attach handler to #convertBtn');
        const btn = document.getElementById('convertBtn');
        if (!btn) {
            console.error('Convert button not found in DOM!');
            return;
        }
        console.log('convertBtn found:', btn);
        btn.onclick = async function() {
            console.log('Convert button clicked.');
            this.disabled = true;
            this.textContent = 'Converting...';
            try {
                // Helper to format a person as email (Name)
                const formatPerson = (p) => {
                    if (!p) return '';
                    if (p.displayName && p.emailAddress) {
                        if (p.displayName === p.emailAddress) return p.emailAddress;
                        return `${p.emailAddress} (${p.displayName})`;
                    }
                    return p.emailAddress || p.displayName || '';
                };
                // Helper to format a list
                const formatList = (arr) => (arr && arr.length ? arr.map(formatPerson).join('; ') : '');
                // Gather metadata
                const item = Office.context.mailbox.item;
                const meta = {
                    from: item.from ? formatPerson(item.from) : '',
                    to: formatList(item.to),
                    cc: formatList(item.cc),
                    subject: item.subject || '',
                    date: (item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : '')
                };
                const metaHtml = `<div style="background:#f5f5f5;padding:12px 18px 12px 18px;margin-bottom:18px;border-radius:6px;font-size:1em;line-height:1.5;">
                    <div><b>From:</b> ${meta.from}</div>
                    <div><b>To:</b> ${meta.to}</div>
                    ${meta.cc ? `<div><b>CC:</b> ${meta.cc}</div>` : ''}
                    <div><b>Subject:</b> ${meta.subject}</div>
                    <div><b>Sent:</b> ${meta.date}</div>
                </div>`;

                // Get selected mode and engine
                const modeSel = document.getElementById('pdfMode');
                const mode = modeSel ? modeSel.value : 'full';
                const engineSel = document.getElementById('pdfEngine');
                const engine = engineSel ? engineSel.value : 'html2pdf';
                console.log('PDF Generation Mode:', mode, 'Engine:', engine);

                // Get email body
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

                        // Wait for pdf-lib to load if needed
                        if (engine === 'pdf-lib' && !PDFDocument) {
                            console.error('PDF library not loaded yet.');
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

                        try {
                            if (engine === 'pdf-lib') {
                                await createPdfLibTextPdf(metaHtml, htmlBody, attachments);
                            } else {
                                if (mode === 'full') {
                                    await createPdf(metaHtml + htmlBody, attachments, imgSources);
                                } else if (mode === 'individual') {
                                    await createIndividualPdfs(metaHtml + htmlBody, attachments, imgSources, false, metaHtml);
                                } else if (mode === 'images') {
                                    await createIndividualPdfs(metaHtml + htmlBody, attachments, imgSources, true, metaHtml);
                                } else {
                                    await createPdf(metaHtml + htmlBody, attachments, imgSources);
                                }
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
    });
});
// PDF generation logic moved into createPdf function
async function createPdf(htmlBody, attachments, imgSources) {
    let msgDiv = document.getElementById('pdf2email-message');

}

