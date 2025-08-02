console.log('taskpane.js loaded');

// Load pdf-lib from CDN for browser compatibility
let PDFDocument, StandardFonts, rgb;
console.log('Loading PDFLib script...');
const pdfLibScript = document.createElement('script');
pdfLibScript.src = 'https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js';
pdfLibScript.onload = () => {
    console.log('PDFLib script loaded.');
    PDFDocument = window.PDFLib.PDFDocument;
    StandardFonts = window.PDFLib.StandardFonts;
    rgb = window.PDFLib.rgb;
};
document.head.appendChild(pdfLibScript);

// Global error handler for debugging
window.addEventListener('error', function(e) {
    console.error('Global JS error:', e.message, e);
});
window.addEventListener('unhandledrejection', function(e) {
    console.error('Unhandled promise rejection:', e.reason);
});

// Helper function to strip HTML tags and clean text for pdf-lib
function stripHtml(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;
    
    // Get text content and clean it up
    let text = temp.textContent || temp.innerText || '';
    
    // Replace HTML entities and special characters that might cause encoding issues
    text = text.replace(/&nbsp;/g, ' ');
    text = text.replace(/&amp;/g, '&');
    text = text.replace(/&lt;/g, '<');
    text = text.replace(/&gt;/g, '>');
    text = text.replace(/&quot;/g, '"');
    text = text.replace(/&#39;/g, "'");
    
    // Replace various Unicode space characters with regular spaces
    text = text.replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000]/g, ' '); // Various space chars
    text = text.replace(/[\u2010-\u2015]/g, '-'); // Various dashes to regular dash
    text = text.replace(/[\u2018\u2019]/g, "'"); // Smart single quotes to regular
    text = text.replace(/[\u201C\u201D]/g, '"'); // Smart double quotes to regular
    text = text.replace(/[\u2026]/g, '...'); // Ellipsis to three dots
    text = text.replace(/[\u2013\u2014]/g, '-'); // En dash and em dash to regular dash
    
    // Remove or replace characters that WinAnsi can't encode (keep only ASCII + basic Latin)
    text = text.replace(/[^\x20-\x7E\n\r\t]/g, '?'); // Replace non-ASCII with ?
    
    // Replace multiple whitespace with single spaces
    text = text.replace(/\s+/g, ' ');
    
    // Normalize line breaks
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    
    return text.trim();
}

// Helper function to clean text specifically for pdf-lib rendering
function cleanTextForPdfLib(text) {
    if (!text) return '';
    
    // Apply all the same cleaning as stripHtml but for plain text
    let cleanText = text;
    
    // Replace various Unicode space characters with regular spaces
    cleanText = cleanText.replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
    cleanText = cleanText.replace(/[\u2010-\u2015]/g, '-'); // Various dashes
    cleanText = cleanText.replace(/[\u2018\u2019]/g, "'"); // Smart quotes
    cleanText = cleanText.replace(/[\u201C\u201D]/g, '"'); // Smart double quotes
    cleanText = cleanText.replace(/[\u2026]/g, '...'); // Ellipsis
    cleanText = cleanText.replace(/[\u2013\u2014]/g, '-'); // En/em dash
    
    // Remove characters that WinAnsi can't encode
    cleanText = cleanText.replace(/[^\x20-\x7E\n\r\t]/g, '?');
    
    // Replace multiple whitespace with single spaces
    cleanText = cleanText.replace(/\s+/g, ' ');
    
    return cleanText.trim();
}

// Helper function to split text into lines
function splitTextIntoLines(text, font, fontSize, maxWidth) {
    const lines = [];
    
    // First, split by newlines to handle paragraph breaks
    const paragraphs = text.split('\n');
    
    for (const paragraph of paragraphs) {
        if (paragraph.trim() === '') {
            lines.push(''); // Add empty line for spacing
            continue;
        }
        
        const words = paragraph.split(' ');
        let currentLine = '';
        
        for (const word of words) {
            // Clean the word of any remaining newline characters
            const cleanWord = word.replace(/[\n\r]/g, '');
            if (cleanWord === '') continue;
            
            const testLine = currentLine ? `${currentLine} ${cleanWord}` : cleanWord;
            
            try {
                const width = font.widthOfTextAtSize(testLine, fontSize);
                
                if (width < maxWidth) {
                    currentLine = testLine;
                } else {
                    if (currentLine) {
                        lines.push(currentLine);
                    }
                    currentLine = cleanWord;
                }
            } catch (err) {
                // If we can't calculate width (e.g., special characters), just add the word
                if (currentLine) {
                    lines.push(currentLine);
                }
                currentLine = cleanWord;
            }
        }
        
        if (currentLine) {
            lines.push(currentLine);
        }
    }
    
    return lines;
}

// PDF generation function definitions
async function createPdf(htmlContent, attachments, imgSources) {
    console.log('Creating PDF with html2pdf');
    const element = document.getElementById('emailToPdfDiv');
    if (!element) {
        console.error('Email to PDF container not found');
        return;
    }

    element.innerHTML = htmlContent;
    element.style.display = 'block';

    const opt = {
        margin: [0.5, 0.5],
        filename: Office.context.mailbox.item.subject + '.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { 
            scale: 2,
            useCORS: true,
            logging: true
        },
        jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
    };

    try {
        const pdf = await html2pdf().set(opt).from(element).save();
        element.style.display = 'none';
        element.innerHTML = '';
        console.log('PDF created successfully');
    } catch (err) {
        console.error('Error creating PDF:', err);
        element.style.display = 'none';
        throw err;
    }
}

async function createPdfLibTextPdf(metaHtml, htmlBody, attachments) {
    console.log('Creating PDF with pdf-lib (text-based)');
    
    try {
        // Create a new PDF document
        const pdfDoc = await PDFDocument.create();
        let currentPage = pdfDoc.addPage([595, 842]); // A4 size in points
        const { width, height } = currentPage.getSize();
        
        // Load fonts
        const helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        const courier = await pdfDoc.embedFont(StandardFonts.Courier);
        
        let y = height - 50;
        const margin = 50;
        const contentWidth = width - (margin * 2);
        
        // Function to add a new page if needed
        const addPageIfNeeded = (requiredHeight = 20) => {
            if (y < margin + requiredHeight) {
                currentPage = pdfDoc.addPage([595, 842]);
                y = height - 50;
                return true;
            }
            return false;
        };
        
        // Function to draw text with wrapping
        const drawWrappedText = (text, font, fontSize, color = rgb(0, 0, 0), lineSpacing = 1.2) => {
            // Clean the text before processing
            const cleanText = cleanTextForPdfLib(text);
            const lines = splitTextIntoLines(cleanText, font, fontSize, contentWidth);
            const lineHeight = fontSize * lineSpacing;
            
            for (const line of lines) {
                addPageIfNeeded(lineHeight);
                if (line.trim()) {
                    // Double-check each line before drawing
                    const cleanLine = cleanTextForPdfLib(line);
                    currentPage.drawText(cleanLine, {
                        x: margin,
                        y,
                        size: fontSize,
                        font,
                        color,
                    });
                }
                y -= lineHeight;
            }
            return lines.length * lineHeight;
        };
        
        // Parse and format the metadata section
        const metaDoc = new DOMParser().parseFromString(metaHtml, 'text/html');
        const metaText = metaDoc.body.textContent || metaDoc.body.innerText || '';
        
        // Header: Email metadata with better formatting
        drawWrappedText('EMAIL DETAILS', helveticaBold, 14, rgb(0.2, 0.2, 0.2));
        y -= 10; // Extra spacing after header
        
        // Parse metadata fields
        const metaLines = metaText.split('\n').filter(line => line.trim());
        for (const line of metaLines) {
            if (line.includes(':')) {
                const [label, ...valueParts] = line.split(':');
                const value = valueParts.join(':').trim();
                
                // Draw label in bold
                const labelWidth = helveticaBold.widthOfTextAtSize(label + ': ', 11);
                addPageIfNeeded(15);
                currentPage.drawText(label + ': ', {
                    x: margin,
                    y,
                    size: 11,
                    font: helveticaBold,
                    color: rgb(0.3, 0.3, 0.3),
                });
                
                // Draw value in regular font, wrapped if needed
                if (value) {
                    const remainingWidth = contentWidth - labelWidth;
                    const valueLines = splitTextIntoLines(value, helvetica, 11, remainingWidth);
                    
                    // First line on same line as label
                    if (valueLines[0]) {
                        currentPage.drawText(valueLines[0], {
                            x: margin + labelWidth,
                            y,
                            size: 11,
                            font: helvetica,
                            color: rgb(0, 0, 0),
                        });
                    }
                    y -= 15;
                    
                    // Additional lines indented
                    for (let i = 1; i < valueLines.length; i++) {
                        addPageIfNeeded(15);
                        currentPage.drawText(valueLines[i], {
                            x: margin + labelWidth,
                            y,
                            size: 11,
                            font: helvetica,
                            color: rgb(0, 0, 0),
                        });
                        y -= 15;
                    }
                } else {
                    y -= 15;
                }
            }
        }
        
        y -= 20; // Extra space before content
        
        // Add separator line
        addPageIfNeeded(10);
        currentPage.drawLine({
            start: { x: margin, y: y },
            end: { x: width - margin, y: y },
            thickness: 1,
            color: rgb(0.8, 0.8, 0.8),
        });
        y -= 20;
        
        // Parse HTML body with better formatting
        const bodyDoc = new DOMParser().parseFromString(htmlBody, 'text/html');
        
        // Process different HTML elements
        const processElement = (element, defaultFont = helvetica, defaultSize = 11, defaultColor = rgb(0, 0, 0)) => {
            const tagName = element.tagName?.toLowerCase();
            
            switch (tagName) {
                case 'h1':
                case 'h2':
                case 'h3':
                case 'h4':
                case 'h5':
                case 'h6':
                    y -= 10; // Space before heading
                    const headingSize = tagName === 'h1' ? 16 : tagName === 'h2' ? 14 : 12;
                    const headingText = cleanTextForPdfLib(element.textContent || element.innerText || '');
                    if (headingText) {
                        drawWrappedText(headingText, helveticaBold, headingSize, rgb(0.1, 0.1, 0.1));
                        y -= 5; // Space after heading
                    }
                    break;
                    
                case 'p':
                    const pText = cleanTextForPdfLib(element.textContent || element.innerText || '');
                    if (pText.trim()) {
                        drawWrappedText(pText, defaultFont, defaultSize, defaultColor);
                        y -= 10; // Paragraph spacing
                    }
                    break;
                    
                case 'strong':
                case 'b':
                    const strongText = cleanTextForPdfLib(element.textContent || element.innerText || '');
                    if (strongText) {
                        drawWrappedText(strongText, helveticaBold, defaultSize, defaultColor);
                    }
                    break;
                    
                case 'code':
                case 'pre':
                    const codeText = cleanTextForPdfLib(element.textContent || element.innerText || '');
                    if (codeText) {
                        y -= 5;
                        // Code background (light gray rectangle)
                        const codeLines = splitTextIntoLines(codeText, courier, 10, contentWidth - 20);
                        const codeHeight = codeLines.length * 12 + 10;
                        addPageIfNeeded(codeHeight);
                        
                        currentPage.drawRectangle({
                            x: margin,
                            y: y - codeHeight + 5,
                            width: contentWidth,
                            height: codeHeight,
                            color: rgb(0.95, 0.95, 0.95),
                        });
                        
                        // Draw code text
                        for (const line of codeLines) {
                            addPageIfNeeded(12);
                            if (line.trim()) {
                                const cleanCodeLine = cleanTextForPdfLib(line);
                                currentPage.drawText(cleanCodeLine, {
                                    x: margin + 10,
                                    y,
                                    size: 10,
                                    font: courier,
                                    color: rgb(0.2, 0.2, 0.2),
                                });
                            }
                            y -= 12;
                        }
                        y -= 10;
                    }
                    break;
                    
                case 'ul':
                case 'ol':
                    const listItems = element.querySelectorAll('li');
                    listItems.forEach((li, index) => {
                        const bullet = tagName === 'ul' ? '• ' : `${index + 1}. `;
                        const liText = cleanTextForPdfLib(li.textContent || li.innerText || '');
                        if (liText.trim()) {
                            addPageIfNeeded(15);
                            // Draw bullet/number
                            currentPage.drawText(bullet, {
                                x: margin + 10,
                                y,
                                size: defaultSize,
                                font: defaultFont,
                                color: defaultColor,
                            });
                            
                            // Draw list item text with proper indentation
                            const bulletWidth = defaultFont.widthOfTextAtSize(bullet, defaultSize);
                            const itemLines = splitTextIntoLines(liText, defaultFont, defaultSize, contentWidth - bulletWidth - 20);
                            
                            for (let i = 0; i < itemLines.length; i++) {
                                if (i > 0) addPageIfNeeded(15);
                                const cleanItemLine = cleanTextForPdfLib(itemLines[i]);
                                currentPage.drawText(cleanItemLine, {
                                    x: margin + 10 + bulletWidth,
                                    y,
                                    size: defaultSize,
                                    font: defaultFont,
                                    color: defaultColor,
                                });
                                if (i < itemLines.length - 1) y -= 15;
                            }
                            y -= 15;
                        }
                    });
                    y -= 10; // Space after list
                    break;
                    
                case 'blockquote':
                    const quoteText = cleanTextForPdfLib(element.textContent || element.innerText || '');
                    if (quoteText) {
                        y -= 5;
                        // Draw quote bar
                        addPageIfNeeded(20);
                        currentPage.drawRectangle({
                            x: margin,
                            y: y - 20,
                            width: 3,
                            height: 20,
                            color: rgb(0.6, 0.6, 0.6),
                        });
                        
                        // Draw quote text with indentation
                        const quoteLines = splitTextIntoLines(quoteText, helvetica, 11, contentWidth - 30);
                        for (const line of quoteLines) {
                            addPageIfNeeded(15);
                            if (line.trim()) {
                                const cleanQuoteLine = cleanTextForPdfLib(line);
                                currentPage.drawText(cleanQuoteLine, {
                                    x: margin + 20,
                                    y,
                                    size: 11,
                                    font: helvetica,
                                    color: rgb(0.4, 0.4, 0.4),
                                });
                            }
                            y -= 15;
                        }
                        y -= 10;
                    }
                    break;
                    
                default:
                    // For other elements, just extract text
                    const text = cleanTextForPdfLib(element.textContent || element.innerText || '');
                    if (text.trim() && !element.querySelector('p, h1, h2, h3, h4, h5, h6, ul, ol, blockquote')) {
                        drawWrappedText(text, defaultFont, defaultSize, defaultColor);
                        y -= 5;
                    }
                    break;
            }
        };
        
        // Process all elements in the body
        const walkElements = (element) => {
            // Process current element
            processElement(element);
            
            // Process children for elements that don't handle their own children
            const tagName = element.tagName?.toLowerCase();
            if (!['ul', 'ol', 'code', 'pre'].includes(tagName)) {
                for (const child of element.children) {
                    walkElements(child);
                }
            }
        };
        
        // Start processing from body
        if (bodyDoc.body) {
            for (const child of bodyDoc.body.children) {
                walkElements(child);
            }
        }
        
        // If no structured content found, fall back to simple text
        if (y > height - 100) {
            const simpleText = stripHtml(htmlBody);
            if (simpleText.trim()) {
                drawWrappedText('EMAIL CONTENT', helveticaBold, 12, rgb(0.2, 0.2, 0.2));
                y -= 10;
                drawWrappedText(simpleText, helvetica, 11);
            }
        }
        
        // Save the PDF
        const pdfBytes = await pdfDoc.save();
        
        // Create a download link
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = (Office.context.mailbox.item.subject || 'email') + '.pdf';
        link.click();
        
        console.log('PDF created successfully with pdf-lib (enhanced formatting)');
    } catch (err) {
        console.error('Error creating PDF with pdf-lib:', err);
        throw err;
    }
}

async function createIndividualPdfs(htmlContent, attachments, imgSources, includeImages, metaHtml) {
    console.log('Creating individual PDFs');
    
    // Create main email PDF
    await createPdf(htmlContent, null, null);
    
    // If we have images and includeImages is true, create PDFs for each image
    if (includeImages && imgSources && imgSources.length) {
        for (let i = 0; i < imgSources.length; i++) {
            const imgSrc = imgSources[i];
            try {
                const img = document.createElement('img');
                img.src = imgSrc;
                await new Promise(resolve => img.onload = resolve);
                
                const imgContainer = document.createElement('div');
                imgContainer.style.padding = '20px';
                imgContainer.appendChild(img);
                
                const imgWithMeta = document.createElement('div');
                imgWithMeta.innerHTML = metaHtml + imgContainer.outerHTML;
                
                await createPdf(imgWithMeta.outerHTML, null, null);
            } catch (err) {
                console.error(`Error creating PDF for image ${i}:`, err);
            }
        }
    }
}

// Main initialization - attach handlers after both Office.js and DOM are ready
Office.onReady(() => {
    console.log('Office.js is ready');
    
    // Make sure DOM is ready as well
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeApp);
    } else {
        // DOM already ready
        initializeApp();
    }
});

// Initialize the application and attach event handlers
function initializeApp() {
    try {
        console.log('Initializing app...');
        
        // Get button and attach click handler
        const btn = document.getElementById('convertBtn');
        if (!btn) {
            console.error('Convert button not found in DOM!');
            return;
        }
        
        console.log('convertBtn found:', btn);
        
        // Attach click handler
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
        
        console.log('Button click handler attached.');
    } catch (err) {
        console.error('Error in initializeApp:', err);
    }
}

