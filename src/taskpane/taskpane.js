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

// Check for html-to-pdfmake and pdfmake availability
let htmlToPdfmake, pdfMake;
window.addEventListener('load', () => {
    htmlToPdfmake = window.htmlToPdfmake;
    pdfMake = window.pdfMake;
    if (htmlToPdfmake && pdfMake) {
        console.log('html-to-pdfmake and pdfMake libraries loaded successfully');
    } else {
        console.warn('html-to-pdfmake or pdfMake not available:', {htmlToPdfmake, pdfMake});
    }
});

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
    text = text.replace(/&apos;/g, "'");
    
    // Replace various Unicode space characters with regular spaces
    text = text.replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000]/g, ' '); // Various space chars
    text = text.replace(/[\u2010-\u2015]/g, '-'); // Various dashes to regular dash
    text = text.replace(/[\u2018\u2019]/g, "'"); // Smart single quotes to regular
    text = text.replace(/[\u201C\u201D]/g, '"'); // Smart double quotes to regular
    text = text.replace(/[\u2026]/g, '...'); // Ellipsis to three dots
    text = text.replace(/[\u2013\u2014]/g, '-'); // En dash and em dash to regular dash
    
    // Remove emojis and problematic symbols, but be more conservative with other characters
    text = text.replace(/[\u1F000-\u1FFFF]/g, ''); // Remove emojis
    text = text.replace(/[\u2600-\u27BF]/g, ''); // Remove misc symbols
    text = text.replace(/[\uFE00-\uFE0F]/g, ''); // Remove variation selectors
    
    // Only remove control characters that are truly problematic
    text = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
    
    // Replace multiple whitespace with single spaces
    text = text.replace(/\s+/g, ' ');
    
    // Normalize line breaks
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    
    return text.trim();
}

// Helper function to clean text specifically for pdf-lib rendering
function cleanTextForPdfLib(text) {
    if (!text) return '';
    
    let cleanText = text;
    
    // Fix common URL encoding issues that might appear in email content
    try {
        // Only decode if it looks like URL encoding (contains %)
        if (cleanText.includes('%')) {
            // First try to decode URL encoding
            cleanText = decodeURIComponent(cleanText);
        }
    } catch (e) {
        // If decoding fails, continue with original text
    }
    
    // Decode HTML entities properly
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = cleanText;
    cleanText = tempDiv.textContent || tempDiv.innerText || cleanText;
    
    // Fix common character encoding issues that appear in email
    cleanText = cleanText.replace(/&nbsp;/g, ' ');
    cleanText = cleanText.replace(/&amp;/g, '&');
    cleanText = cleanText.replace(/&lt;/g, '<');
    cleanText = cleanText.replace(/&gt;/g, '>');
    cleanText = cleanText.replace(/&quot;/g, '"');
    cleanText = cleanText.replace(/&#39;/g, "'");
    cleanText = cleanText.replace(/&apos;/g, "'");
    
    // Replace various Unicode space characters with regular spaces
    cleanText = cleanText.replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
    cleanText = cleanText.replace(/[\u2010-\u2015]/g, '-'); // Various dashes
    cleanText = cleanText.replace(/[\u2018\u2019]/g, "'"); // Smart quotes
    cleanText = cleanText.replace(/[\u201C\u201D]/g, '"'); // Smart double quotes
    cleanText = cleanText.replace(/[\u2026]/g, '...'); // Ellipsis
    cleanText = cleanText.replace(/[\u2013\u2014]/g, '-'); // En/em dash
    
    // Remove only emojis and symbols, but keep extended Latin characters
    cleanText = cleanText.replace(/[\u1F000-\u1FFFF]/g, ''); // Remove emojis
    cleanText = cleanText.replace(/[\u2600-\u27BF]/g, ''); // Remove misc symbols
    cleanText = cleanText.replace(/[\uFE00-\uFE0F]/g, ''); // Remove variation selectors
    
    // More conservative character filtering - keep more text intact
    // Only remove really problematic characters, keep basic Latin and extended
    cleanText = cleanText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, ''); // Remove control chars
    
    // Replace sequences of weird characters that look like encoding artifacts
    cleanText = cleanText.replace(/[%]{2,}/g, ' '); // Multiple % signs
    cleanText = cleanText.replace(/[.]{5,}/g, '...'); // Too many dots
    cleanText = cleanText.replace(/[-]{5,}/g, '---'); // Too many dashes
    cleanText = cleanText.replace(/[,]{2,}/g, ','); // Multiple commas
    
    // Replace multiple whitespace with single spaces, but preserve line breaks
    cleanText = cleanText.replace(/[ \t]+/g, ' ');
    cleanText = cleanText.replace(/\n\s*\n/g, '\n\n'); // Preserve paragraph breaks
    
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
        
        let y = height - 30; // Start higher on page
        const margin = 40;
        const contentWidth = width - (margin * 2);
        
        // Color definitions for better visual appeal
        const colors = {
            darkBlue: rgb(0.2, 0.3, 0.6),
            lightGray: rgb(0.95, 0.95, 0.95),
            mediumGray: rgb(0.7, 0.7, 0.7),
            darkGray: rgb(0.3, 0.3, 0.3),
            black: rgb(0, 0, 0),
            white: rgb(1, 1, 1),
            lightBlue: rgb(0.9, 0.95, 1),
            linkBlue: rgb(0.2, 0.4, 0.8)
        };
        
        // Function to add a new page if needed
        const addPageIfNeeded = (requiredHeight = 20) => {
            if (y < margin + requiredHeight) {
                currentPage = pdfDoc.addPage([595, 842]);
                y = height - 30;
                return true;
            }
            return false;
        };
        
        // Function to draw a styled header box
        const drawHeaderBox = (title, content, bgColor = colors.lightGray) => {
            const titleHeight = 20;
            const contentHeight = content.length * 16 + 20;
            const totalHeight = titleHeight + contentHeight;
            
            addPageIfNeeded(totalHeight + 10);
            
            // Draw background box
            currentPage.drawRectangle({
                x: margin,
                y: y - totalHeight,
                width: contentWidth,
                height: totalHeight,
                color: bgColor,
                borderColor: colors.mediumGray,
                borderWidth: 1,
            });
            
            // Draw title
            currentPage.drawText(title, {
                x: margin + 15,
                y: y - 18,
                size: 12,
                font: helveticaBold,
                color: colors.darkBlue,
            });
            
            y -= titleHeight + 5;
            
            // Draw content lines
            for (const line of content) {
                currentPage.drawText(line, {
                    x: margin + 15,
                    y: y - 12,
                    size: 10,
                    font: helvetica,
                    color: colors.black,
                });
                y -= 16;
            }
            
            y -= 15; // Extra spacing after box
        };
        
        // Function to draw text with wrapping and better styling
        const drawWrappedText = (text, font, fontSize, color = colors.black, lineSpacing = 1.3, indent = 0) => {
            const cleanText = cleanTextForPdfLib(text);
            const lines = splitTextIntoLines(cleanText, font, fontSize, contentWidth - indent);
            const lineHeight = fontSize * lineSpacing;
            
            for (const line of lines) {
                addPageIfNeeded(lineHeight);
                if (line.trim()) {
                    const cleanLine = cleanTextForPdfLib(line);
                    currentPage.drawText(cleanLine, {
                        x: margin + indent,
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
        
        // Function to detect and format links
        const drawTextWithLinks = (text, font, fontSize, defaultColor = colors.black) => {
            const cleanText = cleanTextForPdfLib(text);
            // Simple URL detection
            const urlRegex = /(https?:\/\/[^\s]+)/g;
            const parts = cleanText.split(urlRegex);
            
            let currentX = margin;
            for (let i = 0; i < parts.length; i++) {
                const part = parts[i];
                if (!part) continue;
                
                const isUrl = urlRegex.test(part);
                const color = isUrl ? colors.linkBlue : defaultColor;
                const partFont = isUrl ? helvetica : font;
                
                // Check if we need to wrap to next line
                const textWidth = partFont.widthOfTextAtSize(part, fontSize);
                if (currentX + textWidth > width - margin) {
                    y -= fontSize * 1.3;
                    currentX = margin;
                    addPageIfNeeded(fontSize * 1.3);
                }
                
                currentPage.drawText(part, {
                    x: currentX,
                    y,
                    size: fontSize,
                    font: partFont,
                    color,
                });
                
                currentX += textWidth;
            }
            y -= fontSize * 1.5; // Line spacing after text with links
        };
        
        // Parse and format the metadata section with enhanced styling
        const metaDoc = new DOMParser().parseFromString(metaHtml, 'text/html');
        const metaText = metaDoc.body.textContent || metaDoc.body.innerText || '';
        
        // Create a beautiful email header
        const metaLines = metaText.split('\n').filter(line => line.trim());
        const emailDetails = [];
        
        for (const line of metaLines) {
            if (line.includes(':')) {
                const [label, ...valueParts] = line.split(':');
                const value = valueParts.join(':').trim();
                if (value) {
                    emailDetails.push(`${label.trim()}: ${value}`);
                }
            }
        }
        
        // Draw EMAIL DETAILS header box
        if (emailDetails.length > 0) {
            drawHeaderBox('EMAIL DETAILS', emailDetails, colors.lightBlue);
        }
        
        // Add separator with some visual flair
        y -= 10;
        addPageIfNeeded(20);
        currentPage.drawLine({
            start: { x: margin, y: y },
            end: { x: width - margin, y: y },
            thickness: 2,
            color: colors.darkBlue,
        });
        y -= 25;
        
        // Parse HTML body with enhanced formatting
        const bodyDoc = new DOMParser().parseFromString(htmlBody, 'text/html');
        
        // Enhanced element processing with better visual hierarchy
        // Helper function to safely extract text from an element
        const extractTextContent = (element) => {
            if (!element) return '';
            
            // Try different methods to get clean text
            let text = '';
            
            // First try innerText which handles line breaks better
            if (element.innerText) {
                text = element.innerText;
            } else if (element.textContent) {
                text = element.textContent;
            } else {
                // Fallback: manually extract text from all text nodes
                const walker = document.createTreeWalker(
                    element,
                    NodeFilter.SHOW_TEXT,
                    null,
                    false
                );
                
                const textParts = [];
                let node;
                while (node = walker.nextNode()) {
                    if (node.nodeValue && node.nodeValue.trim()) {
                        textParts.push(node.nodeValue);
                    }
                }
                text = textParts.join(' ');
            }
            
            return cleanTextForPdfLib(text);
        };

        const processElement = (element, defaultFont = helvetica, defaultSize = 11, defaultColor = colors.black, depth = 0) => {
            const tagName = element.tagName?.toLowerCase();
            
            switch (tagName) {
                case 'h1':
                case 'h2':
                case 'h3':
                case 'h4':
                case 'h5':
                case 'h6':
                    y -= 15; // More space before heading
                    const headingSize = tagName === 'h1' ? 18 : tagName === 'h2' ? 16 : tagName === 'h3' ? 14 : 12;
                    const headingText = extractTextContent(element);
                    if (headingText) {
                        // Add background for major headings
                        if (tagName === 'h1' || tagName === 'h2') {
                            const textWidth = helveticaBold.widthOfTextAtSize(headingText, headingSize);
                            const textHeight = headingSize + 8;
                            addPageIfNeeded(textHeight + 10);
                            
                            currentPage.drawRectangle({
                                x: margin - 5,
                                y: y - textHeight + 5,
                                width: Math.min(textWidth + 20, contentWidth + 10),
                                height: textHeight,
                                color: colors.lightGray,
                            });
                        }
                        
                        drawWrappedText(headingText, helveticaBold, headingSize, colors.darkBlue);
                        y -= 10; // Space after heading
                    }
                    break;
                    
                case 'p':
                    const pText = extractTextContent(element);
                    if (pText.trim()) {
                        // Check if paragraph contains links
                        if (pText.includes('http')) {
                            addPageIfNeeded(defaultSize * 1.5);
                            drawTextWithLinks(pText, defaultFont, defaultSize, defaultColor);
                        } else {
                            drawWrappedText(pText, defaultFont, defaultSize, defaultColor);
                        }
                        y -= 12; // Paragraph spacing
                    }
                    break;
                    
                case 'div':
                    // Special handling for Teams-style cards or containers
                    const divText = extractTextContent(element);
                    if (divText.trim()) {
                        // Check if this looks like a Teams conversation card
                        if (divText.includes('Join conversation') || divText.includes('teams.microsoft.com')) {
                            y -= 10;
                            const cardHeight = 60;
                            addPageIfNeeded(cardHeight + 10);
                            
                            // Draw Teams-style card
                            currentPage.drawRectangle({
                                x: margin,
                                y: y - cardHeight,
                                width: contentWidth,
                                height: cardHeight,
                                color: colors.lightBlue,
                                borderColor: colors.darkBlue,
                                borderWidth: 2,
                            });
                            
                            // Add Teams icon representation (using text instead of emoji)
                            currentPage.drawText('[TEAMS]', {
                                x: margin + 15,
                                y: y - 25,
                                size: 12,
                                font: helveticaBold,
                                color: colors.darkBlue,
                            });
                            
                            currentPage.drawText('Join conversation', {
                                x: margin + 75,
                                y: y - 20,
                                size: 12,
                                font: helveticaBold,
                                color: colors.darkBlue,
                            });
                            
                            currentPage.drawText('teams.microsoft.com', {
                                x: margin + 75,
                                y: y - 35,
                                size: 10,
                                font: helvetica,
                                color: colors.linkBlue,
                            });
                            
                            y -= cardHeight + 15;
                        } else {
                            drawWrappedText(divText, defaultFont, defaultSize, defaultColor);
                            y -= 8;
                        }
                    }
                    break;
                    
                case 'strong':
                case 'b':
                    const strongText = extractTextContent(element);
                    if (strongText) {
                        drawWrappedText(strongText, helveticaBold, defaultSize, defaultColor);
                    }
                    break;
                    
                case 'a':
                    const linkText = extractTextContent(element);
                    const href = element.getAttribute('href') || '';
                    if (linkText) {
                        addPageIfNeeded(defaultSize * 1.5);
                        currentPage.drawText(linkText, {
                            x: margin,
                            y,
                            size: defaultSize,
                            font: helvetica,
                            color: colors.linkBlue,
                        });
                        y -= defaultSize * 1.3;
                        
                        // Add URL if different from text
                        if (href && href !== linkText && !linkText.includes(href)) {
                            const shortUrl = href.length > 60 ? href.substring(0, 60) + '...' : href;
                            currentPage.drawText(shortUrl, {
                                x: margin + 10,
                                y,
                                size: 9,
                                font: helvetica,
                                color: colors.mediumGray,
                            });
                            y -= 12;
                        }
                    }
                    break;
                    
                case 'ul':
                case 'ol':
                    y -= 5;
                    const listItems = element.querySelectorAll('li');
                    listItems.forEach((li, index) => {
                        const bullet = tagName === 'ul' ? '• ' : `${index + 1}. `;
                        const liText = extractTextContent(li);
                        if (liText.trim()) {
                            addPageIfNeeded(16);
                            
                            // Draw bullet with background
                            currentPage.drawRectangle({
                                x: margin + 10,
                                y: y - 12,
                                width: 15,
                                height: 12,
                                color: colors.lightGray,
                            });
                            
                            currentPage.drawText(bullet, {
                                x: margin + 12,
                                y: y - 10,
                                size: defaultSize,
                                font: helveticaBold,
                                color: colors.darkBlue,
                            });
                            
                            // Draw list item text with proper indentation
                            const bulletWidth = 25;
                            const itemLines = splitTextIntoLines(liText, defaultFont, defaultSize, contentWidth - bulletWidth - 20);
                            
                            for (let i = 0; i < itemLines.length; i++) {
                                if (i > 0) addPageIfNeeded(16);
                                const cleanItemLine = cleanTextForPdfLib(itemLines[i]);
                                currentPage.drawText(cleanItemLine, {
                                    x: margin + 10 + bulletWidth,
                                    y: y - 10,
                                    size: defaultSize,
                                    font: defaultFont,
                                    color: defaultColor,
                                });
                                if (i < itemLines.length - 1) y -= 16;
                            }
                            y -= 18;
                        }
                    });
                    y -= 10; // Space after list
                    break;
                    
                case 'blockquote':
                    const quoteText = extractTextContent(element);
                    if (quoteText) {
                        y -= 10;
                        const quoteLines = splitTextIntoLines(quoteText, helvetica, 11, contentWidth - 40);
                        const quoteHeight = quoteLines.length * 15 + 10;
                        
                        addPageIfNeeded(quoteHeight + 10);
                        
                        // Draw quote background
                        currentPage.drawRectangle({
                            x: margin + 15,
                            y: y - quoteHeight,
                            width: contentWidth - 20,
                            height: quoteHeight,
                            color: colors.lightGray,
                        });
                        
                        // Draw quote bar
                        currentPage.drawRectangle({
                            x: margin + 10,
                            y: y - quoteHeight,
                            width: 4,
                            height: quoteHeight,
                            color: colors.darkBlue,
                        });
                        
                        // Draw quote text
                        for (const line of quoteLines) {
                            if (line.trim()) {
                                const cleanQuoteLine = cleanTextForPdfLib(line);
                                currentPage.drawText(cleanQuoteLine, {
                                    x: margin + 25,
                                    y: y - 12,
                                    size: 11,
                                    font: helvetica,
                                    color: colors.darkGray,
                                });
                            }
                            y -= 15;
                        }
                        y -= 15;
                    }
                    break;
                    
                default:
                    // For other elements, extract text with better formatting
                    const text = extractTextContent(element);
                    if (text.trim() && !element.querySelector('p, h1, h2, h3, h4, h5, h6, ul, ol, blockquote, div')) {
                        if (text.includes('http')) {
                            addPageIfNeeded(defaultSize * 1.5);
                            drawTextWithLinks(text, defaultFont, defaultSize, defaultColor);
                        } else {
                            drawWrappedText(text, defaultFont, defaultSize, defaultColor);
                        }
                        y -= 8;
                    }
                    break;
            }
        };
        
        // Process all elements in the body
        const walkElements = (element, depth = 0) => {
            processElement(element, helvetica, 11, colors.black, depth);
            
            // Process children for elements that don't handle their own children
            const tagName = element.tagName?.toLowerCase();
            if (!['ul', 'ol', 'code', 'pre'].includes(tagName)) {
                for (const child of element.children) {
                    walkElements(child, depth + 1);
                }
            }
        };
        
        // Add content header
        drawHeaderBox('MESSAGE CONTENT', [], colors.lightGray);
        
        // Start processing from body
        if (bodyDoc.body && bodyDoc.body.children.length > 0) {
            for (const child of bodyDoc.body.children) {
                walkElements(child);
            }
        } else {
            // Fallback to simple text if no structured content
            const simpleText = stripHtml(htmlBody);
            if (simpleText.trim()) {
                drawWrappedText(simpleText, helvetica, 11);
            }
        }
        
        // Add footer with generation info
        y = 30; // Bottom of page
        currentPage.drawLine({
            start: { x: margin, y: y + 10 },
            end: { x: width - margin, y: y + 10 },
            thickness: 1,
            color: colors.mediumGray,
        });
        
        currentPage.drawText(`Generated by Email2PDF • ${new Date().toLocaleDateString()}`, {
            x: margin,
            y: y,
            size: 8,
            font: helvetica,
            color: colors.mediumGray,
        });
        
        // Save the PDF
        const pdfBytes = await pdfDoc.save();
        
        // Create a download link
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = (Office.context.mailbox.item.subject || 'email') + '.pdf';
        link.click();
        
        console.log('PDF created successfully with pdf-lib (enhanced visual formatting)');
    } catch (err) {
        console.error('Error creating PDF with pdf-lib:', err);
        throw err;
    }
}

// Helper function to convert external images to base64 data URLs
async function convertExternalImagesToDataUrls(html) {
    console.log('Converting external images to data URLs...');
    
    // Function to convert a single image URL to base64 data URL with better CORS handling
    function getBase64ImageFromURL(url) {
        return new Promise((resolve, reject) => {
            console.log(`Attempting to load image: ${url}`);
            
            // Try multiple approaches for CORS issues
            const tryLoadImage = (imageUrl, corsMode) => {
                return new Promise((res, rej) => {
                    const img = new Image();
                    
                    if (corsMode) {
                        img.setAttribute("crossOrigin", "anonymous");
                    }
                    
                    img.onload = () => {
                        try {
                            console.log(`Image loaded successfully: ${imageUrl}, dimensions: ${img.width}x${img.height}`);
                            
                            // Validate image dimensions
                            if (img.width === 0 || img.height === 0) {
                                console.warn(`Invalid image dimensions: ${img.width}x${img.height}`);
                                rej(new Error('Invalid image dimensions'));
                                return;
                            }
                            
                            const canvas = document.createElement("canvas");
                            canvas.width = img.width;
                            canvas.height = img.height;
                            
                            const ctx = canvas.getContext("2d");
                            
                            // Clear the canvas and set white background
                            ctx.fillStyle = 'white';
                            ctx.fillRect(0, 0, canvas.width, canvas.height);
                            
                            ctx.drawImage(img, 0, 0);
                            
                            // Try to get data URL with different formats
                            let dataURL;
                            try {
                                dataURL = canvas.toDataURL("image/png");
                                console.log(`Successfully converted to PNG data URL, size: ${dataURL.length} chars`);
                                
                                // Validate the data URL format
                                if (!dataURL.startsWith('data:image/')) {
                                    throw new Error('Invalid data URL format');
                                }
                                
                            } catch (e) {
                                try {
                                    dataURL = canvas.toDataURL("image/jpeg", 0.8);
                                    console.log(`Successfully converted to JPEG data URL, size: ${dataURL.length} chars`);
                                    
                                    // Validate the data URL format
                                    if (!dataURL.startsWith('data:image/')) {
                                        throw new Error('Invalid data URL format');
                                    }
                                    
                                } catch (e2) {
                                    console.warn('Failed to convert canvas to data URL:', e2);
                                    rej(e2);
                                    return;
                                }
                            }
                            
                            res(dataURL);
                        } catch (error) {
                            console.warn('Failed to convert image to canvas:', imageUrl, error);
                            rej(error);
                        }
                    };
                    
                    img.onerror = (error) => {
                        console.warn(`Failed to load image (corsMode: ${corsMode}):`, imageUrl, error);
                        rej(error);
                    };
                    
                    img.src = imageUrl;
                });
            };
            
            // Set a timeout to avoid hanging on slow/unresponsive images
            const timeoutPromise = new Promise((_, rej) => {
                setTimeout(() => {
                    rej(new Error('Image load timeout'));
                }, 15000); // 15 second timeout
            });
            
            // Try with CORS first, then without CORS, then with proxy if that fails
            Promise.race([
                tryLoadImage(url, true).catch(() => {
                    console.log(`CORS failed for ${url}, trying without CORS...`);
                    return tryLoadImage(url, false);
                }).catch(() => {
                    console.log(`Direct loading failed for ${url}, trying with proxy...`);
                    // Try using the image proxy for CORS issues
                    const proxyUrl = `https://email2pdfv2-image-proxy.vercel.app/api/proxy?url=${encodeURIComponent(url)}`;
                    return tryLoadImage(proxyUrl, false);
                }),
                timeoutPromise
            ]).then(resolve).catch(reject);
        });
    }
    
    // Find all image tags with external URLs
    const imageRegex = /<img[^>]*src\s*=\s*["']([^"']*)[^>]*>/gi;
    const imagesToConvert = [];
    let match;
    
    // Collect all external image URLs
    while ((match = imageRegex.exec(html)) !== null) {
        const [fullMatch, src] = match;
        if (src.startsWith('http://') || src.startsWith('https://')) {
            imagesToConvert.push({ fullMatch, src, index: match.index });
        }
    }
    
    console.log(`Found ${imagesToConvert.length} external images to convert`);
    
    // Convert each image to data URL
    let updatedHtml = html;
    let offset = 0; // Track offset changes due to string replacements
    
    for (const imageInfo of imagesToConvert) {
        try {
            console.log(`Converting image: ${imageInfo.src}`);
            const dataUrl = await getBase64ImageFromURL(imageInfo.src);
            
            // Replace the original src with the data URL
            const newImageTag = imageInfo.fullMatch.replace(imageInfo.src, dataUrl);
            
            // Calculate the actual position with offset
            const actualIndex = imageInfo.index + offset;
            const beforeReplacement = updatedHtml.substring(0, actualIndex);
            const afterReplacement = updatedHtml.substring(actualIndex + imageInfo.fullMatch.length);
            
            updatedHtml = beforeReplacement + newImageTag + afterReplacement;
            
            // Update offset for next replacement
            offset += (newImageTag.length - imageInfo.fullMatch.length);
            
            console.log(`Successfully converted image: ${imageInfo.src}`);
        } catch (error) {
            console.warn(`Failed to convert image ${imageInfo.src}, removing it:`, error);
            
            // If conversion fails, remove the image tag
            const actualIndex = imageInfo.index + offset;
            const beforeReplacement = updatedHtml.substring(0, actualIndex);
            const afterReplacement = updatedHtml.substring(actualIndex + imageInfo.fullMatch.length);
            const replacement = '<!-- Image conversion failed, removed -->';
            
            updatedHtml = beforeReplacement + replacement + afterReplacement;
            
            // Update offset for next replacement
            offset += (replacement.length - imageInfo.fullMatch.length);
        }
    }
    
    console.log('Image conversion completed');
    return updatedHtml;
}

async function createHtmlToPdfmakePdf(metaHtml, htmlBody, attachments) {
    console.log('Creating PDF with html-to-pdfmake');
    
    try {
        // Check if libraries are available
        if (!htmlToPdfmake || !pdfMake) {
            throw new Error('html-to-pdfmake or pdfMake libraries not loaded');
        }
        
        // Configure pdfMake fonts to handle system fonts and provide fallbacks
        pdfMake.fonts = {
            // Default Roboto font (always available)
            Roboto: {
                normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf',
                bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf',
                italics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Italic.ttf',
                bolditalics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-MediumItalic.ttf'
            },
            // Map common system fonts to Roboto
            Aptos: {
                normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf',
                bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf',
                italics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Italic.ttf',
                bolditalics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-MediumItalic.ttf'
            },
            Calibri: {
                normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf',
                bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf',
                italics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Italic.ttf',
                bolditalics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-MediumItalic.ttf'
            },
            Arial: {
                normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf',
                bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf',
                italics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Italic.ttf',
                bolditalics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-MediumItalic.ttf'
            },
            'Times New Roman': {
                normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf',
                bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf',
                italics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Italic.ttf',
                bolditalics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-MediumItalic.ttf'
            },
            Helvetica: {
                normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf',
                bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf',
                italics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Italic.ttf',
                bolditalics: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-MediumItalic.ttf'
            }
        };
        
        // Create the email metadata section
        const metaDoc = new DOMParser().parseFromString(metaHtml, 'text/html');
        const metaText = metaDoc.body.textContent || metaDoc.body.innerText || '';
        const metaLines = metaText.split('\n').filter(line => line.trim());
        
        // Build email details for header
        const emailDetails = [];
        for (const line of metaLines) {
            if (line.includes(':')) {
                const [label, ...valueParts] = line.split(':');
                const value = valueParts.join(':').trim();
                if (value) {
                    emailDetails.push(`${label.trim()}: ${value}`);
                }
            }
        }
        
        // Create header HTML for email details with cleaner styling
        let headerHtml = '<div style="margin-bottom: 20px; padding: 12px; background-color: #f8f9fa; border-left: 4px solid #007acc; font-family: Roboto, sans-serif;">';
        headerHtml += '<h2 style="margin: 0 0 8px 0; color: #007acc; font-size: 14px; font-weight: bold; font-family: Roboto, sans-serif;">Email Details</h2>';
        emailDetails.forEach(detail => {
            headerHtml += `<p style="margin: 2px 0; font-size: 10px; font-family: Roboto, sans-serif; color: #333;">${detail}</p>`;
        });
        headerHtml += '</div>';
        
        // Add content header with cleaner styling
        const contentHtml = '<div style="margin: 15px 0 10px 0; padding-bottom: 5px; border-bottom: 1px solid #ddd;"><h2 style="margin: 0; color: #007acc; font-size: 14px; font-weight: bold; font-family: Roboto, sans-serif;">Message Content</h2></div>';
        
        // Combine all HTML and ensure font-family is specified
        let fullHtml = headerHtml + contentHtml + htmlBody;
        
        // Clean HTML by removing problematic elements that html-to-pdfmake can't handle
        console.log('Cleaning HTML before processing...');
        
        // Remove all <style> tags and their contents (html-to-pdfmake doesn't parse CSS)
        fullHtml = fullHtml.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
        
        // Remove all <script> tags and their contents
        fullHtml = fullHtml.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
        
        // Remove HTML comments that might contain CSS or other problematic content
        fullHtml = fullHtml.replace(/<!--[\s\S]*?-->/g, '');
        
        // Remove any remaining CSS rules that might be floating around
        fullHtml = fullHtml.replace(/\{[^{}]*(?:color|font|margin|padding|background)[^{}]*\}/gi, '');
        
        // Clean up excessive borders and styling that create nested appearance
        fullHtml = fullHtml.replace(/border:\s*[^;]*;?/gi, '');
        fullHtml = fullHtml.replace(/border-[^:]*:\s*[^;]*;?/gi, '');
        
        // Improve image formatting - constrain sizes and remove problematic styling
        fullHtml = fullHtml.replace(/<img([^>]*?)style\s*=\s*["']([^"']*)["']([^>]*?)>/gi, (match, before, style, after) => {
            // Remove existing width/height from style and add proper constraints
            let cleanStyle = style
                .replace(/width\s*:\s*[^;]*;?/gi, '')
                .replace(/height\s*:\s*[^;]*;?/gi, '')
                .replace(/max-width\s*:\s*[^;]*;?/gi, '')
                .replace(/max-height\s*:\s*[^;]*;?/gi, '')
                .replace(/border[^:]*:\s*[^;]*;?/gi, '');
            
            // Add proper image constraints
            cleanStyle += '; max-width: 400px; max-height: 300px; width: auto; height: auto; display: block; margin: 10px auto;';
            
            return `<img${before}style="${cleanStyle}"${after}>`;
        });
        
        // Handle images without style attributes
        fullHtml = fullHtml.replace(/<img(?![^>]*style\s*=)([^>]*?)>/gi, 
            '<img$1 style="max-width: 400px; max-height: 300px; width: auto; height: auto; display: block; margin: 10px auto;">');
        
        // Fix problematic table structures that cause pdfMake preprocessing errors
        // Remove empty table elements that can cause issues
        fullHtml = fullHtml.replace(/<table[^>]*>\s*<\/table>/gi, '');
        fullHtml = fullHtml.replace(/<tr[^>]*>\s*<\/tr>/gi, '');
        fullHtml = fullHtml.replace(/<td[^>]*>\s*<\/td>/gi, '<td>&nbsp;</td>');
        fullHtml = fullHtml.replace(/<th[^>]*>\s*<\/th>/gi, '<th>&nbsp;</th>');
        
        // Ensure all table rows have at least one cell
        fullHtml = fullHtml.replace(/<tr([^>]*)>(?!.*<t[dh])/gi, '<tr$1><td>&nbsp;</td>');
        
        // Remove nested tables that can cause issues
        fullHtml = fullHtml.replace(/<table[^>]*>[\s\S]*?<table[^>]*>[\s\S]*?<\/table>[\s\S]*?<\/table>/gi, '<!-- Nested table removed -->');
        
        console.log('HTML cleaning completed');
        
        // Pre-process HTML to replace all font-family references with Roboto
        // This ensures no unknown fonts reach pdfMake
        fullHtml = fullHtml.replace(/font-family\s*:\s*[^;}"]+/gi, 'font-family: Roboto, sans-serif');
        // Also handle inline style attributes
        fullHtml = fullHtml.replace(/style\s*=\s*"([^"]*?)"/gi, (match, styleContent) => {
            const updatedStyle = styleContent.replace(/font-family\s*:\s*[^;]+/gi, 'font-family: Roboto, sans-serif');
            return `style="${updatedStyle}"`;
        });
        // Handle any remaining font references in CSS
        fullHtml = fullHtml.replace(/(Aptos|Calibri|Arial|Times New Roman|Helvetica|serif|sans-serif|monospace)/gi, 'Roboto');
        
        // Try imagesByReference first, then fallback to base64 conversion
        console.log('Attempting to use imagesByReference for better image handling...');
        
        try {
            // First attempt: Use imagesByReference which lets pdfMake handle images directly
            const pdfContentWithRefs = htmlToPdfmake(fullHtml, {
                tableAutoSize: true,
                removeExtraBlanks: true,
                imagesByReference: true, // Let pdfMake handle images by reference
                fontMapping: {
                    'Aptos': 'Roboto',
                    'Calibri': 'Roboto', 
                    'Arial': 'Roboto',
                    'Times New Roman': 'Roboto',
                    'Helvetica': 'Roboto',
                    'serif': 'Roboto',
                    'sans-serif': 'Roboto',
                    'monospace': 'Roboto'
                }
            });
            
            // Check if we have images that need to be processed
            if (pdfContentWithRefs.images && Object.keys(pdfContentWithRefs.images).length > 0) {
                console.log(`Found ${Object.keys(pdfContentWithRefs.images).length} images to process with imagesByReference`);
                console.log('Images found:', Object.keys(pdfContentWithRefs.images));
                
                // Create document definition with images
                const docDefinition = {
                    content: pdfContentWithRefs.content,
                    images: pdfContentWithRefs.images,
                    styles: {
                        header: {
                            fontSize: 12,
                            bold: true,
                            margin: [0, 0, 0, 8],
                            font: 'Roboto',
                            color: '#007acc'
                        },
                        subheader: {
                            fontSize: 11,
                            bold: true,
                            margin: [0, 5, 0, 5],
                            font: 'Roboto',
                            color: '#333'
                        }
                    },
                    defaultStyle: {
                        font: 'Roboto'
                    }
                };
                
                console.log('Creating PDF with images via imagesByReference...');
                pdfMake.createPdf(docDefinition).download('email.pdf');
                console.log('PDF generated successfully with imagesByReference!');
                return;
            } else {
                console.log('No images found with imagesByReference, proceeding with base64 conversion...');
            }
        } catch (error) {
            console.warn('imagesByReference approach failed, falling back to base64 conversion:', error);
        }
        
        // Fallback: Pre-process HTML to handle external images by converting them to data URLs
        // This allows pdfMake to handle images properly instead of throwing VFS errors
        console.log('Converting external images to base64 data URLs...');
        fullHtml = await convertExternalImagesToDataUrls(fullHtml);
        
        // Configure html-to-pdfmake options with font mapping and conservative image handling
        const options = {
            tableAutoSize: true,
            removeExtraBlanks: true,
            imagesByReference: false, // Keep disabled as we're now using data URLs
            ignoreExtraColumns: true, // Ignore extra columns in tables
            ignoreExtraRows: true, // Ignore extra rows in tables
            replaceText: function(text, node) {
                // Clean text and handle encoding issues
                if (typeof text === 'string') {
                    return text.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
                }
                return text;
            },
            fontMapping: {
                'Aptos': 'Roboto',
                'Calibri': 'Roboto', 
                'Arial': 'Roboto',
                'Times New Roman': 'Roboto',
                'Helvetica': 'Roboto',
                'sans-serif': 'Roboto',
                'serif': 'Roboto',
                'monospace': 'Roboto'
            },
            defaultStyles: {
                h1: { fontSize: 16, bold: true, marginBottom: 6, marginTop: 10, color: '#007acc', font: 'Roboto' },
                h2: { fontSize: 14, bold: true, marginBottom: 5, marginTop: 8, color: '#007acc', font: 'Roboto' },
                h3: { fontSize: 12, bold: true, marginBottom: 4, marginTop: 6, color: '#333', font: 'Roboto' },
                h4: { fontSize: 11, bold: true, marginBottom: 3, marginTop: 5, color: '#333', font: 'Roboto' },
                h5: { fontSize: 10, bold: true, marginBottom: 3, marginTop: 4, color: '#333', font: 'Roboto' },
                h6: { fontSize: 10, bold: true, marginBottom: 2, marginTop: 3, color: '#333', font: 'Roboto' },
                p: { margin: [0, 2, 0, 4], fontSize: 10, lineHeight: 1.3, font: 'Roboto' },
                div: { margin: [0, 1, 0, 2], fontSize: 10, font: 'Roboto' },
                span: { fontSize: 10, font: 'Roboto' },
                a: { color: '#0066cc', decoration: 'underline', fontSize: 10, font: 'Roboto' },
                strong: { bold: true, font: 'Roboto' },
                b: { bold: true, font: 'Roboto' },
                em: { italics: true, font: 'Roboto' },
                i: { italics: true, font: 'Roboto' },
                ul: { marginBottom: 4, marginLeft: 8, marginTop: 3, font: 'Roboto' },
                ol: { marginBottom: 4, marginLeft: 8, marginTop: 3, font: 'Roboto' },
                li: { marginBottom: 1, fontSize: 10, font: 'Roboto' },
                table: { marginBottom: 8, marginTop: 5, font: 'Roboto' },
                th: { bold: true, fillColor: '#f0f8ff', alignment: 'left', fontSize: 10, padding: [3, 3, 3, 3], font: 'Roboto' },
                td: { margin: [3, 2, 3, 2], fontSize: 10, font: 'Roboto' },
                img: { margin: [0, 5, 0, 5], alignment: 'center' },
                blockquote: { margin: [10, 5, 0, 5], italics: true, color: '#666', fontSize: 9, font: 'Roboto' }
            }
        };
        
        // Convert HTML to pdfmake format
        console.log('Converting HTML to pdfmake format...');
        const pdfmakeContent = htmlToPdfmake(fullHtml, options);
        
        // Validate and clean pdfmake content to prevent preprocessing errors
        const validateAndCleanContent = (content) => {
            if (Array.isArray(content)) {
                return content.map(validateAndCleanContent).filter(item => item !== null);
            }
            
            if (typeof content === 'object' && content !== null) {
                // Handle table objects
                if (content.table) {
                    // Ensure table has valid structure
                    if (!content.table.body || !Array.isArray(content.table.body) || content.table.body.length === 0) {
                        console.warn('Removing invalid table with no body');
                        return null;
                    }
                    
                    // Validate each row has cells
                    content.table.body = content.table.body.filter(row => {
                        return Array.isArray(row) && row.length > 0;
                    });
                    
                    // If no valid rows remain, remove the table
                    if (content.table.body.length === 0) {
                        console.warn('Removing table with no valid rows');
                        return null;
                    }
                    
                    // Ensure all rows have the same number of columns
                    const maxCols = Math.max(...content.table.body.map(row => row.length));
                    content.table.body = content.table.body.map(row => {
                        while (row.length < maxCols) {
                            row.push('');
                        }
                        return row;
                    });
                }
                
                // Recursively validate other properties
                const cleanedContent = {};
                for (const [key, value] of Object.entries(content)) {
                    const cleanedValue = validateAndCleanContent(value);
                    if (cleanedValue !== null) {
                        cleanedContent[key] = cleanedValue;
                    }
                }
                return cleanedContent;
            }
            
            return content;
        };
        
        const cleanedContent = validateAndCleanContent(pdfmakeContent);
        console.log('Content validation and cleaning completed');
        
        // Create the document definition
        const docDefinition = {
            content: cleanedContent,
            defaultStyle: {
                fontSize: 10,
                lineHeight: 1.2,
                font: 'Roboto',
                color: '#333'
            },
            styles: {
                header: {
                    fontSize: 12,
                    bold: true,
                    margin: [0, 0, 0, 8],
                    font: 'Roboto',
                    color: '#007acc'
                },
                subheader: {
                    fontSize: 11,
                    bold: true,
                    margin: [0, 5, 0, 5],
                    font: 'Roboto',
                    color: '#333'
                }
            },
            pageSize: 'A4',
            pageOrientation: 'portrait',
            pageMargins: [40, 50, 40, 60],
            footer: function(currentPage, pageCount) {
                return {
                    text: `Generated by Email2PDF • ${new Date().toLocaleDateString()} • Page ${currentPage} of ${pageCount}`,
                    alignment: 'center',
                    fontSize: 8,
                    color: '#888',
                    margin: [0, 10, 0, 0],
                    font: 'Roboto'
                };
            },
            info: {
                title: (Office.context.mailbox.item.subject || 'Email') + ' - PDF',
                author: 'Email2PDF',
                subject: Office.context.mailbox.item.subject || 'Email Export',
                creator: 'Email2PDF v1.6.0'
            }
        };
        
        // Generate and download the PDF
        try {
            console.log('Creating PDF document...');
            pdfMake.createPdf(docDefinition).download((Office.context.mailbox.item.subject || 'email') + '.pdf');
            console.log('PDF created successfully with html-to-pdfmake');
        } catch (pdfError) {
            console.error('PDF creation failed with cleaned content, trying simplified fallback:', pdfError);
            
            // Create a simplified fallback document with just text content
            const fallbackContent = [
                { text: 'Email Details', style: 'header' },
                ...emailDetails.map(detail => ({ text: detail, fontSize: 10, margin: [0, 2] })),
                { text: '\nMessage Content', style: 'header', marginTop: 15 },
                { text: stripHtml(htmlBody), fontSize: 10, margin: [0, 5], lineHeight: 1.3 }
            ];
            
            const fallbackDocDefinition = {
                content: fallbackContent,
                defaultStyle: {
                    fontSize: 10,
                    lineHeight: 1.2,
                    font: 'Roboto',
                    color: '#333'
                },
                styles: {
                    header: {
                        fontSize: 12,
                        bold: true,
                        margin: [0, 0, 0, 8],
                        font: 'Roboto',
                        color: '#007acc'
                    }
                },
                pageSize: 'A4',
                pageMargins: [40, 50, 40, 60],
                footer: function(currentPage, pageCount) {
                    return {
                        text: `Generated by Email2PDF • ${new Date().toLocaleDateString()} • Page ${currentPage} of ${pageCount}`,
                        alignment: 'center',
                        fontSize: 8,
                        color: '#888',
                        margin: [0, 10, 0, 0],
                        font: 'Roboto'
                    };
                }
            };
            
            try {
                console.log('Creating fallback PDF...');
                pdfMake.createPdf(fallbackDocDefinition).download((Office.context.mailbox.item.subject || 'email') + '.pdf');
                console.log('Fallback PDF created successfully');
            } catch (fallbackError) {
                console.error('Both PDF creation attempts failed:', fallbackError);
                throw fallbackError;
            }
        }
    } catch (err) {
        console.error('Error creating PDF with html-to-pdfmake:', err);
        throw err;
    }
}

async function createServerPdf(metaHtml, htmlBody, attachments) {
    console.log('Creating PDF with server-side generation');
    
    try {
        // Check if server is available
        const serverUrl = 'http://localhost:5000';
        
        // Test server connection first
        try {
            const healthResponse = await fetch(`${serverUrl}/health`, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                },
                signal: AbortSignal.timeout(5000) // 5 second timeout
            });
            
            if (!healthResponse.ok) {
                throw new Error(`Server health check failed: ${healthResponse.status}`);
            }
        } catch (error) {
            console.error('Server connection failed:', error);
            throw new Error('PDF server is not running. Please start the server using: docker-compose up or python pdf-server/app.py');
        }
        
        // Combine metadata and content into a single HTML document
        const fullHtml = `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Email PDF</title>
            <style>
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
                    line-height: 1.6;
                    color: #333;
                    margin: 0;
                    padding: 20px;
                    max-width: 800px;
                }
                
                .email-header {
                    background-color: #f8f9fa;
                    border-left: 4px solid #007acc;
                    padding: 15px;
                    margin-bottom: 20px;
                    border-radius: 4px;
                }
                
                .email-header h2 {
                    margin: 0 0 10px 0;
                    color: #007acc;
                    font-size: 16px;
                }
                
                .email-header p {
                    margin: 4px 0;
                    font-size: 14px;
                }
                
                .email-content {
                    border-top: 2px solid #007acc;
                    padding-top: 20px;
                }
                
                .email-content h2 {
                    color: #007acc;
                    font-size: 16px;
                    margin: 0 0 15px 0;
                }
                
                img {
                    max-width: 100%;
                    height: auto;
                    display: block;
                    margin: 10px 0;
                }
                
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin: 15px 0;
                }
                
                table, th, td {
                    border: 1px solid #ddd;
                }
                
                th, td {
                    padding: 8px;
                    text-align: left;
                }
                
                th {
                    background-color: #f5f5f5;
                    font-weight: bold;
                }
                
                ul, ol {
                    margin: 10px 0;
                    padding-left: 20px;
                }
                
                li {
                    margin: 5px 0;
                }
                
                h1, h2, h3, h4, h5, h6 {
                    margin: 20px 0 10px 0;
                    line-height: 1.2;
                }
                
                p {
                    margin: 10px 0;
                }
                
                a {
                    color: #0066cc;
                    text-decoration: underline;
                }
                
                blockquote {
                    border-left: 4px solid #ccc;
                    margin: 15px 0;
                    padding: 10px 15px;
                    background-color: #f9f9f9;
                    font-style: italic;
                }
            </style>
        </head>
        <body>
            <div class="email-header">
                <h2>Email Details</h2>
                ${metaHtml}
            </div>
            <div class="email-content">
                <h2>Message Content</h2>
                ${htmlBody}
            </div>
        </body>
        </html>`;
        
        // Prepare the request payload
        const payload = {
            html: fullHtml,
            options: {
                page_size: 'A4',
                margin: '20mm',
                orientation: 'portrait'
            }
        };
        
        console.log('Sending request to PDF server...');
        
        // Send request to server
        const response = await fetch(`${serverUrl}/convert`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload),
            signal: AbortSignal.timeout(30000) // 30 second timeout for PDF generation
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Server error: ${response.status} - ${errorText}`);
        }
        
        // Get the PDF blob
        const pdfBlob = await response.blob();
        
        if (pdfBlob.size === 0) {
            throw new Error('Server returned empty PDF');
        }
        
        // Create download link
        const link = document.createElement('a');
        link.href = URL.createObjectURL(pdfBlob);
        link.download = (Office.context.mailbox.item.subject || 'email') + '.pdf';
        link.click();
        
        // Clean up the blob URL
        setTimeout(() => {
            URL.revokeObjectURL(link.href);
        }, 1000);
        
        console.log('PDF created successfully with server-side generation');
        
    } catch (err) {
        console.error('Error creating PDF with server:', err);
        
        // Provide helpful error messages
        if (err.message.includes('server is not running')) {
            alert('PDF Server Error: The PDF server is not running.\n\nTo start the server:\n1. Open a terminal\n2. Navigate to the project folder\n3. Run: docker-compose up\n\nOr see the README for detailed instructions.');
        } else if (err.message.includes('timeout')) {
            alert('PDF Generation Timeout: The server took too long to respond. Please try again with a shorter email or check your server logs.');
        } else {
            alert(`PDF Generation Error: ${err.message}\n\nPlease check that the PDF server is running and try again.`);
        }
        
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

                        // Wait for libraries to load if needed
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

                        if (engine === 'html-to-pdfmake' && (!htmlToPdfmake || !pdfMake)) {
                            console.error('html-to-pdfmake libraries not loaded yet.');
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
                            msgDiv.textContent = 'html-to-pdfmake libraries not loaded yet. Please try again in a moment.';
                            this.disabled = false;
                            this.textContent = 'Convert Email to PDF';
                            return;
                        }

                        try {
                            if (engine === 'pdf-lib') {
                                await createPdfLibTextPdf(metaHtml, htmlBody, attachments);
                            } else if (engine === 'html-to-pdfmake') {
                                await createHtmlToPdfmakePdf(metaHtml, htmlBody, attachments);
                            } else if (engine === 'server-pdf') {
                                await createServerPdf(metaHtml, htmlBody, attachments);
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
        
        // Get test button and attach click handler for testing PDF engines
        const testBtn = document.getElementById('testBtn');
        if (testBtn) {
            console.log('testBtn found:', testBtn);
            
            testBtn.onclick = async function() {
                console.log('Test button clicked.');
                this.disabled = true;
                this.textContent = 'Testing...';
                
                try {
                    // Get selected engine
                    const engineSel = document.getElementById('pdfEngine');
                    const engine = engineSel ? engineSel.value : 'html-to-pdfmake';
                    console.log('Testing PDF Engine:', engine);
                    
                    // Sample test HTML with different fonts and images
                    const testMetaHtml = `<div style="background:#f5f5f5;padding:12px 18px 12px 18px;margin-bottom:18px;border-radius:6px;font-size:1em;line-height:1.5;font-family:Aptos,Calibri,Arial,sans-serif;">
                        <div><b>From:</b> test@example.com (Test User)</div>
                        <div><b>To:</b> recipient@example.com (Recipient)</div>
                        <div><b>Subject:</b> Test Email with Various Fonts</div>
                        <div><b>Sent:</b> ${new Date().toLocaleString()}</div>
                    </div>`;
                    
                    const testHtmlBody = `
                        <div style="font-family: Roboto, sans-serif; font-size: 12px; line-height: 1.4; color: #333;">
                            <h1 style="font-family: Roboto, sans-serif; color: #007acc;">Test Email with Images and Formatting</h1>
                            <p style="font-family: Roboto, sans-serif;">This paragraph demonstrates <strong>bold text</strong> and <em>italic text</em> formatting.</p>
                            
                            <h2 style="font-family: Roboto, sans-serif; color: #007acc;">Image Gallery</h2>
                            <p style="font-family: Roboto, sans-serif;">The images below should display properly sized and centered:</p>
                            
                            <div style="text-align: center; margin: 15px 0;">
                                <img src="https://via.placeholder.com/200x150/0066cc/ffffff?text=Sample+Image+1" 
                                     alt="Test Image 1" 
                                     style="max-width: 200px; max-height: 150px; margin: 10px; border-radius: 4px;">
                            </div>
                            
                            <div style="text-align: center; margin: 15px 0;">
                                <img src="https://via.placeholder.com/250x100/cc6600/ffffff?text=Wide+Image+2" 
                                     alt="Test Image 2" 
                                     style="max-width: 250px; max-height: 100px; margin: 10px; border-radius: 4px;">
                            </div>
                            
                            <h2 style="font-family: Roboto, sans-serif; color: #007acc;">Content Formatting Test</h2>
                            
                            <h3 style="font-family: Roboto, sans-serif;">Lists and Structure</h3>
                            <ul style="font-family: Roboto, sans-serif; margin-left: 20px;">
                                <li>First list item with proper spacing</li>
                                <li>Second item with <strong>bold text</strong></li>
                                <li>Third item with <a href="#" style="color: #0066cc;">a sample link</a></li>
                            </ul>
                            
                            <h3 style="font-family: Roboto, sans-serif;">Table Example</h3>
                            <table style="width: 100%; margin: 15px 0; border-collapse: collapse; font-family: Roboto, sans-serif;">
                                <tr style="background-color: #f0f8ff;">
                                    <th style="padding: 8px; text-align: left; border: 1px solid #ddd;">Feature</th>
                                    <th style="padding: 8px; text-align: left; border: 1px solid #ddd;">Status</th>
                                </tr>
                                <tr>
                                    <td style="padding: 8px; border: 1px solid #ddd;">Font Mapping</td>
                                    <td style="padding: 8px; border: 1px solid #ddd;">✓ Working</td>
                                </tr>
                                <tr style="background-color: #f9f9f9;">
                                    <td style="padding: 8px; border: 1px solid #ddd;">Image Display</td>
                                    <td style="padding: 8px; border: 1px solid #ddd;">✓ Improved</td>
                                </tr>
                            </table>
                            
                            <div style="text-align: center; margin: 20px 0;">
                                <img src="https://via.placeholder.com/180x120/00cc66/ffffff?text=Final+Image" 
                                     alt="Final Test Image" 
                                     style="max-width: 180px; max-height: 120px; border-radius: 4px;">
                            </div>
                            
                            <p style="font-family: Roboto, sans-serif; color: #666; font-size: 11px; margin-top: 20px;">
                                This test demonstrates improved formatting with:
                                <br>• Better image sizing and positioning
                                <br>• Cleaner layout without excessive borders
                                <br>• Proper font mapping and consistent styling
                                <br>• Optimized spacing and visual hierarchy
                            </p>
                        </div>
                    `;
                    
                    // Test the selected engine
                    if (engine === 'html-to-pdfmake') {
                        if (!htmlToPdfmake || !pdfMake) {
                            throw new Error('html-to-pdfmake libraries not loaded yet. Please try again in a moment.');
                        }
                        await createHtmlToPdfmakePdf(testMetaHtml, testHtmlBody, []);
                        console.log('Test completed successfully with html-to-pdfmake');
                    } else if (engine === 'pdf-lib') {
                        if (!PDFDocument) {
                            throw new Error('PDF library not loaded yet. Please try again in a moment.');
                        }
                        await createPdfLibTextPdf(testMetaHtml, testHtmlBody, []);
                        console.log('Test completed successfully with pdf-lib');
                    } else if (engine === 'server-pdf') {
                        await createServerPdf(testMetaHtml, testHtmlBody, []);
                        console.log('Test completed successfully with server-pdf');
                    } else {
                        // html2pdf test
                        await createPdf(testMetaHtml + testHtmlBody, [], []);
                        console.log('Test completed successfully with html2pdf');
                    }
                    
                    // Show success message
                    let msgDiv = document.getElementById('pdf2email-message');
                    if (!msgDiv) {
                        msgDiv = document.createElement('div');
                        msgDiv.id = 'pdf2email-message';
                        msgDiv.style.color = '#007acc';
                        msgDiv.style.fontSize = '1em';
                        msgDiv.style.margin = '12px 0';
                        msgDiv.style.display = 'block';
                        const appDiv = document.getElementById('app');
                        if (appDiv) appDiv.parentNode.insertBefore(msgDiv, appDiv.nextSibling);
                        else document.body.insertBefore(msgDiv, document.body.firstChild);
                    }
                    msgDiv.style.color = '#007acc';
                    msgDiv.textContent = `Test PDF generated successfully using ${engine} engine! Check your downloads folder.`;
                    
                } catch (err) {
                    console.error('Error in test PDF generation:', err);
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
                    msgDiv.style.color = '#b00';
                    msgDiv.textContent = 'Test PDF generation failed: ' + err.message;
                }
                
                this.disabled = false;
                this.textContent = 'Test PDF Engine';
            };
        }
        
        console.log('Button click handler attached.');
    } catch (err) {
        console.error('Error in initializeApp:', err);
    }
}

