console.log('taskpane.js loaded - Server PDF engine only');

// Global variables for performance optimization
let serverHealthStatus = { healthy: null, lastCheck: 0 };
const HEALTH_CHECK_CACHE_DURATION = 30000; // 30 seconds

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

// Helper function to get attachment content as base64
async function getAttachmentContent(attachment) {
    return new Promise((resolve, reject) => {
        console.log(`Fetching content for attachment: ${attachment.name}`);
        
        Office.context.mailbox.item.getAttachmentContentAsync(
            attachment.id,
            { asyncContext: attachment },
            function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log(`Successfully retrieved content for attachment: ${attachment.name}`);
                    // The content is already base64 encoded from the API
                    resolve(result.value.content);
                } else {
                    console.error(`Failed to get attachment content for ${attachment.name}:`, result.error);
                    reject(new Error(`Failed to get attachment content: ${result.error.message}`));
                }
            }
        );
    });
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
        
