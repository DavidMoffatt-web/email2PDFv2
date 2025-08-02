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
