console.log('taskpane.js loaded - Server PDF engine only v4.1.6');

// Global variables for performance optimization
let serverHealthStatus = { healthy: null, lastCheck: 0 };
const HEALTH_CHECK_CACHE_DURATION = 30000; // 30 seconds

// Cache for converted images to avoid redundant processing
const imageCache = new Map();
const IMAGE_CACHE_SIZE = 50; // Maximum number of images to cache
const IMAGE_CACHE_DURATION = 60 * 60 * 1000; // 1 hour cache duration

// Global error handler for debugging
window.addEventListener('error', function(e) {
    console.error('Global JS error:', e.message, e);
});
window.addEventListener('unhandledrejection', function(e) {
    console.error('Unhandled promise rejection:', e.reason);
});

// Helper function to strip HTML tags and clean text
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
    text = text.replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000]/g, ' ');
    text = text.replace(/[\u2010-\u2015]/g, '-');
    text = text.replace(/[\u2018\u2019]/g, "'");
    text = text.replace(/[\u201C\u201D]/g, '"');
    text = text.replace(/[\u2026]/g, '...');
    text = text.replace(/[\u2013\u2014]/g, '-');
    
    // Remove emojis and problematic symbols
    text = text.replace(/[\u1F000-\u1FFFF]/g, '');
    text = text.replace(/[\u2600-\u27BF]/g, '');
    text = text.replace(/[\uFE00-\uFE0F]/g, '');
    
    // Only remove control characters that are truly problematic
    text = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
    
    // Replace multiple whitespace with single spaces
    text = text.replace(/\s+/g, ' ');
    
    return text.trim();
}

// Pre-optimize HTML content before sending to server for faster processing
function preOptimizeHtmlForServer(html) {
    // Don't process data URLs (already optimized)
    if (!html || html.length < 1000) {
        return html;
    }
    
    console.log('Pre-optimizing HTML content before sending to server');
    
    try {
        // Parse HTML to modify it
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        
        // 1. Remove comments
        const commentIterator = document.createNodeIterator(
            doc, NodeFilter.SHOW_COMMENT, null, false
        );
        let currentNode;
        const commentsToRemove = [];
        while (currentNode = commentIterator.nextNode()) {
            commentsToRemove.push(currentNode);
        }
        commentsToRemove.forEach(comment => {
            comment.parentNode.removeChild(comment);
        });
        
        // 2. Remove unused attributes that don't affect rendering
        const allElements = doc.querySelectorAll('*');
        const attributesToRemove = ['data-', 'aria-', 'role', 'tabindex', 'title'];
        
        allElements.forEach(element => {
            // Skip essential elements
            if (element.tagName === 'IMG' || element.tagName === 'A') {
                return;
            }
            
            Array.from(element.attributes).forEach(attr => {
                const attrName = attr.name.toLowerCase();
                // Remove data attributes and other non-visual attributes
                if (attributesToRemove.some(prefix => attrName.startsWith(prefix)) || 
                    attrName === 'id' || attrName === 'name') {
                    element.removeAttribute(attrName);
                }
            });
        });
        
        // 3. Simplify inline styles that don't affect PDF rendering
        const elementsWithStyle = doc.querySelectorAll('[style]');
        elementsWithStyle.forEach(element => {
            const style = element.getAttribute('style');
            // Keep only essential styles
            const essentialStyles = style
                .split(';')
                .filter(s => {
                    const prop = s.split(':')[0]?.trim().toLowerCase();
                    // Keep only these essential visual properties
                    return prop && [
                        'color', 'background', 'font', 'margin', 'padding', 
                        'border', 'text-align', 'width', 'height', 'display',
                        'line-height', 'letter-spacing', 'vertical-align'
                    ].some(p => prop.includes(p));
                })
                .join(';');
                
            element.setAttribute('style', essentialStyles);
        });
        
        // Convert back to string
        const serializer = new XMLSerializer();
        let optimizedHtml = serializer.serializeToString(doc);
        
        // Clean up any excessive whitespace
        optimizedHtml = optimizedHtml.replace(/>\s+</g, '><');
        
        console.log(`HTML optimization complete: ${Math.round((html.length - optimizedHtml.length) / html.length * 100)}% size reduction`);
        
        return optimizedHtml;
    } catch (err) {
        console.warn('Error optimizing HTML, using original:', err);
        return html;
    }
}

// Optimized function to get attachment content with parallel processing
async function getAttachmentContentOptimized(attachments) {
    if (!attachments || attachments.length === 0) return [];
    
    console.log(`Processing ${attachments.length} attachments in parallel`);
    
    const attachmentPromises = attachments.map(attachment => 
        new Promise((resolve) => {
            Office.context.mailbox.item.getAttachmentContentAsync(
                attachment.id,
                { asyncContext: { attachment } },
                function(result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve({
                            success: true,
                            name: result.asyncContext.attachment.name,
                            contentType: result.asyncContext.attachment.contentType,
                            size: result.asyncContext.attachment.size,
                            content: result.value.content
                        });
                    } else {
                        console.error('Failed to get attachment content:', result.error);
                        resolve({
                            success: false,
                            name: result.asyncContext.attachment.name,
                            error: result.error
                        });
                    }
                }
            );
        })
    );
    
    const results = await Promise.allSettled(attachmentPromises);
    return results
        .filter(result => result.status === 'fulfilled' && result.value.success)
        .map(result => result.value);
}

// Optimized function to convert external images to data URLs with caching, timeout and parallel processing
async function convertExternalImagesToDataUrlsOptimized(html, timeout = 2000) {
    console.log('Converting external images to data URLs with caching and 2s timeout');
    
    const imgRegex = /<img[^>]+src\s*=\s*["']([^"']+)["'][^>]*>/gi;
    const matches = [];
    let match;
    
    while ((match = imgRegex.exec(html)) !== null) {
        const fullMatch = match[0];
        const url = match[1];
        if (url.startsWith('http')) {
            matches.push({ fullMatch, url });
        }
    }
    
    if (matches.length === 0) return html;
    
    console.log(`Found ${matches.length} external images to convert`);
    
    // Clean up old cache entries
    const now = Date.now();
    for (const [key, entry] of imageCache.entries()) {
        if (now - entry.timestamp > IMAGE_CACHE_DURATION) {
            imageCache.delete(key);
        }
    }
    
    const imagePromises = matches.map(async ({ fullMatch, url }) => {
        // Check cache first
        if (imageCache.has(url)) {
            console.log(`Using cached image for ${url}`);
            return { fullMatch, dataURL: imageCache.get(url).dataURL };
        }
        
        try {
            const timeoutPromise = new Promise((_, reject) => 
                setTimeout(() => reject(new Error('Image load timeout')), timeout)
            );
            
            const imagePromise = new Promise((resolve, reject) => {
                const img = new Image();
                img.crossOrigin = 'anonymous';
                img.onload = () => {
                    try {
                        // Use a smaller canvas size for performance
                        const maxWidth = 1200; // Reasonable max width for email images
                        const width = Math.min(img.width, maxWidth);
                        const scaleFactor = width / img.width;
                        const height = img.height * scaleFactor;
                        
                        const canvas = document.createElement('canvas');
                        const ctx = canvas.getContext('2d');
                        canvas.width = width;
                        canvas.height = height;
                        ctx.drawImage(img, 0, 0, width, height);
                        
                        // Use lower quality for performance
                        const dataURL = canvas.toDataURL('image/jpeg', 0.7);
                        
                        // Cache the result
                        if (imageCache.size >= IMAGE_CACHE_SIZE) {
                            // Remove oldest entry if cache is full
                            const oldestKey = imageCache.keys().next().value;
                            imageCache.delete(oldestKey);
                        }
                        imageCache.set(url, { dataURL, timestamp: Date.now() });
                        
                        resolve({ fullMatch, dataURL });
                    } catch (err) {
                        reject(err);
                    }
                };
                img.onerror = () => reject(new Error('Image load failed'));
                img.src = url;
            });
            
            return await Promise.race([imagePromise, timeoutPromise]);
        } catch (err) {
            console.warn(`Failed to convert image: ${url}`, err);
            return { fullMatch, dataURL: null };
        }
    });
    
    const results = await Promise.allSettled(imagePromises);
    let convertedHtml = html;
    
    results.forEach(result => {
        if (result.status === 'fulfilled' && result.value.dataURL) {
            const { fullMatch, dataURL } = result.value;
            const newImg = fullMatch.replace(/src\s*=\s*["'][^"']+["']/, `src="${dataURL}"`);
            convertedHtml = convertedHtml.replace(fullMatch, newImg);
        }
    });
    
    return convertedHtml;
}

// Cached server health check to avoid redundant requests
async function checkServerHealthCached() {
    const now = Date.now();
    
    // Return cached result if recent
    if (serverHealthStatus.lastCheck && (now - serverHealthStatus.lastCheck) < HEALTH_CHECK_CACHE_DURATION) {
        return serverHealthStatus.healthy;
    }
    
    try {
        // Use AbortController for timeout
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 2000); // 2 second timeout
        
        const response = await fetch('http://localhost:5000/health', {
            method: 'GET',
            mode: 'cors',
            headers: {
                'Accept': 'application/json'
            },
            signal: controller.signal
        });
        
        // Clear the timeout
        clearTimeout(timeoutId);
        
        const isHealthy = response.ok;
        serverHealthStatus = { healthy: isHealthy, lastCheck: now };
        return isHealthy;
    } catch (err) {
        console.warn('Server health check failed:', err);
        serverHealthStatus = { healthy: false, lastCheck: now };
        return false;
    }
}

// Optimized Server PDF creation with parallel processing and caching
async function createServerPdf(metaHtml, htmlBody, attachments) {
    console.log('Creating PDF with optimized server-side generation');
    
    try {
        const serverUrl = 'http://localhost:5000';
        
        // Check server health first (cached)
        const isServerHealthy = await checkServerHealthCached();
        if (!isServerHealthy) {
            throw new Error('PDF server is not available. Please ensure the server is running on localhost:5000');
        }
        
        // Process attachments and images in parallel with reduced timeout
        const [processedAttachments, convertedHtml] = await Promise.all([
            getAttachmentContentOptimized(attachments),
            convertExternalImagesToDataUrlsOptimized(htmlBody, 2000) // Reduced timeout to 2 seconds
        ]);
        
        // Combine meta HTML and body HTML into a single HTML document
        const fullHtml = metaHtml + convertedHtml;
        
        // Pre-optimize HTML content to reduce size
        const optimizedHtml = preOptimizeHtmlForServer(fullHtml);
        
        // Further optimize attachments to reduce payload size
        const optimizedAttachments = processedAttachments.map(attachment => ({
            name: attachment.name,
            contentType: attachment.contentType,
            content: attachment.content,
            size: attachment.size
        }));
        
        const requestData = {
            html: optimizedHtml,
            attachments: optimizedAttachments,
            mode: 'full',
            extractImages: false
        };
        
        console.log('Sending optimized request to server...');
        
        // Use AbortController for timeout
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout
        
        const response = await fetch(`${serverUrl}/convert-with-attachments`, {
            method: 'POST',
            mode: 'cors',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/pdf'
            },
            body: JSON.stringify(requestData),
            signal: controller.signal
        });
        
        // Clear the timeout
        clearTimeout(timeoutId);
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Server error: ${response.status} ${response.statusText} - ${errorText}`);
        }
        
        const pdfBlob = await response.blob();
        const url = URL.createObjectURL(pdfBlob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = (Office.context.mailbox.item.subject || 'email') + '.pdf';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        console.log('PDF created and downloaded successfully with optimized server processing');
        
    } catch (err) {
        console.error('Error creating PDF with server:', err);
        throw err;
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
        
        // Attach optimized click handler
        btn.onclick = async function() {
            console.log('Convert button clicked.');
            this.disabled = true;
            this.textContent = 'Processing...';
            
            // Add progress indicator
            let progressDiv = document.getElementById('progress-indicator');
            if (!progressDiv) {
                progressDiv = document.createElement('div');
                progressDiv.id = 'progress-indicator';
                progressDiv.style.cssText = 'margin: 10px 0; font-size: 0.9em; color: #007acc; background: #f0f8ff; padding: 8px; border-radius: 4px; border-left: 4px solid #007acc;';
                btn.parentNode.insertBefore(progressDiv, btn.nextSibling);
            }
            progressDiv.textContent = '🔄 Initializing PDF generation...';
            
            try {
                // Check server health first
                progressDiv.textContent = '🔄 Checking server availability...';
                const isServerHealthy = await checkServerHealthCached();
                if (!isServerHealthy) {
                    throw new Error('PDF server is not available. Please ensure the server is running on localhost:5000');
                }
                
                progressDiv.textContent = '🔄 Getting email content...';
                
                // Get the current email
                const item = Office.context.mailbox.item;
                if (!item) {
                    throw new Error('No email item found');
                }
                
                // Get email content and metadata
                const getBodyPromise = new Promise((resolve, reject) => {
                    item.body.getAsync(Office.CoercionType.Html, (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            resolve(result.value);
                        } else {
                            reject(new Error('Failed to get email body: ' + result.error.message));
                        }
                    });
                });
                
                const getAttachmentsPromise = new Promise((resolve) => {
                    if (item.attachments && item.attachments.length > 0) {
                        resolve(item.attachments);
                    } else {
                        resolve([]);
                    }
                });
                
                progressDiv.textContent = '🔄 Processing email content...';
                const [htmlBody, attachments] = await Promise.all([getBodyPromise, getAttachmentsPromise]);
                
                // Create meta HTML with email details
                const metaHtml = `<div style="background:#f5f5f5;padding:12px 18px 12px 18px;margin-bottom:18px;border-radius:6px;font-size:1em;line-height:1.5;font-family:Aptos,Calibri,Arial,sans-serif;">
                    <div><b>From:</b> ${item.from ? item.from.displayName + ' (' + item.from.emailAddress + ')' : 'Unknown'}</div>
                    <div><b>To:</b> ${item.to ? item.to.map(recipient => recipient.displayName || recipient.emailAddress).join(', ') : 'Unknown'}</div>
                    <div><b>Subject:</b> ${item.subject || 'No Subject'}</div>
                    <div><b>Sent:</b> ${item.dateTimeCreated ? item.dateTimeCreated.toLocaleString() : 'Unknown'}</div>
                </div>`;
                
                progressDiv.textContent = '🔄 Generating PDF with optimized server processing...';
                
                // Generate PDF using optimized Server PDF
                await createServerPdf(metaHtml, htmlBody, attachments);
                
                progressDiv.textContent = '✅ PDF generated successfully! Check your downloads folder.';
                progressDiv.style.color = '#28a745';
                progressDiv.style.borderLeftColor = '#28a745';
                
                // Remove progress indicator after 3 seconds
                setTimeout(() => {
                    if (progressDiv && progressDiv.parentNode) {
                        progressDiv.parentNode.removeChild(progressDiv);
                    }
                }, 3000);
                
            } catch (err) {
                console.error('Error in PDF generation:', err);
                
                // Update progress indicator with error
                if (progressDiv) {
                    progressDiv.textContent = '❌ Error: ' + err.message;
                    progressDiv.style.color = '#dc3545';
                    progressDiv.style.borderLeftColor = '#dc3545';
                    progressDiv.style.backgroundColor = '#f8d7da';
                }
                
                // Also show error in message div for compatibility
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
                this.textContent = 'Testing Server PDF...';
                
                try {
                    console.log('Testing optimized Server PDF engine');
                    
                    // Sample test HTML with different fonts and images
                    const testMetaHtml = `<div style="background:#f5f5f5;padding:12px 18px 12px 18px;margin-bottom:18px;border-radius:6px;font-size:1em;line-height:1.5;font-family:Aptos,Calibri,Arial,sans-serif;">
                        <div><b>From:</b> test@example.com (Test User)</div>
                        <div><b>To:</b> recipient@example.com (Recipient)</div>
                        <div><b>Subject:</b> Test Email with Optimized Performance</div>
                        <div><b>Sent:</b> ${new Date().toLocaleString()}</div>
                    </div>`;
                    
                    const testHtmlBody = '<div style="font-family: Roboto, sans-serif; font-size: 12px; line-height: 1.4; color: #333;">' +
                        '<h1 style="font-family: Roboto, sans-serif; color: #007acc;">Optimized Server PDF Test</h1>' +
                        '<p style="font-family: Roboto, sans-serif;">This test demonstrates the <strong>improved performance</strong> of the optimized server PDF engine.</p>' +
                        '<h2 style="font-family: Roboto, sans-serif; color: #007acc;">Performance Improvements</h2>' +
                        '<ul style="font-family: Roboto, sans-serif; margin-left: 20px;">' +
                        '<li>✅ Parallel attachment processing</li>' +
                        '<li>✅ Cached server health checks</li>' +
                        '<li>✅ Optimized image conversion (5s timeout)</li>' +
                        '<li>✅ Real-time progress indicators</li>' +
                        '<li>✅ Simplified single-engine architecture</li>' +
                        '</ul>' +
                        '<p style="font-family: Roboto, sans-serif; color: #666; font-size: 11px; margin-top: 20px;">' +
                        'Server PDF engine now provides the best balance of speed, quality, and reliability.' +
                        '</p></div>';
                    
                    // Test the optimized server-pdf engine only
                    await createServerPdf(testMetaHtml, testHtmlBody, []);
                    console.log('Test completed successfully with optimized server-pdf');
                    
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
                    msgDiv.textContent = 'Test PDF generated successfully using optimized Server PDF engine! Check your downloads folder.';
                    
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

console.log('File loaded successfully');
