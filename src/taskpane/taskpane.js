console.log('taskpane.js loaded - Server metadata sender');

async function sendEmailToServer() {
    console.log('Sending email data to server');
    // Read user-selected output mode and engine from the UI
    const modeEl = document.getElementById('pdfMode');
    const engineEl = document.getElementById('pdfEngine');
    const outputMode = modeEl ? modeEl.value : 'full'; // 'full' | 'individual' | 'images'
    const pdfEngine = engineEl ? engineEl.value : 'server-pdf';
    const separateAttachments = outputMode !== 'full';
    const splitInlineImages = outputMode === 'images';
    
    try {
        const item = Office.context.mailbox.item;
        if (!item) {
            throw new Error('No email item found');
        }

        // Get the email body with enhanced options for better content capture
        const htmlBody = await new Promise((resolve, reject) => {
            // Use FullBody mode to get complete conversation content for better formatting
            const options = {
                bodyMode: Office.MailboxEnums.BodyMode.FullBody,
                asyncContext: 'bodyRetrieval'
            };
            
            item.body.getAsync(Office.CoercionType.Html, options, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error('Failed to get email body'));
                }
            });
        });

        // Get attachments with enhanced content capture
        const attachmentPromises = (item.attachments || []).map(attachment => {
            return new Promise((resolve) => {
                // Prefer retrieving content for file attachments when API is available
                const canGetContent =
                    typeof item.getAttachmentContentAsync === 'function' &&
                    (!Office.context.requirements || Office.context.requirements.isSetSupported?.('Mailbox', '1.8') || true);

                if (attachment.attachmentType === Office.MailboxEnums.AttachmentType.File && canGetContent) {
                    const options = { asyncContext: attachment.id };
                    item.getAttachmentContentAsync(attachment.id, options, (result) => {
                        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                            resolve({
                                id: attachment.id,
                                name: attachment.name,
                                contentType: attachment.contentType,
                                size: attachment.size,
                                isInline: attachment.isInline,
                                contentId: attachment.contentId,
                                content: result.value.content, // Base64 encoded content
                                format: result.value.format,
                                // Additional metadata for better processing
                                attachmentType: 'file',
                                lastModifiedTime: attachment.lastModifiedTime || null
                            });
                        } else {
                            // Still return basic info even if content fails
                            resolve({
                                id: attachment.id,
                                name: attachment.name,
                                contentType: attachment.contentType,
                                size: attachment.size,
                                isInline: attachment.isInline,
                                contentId: attachment.contentId,
                                attachmentType: 'file',
                                error: `Failed to get content: ${result.error?.message || 'Unknown error'}`
                            });
                        }
                    });
                } else {
                    // For item attachments (emails/meetings) or when API isn't available, return metadata only
                    resolve({
                        id: attachment.id,
                        name: attachment.name,
                        attachmentType: attachment.attachmentType,
                        size: attachment.size,
                        isInline: attachment.isInline,
                        contentId: attachment.contentId,
                        contentType: attachment.contentType || 'application/octet-stream'
                    });
                }
            });
        });

        const attachments = await Promise.all(attachmentPromises);

        // Create comprehensive email data with enhanced metadata
        const emailData = {
            // Basic email properties
            subject: item.subject || 'No Subject',
            from: {
                displayName: item.from ? item.from.displayName : 'Unknown',
                emailAddress: item.from ? item.from.emailAddress : 'unknown@example.com'
            },
            to: item.to ? item.to.map(r => ({
                displayName: r.displayName || r.emailAddress,
                emailAddress: r.emailAddress
            })) : [],
            cc: item.cc ? item.cc.map(r => ({
                displayName: r.displayName || r.emailAddress,
                emailAddress: r.emailAddress
            })) : [],
            bcc: item.bcc ? item.bcc.map(r => ({
                displayName: r.displayName || r.emailAddress,
                emailAddress: r.emailAddress
            })) : [],
            
            // Timing information
            dateTimeCreated: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString(),
            dateTimeModified: item.dateTimeModified ? item.dateTimeModified.toISOString() : null,
            
            // Content
            bodyHtml: htmlBody,
            attachments: attachments,
            
            // Client-side mode and engine selections
            pdfMode: outputMode,
            outputMode: outputMode,
            pdfEngine: pdfEngine,
            
            // Additional metadata for better processing
            importance: item.importance,
            categories: item.categories || [],
            conversationId: item.conversationId,
            itemType: item.itemType,
            internetMessageId: item.internetMessageId,
            normalizedSubject: item.normalizedSubject,
            
            // Processing hints for server
            processingHints: {
                bodyMode: 'FullBody',
                includeInlineImages: true,
                preserveFormatting: true,
                // Server-facing hints for output behavior
                separateAttachments: separateAttachments,
                splitInlineImages: splitInlineImages,
                mergeAttachments: outputMode === 'full',
                outputMode: outputMode,
                pdfEngine: pdfEngine,
                clientInfo: {
                    platform: 'Outlook',
                    version: Office.context.diagnostics.version,
                    host: Office.context.diagnostics.host
                }
            }
        };

        console.log('Email data prepared:', {
            subject: emailData.subject,
            attachmentCount: attachments.length,
            inlineAttachments: attachments.filter(a => a.isInline).length,
            bodyLength: htmlBody.length,
            pdfMode: outputMode,
            pdfEngine: pdfEngine
        });

        // Send to server
        const response = await fetch('http://localhost:5000/convert-email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(emailData)
        });

        if (!response.ok) {
            throw new Error(`Server error: ${response.status}`);
        }

        console.log('Email data sent successfully to server');
        alert('Email data sent to server successfully!');
        
    } catch (err) {
        console.error('Error sending email data:', err);
        alert('Error: ' + err.message);
    }
}

// Initialize when Office is ready
Office.onReady(() => {
    console.log('Office.js is ready');
    
    // Wait for DOM
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeApp);
    } else {
        initializeApp();
    }
});

// Attach button handler
function initializeApp() {
    const btn = document.getElementById('convertBtn');
    if (btn) {
        btn.onclick = async function() {
            this.disabled = true;
            this.textContent = 'Sending...';
            
            try {
                await sendEmailToServer();
                this.textContent = 'Convert Email to PDF';
            } catch (err) {
                this.textContent = 'Convert Email to PDF';
            } finally {
                this.disabled = false;
            }
        };
    }
}