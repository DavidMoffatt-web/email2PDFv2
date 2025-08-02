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

                // Get selected mode - engine is now fixed to server-pdf
                const modeSel = document.getElementById('pdfMode');
                const mode = modeSel ? modeSel.value : 'full';
                console.log('PDF Generation Mode:', mode, 'Engine: server-pdf (optimized)');
                
                progressDiv.textContent = 'ðŸ“¨ Reading email content...';

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

                        progressDiv.textContent = `ðŸ“„ Generating PDF (${mode} mode)...`;

                        try {
                            // All modes now use optimized server-pdf engine
                            await createServerPdf(metaHtml, htmlBody, attachments);
                            
                            progressDiv.textContent = 'âœ… PDF generated successfully!';
                            progressDiv.style.color = '#28a745';
                            progressDiv.style.borderColor = '#28a745';
                            progressDiv.style.background = '#d4edda';
                            
                            // Hide progress after 3 seconds
                            setTimeout(() => {
                                if (progressDiv.parentNode) {
                                    progressDiv.parentNode.removeChild(progressDiv);
                                }
                            }, 3000);
                            
                        } catch (err) {
                            console.error('Error in PDF generation:', err);
                            progressDiv.textContent = `âŒ PDF generation failed: ${err.message}`;
                            progressDiv.style.color = '#dc3545';
                            progressDiv.style.borderColor = '#dc3545';
                            progressDiv.style.background = '#f8d7da';
                            
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
                        '<li>âœ… Parallel attachment processing</li>' +
                        '<li>âœ… Cached server health checks</li>' +
                        '<li>âœ… Optimized image conversion (5s timeout)</li>' +
                        '<li>âœ… Real-time progress indicators</li>' +
                        '<li>âœ… Simplified single-engine architecture</li>' +
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

