console.log('taskpane.js loaded - Server metadata sender');

// Simple function to send email data to server
async function sendEmailToServer() {
    console.log('Sending email data to server');
    
    try {
        // Get the current email
        const item = Office.context.mailbox.item;
        if (!item) {
            throw new Error('No email item found');
        }
        
        // Get email body
        const htmlBody = await new Promise((resolve, reject) => {
            item.body.getAsync(Office.CoercionType.Html, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error('Failed to get email body'));
                }
            });
        });
        
        // Create email data object for server
        const emailData = {
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
            dateTimeCreated: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : new Date().toISOString(),
            bodyHtml: htmlBody,
            attachments: []
        };
        
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