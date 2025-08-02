# Lightweight PDF Generation Server

A minimal Python Flask server that provides high-quality HTML-to-PDF conversion for the Outlook Email2PDF add-in.

## Features

- ✅ **Lightweight**: ~100MB Docker image with minimal dependencies
- ✅ **High Quality**: Uses WeasyPrint for professional PDF output
- ✅ **Selectable Text**: Perfect text selection and copy/paste
- ✅ **CSS Support**: Full CSS styling including fonts, tables, images
- ✅ **CORS Enabled**: Works with web-based Outlook add-ins
- ✅ **Health Checks**: Built-in monitoring and health endpoints
- ✅ **Security**: Runs as non-root user in container

## Quick Start

### Option 1: Docker Compose (Recommended)

```bash
# Start the server
docker-compose up -d

# View logs
docker-compose logs -f

# Stop the server
docker-compose down
```

### Option 2: Docker Run

```bash
# Build the image
docker build -t email2pdf-server .

# Run the container
docker run -d -p 5000:5000 --name pdf-server email2pdf-server

# View logs
docker logs -f pdf-server

# Stop the container
docker stop pdf-server && docker rm pdf-server
```

### Option 3: Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the server
python app.py
```

## Server Endpoints

### Health Check
```
GET /health
```
Returns server status and timestamp.

### Convert to PDF (Base64)
```
POST /convert
Content-Type: application/json

{
  "html": "<html>...</html>",
  "css": "body { font-family: Arial; }",  // Optional
  "options": {                           // Optional
    "page_size": "A4",                   // A4, Letter, Legal
    "margin": "20mm",                    // CSS margin value
    "orientation": "portrait"            // portrait, landscape
  }
}
```

Response:
```json
{
  "success": true,
  "pdf": "base64-encoded-pdf-data",
  "size": 1234,
  "message": "PDF generated successfully"
}
```

### Convert to PDF (File Download)
```
POST /convert-file
Content-Type: application/json

{
  "html": "<html>...</html>",
  "options": {
    "filename": "email.pdf"             // Optional filename
  }
}
```

Returns PDF file for direct download.

## Testing

Test the server with the included test script:

```bash
# Make sure server is running first
docker-compose up -d

# Install requests library for testing
pip install requests

# Run test
python test_server.py
```

This will generate a `test_output.pdf` file to verify the PDF quality.

## Integration with Outlook Add-in

Update your Outlook add-in JavaScript to use this server instead of client-side PDF generation:

```javascript
async function generatePDFViaServer(htmlContent, options = {}) {
    try {
        const response = await fetch('http://localhost:5000/convert', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                html: htmlContent,
                options: {
                    page_size: 'A4',
                    margin: '20mm',
                    orientation: 'portrait',
                    ...options
                }
            })
        });

        const result = await response.json();
        
        if (result.success) {
            // Convert base64 to blob
            const pdfData = atob(result.pdf);
            const pdfArray = new Uint8Array(pdfData.length);
            for (let i = 0; i < pdfData.length; i++) {
                pdfArray[i] = pdfData.charCodeAt(i);
            }
            
            const pdfBlob = new Blob([pdfArray], { type: 'application/pdf' });
            return pdfBlob;
        } else {
            throw new Error(result.error || 'PDF generation failed');
        }
    } catch (error) {
        console.error('Server PDF generation error:', error);
        throw error;
    }
}
```

## Production Deployment

For production, consider:

1. **Environment Variables**:
   ```bash
   DEBUG=false
   PORT=5000
   ```

2. **Reverse Proxy**: Use nginx or Traefik for SSL termination

3. **Resource Limits**: The Docker Compose file includes sensible defaults

4. **Scaling**: Can be easily scaled horizontally with multiple containers

## Performance

- **Memory Usage**: ~256-512MB per container
- **Startup Time**: ~5-10 seconds
- **PDF Generation**: ~500ms-2s depending on content complexity
- **Concurrent Requests**: Handles 10-50 concurrent requests efficiently

## Why This Approach?

1. **Better PDF Quality**: WeasyPrint produces superior PDFs compared to browser-based libraries
2. **Selectable Text**: Perfect text selection and copy/paste functionality
3. **Image Handling**: Excellent support for embedded images and complex layouts
4. **CSS Support**: Full CSS3 support including print media queries
5. **Lightweight**: Much smaller than .NET alternatives
6. **Scalable**: Easy to deploy and scale in any environment
7. **Cross-Platform**: Works with all Outlook clients (web, desktop, mobile)

## Dependencies

- **Flask**: Lightweight web framework
- **WeasyPrint**: High-quality HTML/CSS to PDF converter
- **System Libraries**: Cairo, Pango for text rendering (included in Docker image)

Total Docker image size: ~95MB (compressed)
