# Email2PDF Outlook Add-in

A powerful Outlook add-in that converts emails to high-quality PDF documents with multiple generation engines and advanced features.

## 🚀 Features

- ✅ **Multiple PDF Engines**: Choose from 4 different PDF generation methods
- ✅ **Server-Side Processing**: High-quality PDF generation with our lightweight Python server
- ✅ **Selectable Text**: Perfect text selection and copy/paste in generated PDFs
- ✅ **Multiple Modes**: Full PDF, individual PDFs, or individual images
- ✅ **Cross-Platform**: Works on Outlook Web, Desktop (Windows/Mac), and Mobile
- ✅ **Professional Formatting**: Maintains email formatting, tables, images, and styling

## 📋 PDF Engine Options

1. **🚀 Server PDF** (Recommended) - High-quality server-side generation
2. **html2pdf** - Best visual match, but text not selectable
3. **pdf-lib** - Selectable text with basic formatting
4. **html-to-pdfmake** - Selectable text preserving HTML formatting

## 🏗️ Project Structure

```
email2PDFv2/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Main UI
│   │   ├── taskpane.js        # Core functionality
│   │   └── styles.css         # Styling
│   └── manifest.xml           # Add-in manifest
├── pdf-server/                # Lightweight Python PDF server
│   ├── app.py                 # Flask server
│   ├── Dockerfile             # Docker configuration
│   ├── docker-compose.yml     # Easy deployment
│   ├── requirements.txt       # Python dependencies
│   ├── test_server.py         # Server testing
│   ├── start-server.bat       # Windows startup
│   └── start-server.sh        # Linux/Mac startup
├── package.json
├── webpack.config.js
└── README.md
```

## 🚀 Quick Start

### Option 1: Full Setup (Recommended)

1. **Clone and setup the add-in**:
   ```bash
   git clone <repository-url>
   cd email2PDFv2
   npm install
   npm start
   ```

2. **Start the PDF server** (for best quality):
   ```bash
   # Windows
   cd pdf-server
   start-server.bat
   
   # Linux/Mac
   cd pdf-server
   ./start-server.sh
   ```

3. **Load the add-in in Outlook** and select "🚀 Server PDF" for best results

### Option 2: Add-in Only (Client-side PDF)

If you don't want to run the server, the add-in works standalone with JavaScript PDF libraries:

```bash
npm install
npm start
```

Then use html2pdf, pdf-lib, or html-to-pdfmake engines in the add-in.

## 🐳 PDF Server Setup

The PDF server provides the highest quality PDF generation using Python and WeasyPrint.

### Prerequisites

- Docker and Docker Compose
- Or Python 3.11+ for local development

### Starting the Server

**Docker (Recommended)**:
```bash
cd pdf-server
docker-compose up -d
```

**Local Development**:
```bash
cd pdf-server
pip install -r requirements.txt
python app.py
```

### Server Features

- **Lightweight**: ~95MB Docker image
- **Fast**: 500ms-2s PDF generation
- **High Quality**: Professional PDF output with selectable text
- **Health Checks**: Built-in monitoring at `/health`
- **CORS Enabled**: Works with web-based Outlook add-ins
- **Secure**: Runs as non-root user

### Testing the Server

```bash
cd pdf-server
pip install requests  # If not already installed
python test_server.py
```

This generates a `test_output.pdf` to verify quality.

## 🛠️ Development

### Building the Add-in

```bash
npm run build
```

### Development Mode

```bash
npm run dev
```

### Sideloading in Outlook

1. Start the development server: `npm start`
2. In Outlook, go to Add-ins → Manage Add-ins → Upload Manifest
3. Select `src/manifest.xml`

## 🔧 Configuration

### PDF Server Configuration

Environment variables for the server:

```bash
DEBUG=false          # Set to true for development
PORT=5000           # Server port
```

### Add-in Configuration

The add-in automatically detects if the server is running and enables the "Server PDF" option.

## 📊 Performance Comparison

| Engine | Quality | Text Selection | Speed | File Size |
|--------|---------|----------------|-------|-----------|
| 🚀 Server PDF | ⭐⭐⭐⭐⭐ | ✅ Perfect | ⭐⭐⭐⭐ | Optimal |
| html2pdf | ⭐⭐⭐⭐ | ❌ None | ⭐⭐⭐ | Larger |
| pdf-lib | ⭐⭐⭐ | ✅ Good | ⭐⭐⭐⭐⭐ | Small |
| html-to-pdfmake | ⭐⭐⭐ | ✅ Good | ⭐⭐ | Medium |

## 🚀 Production Deployment

### Server Deployment

The PDF server can be deployed anywhere Docker is supported:

```bash
# Build and deploy
docker build -t email2pdf-server ./pdf-server
docker run -d -p 5000:5000 email2pdf-server
```

For production, consider:
- Reverse proxy (nginx/Traefik) for SSL
- Load balancing for multiple instances
- Health check monitoring

### Add-in Deployment

1. Update `manifest.xml` with your production URLs
2. Host the add-in files on a web server with HTTPS
3. Submit to Microsoft AppSource or deploy privately

## 🔍 Troubleshooting

### Common Issues

1. **Server not starting**: Ensure Docker is running and port 5000 is available
2. **PDF generation fails**: Check server logs with `docker-compose logs`
3. **Add-in not loading**: Verify manifest.xml URLs are accessible
4. **CORS errors**: Ensure server is running and CORS is enabled

### Debug Mode

Enable debug mode in the server:
```bash
DEBUG=true docker-compose up
```

## 📈 Roadmap

- [ ] Azure deployment templates
- [ ] Batch processing for multiple emails
- [ ] PDF password protection
- [ ] Custom watermarks and headers
- [ ] Integration with cloud storage (OneDrive, SharePoint)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is proprietary and not licensed for use, modification, or distribution by others.
All rights reserved.