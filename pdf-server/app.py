#!/usr/bin/env python3
"""
Lightweight Python PDF Generation Server
Provides high-quality HTML-to-PDF conversion for Outlook add-in
"""

import os
import json
import base64
import tempfile
import logging
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import weasyprint
from weasyprint import HTML, CSS
from werkzeug.exceptions import BadRequest, InternalServerError

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max request size
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'service': 'pdf-generator'
    })

@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    """
    Convert HTML content to PDF
    
    Expected JSON payload:
    {
        "html": "HTML content",
        "css": "Optional CSS content",
        "options": {
            "page_size": "A4",
            "margin": "20mm",
            "orientation": "portrait"
        }
    }
    """
    try:
        # Validate request
        if not request.is_json:
            raise BadRequest("Content-Type must be application/json")
        
        data = request.get_json()
        if not data or 'html' not in data:
            raise BadRequest("Missing 'html' field in request body")
        
        html_content = data['html']
        css_content = data.get('css', '')
        options = data.get('options', {})
        
        logger.info(f"Processing PDF conversion request, HTML length: {len(html_content)}")
        
        # Default options
        page_size = options.get('page_size', 'A4')
        margin = options.get('margin', '20mm')
        orientation = options.get('orientation', 'portrait')
        
        # Enhanced CSS for better PDF rendering
        base_css = f"""
        @page {{
            size: {page_size} {orientation};
            margin: {margin};
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 0;
        }}
        
        /* Email content styling */
        .email-content {{
            max-width: 100%;
            word-wrap: break-word;
        }}
        
        /* Image handling */
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 10px 0;
        }}
        
        /* Table styling */
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
        }}
        
        table, th, td {{
            border: 1px solid #ddd;
        }}
        
        th, td {{
            padding: 8px;
            text-align: left;
        }}
        
        th {{
            background-color: #f5f5f5;
            font-weight: bold;
        }}
        
        /* Lists */
        ul, ol {{
            margin: 10px 0;
            padding-left: 20px;
        }}
        
        li {{
            margin: 5px 0;
        }}
        
        /* Headings */
        h1, h2, h3, h4, h5, h6 {{
            margin: 20px 0 10px 0;
            line-height: 1.2;
        }}
        
        /* Paragraphs */
        p {{
            margin: 10px 0;
        }}
        
        /* Links */
        a {{
            color: #0066cc;
            text-decoration: underline;
        }}
        
        /* Quote blocks */
        blockquote {{
            margin: 15px 0;
            padding: 10px 20px;
            border-left: 4px solid #ddd;
            background-color: #f9f9f9;
        }}
        
        /* Code blocks */
        pre, code {{
            font-family: 'Courier New', monospace;
            background-color: #f5f5f5;
            padding: 2px 4px;
            border-radius: 3px;
        }}
        
        pre {{
            padding: 10px;
            overflow-x: auto;
        }}
        """
        
        # Combine CSS
        final_css = base_css + "\n" + css_content
        
        # Create temporary file for PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            try:
                # Generate PDF using WeasyPrint
                html_doc = HTML(string=html_content)
                css_doc = CSS(string=final_css)
                
                html_doc.write_pdf(
                    temp_pdf.name,
                    stylesheets=[css_doc],
                    optimize_images=True
                )
                
                # Read PDF content
                with open(temp_pdf.name, 'rb') as pdf_file:
                    pdf_content = pdf_file.read()
                
                # Return PDF as base64
                pdf_base64 = base64.b64encode(pdf_content).decode('utf-8')
                
                logger.info(f"PDF generated successfully, size: {len(pdf_content)} bytes")
                
                return jsonify({
                    'success': True,
                    'pdf': pdf_base64,
                    'size': len(pdf_content),
                    'message': 'PDF generated successfully'
                })
                
            finally:
                # Clean up temporary file
                try:
                    os.unlink(temp_pdf.name)
                except OSError:
                    pass
                    
    except BadRequest as e:
        logger.warning(f"Bad request: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400
    
    except Exception as e:
        logger.error(f"Error generating PDF: {str(e)}")
        return jsonify({
            'success': False,
            'error': 'Internal server error during PDF generation'
        }), 500

@app.route('/convert-file', methods=['POST'])
def convert_to_pdf_file():
    """
    Convert HTML content to PDF and return as file download
    Same input as /convert but returns PDF file directly
    """
    try:
        # Validate request
        if not request.is_json:
            raise BadRequest("Content-Type must be application/json")
        
        data = request.get_json()
        if not data or 'html' not in data:
            raise BadRequest("Missing 'html' field in request body")
        
        html_content = data['html']
        css_content = data.get('css', '')
        options = data.get('options', {})
        filename = options.get('filename', 'email.pdf')
        
        logger.info(f"Processing PDF file conversion request")
        
        # Default options
        page_size = options.get('page_size', 'A4')
        margin = options.get('margin', '20mm')
        orientation = options.get('orientation', 'portrait')
        
        # Use same CSS as convert endpoint
        base_css = f"""
        @page {{
            size: {page_size} {orientation};
            margin: {margin};
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            color: #333;
        }}
        img {{ max-width: 100%; height: auto; }}
        table {{ width: 100%; border-collapse: collapse; }}
        table, th, td {{ border: 1px solid #ddd; }}
        th, td {{ padding: 8px; text-align: left; }}
        """
        
        final_css = base_css + "\n" + css_content
        
        # Create temporary file for PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
            try:
                # Generate PDF
                html_doc = HTML(string=html_content)
                css_doc = CSS(string=final_css)
                
                html_doc.write_pdf(
                    temp_pdf.name,
                    stylesheets=[css_doc],
                    optimize_images=True
                )
                
                logger.info(f"PDF file generated successfully")
                
                return send_file(
                    temp_pdf.name,
                    as_attachment=True,
                    download_name=filename,
                    mimetype='application/pdf'
                )
                
            except Exception as e:
                try:
                    os.unlink(temp_pdf.name)
                except OSError:
                    pass
                raise e
                
    except BadRequest as e:
        logger.warning(f"Bad request: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400
    
    except Exception as e:
        logger.error(f"Error generating PDF file: {str(e)}")
        return jsonify({
            'success': False,
            'error': 'Internal server error during PDF generation'
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'false').lower() == 'true'
    
    logger.info(f"Starting PDF server on port {port}, debug={debug}")
    app.run(host='0.0.0.0', port=port, debug=debug)
