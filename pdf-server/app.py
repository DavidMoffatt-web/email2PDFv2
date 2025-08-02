from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import weasyprint
import tempfile
import os
import requests
import io
import json
from PyPDF2 import PdfMerger
import base64
import mimetypes
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Gotenberg service URL
GOTENBERG_URL = os.environ.get('GOTENBERG_URL', 'http://gotenberg:3000')

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    try:
        # Check if Gotenberg is available
        response = requests.get(f"{GOTENBERG_URL}/health", timeout=5)
        gotenberg_status = response.status_code == 200
    except:
        gotenberg_status = False
    
    return jsonify({
        'status': 'healthy',
        'gotenberg_available': gotenberg_status,
        'message': 'PDF server is running'
    })

@app.route('/convert', methods=['POST'])
def convert_html_to_pdf():
    """Convert HTML to PDF using WeasyPrint"""
    try:
        data = request.get_json()
        html_content = data.get('html', '')
        
        if not html_content:
            return jsonify({'error': 'No HTML content provided'}), 400
        
        # Create temporary file for the PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            # Use WeasyPrint to convert HTML to PDF
            weasyprint.HTML(string=html_content).write_pdf(temp_file.name)
            
            # Read the PDF file
            with open(temp_file.name, 'rb') as pdf_file:
                pdf_content = pdf_file.read()
            
            # Clean up temporary file
            os.unlink(temp_file.name)
        
        return send_file(
            io.BytesIO(pdf_content),
            as_attachment=True,
            download_name='email.pdf',
            mimetype='application/pdf'
        )
    
    except Exception as e:
        logger.error(f"Error converting HTML to PDF: {str(e)}")
        return jsonify({'error': f'Failed to convert HTML to PDF: {str(e)}'}), 500

def convert_file_to_pdf_with_gotenberg(file_content, filename, content_type):
    """Convert any file to PDF using Gotenberg's LibreOffice route"""
    try:
        # Determine the file extension
        file_ext = Path(filename).suffix.lower()
        
        # Supported file types that LibreOffice can convert
        supported_extensions = {
            '.docx', '.doc', '.odt', '.rtf',  # Word documents
            '.xlsx', '.xls', '.ods',          # Excel spreadsheets
            '.pptx', '.ppt', '.odp',          # PowerPoint presentations
            '.txt', '.csv',                   # Text files
            '.html', '.htm'                   # HTML files
        }
        
        if file_ext not in supported_extensions:
            logger.warning(f"File type {file_ext} not supported for conversion")
            return None
        
        # Prepare the request to Gotenberg
        files = {
            'files': (filename, file_content, content_type)
        }
        
        # Send request to Gotenberg LibreOffice route
        response = requests.post(
            f"{GOTENBERG_URL}/forms/libreoffice/convert",
            files=files,
            timeout=30
        )
        
        if response.status_code == 200:
            logger.info(f"Successfully converted {filename} to PDF")
            return response.content
        else:
            logger.error(f"Gotenberg conversion failed: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        logger.error(f"Error converting file {filename} to PDF: {str(e)}")
        return None

def merge_pdfs_with_gotenberg(pdf_files):
    """Merge multiple PDFs using Gotenberg's merge endpoint"""
    try:
        if len(pdf_files) == 1:
            return pdf_files[0]['content']
        
        # Prepare files for Gotenberg merge endpoint
        files = {}
        for i, pdf_file in enumerate(pdf_files):
            files[f'file{i}.pdf'] = (f'file{i}.pdf', pdf_file['content'], 'application/pdf')
        
        # Send request to Gotenberg merge endpoint
        response = requests.post(
            f"{GOTENBERG_URL}/forms/pdfengines/merge",
            files=files,
            timeout=30
        )
        
        if response.status_code == 200:
            logger.info(f"Successfully merged {len(pdf_files)} PDFs")
            return response.content
        else:
            logger.error(f"Gotenberg merge failed: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        logger.error(f"Error merging PDFs: {str(e)}")
        return None

def fallback_merge_pdfs(pdf_files):
    """Fallback PDF merge using PyPDF2"""
    try:
        merger = PdfMerger()
        
        for pdf_file in pdf_files:
            pdf_stream = io.BytesIO(pdf_file['content'])
            merger.append(pdf_stream)
        
        output_stream = io.BytesIO()
        merger.write(output_stream)
        merger.close()
        
        return output_stream.getvalue()
        
    except Exception as e:
        logger.error(f"Fallback PDF merge failed: {str(e)}")
        return None

@app.route('/convert-with-attachments', methods=['POST'])
def convert_with_attachments():
    """Convert HTML email to PDF and include converted attachments"""
    try:
        data = request.get_json()
        html_content = data.get('html', '')
        attachments = data.get('attachments', [])
        
        if not html_content:
            return jsonify({'error': 'No HTML content provided'}), 400
        
        # Convert main email HTML to PDF using WeasyPrint
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            weasyprint.HTML(string=html_content).write_pdf(temp_file.name)
            
            with open(temp_file.name, 'rb') as pdf_file:
                email_pdf_content = pdf_file.read()
            
            os.unlink(temp_file.name)
        
        # Prepare list of PDFs to merge (starting with email PDF)
        pdfs_to_merge = [{
            'name': 'email.pdf',
            'content': email_pdf_content
        }]
        
        # Convert each attachment to PDF if possible
        converted_attachments = []
        for attachment in attachments:
            try:
                # Decode base64 content
                file_content = base64.b64decode(attachment['content'])
                filename = attachment['name']
                content_type = attachment.get('contentType', 'application/octet-stream')
                
                logger.info(f"Processing attachment: {filename} ({content_type})")
                
                # Try to convert the file to PDF using Gotenberg
                pdf_content = convert_file_to_pdf_with_gotenberg(file_content, filename, content_type)
                
                if pdf_content:
                    pdfs_to_merge.append({
                        'name': f"{Path(filename).stem}_converted.pdf",
                        'content': pdf_content
                    })
                    converted_attachments.append({
                        'name': filename,
                        'converted': True,
                        'size': len(pdf_content)
                    })
                    logger.info(f"Successfully converted {filename} to PDF")
                else:
                    logger.info(f"Could not convert {filename} - will be listed as attachment")
                    converted_attachments.append({
                        'name': filename,
                        'converted': False,
                        'reason': 'Unsupported file type or conversion failed'
                    })
                    
            except Exception as e:
                logger.error(f"Error processing attachment {attachment.get('name', 'unknown')}: {str(e)}")
                converted_attachments.append({
                    'name': attachment.get('name', 'unknown'),
                    'converted': False,
                    'reason': f'Processing error: {str(e)}'
                })
        
        # Merge all PDFs
        if len(pdfs_to_merge) > 1:
            logger.info(f"Merging {len(pdfs_to_merge)} PDFs...")
            
            # Try Gotenberg merge first
            merged_pdf = merge_pdfs_with_gotenberg(pdfs_to_merge)
            
            # Fallback to PyPDF2 if Gotenberg fails
            if not merged_pdf:
                logger.info("Gotenberg merge failed, trying PyPDF2 fallback...")
                merged_pdf = fallback_merge_pdfs(pdfs_to_merge)
            
            if not merged_pdf:
                logger.error("Both merge methods failed, returning email PDF only")
                merged_pdf = email_pdf_content
        else:
            merged_pdf = email_pdf_content
        
        # Return the merged PDF directly as a file
        return send_file(
            io.BytesIO(merged_pdf),
            as_attachment=True,
            download_name='email_with_attachments.pdf',
            mimetype='application/pdf'
        )
        
    except Exception as e:
        logger.error(f"Error in convert_with_attachments: {str(e)}")
        return jsonify({'error': f'Failed to convert email with attachments: {str(e)}'}), 500

@app.route('/convert-with-attachments-debug', methods=['POST'])
def convert_with_attachments_debug():
    """Convert HTML email to PDF and include converted attachments - returns JSON for testing"""
    try:
        data = request.get_json()
        html_content = data.get('html', '')
        attachments = data.get('attachments', [])
        
        if not html_content:
            return jsonify({'error': 'No HTML content provided'}), 400
        
        # Convert main email HTML to PDF using WeasyPrint
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            weasyprint.HTML(string=html_content).write_pdf(temp_file.name)
            
            with open(temp_file.name, 'rb') as pdf_file:
                email_pdf_content = pdf_file.read()
            
            os.unlink(temp_file.name)
        
        # Prepare list of PDFs to merge (starting with email PDF)
        pdfs_to_merge = [{
            'name': 'email.pdf',
            'content': email_pdf_content
        }]
        
        # Convert each attachment to PDF if possible
        converted_attachments = []
        for attachment in attachments:
            try:
                # Decode base64 content
                file_content = base64.b64decode(attachment['content'])
                filename = attachment['name']
                content_type = attachment.get('contentType', 'application/octet-stream')
                
                logger.info(f"Processing attachment: {filename} ({content_type})")
                
                # Try to convert the file to PDF using Gotenberg
                pdf_content = convert_file_to_pdf_with_gotenberg(file_content, filename, content_type)
                
                if pdf_content:
                    pdfs_to_merge.append({
                        'name': f"{Path(filename).stem}_converted.pdf",
                        'content': pdf_content
                    })
                    converted_attachments.append({
                        'name': filename,
                        'converted': True,
                        'size': len(pdf_content)
                    })
                    logger.info(f"Successfully converted {filename} to PDF")
                else:
                    logger.info(f"Could not convert {filename} - will be listed as attachment")
                    converted_attachments.append({
                        'name': filename,
                        'converted': False,
                        'reason': 'Unsupported file type or conversion failed'
                    })
                    
            except Exception as e:
                logger.error(f"Error processing attachment {attachment.get('name', 'unknown')}: {str(e)}")
                converted_attachments.append({
                    'name': attachment.get('name', 'unknown'),
                    'converted': False,
                    'reason': f'Processing error: {str(e)}'
                })
        
        # Merge all PDFs
        if len(pdfs_to_merge) > 1:
            logger.info(f"Merging {len(pdfs_to_merge)} PDFs...")
            
            # Try Gotenberg merge first
            merged_pdf = merge_pdfs_with_gotenberg(pdfs_to_merge)
            
            # Fallback to PyPDF2 if Gotenberg fails
            if not merged_pdf:
                logger.info("Gotenberg merge failed, trying PyPDF2 fallback...")
                merged_pdf = fallback_merge_pdfs(pdfs_to_merge)
            
            if not merged_pdf:
                logger.error("Both merge methods failed, returning email PDF only")
                merged_pdf = email_pdf_content
        else:
            merged_pdf = email_pdf_content
        
        # Return the merged PDF as base64 JSON (for testing)
        response_data = {
            'pdf': base64.b64encode(merged_pdf).decode('utf-8'),
            'attachments_processed': converted_attachments,
            'total_pages_estimated': len(pdfs_to_merge)
        }
        
        return jsonify(response_data)
        
    except Exception as e:
        logger.error(f"Error in convert_with_attachments: {str(e)}")
        return jsonify({'error': f'Failed to convert email with attachments: {str(e)}'}), 500

@app.route('/test-gotenberg', methods=['GET'])
def test_gotenberg():
    """Test endpoint to check Gotenberg connectivity and capabilities"""
    try:
        # Test basic connectivity
        health_response = requests.get(f"{GOTENBERG_URL}/health", timeout=5)
        
        if health_response.status_code != 200:
            return jsonify({
                'error': 'Gotenberg health check failed',
                'status_code': health_response.status_code
            }), 500
        
        # Test a simple conversion
        test_html = """
        <!DOCTYPE html>
        <html>
        <head><title>Test</title></head>
        <body><h1>Gotenberg Test</h1><p>This is a test document.</p></body>
        </html>
        """
        
        files = {
            'files': ('test.html', test_html, 'text/html')
        }
        
        convert_response = requests.post(
            f"{GOTENBERG_URL}/forms/libreoffice/convert",
            files=files,
            timeout=10
        )
        
        if convert_response.status_code == 200:
            return jsonify({
                'status': 'success',
                'message': 'Gotenberg is working correctly',
                'test_pdf_size': len(convert_response.content)
            })
        else:
            return jsonify({
                'error': 'Gotenberg conversion test failed',
                'status_code': convert_response.status_code,
                'response': convert_response.text
            }), 500
            
    except Exception as e:
        return jsonify({
            'error': f'Gotenberg test failed: {str(e)}'
        }), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
