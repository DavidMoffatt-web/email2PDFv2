from flask import Flask, request, jsonify, send_file, make_response
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
import zipfile
import time
import re
from urllib.parse import urlparse

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Gotenberg service URL
GOTENBERG_URL = os.environ.get('GOTENBERG_URL', 'http://gotenberg:3000')

# Temporary storage for individual PDFs
temp_pdfs = {}

def cleanup_old_pdfs():
    """Remove PDFs older than 1 hour to prevent memory leaks"""
    current_time = time.time()
    expired_ids = []
    
    for pdf_id, pdf_data in temp_pdfs.items():
        # Handle both 'timestamp' and 'created_at' keys for backward compatibility
        pdf_time = pdf_data.get('timestamp', pdf_data.get('created_at', 0))
        if current_time - pdf_time > 3600:  # 1 hour
            expired_ids.append(pdf_id)
    
    for pdf_id in expired_ids:
        del temp_pdfs[pdf_id]
        logger.info(f"Cleaned up expired PDF: {pdf_id}")

def generate_pdf_id():
    """Generate a unique PDF ID"""
    return f"pdf_{int(time.time() * 1000)}_{len(temp_pdfs)}"

def extract_images_from_html(html_content):
    """Extract embedded images from HTML content"""
    try:
        logger.info("Extracting embedded images from HTML content...")
        images = []
        
        # Pattern to match img tags with src attributes (data URLs and regular URLs)
        img_pattern = r'<img[^>]*src=["\']([^"\']+)["\'][^>]*>'
        img_matches = re.findall(img_pattern, html_content, re.IGNORECASE)
        
        for i, src in enumerate(img_matches):
            try:
                if src.startswith('data:'):
                    # Handle data URLs (base64 embedded images)
                    logger.info(f"Found embedded data URL image {i+1}")
                    
                    # Extract the data URL components
                    data_url_pattern = r'data:([^;]+);base64,(.+)'
                    match = re.match(data_url_pattern, src)
                    
                    if match:
                        mime_type = match.group(1)
                        base64_data = match.group(2)
                        
                        try:
                            # Decode base64 data
                            image_data = base64.b64decode(base64_data)
                            
                            # Determine file extension from MIME type
                            extension = {
                                'image/jpeg': '.jpg',
                                'image/jpg': '.jpg',
                                'image/png': '.png',
                                'image/gif': '.gif',
                                'image/bmp': '.bmp',
                                'image/webp': '.webp'
                            }.get(mime_type, '.jpg')
                            
                            filename = f"embedded_image_{i+1}{extension}"
                            
                            images.append({
                                'filename': filename,
                                'data': image_data,
                                'mime_type': mime_type,
                                'size': len(image_data)
                            })
                            
                            logger.info(f"✓ Extracted embedded image: {filename} ({len(image_data)} bytes, {mime_type})")
                            
                        except Exception as e:
                            logger.warning(f"Failed to decode base64 image {i+1}: {str(e)}")
                            
                elif src.startswith('cid:'):
                    # Handle CID references (Content-ID attachments)
                    logger.info(f"Found CID reference: {src} - skipping (should be handled as attachment)")
                    
                else:
                    # Handle external URLs - we'll skip these for now since they're not embedded
                    logger.info(f"Found external image URL: {src[:100]}... - skipping external images")
                    
            except Exception as e:
                logger.warning(f"Error processing image {i+1}: {str(e)}")
                
        logger.info(f"Extracted {len(images)} embedded images from HTML")
        return images
        
    except Exception as e:
        logger.error(f"Error extracting images from HTML: {str(e)}")
        return []

def convert_image_to_pdf_with_gotenberg(image_data, filename, mime_type):
    """Convert an image to PDF using Gotenberg's Chromium route"""
    try:
        logger.info(f"Converting image {filename} to PDF using Gotenberg")
        
        # Handle different input types - image_data could be binary data or already base64
        if isinstance(image_data, bytes):
            # Binary image data - encode to base64
            base64_image = base64.b64encode(image_data).decode('utf-8')
        else:
            # Assume it's already base64 encoded
            base64_image = image_data
        
        # If mime_type is not provided or doesn't look like an image type, try to detect it
        if not mime_type or not mime_type.startswith('image/'):
            # Try to detect from filename extension
            file_ext = Path(filename).suffix.lower()
            mime_type = {
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg', 
                '.png': 'image/png',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp',
                '.webp': 'image/webp'
            }.get(file_ext, 'image/jpeg')  # Default to JPEG if unknown
            
        logger.info(f"Using MIME type: {mime_type} for image {filename}")
        
        # Create HTML wrapper for the image
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                @page {{
                    size: A4;
                    margin: 0;
                }}
                body {{
                    margin: 0;
                    padding: 0;
                    width: 100vw;
                    height: 100vh;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                }}
                img {{
                    max-width: 100vw;
                    max-height: 100vh;
                    width: auto;
                    height: auto;
                    object-fit: contain;
                }}
            </style>
        </head>
        <body>
            <img src="data:{mime_type};base64,{base64_image}" alt="{filename}" />
        </body>
        </html>
        """
        
        # Use Gotenberg to convert HTML to PDF
        files = {
            'files': ('index.html', html_content, 'text/html')
        }
        
        data = {
            'paperWidth': '8.5',
            'paperHeight': '11',
            'marginTop': '0.5',
            'marginBottom': '0.5',
            'marginLeft': '0.5',
            'marginRight': '0.5',
            'printBackground': 'true',
            'preferCSSPageSize': 'false'
        }
        
        response = requests.post(
            f"{GOTENBERG_URL}/forms/chromium/convert/html",
            files=files,
            data=data,
            timeout=30
        )
        
        if response.status_code == 200:
            logger.info(f"✓ Successfully converted image {filename} to PDF ({len(response.content)} bytes)")
            return response.content
        else:
            logger.error(f"✗ Failed to convert image {filename} to PDF: {response.status_code}")
            return None
            
    except Exception as e:
        logger.error(f"✗ Exception converting image {filename} to PDF: {str(e)}")
        return None

def resolve_cid_images_in_html(html_content, attachments):
    """Replace CID references in HTML with base64 data URLs"""
    try:
        logger.info("Resolving CID image references in HTML...")
        
        # Find all CID references in the HTML
        cid_pattern = r'src=["\']cid:([^"\']+)["\']'
        cid_matches = re.findall(cid_pattern, html_content, re.IGNORECASE)
        
        resolved_count = 0
        for cid in cid_matches:
            logger.info(f"Attempting to resolve CID: {cid}")
            
            # Look for matching attachment
            for attachment in attachments:
                attachment_name = attachment.get('name', '')
                content_type = attachment.get('contentType', '')
                
                # Check if this attachment matches the CID or is an image
                if (cid in attachment_name or 
                    content_type.startswith('image/') or
                    attachment_name.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'))):
                    
                    try:
                        # Get the image data
                        image_data = attachment.get('content', '')
                        if not image_data:
                            continue
                            
                        # If content_type is not set, try to detect it
                        if not content_type or not content_type.startswith('image/'):
                            file_ext = Path(attachment_name).suffix.lower()
                            content_type = {
                                '.jpg': 'image/jpeg',
                                '.jpeg': 'image/jpeg',
                                '.png': 'image/png',
                                '.gif': 'image/gif',
                                '.bmp': 'image/bmp',
                                '.webp': 'image/webp'
                            }.get(file_ext, 'image/jpeg')
                        
                        # Create data URL
                        data_url = f"data:{content_type};base64,{image_data}"
                        
                        # Replace the CID reference with the data URL
                        old_src = f'src="cid:{cid}"'
                        new_src = f'src="{data_url}"'
                        html_content = html_content.replace(old_src, new_src)
                        
                        old_src_single = f"src='cid:{cid}'"
                        new_src_single = f"src='{data_url}'"
                        html_content = html_content.replace(old_src_single, new_src_single)
                        
                        logger.info(f"✓ Resolved CID {cid} to data URL using attachment {attachment_name}")
                        resolved_count += 1
                        break
                        
                    except Exception as e:
                        logger.warning(f"Error resolving CID {cid} with attachment {attachment_name}: {str(e)}")
        
        logger.info(f"Resolved {resolved_count} out of {len(cid_matches)} CID references")
        return html_content
        
    except Exception as e:
        logger.error(f"Error resolving CID images: {str(e)}")
        return html_content

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

def convert_html_to_pdf_with_gotenberg(html_content):
    """Convert HTML to PDF using Gotenberg's Chromium route with print-optimized settings"""
    try:
        logger.info(f"Converting HTML email to PDF using Gotenberg ({len(html_content)} characters)")
        
        # Enhance HTML content with print-specific CSS and settings
        enhanced_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                @media print {{
                    body {{ 
                        font-family: Arial, sans-serif;
                        font-size: 12pt;
                        line-height: 1.6;
                        color: #000;
                        background: white;
                        margin: 0;
                        padding: 20px;
                    }}
                    .no-print {{ display: none !important; }}
                    img {{ max-width: 100%; height: auto; }}
                    table {{ 
                        border-collapse: collapse; 
                        width: 100%; 
                        border: none !important;
                        margin: 1em 0;
                    }}
                    th, td {{ 
                        border: none !important; 
                        padding: 12px 8px; 
                        vertical-align: top;
                    }}
                    h1, h2, h3, h4, h5, h6 {{ 
                        color: #000; 
                        page-break-after: avoid;
                        margin: 1.2em 0 0.8em 0;
                        font-weight: bold;
                    }}
                    p {{ margin: 1em 0; }}
                    blockquote {{ 
                        margin: 1.5em 0; 
                        padding-left: 1em; 
                        border-left: none;
                    }}
                    /* Remove any default borders from email content */
                    * {{ border: none !important; }}
                    /* Preserve background colors but remove borders */
                    [style*="border"] {{ border: none !important; }}
                }}
                
                /* Apply same styles to screen for consistency */
                body {{ 
                    font-family: Arial, sans-serif;
                    font-size: 12pt;
                    line-height: 1.6;
                    color: #000;
                    background: white;
                    margin: 0;
                    padding: 20px;
                    max-width: 8.5in;
                }}
                img {{ max-width: 100%; height: auto; }}
                table {{ 
                    border-collapse: collapse; 
                    width: 100%; 
                    border: none !important;
                    margin: 1em 0;
                }}
                th, td {{ 
                    border: none !important; 
                    padding: 12px 8px; 
                    vertical-align: top;
                }}
                h1, h2, h3, h4, h5, h6 {{ 
                    color: #000; 
                    margin: 1.2em 0 0.8em 0;
                    font-weight: bold;
                }}
                p {{ margin: 1em 0; }}
                blockquote {{ 
                    margin: 1.5em 0; 
                    padding-left: 1em; 
                    border-left: none;
                }}
                /* Remove any borders from email content */
                * {{ border: none !important; }}
                [style*="border"] {{ border: none !important; }}
            </style>
        </head>
        <body>
            {html_content}
        </body>
        </html>
        """
        
        # Prepare the HTML file for Gotenberg - must be named 'index.html'
        files = {
            'files': ('index.html', enhanced_html, 'text/html')
        }
        
        # Form data for print-like settings
        data = {
            'paperWidth': '8.5',      # US Letter width in inches
            'paperHeight': '11',      # US Letter height in inches
            'marginTop': '0.5',       # Top margin in inches
            'marginBottom': '0.5',    # Bottom margin in inches
            'marginLeft': '0.5',      # Left margin in inches
            'marginRight': '0.5',     # Right margin in inches
            'printBackground': 'true', # Include background colors/images
            'preferCSSPageSize': 'false', # Use our paper size settings
            'emulateMediaType': 'print'   # Force print media type
        }
        
        # Send request to Gotenberg Chromium route for HTML conversion
        response = requests.post(
            f"{GOTENBERG_URL}/forms/chromium/convert/html",
            files=files,
            data=data,
            timeout=30
        )
        
        if response.status_code == 200:
            pdf_size = len(response.content)
            logger.info(f"✓ Successfully converted HTML email to PDF with print settings ({pdf_size} bytes)")
            return response.content
        else:
            logger.error(f"✗ Gotenberg HTML conversion failed: {response.status_code}")
            logger.error(f"Response text: {response.text[:500]}")
            return None
            
    except Exception as e:
        logger.error(f"✗ Exception during HTML conversion: {str(e)}")
        return None

def convert_html_to_pdf_with_weasyprint(html_content):
    """Fallback: Convert HTML to PDF using WeasyPrint"""
    try:
        logger.info("Using WeasyPrint fallback for HTML conversion")
        
        # Create temporary file for the PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            # Use WeasyPrint to convert HTML to PDF
            weasyprint.HTML(string=html_content).write_pdf(temp_file.name)
            
            # Read the PDF file
            with open(temp_file.name, 'rb') as pdf_file:
                pdf_content = pdf_file.read()
            
            # Clean up temporary file
            os.unlink(temp_file.name)
            
        logger.info(f"✓ WeasyPrint conversion successful ({len(pdf_content)} bytes)")
        return pdf_content
            
    except Exception as e:
        logger.error(f"✗ WeasyPrint conversion failed: {str(e)}")
        return None

def convert_file_to_pdf_with_gotenberg(file_content, filename, content_type):
    """Convert any file to PDF using Gotenberg's LibreOffice route"""
    try:
        logger.info(f"Starting conversion of {filename} ({len(file_content)} bytes, type: {content_type})")
        
        # Determine the file extension
        file_ext = Path(filename).suffix.lower()
        logger.info(f"File extension detected: {file_ext}")
        
        # Supported file types that LibreOffice can convert
        supported_extensions = {
            '.docx', '.doc', '.odt', '.rtf',  # Word documents
            '.xlsx', '.xls', '.ods',          # Excel spreadsheets
            '.pptx', '.ppt', '.odp',          # PowerPoint presentations
            '.txt', '.csv',                   # Text files
            '.html', '.htm'                   # HTML files
        }
        
        # Image file types that we can convert directly to PDF
        image_extensions = {
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'
        }
        
        if file_ext in image_extensions:
            logger.info(f"Image file {file_ext} detected - converting directly to PDF using image conversion")
            # Use our image-to-PDF conversion for image files
            return convert_image_to_pdf_with_gotenberg(file_content, filename, content_type)
        elif file_ext not in supported_extensions:
            logger.warning(f"File type {file_ext} not supported for conversion")
            return None
        
        logger.info(f"File type {file_ext} is supported, proceeding with Gotenberg conversion...")
        
        # Prepare the request to Gotenberg
        files = {
            'files': (filename, file_content, content_type)
        }
        
        logger.info(f"Sending {filename} to Gotenberg at {GOTENBERG_URL}/forms/libreoffice/convert")
        
        # Send request to Gotenberg LibreOffice route
        response = requests.post(
            f"{GOTENBERG_URL}/forms/libreoffice/convert",
            files=files,
            timeout=30
        )
        
        logger.info(f"Gotenberg response for {filename}: Status {response.status_code}")
        
        if response.status_code == 200:
            pdf_size = len(response.content)
            logger.info(f"✓ Successfully converted {filename} to PDF ({pdf_size} bytes)")
            return response.content
        else:
            logger.error(f"✗ Gotenberg conversion failed for {filename}: {response.status_code}")
            logger.error(f"Response headers: {dict(response.headers)}")
            logger.error(f"Response text: {response.text[:500]}")  # First 500 chars of error
            return None
            
    except Exception as e:
        logger.error(f"✗ Exception during conversion of {filename}: {str(e)}")
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

@app.route('/convert', methods=['POST'])
def convert():
    """Convert HTML email to PDF (simple conversion without attachments)"""
    logger.info("=== EMAIL CONVERSION REQUEST RECEIVED ===")
    try:
        data = request.get_json()
        html_content = data.get('html', '')
        
        if not html_content:
            logger.error("No HTML content received")
            return jsonify({'success': False, 'error': 'No HTML content provided'}), 400
        
        logger.info(f"Converting HTML content ({len(html_content)} characters)")
        
        # Try Gotenberg first for better HTML rendering
        pdf_content = convert_html_to_pdf_with_gotenberg(html_content)
        
        # Fallback to WeasyPrint if Gotenberg fails
        if pdf_content is None:
            logger.warning("Gotenberg failed, trying WeasyPrint fallback")
            pdf_content = convert_html_to_pdf_with_weasyprint(html_content)
        
        if pdf_content is None:
            raise Exception("Both Gotenberg and WeasyPrint conversion failed")
        
        logger.info(f"✓ PDF conversion successful ({len(pdf_content)} bytes)")
        
        # Return the PDF as bytes
        response = make_response(pdf_content)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = 'attachment; filename="email.pdf"'
        
        return response
        
    except Exception as e:
        logger.error(f"✗ Conversion failed: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/convert-with-attachments', methods=['POST'])
def convert_with_attachments():
    """Convert HTML email to PDF and include converted attachments"""
    try:
        data = request.get_json()
        html_content = data.get('html', '')
        attachments = data.get('attachments', [])
        mode = data.get('mode', 'full')  # 'full' or 'individual'
        extract_images = data.get('extractImages', False)  # Whether to extract embedded images
        
        if not html_content:
            return jsonify({'error': 'No HTML content provided'}), 400
        
        # Resolve CID image references in the HTML before converting to PDF
        if attachments:
            html_content = resolve_cid_images_in_html(html_content, attachments)
        
        # Convert main email HTML to PDF using Gotenberg (with WeasyPrint fallback)
        logger.info("Converting email HTML to PDF...")
        email_pdf_content = convert_html_to_pdf_with_gotenberg(html_content)
        
        if email_pdf_content is None:
            logger.warning("Gotenberg failed for email HTML, trying WeasyPrint fallback")
            email_pdf_content = convert_html_to_pdf_with_weasyprint(html_content)
        
        if email_pdf_content is None:
            raise Exception("Failed to convert email HTML to PDF with both methods")
        
        logger.info(f"✓ Email HTML converted to PDF ({len(email_pdf_content)} bytes)")
        
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
                
                logger.info(f"Processing attachment: {filename} ({content_type}, {len(file_content)} bytes)")
                
                # Try to convert the file to PDF using Gotenberg
                pdf_content = convert_file_to_pdf_with_gotenberg(file_content, filename, content_type)
                
                if pdf_content:
                    converted_pdf_name = f"{Path(filename).stem}_converted.pdf"
                    pdfs_to_merge.append({
                        'name': converted_pdf_name,
                        'content': pdf_content
                    })
                    converted_attachments.append({
                        'name': filename,
                        'converted': True,
                        'size': len(pdf_content)
                    })
                    logger.info(f"✓ Successfully converted {filename} to PDF ({len(pdf_content)} bytes)")
                else:
                    logger.warning(f"✗ Could not convert {filename} - conversion failed")
                    converted_attachments.append({
                        'name': filename,
                        'converted': False,
                        'reason': 'Conversion failed - unsupported file type or processing error'
                    })
                    
            except Exception as e:
                logger.error(f"Error processing attachment {attachment.get('name', 'unknown')}: {str(e)}")
                converted_attachments.append({
                    'name': attachment.get('name', 'unknown'),
                    'converted': False,
                    'reason': f'Processing error: {str(e)}'
                })
        
        # Extract and convert embedded images if requested
        image_pdfs = []
        if extract_images:
            logger.info("Image extraction requested - scanning HTML for embedded images...")
            extracted_images = extract_images_from_html(html_content)
            
            # Also look for CID references that match attachments
            cid_pattern = r'cid:([^"\'>\s]+)'
            cid_matches = re.findall(cid_pattern, html_content, re.IGNORECASE)
            
            for cid in cid_matches:
                logger.info(f"Found CID reference: {cid}")
                # Look for matching attachment by CID or similar filename patterns
                for attachment in attachments:
                    attachment_name = attachment.get('name', '')
                    # Check if the attachment name contains the CID or is an image
                    if (cid in attachment_name or 
                        attachment_name.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'))):
                        try:
                            logger.info(f"Found matching CID attachment: {attachment_name}")
                            file_content = base64.b64decode(attachment['content'])
                            content_type = attachment.get('contentType', 'image/jpeg')
                            
                            # Add this as an extracted image for individual PDF creation
                            extracted_images.append({
                                'filename': attachment_name,
                                'data': file_content,
                                'mime_type': content_type,
                                'size': len(file_content),
                                'is_cid': True,
                                'cid': cid
                            })
                            logger.info(f"Added CID image {attachment_name} for individual PDF conversion")
                            break
                        except Exception as e:
                            logger.warning(f"Error processing CID attachment {attachment_name}: {str(e)}")
            
            for image_info in extracted_images:
                try:
                    logger.info(f"Converting extracted image {image_info['filename']} to PDF...")
                    image_pdf_content = convert_image_to_pdf_with_gotenberg(
                        image_info['data'], 
                        image_info['filename'], 
                        image_info['mime_type']
                    )
                    
                    if image_pdf_content:
                        # Use a cleaner name for the PDF
                        base_name = Path(image_info['filename']).stem
                        image_pdf_name = f"{base_name}.pdf"
                        pdfs_to_merge.append({
                            'name': image_pdf_name,
                            'content': image_pdf_content
                        })
                        image_pdfs.append({
                            'name': image_info['filename'],
                            'pdf_name': image_pdf_name,
                            'converted': True,
                            'size': len(image_pdf_content),
                            'is_cid': image_info.get('is_cid', False)
                        })
                        logger.info(f"✓ Successfully converted image {image_info['filename']} to PDF")
                    else:
                        logger.warning(f"✗ Failed to convert image {image_info['filename']} to PDF")
                        image_pdfs.append({
                            'name': image_info['filename'],
                            'converted': False,
                            'reason': 'Image to PDF conversion failed'
                        })
                        
                except Exception as e:
                    logger.error(f"Error converting image {image_info['filename']}: {str(e)}")
                    image_pdfs.append({
                        'name': image_info['filename'],
                        'converted': False,
                        'reason': f'Processing error: {str(e)}'
                    })
            
            logger.info(f"Image extraction complete. Converted {len([p for p in image_pdfs if p['converted']])} out of {len(extracted_images)} images to PDF")
        
        logger.info(f"Attachment processing complete. Total PDFs ready: {len(pdfs_to_merge)} (1 email + {len(pdfs_to_merge)-1} converted files)")
        
        # Handle different modes
        if mode == 'individual':
            # Individual mode: prepare separate PDFs and return download info
            logger.info(f"Preparing individual PDFs with {len(pdfs_to_merge)} files...")
            
            # Store PDFs temporarily with unique IDs for download
            pdf_downloads = []
            
            # Add email PDF
            email_id = f"email_{int(time.time())}"
            temp_pdfs[email_id] = {
                'content': email_pdf_content,
                'filename': 'email.pdf',
                'created_at': time.time()
            }
            pdf_downloads.append({
                'id': email_id,
                'filename': 'email.pdf',
                'size': len(email_pdf_content),
                'download_url': f'/download-pdf/{email_id}'
            })
            
            # Add each converted attachment PDF
            for i, pdf_file in enumerate(pdfs_to_merge[1:], 1):  # Skip the first one (email)
                pdf_id = f"attachment_{i}_{int(time.time())}"
                temp_pdfs[pdf_id] = {
                    'content': pdf_file['content'],
                    'filename': pdf_file['name'],
                    'created_at': time.time()
                }
                pdf_downloads.append({
                    'id': pdf_id,
                    'filename': pdf_file['name'],
                    'size': len(pdf_file['content']),
                    'download_url': f'/download-pdf/{pdf_id}'
                })
            
            # Return JSON with download information
            return jsonify({
                'mode': 'individual',
                'pdfs': pdf_downloads,
                'total_count': len(pdf_downloads),
                'message': 'Individual PDFs prepared for download'
            })
            
        else:
            # Full mode: merge all PDFs into one
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
        
        # Convert main email HTML to PDF using Gotenberg (with WeasyPrint fallback)
        logger.info("Converting email HTML to PDF...")
        email_pdf_content = convert_html_to_pdf_with_gotenberg(html_content)
        
        if email_pdf_content is None:
            logger.warning("Gotenberg failed for email HTML, trying WeasyPrint fallback")
            email_pdf_content = convert_html_to_pdf_with_weasyprint(html_content)
        
        if email_pdf_content is None:
            raise Exception("Failed to convert email HTML to PDF with both methods")
        
        logger.info(f"✓ Email HTML converted to PDF ({len(email_pdf_content)} bytes)")
        
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

@app.route('/download-pdf/<pdf_id>', methods=['GET'])
def download_pdf(pdf_id):
    """Download a specific PDF by ID"""
    if pdf_id not in temp_pdfs:
        return jsonify({'error': 'PDF not found or expired'}), 404
    
    pdf_data = temp_pdfs[pdf_id]
    pdf_content = pdf_data['content']
    filename = pdf_data['filename']
    
    # Create a BytesIO object with the PDF content
    pdf_io = io.BytesIO(pdf_content)
    
    return send_file(
        pdf_io,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
