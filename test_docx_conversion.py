#!/usr/bin/env python3
"""Test script to verify DOCX conversion specifically."""

import requests
import json
import base64
import os

def test_docx_conversion():
    """Test DOCX file conversion via the server."""
    server_url = 'http://localhost:5000'
    
    # Test HTML content (simulating email)
    test_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Test Email with DOCX</title>
        <meta charset="UTF-8">
    </head>
    <body>
        <h1>Test Email</h1>
        <p>This email contains a DOCX attachment that should be converted to PDF.</p>
        <p>Date: August 2, 2025</p>
        <p>From: test@example.com</p>
        <p>To: recipient@example.com</p>
    </body>
    </html>
    """
    
    # Create a minimal DOCX file for testing
    # This is a very basic DOCX structure that LibreOffice should be able to handle
    docx_content = """PK\x03\x04\x14\x00\x00\x00\x08\x00\x00\x00!\x00\xb5U+\x17\xf4\x00\x00\x00L\x02\x00\x00\x13\x00\x00\x00[Content_Types].xml\x9c\x91MO\xc30\x10\x86\xef\xf9\n\x83\x11[8\x07\x14D\x85^\x80\x88\x83\x12\x8c\x14\xc7\x06\xd6\xd6\x8e\x93X\xf6\xdf\xce\x9e\xd8\x80\x89\xe3\xc4s\xf2\xcd\xcc\xf7vN\xa6\xfc\xdaD0\xcb\x80\xb8\x13x\xfa\x04J#\xa9\xe1\xa8\x98B\xc7\x04\x8c \x82"\x91\x90f\x0b\x83\x11\xa5q\x06\xd5&\x8d\'\x0e\x9f\x0b\x97\x93\x12\x1a\x0c\x84q\xde\xb5c\x8a\x88L\xd8\x14\xac\xda\x07\xa5\x86(\x81\xa6\xf7\xa9\xd1M\xc7l\x80\x80\xf4Q\x02\x8e2\x02\xeeJ\x06\x8e\x06\xb5m\x9a\xba\x86,\x97\xe9\xe2\xca\xec\x08\xeb4~\xb9\xf7.\x83\xa7\xf0\x9eD\xb8\xfa\x1d\xd5\xddg\xe4\x9c\x07\xdf`B\x1a\xc2q)\xac\xb1\xb8\x03m(\xcb\x9c\x8e\xa5I\x86C\x19\x14\xc7M\xc3L\x93\xb1|\x1a\xc2\xae\x9b\xfc\xd3Wx\xb8\x8f$\xc6\xa3\x96\x85\x03^\x96\x10\x01_\x9b\xdb\xd4`1\xde\x9d\x85\xb5\xc0\xc0\xd6\x82\x94\x0b\x80\x8b\x1f\xd3gZ\x84\x11R"z\xacQ\xa3\x18\x8e\x1e\x1a\xae-\xd8!\x85\xa5\xf4\x8c\x1f\xe4Y\x10\x05\xc4\xcb\xf4\x9c\x97\x04\x1eu\x8b\xcc\xd2\xd4\x9b\xf5\x8a\xef\x91\xf8\x8b\xe8\xa5:\xf5\x9bb\xf7\x85\x9f\xf0{\x1e\x19\x1c\xb5\xd3\x01\xc9\xb4\x1e\x9a2s\xab\x8f\xb4\x1a\xb8\x9c\xa7\x1b\xff\x80\x02\x00\x00PK\x03\x04\x14\x00\x00\x00\x08\x00\x00\x00!\x00\xc1\xba[.\xf0\x00\x00\x00\xa0\x01\x00\x00\x0b\x00\x00\x00_rels/.rels\x9c\x90\xcbn\xc3 \x10D\xef\xf9\n\x83\x0f\x05\xdb\x8a\x8a\xe1\x8a\xbaP\x01\x1e\x08F\xb1\x8eE\xe9\xe6\xdew\xd9\xa9\x08\xfc\x91\xf6\x9a3\xbe\x19\xc3\xc4o\xfeM\x19\xbe"\xf5\x9b\x8f\x05]\x12w\xbf\xd3\x89\x81:\'\xbf\xd0\x97X\xd4\x86B\xc9$T\x04\xc0\x15@\xa3\xbd\x80\xa8\x00\xd4\x82G\x86\x01\x13\x11\xe0\x11\xfe\x1c\x1a\xbb\xd8\x1f\x9a\xdb\xc8\xe6`\xdc\x8b\xbc\xd52\xb4Cp\xdf\xa0\x1a\xa8\xc3\xb8\xdb\x9a\xe5\xf5S\x1e8v\x0c\x9e\xc80|\\8\x0e\xd9\x8cH\xa0y\x14\r\x9d\x1c\x1a\xfbwg2\x01\x9a\x9a\xe8\xa4\x90\xee\xa4\xea\x9b\xbc\xc5>\xb6\x7f\xd8\x96*{\xe2\x96\xe8\xa8k\xfa\x8d\xe1f\xa5y\xe4\x97\x87\x01\xc6\x9b\x99\xb4Q\xb0\x98\x82\xf1\x12\x18\xc2\x89\\\x10\x85\x9eT\xd3\xc3]\xbb\x91Y\xcc\x19\'\x92IG\xd8\x0f\x91J\xe5n\xf0\x06\x04\x00\x00PK\x03\x04\x14\x00\x00\x00\x08\x00\x00\x00!\x00\x98\x8dM\x0e\x8d\x00\x00\x00\xe8\x00\x00\x00\x10\x00\x00\x00word/document.xml\x95\x90\xc1\n\x830\x0c\x86\xef\xf9\n\x83;\x8f\xa0\x07\x8f\x1e]=\x14\x82\x07\x0f\x1eB\x8a\x84$r\xa9\xff\xbe\x8c\xa6:\t\xa5C\x1b\xbe\xf9\xe1\x7f\x92\x86\x11S\x08`0\xe4\x83\x9a\x9b\x94a\x81\x90\xc8\x08\x9c\x86\xa6\xc1\x01\x16\x9b\xbb\xc8r\x94)\x94"C!\x04@\xa0B\n\xd5\x81\xdb\x81n\xa3\x1a\xe1\x8d\xfd\xa1\xa2\x13\xd6\xf0!\xde\x11\xf3\x18\xa3J<D\x88\x94\\\x9a\xcf:\x8e\x9e\x97e\x9e\x1fl\x1d\xa3K\x86\xe3\xca\xd1\xa0|\x94\xd4\x1b\xf3;\x85\x9f\xcey\x94\x08\xe5Bs\x9e\x02\x8b\xa6\xab\xf7\x92\x1e\x1e\x95\x1a\xf5\x8a-\x9d\x98\xd4:\x18\xfd\x02\xe4\xd8\xd3\xe6\x14\xd3\xc5y\x98\x95.Cv]\x14\x82\xa4Y\x1c3;\xe0\x14H\x84\x84\xd4R\x04\'\x88\x00\x92P\x1a\x95\x88\x9c\x94\x88\xb6\x9c\x89\x88[\x1a\xf1\x9e\x04\x02\x00\x00PK\x01\x02\x14\x00\x14\x00\x00\x00\x08\x00\x00\x00!\x00\xb5U+\x17\xf4\x00\x00\x00L\x02\x00\x00\x13\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x80\x01\x00\x00\x00\x00[Content_Types].xmlPK\x01\x02\x14\x00\x14\x00\x00\x00\x08\x00\x00\x00!\x00\xc1\xba[.\xf0\x00\x00\x00\xa0\x01\x00\x00\x0b\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x80\x01%\x01\x00\x00_rels/.relsPK\x01\x02\x14\x00\x14\x00\x00\x00\x08\x00\x00\x00!\x00\x98\x8dM\x0e\x8d\x00\x00\x00\xe8\x00\x00\x00\x10\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x80\x015\x02\x00\x00word/document.xmlPK\x05\x06\x00\x00\x00\x00\x03\x00\x03\x00\x9f\x00\x00\x00\xf2\x02\x00\x00\x00\x00"""
    
    # Convert to base64
    docx_b64 = base64.b64encode(docx_content.encode('latin1')).decode('utf-8')
    
    # Prepare test payload for individual mode
    payload = {
        'html': test_html,
        'mode': 'individual',
        'attachments': [
            {
                'name': 'test_document.docx',
                'content': docx_b64,
                'contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        ],
        'options': {
            'page_size': 'A4',
            'margin': '20mm'
        }
    }
    
    print("Testing DOCX conversion in individual mode...")
    print(f"Sending request to {server_url}/convert-with-attachments")
    print(f"DOCX file size: {len(docx_content)} bytes")
    
    try:
        # Send request
        response = requests.post(
            f'{server_url}/convert-with-attachments',
            json=payload,
            timeout=60  # Longer timeout for DOCX conversion
        )
        
        print(f"Response status: {response.status_code}")
        print(f"Response headers: {dict(response.headers)}")
        
        if response.status_code == 200:
            content_type = response.headers.get('content-type', '')
            
            if 'application/json' in content_type:
                # Individual mode - should receive JSON with download URLs
                response_data = response.json()
                print(f"Response data: {json.dumps(response_data, indent=2)}")
                
                if response_data.get('mode') == 'individual' and 'pdfs' in response_data:
                    print(f"\nReceived {len(response_data['pdfs'])} PDF(s):")
                    
                    for i, pdf_info in enumerate(response_data['pdfs']):
                        print(f"  {i+1}. {pdf_info['filename']} ({pdf_info['size']} bytes)")
                        
                        # Test downloading each PDF
                        download_url = f"{server_url}{pdf_info['download_url']}"
                        pdf_response = requests.get(download_url, timeout=10)
                        
                        if pdf_response.status_code == 200:
                            print(f"     ✓ Downloaded successfully")
                            if pdf_response.content.startswith(b'%PDF-'):
                                print(f"     ✓ Valid PDF format")
                            else:
                                print(f"     ✗ Invalid PDF format")
                        else:
                            print(f"     ✗ Download failed: {pdf_response.status_code}")
                    
                    # Check if DOCX was converted
                    docx_converted = any('docx' in pdf['filename'].lower() or 'document' in pdf['filename'].lower() 
                                       for pdf in response_data['pdfs'])
                    
                    if docx_converted:
                        print(f"\n✓ DOCX file was successfully converted to PDF")
                    else:
                        print(f"\n✗ DOCX file was NOT converted to PDF")
                        print("Only email PDF was created")
                        
                else:
                    print("✗ Unexpected response format")
            else:
                print(f"✗ Unexpected content type: {content_type}")
        else:
            print(f"✗ Request failed: {response.status_code}")
            error_text = response.text
            print(f"Error response: {error_text}")
            
    except Exception as e:
        print(f"✗ Error testing DOCX conversion: {e}")

if __name__ == '__main__':
    test_docx_conversion()
