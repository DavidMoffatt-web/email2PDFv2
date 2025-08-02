#!/usr/bin/env python3
"""Test script to verify individual PDF download functionality."""

import requests
import json
import base64
import os

def test_individual_pdf_creation():
    """Test creating individual PDFs via the server."""
    server_url = 'http://localhost:5000'
    
    # Test HTML content
    test_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Test Email</title>
        <meta charset="UTF-8">
    </head>
    <body>
        <h1>Test Email</h1>
        <p>This is a test email for individual PDF creation.</p>
        <p>Date: January 1, 2024</p>
        <p>From: test@example.com</p>
        <p>To: recipient@example.com</p>
    </body>
    </html>
    """
    
    # Create a simple test attachment (text file)
    test_attachment_content = "This is a test document content.\n\nLine 2\nLine 3"
    test_attachment_b64 = base64.b64encode(test_attachment_content.encode('utf-8')).decode('utf-8')
    
    # Prepare test payload for individual mode
    payload = {
        'html': test_html,
        'mode': 'individual',
        'attachments': [
            {
                'name': 'test_document.txt',
                'content': test_attachment_b64,
                'contentType': 'text/plain'
            }
        ],
        'options': {
            'page_size': 'A4',
            'margin': '20mm'
        }
    }
    
    print("Testing individual PDF creation...")
    print(f"Sending request to {server_url}/convert-with-attachments")
    
    try:
        # Send request
        response = requests.post(
            f'{server_url}/convert-with-attachments',
            json=payload,
            timeout=30
        )
        
        print(f"Response status: {response.status_code}")
        print(f"Response headers: {dict(response.headers)}")
        
        if response.status_code == 200:
            # Should receive JSON with download URLs
            response_data = response.json()
            print(f"Response data: {json.dumps(response_data, indent=2)}")
            
            if response_data.get('mode') == 'individual' and 'pdfs' in response_data:
                print(f"\nSuccess! Received {len(response_data['pdfs'])} PDF download URLs:")
                
                for i, pdf_info in enumerate(response_data['pdfs']):
                    print(f"  {i+1}. {pdf_info['filename']} ({pdf_info['size']} bytes)")
                    print(f"     Download URL: {pdf_info['download_url']}")
                    
                    # Test downloading the PDF
                    download_url = f"{server_url}{pdf_info['download_url']}"
                    print(f"     Testing download from: {download_url}")
                    
                    try:
                        pdf_response = requests.get(download_url, timeout=10)
                        if pdf_response.status_code == 200:
                            pdf_size = len(pdf_response.content)
                            print(f"     ✓ Downloaded successfully ({pdf_size} bytes)")
                            
                            # Verify it's a PDF by checking magic bytes
                            if pdf_response.content.startswith(b'%PDF-'):
                                print(f"     ✓ Valid PDF format")
                            else:
                                print(f"     ✗ Invalid PDF format")
                        else:
                            print(f"     ✗ Download failed: {pdf_response.status_code}")
                    except Exception as e:
                        print(f"     ✗ Download error: {e}")
            else:
                print("✗ Unexpected response format")
        else:
            print(f"✗ Request failed: {response.status_code}")
            print(f"Response: {response.text}")
            
    except Exception as e:
        print(f"✗ Error testing individual PDFs: {e}")

if __name__ == '__main__':
    test_individual_pdf_creation()
