#!/usr/bin/env python3
"""Test script to verify merged PDF functionality still works."""

import requests
import json
import base64

def test_merged_pdf_creation():
    """Test creating merged PDF via the server."""
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
        <p>This is a test email for merged PDF creation.</p>
        <p>Date: January 1, 2024</p>
        <p>From: test@example.com</p>
        <p>To: recipient@example.com</p>
    </body>
    </html>
    """
    
    # Create a simple test attachment (text file)
    test_attachment_content = "This is a test document content.\n\nLine 2\nLine 3"
    test_attachment_b64 = base64.b64encode(test_attachment_content.encode('utf-8')).decode('utf-8')
    
    # Prepare test payload for full (merged) mode
    payload = {
        'html': test_html,
        'mode': 'full',
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
    
    print("Testing merged PDF creation...")
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
            content_type = response.headers.get('content-type', '')
            if 'application/pdf' in content_type:
                pdf_size = len(response.content)
                print(f"✓ Received merged PDF ({pdf_size} bytes)")
                
                # Verify it's a valid PDF
                if response.content.startswith(b'%PDF-'):
                    print(f"✓ Valid PDF format")
                else:
                    print(f"✗ Invalid PDF format")
            else:
                print(f"✗ Unexpected content type: {content_type}")
                print(f"Response text: {response.text[:500]}")
        else:
            print(f"✗ Request failed: {response.status_code}")
            print(f"Response: {response.text}")
            
    except Exception as e:
        print(f"✗ Error testing merged PDF: {e}")

if __name__ == '__main__':
    test_merged_pdf_creation()
