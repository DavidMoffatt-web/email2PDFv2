#!/usr/bin/env python3
"""
Test script to verify universal file-to-PDF conversion using our new system
"""

import requests
import base64
import json
import tempfile
import os

# Test data - a simple DOCX content (simulated)
test_html = """
<!DOCTYPE html>
<html>
<head><title>Email Test</title></head>
<body>
    <h1>Test Email</h1>
    <p>This is a test email with an attachment.</p>
    <p>The attachment should be converted to PDF and merged with this email.</p>
</body>
</html>
"""

# Simple text content that could represent a converted file
test_docx_content = """
Test Document Content
====================

This is a test document that would normally be a DOCX file.
It contains:
- Sample text
- Multiple paragraphs
- Basic formatting

This content should be converted to PDF by Gotenberg and merged with the email PDF.
"""

def test_conversion_with_attachments():
    """Test the /convert-with-attachments endpoint"""
    
    # Simulate attachment data
    attachments = [
        {
            "name": "test_document.txt",  # Using txt since it's supported by LibreOffice
            "content": base64.b64encode(test_docx_content.encode('utf-8')).decode('utf-8'),
            "contentType": "text/plain"
        }
    ]
    
    payload = {
        "html": test_html,
        "attachments": attachments
    }
    
    print("Testing universal file-to-PDF conversion...")
    print(f"Email HTML length: {len(test_html)} characters")
    print(f"Number of attachments: {len(attachments)}")
    print(f"Attachment: {attachments[0]['name']} ({attachments[0]['contentType']})")
    
    try:
        response = requests.post(
            'http://localhost:5000/convert-with-attachments',
            json=payload,
            timeout=30
        )
        
        print(f"Response status: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ SUCCESS!")
            print(f"  - PDF generated successfully")
            print(f"  - PDF size: {len(base64.b64decode(result['pdf']))} bytes")
            print(f"  - Estimated total pages: {result.get('total_pages_estimated', 'unknown')}")
            print(f"  - Attachments processed: {len(result.get('attachments_processed', []))}")
            
            for attachment in result.get('attachments_processed', []):
                status = "✅ Converted" if attachment['converted'] else "❌ Not converted"
                print(f"    - {attachment['name']}: {status}")
                if not attachment['converted']:
                    print(f"      Reason: {attachment.get('reason', 'Unknown')}")
            
            # Save the PDF for inspection
            pdf_content = base64.b64decode(result['pdf'])
            with open('test_output.pdf', 'wb') as f:
                f.write(pdf_content)
            print(f"  - PDF saved as: test_output.pdf")
            
            return True
        else:
            print(f"❌ FAILED: {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        return False

def test_simple_conversion():
    """Test the /convert endpoint (email only)"""
    
    payload = {"html": test_html}
    
    print("\nTesting simple HTML-to-PDF conversion...")
    
    try:
        response = requests.post(
            'http://localhost:5000/convert',
            json=payload,
            timeout=30
        )
        
        print(f"Response status: {response.status_code}")
        
        if response.status_code == 200:
            print("✅ Simple conversion successful!")
            print(f"  - PDF size: {len(response.content)} bytes")
            
            # Save the PDF
            with open('test_simple_output.pdf', 'wb') as f:
                f.write(response.content)
            print(f"  - PDF saved as: test_simple_output.pdf")
            
            return True
        else:
            print(f"❌ FAILED: {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        return False

if __name__ == "__main__":
    print("🧪 Testing Universal File-to-PDF Conversion System")
    print("=" * 50)
    
    # Test simple conversion first
    simple_success = test_simple_conversion()
    
    # Test conversion with attachments
    attachment_success = test_conversion_with_attachments()
    
    print("\n" + "=" * 50)
    print("📊 Test Results:")
    print(f"  - Simple conversion: {'✅ PASS' if simple_success else '❌ FAIL'}")
    print(f"  - Attachment conversion: {'✅ PASS' if attachment_success else '❌ FAIL'}")
    
    if simple_success and attachment_success:
        print("🎉 All tests passed! Universal file conversion is working!")
    else:
        print("⚠️  Some tests failed. Check the logs above for details.")
