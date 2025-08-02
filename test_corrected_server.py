#!/usr/bin/env python3
"""
Test the corrected server that returns PDF blob directly (not JSON)
"""

import requests
import base64

# Simple test content
test_html = """
<!DOCTYPE html>
<html>
<head><title>Test Email</title></head>
<body>
    <h1>Test Email</h1>
    <p>This is a test email to verify the corrected PDF response format.</p>
</body>
</html>
"""

def test_corrected_server():
    """Test the corrected server that returns PDF blob directly"""
    
    # Test simple conversion first
    print("🧪 Testing corrected server - PDF blob response")
    print("=" * 50)
    
    payload = {"html": test_html}
    
    try:
        response = requests.post(
            'http://localhost:5000/convert',
            json=payload,
            timeout=30
        )
        
        print(f"Simple conversion status: {response.status_code}")
        print(f"Content-Type: {response.headers.get('content-type', 'unknown')}")
        print(f"Content-Length: {len(response.content)} bytes")
        
        if response.status_code == 200:
            # Save the PDF
            with open('corrected_test_simple.pdf', 'wb') as f:
                f.write(response.content)
            print("✅ Simple PDF saved as: corrected_test_simple.pdf")
        else:
            print(f"❌ Simple conversion failed: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Simple conversion error: {str(e)}")
        return False
    
    # Test conversion with attachments
    print("\nTesting attachment conversion...")
    
    test_attachment_content = "This is a test document that will be converted to PDF"
    
    attachments = [
        {
            "name": "test_attachment.txt",
            "content": base64.b64encode(test_attachment_content.encode('utf-8')).decode('utf-8'),
            "contentType": "text/plain"
        }
    ]
    
    payload_with_attachments = {
        "html": test_html,
        "attachments": attachments
    }
    
    try:
        response = requests.post(
            'http://localhost:5000/convert-with-attachments',
            json=payload_with_attachments,
            timeout=60
        )
        
        print(f"Attachment conversion status: {response.status_code}")
        print(f"Content-Type: {response.headers.get('content-type', 'unknown')}")
        print(f"Content-Length: {len(response.content)} bytes")
        
        if response.status_code == 200:
            # Check if response is PDF
            if response.content.startswith(b'%PDF'):
                with open('corrected_test_with_attachments.pdf', 'wb') as f:
                    f.write(response.content)
                print("✅ Attachment PDF saved as: corrected_test_with_attachments.pdf")
                print("🎉 SUCCESS! Server now returns PDF blob directly!")
                return True
            else:
                print("❌ Response is not a valid PDF")
                print(f"First 100 bytes: {response.content[:100]}")
                return False
        else:
            print(f"❌ Attachment conversion failed: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Attachment conversion error: {str(e)}")
        return False

if __name__ == "__main__":
    success = test_corrected_server()
    
    print("\n" + "=" * 50)
    if success:
        print("🏆 FIXED! Server now returns PDF blobs correctly!")
        print("✅ Client should be able to open the PDFs now!")
    else:
        print("❌ Still having issues - check the errors above.")
