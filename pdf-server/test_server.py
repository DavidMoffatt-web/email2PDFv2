#!/usr/bin/env python3
"""
Test script for the PDF server
"""

import requests
import json
import base64

# Test HTML content
test_html = """
<!DOCTYPE html>
<html>
<head>
    <title>Test Email</title>
</head>
<body>
    <h1>Test Email Subject</h1>
    <p><strong>From:</strong> test@example.com</p>
    <p><strong>To:</strong> recipient@example.com</p>
    <p><strong>Date:</strong> August 2, 2025</p>
    
    <hr>
    
    <div class="email-content">
        <p>This is a test email with <strong>bold text</strong> and <em>italic text</em>.</p>
        
        <h2>Features to Test:</h2>
        <ul>
            <li>Selectable text ✓</li>
            <li>Proper formatting ✓</li>
            <li>Lists and headings ✓</li>
            <li>Tables (see below) ✓</li>
        </ul>
        
        <table>
            <thead>
                <tr>
                    <th>Feature</th>
                    <th>Status</th>
                    <th>Notes</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Text Selection</td>
                    <td>✅ Working</td>
                    <td>Much better than html2pdf</td>
                </tr>
                <tr>
                    <td>Image Support</td>
                    <td>✅ Working</td>
                    <td>Automatic scaling</td>
                </tr>
                <tr>
                    <td>CSS Styling</td>
                    <td>✅ Working</td>
                    <td>Professional appearance</td>
                </tr>
            </tbody>
        </table>
        
        <blockquote>
            This is a quoted section that should be nicely formatted with proper indentation and styling.
        </blockquote>
        
        <p>Links should work too: <a href="https://example.com">Visit Example.com</a></p>
    </div>
</body>
</html>
"""

def test_server(base_url="http://localhost:5000"):
    """Test the PDF server"""
    
    print("Testing PDF Server...")
    print(f"Base URL: {base_url}")
    
    # Test health endpoint
    try:
        response = requests.get(f"{base_url}/health")
        print(f"✓ Health check: {response.status_code}")
        if response.status_code == 200:
            print(f"  Response: {response.json()}")
    except Exception as e:
        print(f"✗ Health check failed: {e}")
        return
    
    # Test PDF conversion
    try:
        payload = {
            "html": test_html,
            "options": {
                "page_size": "A4",
                "margin": "20mm",
                "orientation": "portrait"
            }
        }
        
        response = requests.post(
            f"{base_url}/convert",
            json=payload,
            headers={"Content-Type": "application/json"}
        )
        
        print(f"✓ PDF conversion: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                pdf_size = result.get('size', 0)
                print(f"  PDF generated successfully!")
                print(f"  Size: {pdf_size:,} bytes ({pdf_size/1024:.1f} KB)")
                
                # Save PDF for manual inspection
                if result.get('pdf'):
                    pdf_data = base64.b64decode(result['pdf'])
                    with open('test_output.pdf', 'wb') as f:
                        f.write(pdf_data)
                    print(f"  ✓ Saved test PDF as 'test_output.pdf'")
            else:
                print(f"  ✗ Error: {result.get('error')}")
        else:
            print(f"  ✗ HTTP Error: {response.text}")
            
    except Exception as e:
        print(f"✗ PDF conversion failed: {e}")

if __name__ == "__main__":
    test_server()
