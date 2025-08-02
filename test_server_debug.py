#!/usr/bin/env python3
"""Test the exact server workflow with debugging."""

import requests
import json
import base64
import os
from pathlib import Path

def test_server_workflow_debug():
    """Test the exact workflow that happens in the server with debugging."""
    server_url = 'http://localhost:5000'
    gotenberg_url = 'http://localhost:3000'
    
    # Create a proper DOCX file for testing
    docx_content = create_minimal_valid_docx()
    docx_b64 = base64.b64encode(docx_content).decode('utf-8')
    
    print(f"Created DOCX file: {len(docx_content)} bytes")
    print(f"Base64 encoded size: {len(docx_b64)} characters")
    
    # Test the base64 decoding (what server does)
    print("\n--- Testing Base64 Decoding ---")
    try:
        decoded_content = base64.b64decode(docx_b64)
        print(f"✓ Base64 decoded successfully: {len(decoded_content)} bytes")
        
        # Verify it's the same as original
        if decoded_content == docx_content:
            print("✓ Decoded content matches original")
        else:
            print("✗ Decoded content differs from original")
            
    except Exception as e:
        print(f"✗ Base64 decoding failed: {e}")
        return
    
    # Test direct Gotenberg call with decoded content (what server does)
    print("\n--- Testing Direct Gotenberg Call ---")
    try:
        files = {
            'files': ('test_document.docx', decoded_content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        }
        
        response = requests.post(
            f"{gotenberg_url}/forms/libreoffice/convert",
            files=files,
            timeout=30
        )
        
        print(f"Gotenberg response status: {response.status_code}")
        if response.status_code == 200:
            print(f"✓ Gotenberg conversion successful: {len(response.content)} bytes")
            if response.content.startswith(b'%PDF-'):
                print("✓ Output is valid PDF")
            else:
                print("✗ Output is not valid PDF")
        else:
            print(f"✗ Gotenberg conversion failed: {response.text}")
            
    except Exception as e:
        print(f"✗ Error testing Gotenberg directly: {e}")
        return
    
    # Test full server workflow
    print("\n--- Testing Full Server Workflow ---")
    
    test_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Debug Test Email</title>
        <meta charset="UTF-8">
    </head>
    <body>
        <h1>Debug Test Email</h1>
        <p>Testing DOCX conversion through full server workflow.</p>
    </body>
    </html>
    """
    
    payload = {
        'html': test_html,
        'mode': 'individual',
        'attachments': [
            {
                'name': 'test_document.docx',
                'content': docx_b64,
                'contentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }
        ]
    }
    
    try:
        response = requests.post(
            f'{server_url}/convert-with-attachments',
            json=payload,
            timeout=60
        )
        
        print(f"Server response status: {response.status_code}")
        
        if response.status_code == 200:
            if 'application/json' in response.headers.get('content-type', ''):
                response_data = response.json()
                print(f"Server returned {len(response_data.get('pdfs', []))} PDFs")
                
                # Check if DOCX was converted
                for pdf in response_data.get('pdfs', []):
                    print(f"  - {pdf['filename']} ({pdf['size']} bytes)")
                    
                docx_found = any('docx' in pdf['filename'].lower() or 'document' in pdf['filename'].lower() 
                               for pdf in response_data.get('pdfs', []))
                
                if docx_found:
                    print("✓ DOCX was converted by server")
                else:
                    print("✗ DOCX was NOT converted by server")
            else:
                print("✗ Server returned non-JSON response")
        else:
            print(f"✗ Server request failed: {response.text}")
            
    except Exception as e:
        print(f"✗ Error testing server workflow: {e}")

def create_minimal_valid_docx():
    """Create a minimal but completely valid DOCX file."""
    import zipfile
    import io
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as zip_file:
        # [Content_Types].xml
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
        zip_file.writestr('[Content_Types].xml', content_types)
        
        # _rels/.rels
        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
        zip_file.writestr('_rels/.rels', rels)
        
        # word/document.xml - More realistic content
        document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>DOCX Conversion Test Document</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>This document contains multiple paragraphs to test the conversion process.</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Line 1: This is the first line of content.</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Line 2: This is the second line of content.</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Line 3: This document should convert successfully to PDF.</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>'''
        zip_file.writestr('word/document.xml', document)
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

if __name__ == '__main__':
    test_server_workflow_debug()
