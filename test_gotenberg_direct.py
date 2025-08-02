#!/usr/bin/env python3
"""Test Gotenberg LibreOffice conversion directly."""

import requests
import base64

def test_gotenberg_direct():
    """Test Gotenberg LibreOffice route directly."""
    gotenberg_url = 'http://localhost:3000'
    
    # Create a simple DOCX content (just text that should be convertible)
    simple_doc_content = "Test document content for conversion."
    
    # Test with a simple text file first to see if LibreOffice route works
    print("Testing Gotenberg LibreOffice route with text file...")
    
    try:
        files = {
            'files': ('test.txt', simple_doc_content, 'text/plain')
        }
        
        response = requests.post(
            f"{gotenberg_url}/forms/libreoffice/convert",
            files=files,
            timeout=30
        )
        
        print(f"Text file conversion - Status: {response.status_code}")
        if response.status_code == 200:
            print(f"✓ Text file converted successfully ({len(response.content)} bytes)")
        else:
            print(f"✗ Text file conversion failed: {response.text}")
            
    except Exception as e:
        print(f"✗ Error testing text file: {e}")
    
    # Now test with a real DOCX file
    print("\nTesting Gotenberg LibreOffice route with DOCX...")
    
    # Create a properly formatted minimal DOCX
    docx_content = create_minimal_docx()
    
    try:
        files = {
            'files': ('test.docx', docx_content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        }
        
        response = requests.post(
            f"{gotenberg_url}/forms/libreoffice/convert",
            files=files,
            timeout=30
        )
        
        print(f"DOCX conversion - Status: {response.status_code}")
        if response.status_code == 200:
            print(f"✓ DOCX converted successfully ({len(response.content)} bytes)")
            # Verify it's a PDF
            if response.content.startswith(b'%PDF-'):
                print("✓ Output is valid PDF")
            else:
                print("✗ Output is not a valid PDF")
        else:
            print(f"✗ DOCX conversion failed: {response.text}")
            
    except Exception as e:
        print(f"✗ Error testing DOCX: {e}")

def create_minimal_docx():
    """Create a minimal valid DOCX file."""
    # This is a properly structured minimal DOCX file
    import zipfile
    import io
    
    # Create ZIP structure in memory
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # Add [Content_Types].xml
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
        zip_file.writestr('[Content_Types].xml', content_types)
        
        # Add _rels/.rels
        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
        zip_file.writestr('_rels/.rels', rels)
        
        # Add word/document.xml
        document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>This is a test DOCX document that should be converted to PDF.</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>It contains simple text content for testing conversion.</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>'''
        zip_file.writestr('word/document.xml', document)
    
    zip_buffer.seek(0)
    return zip_buffer.read()

if __name__ == '__main__':
    test_gotenberg_direct()
