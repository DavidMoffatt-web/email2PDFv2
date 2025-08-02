#!/usr/bin/env python3
"""
Test the DOCX server endpoint with actual data
"""

import requests
import json

def test_docx_endpoint():
    """Test the DOCX endpoint with real data"""
    
    # Read the test base64 content
    try:
        with open('test_docx_base64.txt', 'r') as f:
            base64_content = f.read().strip()
    except FileNotFoundError:
        print("Please run debug_docx.py first to generate test data")
        return
    
    # Prepare the request
    url = 'http://localhost:5001/test-docx'
    payload = {
        'filename': 'test_document.docx',
        'fileData': base64_content
    }
    
    try:
        print(f"Sending request to {url}")
        print(f"Payload size: {len(json.dumps(payload))} bytes")
        
        response = requests.post(url, json=payload, timeout=30)
        
        print(f"Response status: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            print(f"Success: {result['success']}")
            print(f"Filename: {result['filename']}")
            print(f"Total paragraphs: {result['total_paragraphs']}")
            print(f"Content paragraphs: {result['content_paragraphs']}")
            print("\nExtracted content:")
            
            for item in result['extracted_content']:
                print(f"  {item['index']}. [{item['style']}] ({item['type']}): {item['text'][:50]}...")
                
        else:
            print(f"Error: {response.status_code}")
            print(response.text)
            
    except Exception as e:
        print(f"Request failed: {e}")

if __name__ == "__main__":
    test_docx_endpoint()
