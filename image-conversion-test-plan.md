# Image Conversion Test Plan for html-to-pdfmake

## Summary
This test plan validates that external images are successfully converted to base64 data URLs when using the html-to-pdfmake PDF engine, while maintaining selectable text and proper formatting.

## Test Setup
1. Open the add-in at https://localhost:3000
2. Set PDF engine to "html-to-pdfmake"
3. Click "Test PDF Engine" button

## Expected Results

### ✅ Font Mapping
- All fonts (Aptos, Calibri, Arial, Times New Roman, Helvetica) should be mapped to Roboto
- Text should remain selectable in the generated PDF
- No font-related errors in console

### ✅ Image Conversion 
- 3 placeholder images from via.placeholder.com should appear in the PDF:
  - Image 1: Blue 150x100 with "Image 1" text
  - Image 2: Orange 150x100 with "Image 2" text  
  - Image 3: Green 200x80 with "Image 3 Wider" text
- Console should show image conversion progress messages
- No CORS or VFS errors related to images

### ✅ HTML Structure
- Headings, lists, tables, and blockquotes should be preserved
- CSS styles should be properly interpreted
- No CSS code appearing as text in the PDF

### ✅ Error Handling
- If any image fails to load, it should be gracefully removed with a comment
- PDF generation should not fail due to image conversion errors
- Timeout protection should prevent hanging on slow images

## Console Log Analysis
Look for these messages in browser console:
- "Converting external images to data URLs..."
- "Found X external images to convert"
- "Converting image: https://via.placeholder.com/..."
- "Successfully converted image: https://via.placeholder.com/..."
- "Image conversion completed"

## Success Criteria
1. PDF downloads successfully without errors
2. PDF contains all 3 test images
3. Text is selectable and properly formatted
4. Console shows successful image conversion messages
5. No JavaScript errors or CORS failures

## Troubleshooting
If images don't appear:
1. Check browser console for CORS or conversion errors
2. Verify via.placeholder.com is accessible
3. Test with alternative image URLs if needed
4. Confirm the convertExternalImagesToDataUrls function is being called

## Implementation Details
- Images are converted using HTML5 Canvas API
- 10-second timeout prevents hanging on unresponsive images
- Failed conversions result in image removal, not PDF failure
- Base64 data URLs are embedded directly in the PDF document definition
