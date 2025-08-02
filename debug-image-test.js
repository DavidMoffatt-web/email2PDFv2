// Quick test to verify imagesByReference functionality
const testHtml = `
<div>
    <h2>Image Test</h2>
    <p>Testing image handling with html-to-pdfmake:</p>
    <img src="https://via.placeholder.com/150x100/0066cc/ffffff?text=Test+Image" alt="Test Image">
</div>
`;

console.log('Testing imagesByReference with html-to-pdfmake...');

// Test with imagesByReference
const result = htmlToPdfmake(testHtml, { imagesByReference: true });
console.log('Result with imagesByReference:', result);

if (result.images) {
    console.log('Images found:', Object.keys(result.images));
    console.log('Image URLs:', Object.values(result.images));
} else {
    console.log('No images object found in result');
}

// Test without imagesByReference
const resultWithoutRef = htmlToPdfmake(testHtml, { imagesByReference: false });
console.log('Result without imagesByReference:', resultWithoutRef);

// Save this as debug-image-test.js for browser console testing
