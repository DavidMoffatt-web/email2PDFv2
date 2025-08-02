// Test structure
Office.onReady(() => {
    console.log('Office.js is ready');
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeApp);
    } else {
        initializeApp();
    }
});

function initializeApp() {
    try {
        console.log('Initializing app...');
        
        const btn = document.getElementById('convertBtn');
        if (!btn) {
            console.error('Convert button not found in DOM!');
            return;
        }
        
        btn.onclick = async function() {
            console.log('test');
        };
        
        console.log('Button click handler attached.');
    } catch (err) {
        console.error('Error in initializeApp:', err);
    }
}
