const fs = require('fs');

function checkBracketBalance(filename) {
    const content = fs.readFileSync(filename, 'utf8');
    const lines = content.split('\n');
    let depth = 0;
    let brackets = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        for (let j = 0; j < line.length; j++) {
            const char = line[j];
            if (char === '{') {
                depth++;
                brackets.push({ type: 'open', line: i + 1, col: j + 1, depth });
            } else if (char === '}') {
                if (depth <= 0) {
                    console.log(`Extra closing brace at line ${i + 1}, col ${j + 1}`);
                    return false;
                }
                depth--;
                brackets.push({ type: 'close', line: i + 1, col: j + 1, depth });
            }
        }
    }
    
    if (depth > 0) {
        console.log(`Missing ${depth} closing brace(s)`);
        // Show the last few open braces
        const openBraces = brackets.filter(b => b.type === 'open').slice(-depth);
        console.log('Unmatched opening braces:');
        openBraces.forEach(b => console.log(`  Line ${b.line}, col ${b.col}`));
        return false;
    } else if (depth < 0) {
        console.log(`${-depth} extra closing brace(s)`);
        return false;
    }
    
    console.log('Brackets are balanced!');
    return true;
}

checkBracketBalance('src/taskpane/taskpane.js');
