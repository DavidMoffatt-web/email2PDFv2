const fs = require('fs');
const path = 'src/taskpane/taskpane.html';

let html = fs.readFileSync(path, 'utf8');
const versionRegex = /v(\d+)\.(\d+)\.(\d+)/g;
let match = versionRegex.exec(html);
if (!match) {
  console.error('No version string found!');
  process.exit(1);
}
let [major, minor, patch] = match.slice(1, 4).map(Number);
patch++;
const newVersion = `v${major}.${minor}.${patch}`;
html = html.replace(versionRegex, newVersion);
fs.writeFileSync(path, html);
console.log(`Version bumped to ${newVersion}`);
