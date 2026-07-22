const fs = require('fs');
const transcript = 'C:/Users/jayda/.cursor/projects/c-Users-jayda-OneDrive-Desktop-AI-Agents-A-web-app-for-s-Attempt-1/agent-transcripts/8822d2a3-15ea-4084-ba6d-cb3eb839251e/8822d2a3-15ea-4084-ba6d-cb3eb839251e.jsonl';
const lines = fs.readFileSync(transcript, 'utf8').split(/\r?\n/);
const text = JSON.parse(lines[516]).message.content[0].text;
const i1 = text.indexOf('## 1.');
const i2 = text.indexOf('## 2.');
const i3 = text.indexOf('## 3.');
console.log('===== ## 1. through ## 2. =====\n');
console.log(text.slice(i1, i2));
console.log('\n===== FULL ## 2. through ## 3. =====\n');
console.log(text.slice(i2, i3));
const re = /Slot A/gi;
let m;
const hits = [];
while ((m = re.exec(text))) hits.push(m.index);
console.log('\n===== Slot A hit offsets =====', hits);
for (const i of hits) {
  console.log('\n--- at', i, '---\n');
  console.log(text.slice(Math.max(0, i - 300), Math.min(text.length, i + 1200)));
}
