const fs = require('fs');
const path = require('path');

const transcript = String.raw`C:\Users\jayda\.cursor\projects\c-Users-jayda-OneDrive-Desktop-AI-Agents-A-web-app-for-s-Attempt-1\agent-transcripts\8822d2a3-15ea-4084-ba6d-cb3eb839251e\8822d2a3-15ea-4084-ba6d-cb3eb839251e.jsonl`;
const outDir = String.raw`C:\Users\jayda\OneDrive\Desktop\AIAGEN~1\AWEBAP~1`;
const sectionOut = path.join(outDir, '_r6_section2.txt');
const slotAOut = path.join(outDir, '_r6_slotA.txt');

const lines = fs.readFileSync(transcript, 'utf8').split(/\r?\n/);
if (lines.length <= 516) {
  console.error('Not enough lines:', lines.length);
  process.exit(1);
}

const obj = JSON.parse(lines[516]);
const text = obj?.message?.content?.[0]?.text;
if (typeof text !== 'string') {
  console.error('No message.content[0].text at line 516');
  console.error('Keys:', Object.keys(obj || {}));
  process.exit(1);
}

const start = text.indexOf('## 2.');
if (start < 0) {
  console.error('## 2. not found');
  process.exit(1);
}
const after = text.slice(start + 4);
const nextH2 = after.search(/\n## 3\./);
let section;
if (nextH2 >= 0) {
  section = text.slice(start, start + 4 + nextH2);
} else {
  section = text.slice(start, start + 15000);
}

fs.writeFileSync(sectionOut, section, 'utf8');
console.log('Wrote section2:', sectionOut, 'chars:', section.length);

// Slot A / 2.2 / completed / Pivot by Job related content from full text
const keywords = [/Slot A/i, /2\.2/, /§\s*2\.2/, /completed/i, /Pivot by Job/i, /payroll upload/i, /parser/i];
const linesOfText = text.split(/\r?\n/);
const relatedBlocks = [];
const seen = new Set();

// Prefer contiguous sections that mention these topics
function addBlock(label, block) {
  const key = block.slice(0, 200);
  if (seen.has(key)) return;
  seen.add(key);
  relatedBlocks.push('=== ' + label + ' ===\n' + block.trim() + '\n');
}

// Extract §2.2 / ## 2.2 sections more broadly
const sectionPatterns = [
  /(?:^|\n)(#{1,4}\s*2\.2[\s\S]*?)(?=\n#{1,4}\s*\d|\n##\s|\n---\s*$|$)/g,
  /(?:^|\n)((?:#{1,4}\s*)?Slot\s*A[\s\S]*?)(?=\n#{1,4}\s*(?:Slot\s*B|2\.|\d+\.)|\n##\s|\n---\s*$|$)/gi,
];

for (const re of sectionPatterns) {
  let m;
  while ((m = re.exec(text)) !== null) {
    addBlock('section-match', m[1]);
  }
}

// Also grab windows around keyword hits for anything missed
for (let i = 0; i < linesOfText.length; i++) {
  const line = linesOfText[i];
  if (keywords.some((k) => k.test(line))) {
    const from = Math.max(0, i - 5);
    const to = Math.min(linesOfText.length, i + 40);
    addBlock('context@' + i, linesOfText.slice(from, to).join('\n'));
  }
}

// Full Slot A / payroll upload parser requirements: include section 2 content plus any Slot A blocks
let slotAContent = '';
if (relatedBlocks.length) {
  slotAContent = relatedBlocks.join('\n');
} else {
  slotAContent = 'No Slot A / 2.2 matches found beyond section 2.\n\n' + section;
}

// Also append a dedicated search for payroll parser requirements in section 2 and nearby
const payrollIdx = text.search(/payroll|Slot A|Pivot by Job|upload parser/i);
if (payrollIdx >= 0) {
  const chunk = text.slice(Math.max(0, payrollIdx - 500), Math.min(text.length, payrollIdx + 8000));
  slotAContent += '\n\n=== payroll/slotA neighborhood ===\n' + chunk;
}

fs.writeFileSync(slotAOut, slotAContent, 'utf8');
console.log('Wrote slotA:', slotAOut, 'chars:', slotAContent.length);

// Print full extracted content for parent
console.log('\n========== _r6_section2.txt ==========\n');
console.log(section);
console.log('\n========== _r6_slotA.txt ==========\n');
console.log(slotAContent);
