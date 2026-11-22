'use strict';
const fs = require('fs');
const path = require('path');

const targets = [
  path.join(__dirname, '..', 'admin', 'main.js'),
  path.join(__dirname, '..', 'worker', 'main.js')
];

const replacements = [
  ['\u00e2\u20ac\u201c', '\u2013'], // â€“ -> – (en dash)
  ['\u00c2\u00b7', '\u00b7'],         // Â· -> · (middle dot)
  ['\u00e2\u2020\u2019', '\u2192'], // â†’ -> →
  ['\u00e2\u20ac\u00a2', '\u2022'], // â€¢ -> •
  ['\u00e2\u2030\u02c6', '\u2248'], // â‰ˆ -> ≈
  ['\u00e2\u2020\u2014', '\u2197'], // â†— -> ↗
  ['\u00e2\u0153\u201c', '\u2713'], // âœ“ -> ✓
  ['\u00c3\u2014', '\u00d7'],         // Ã— -> ×
  ['\u00e2\u20ac\u2122', '\''],     // â€™ -> '
  ['\u00e2\u20ac\u00a6', '\u2026'], // â€¦ -> … (ellipsis)
];

for (const target of targets) {
  if (!fs.existsSync(target)) continue;
  let code = fs.readFileSync(target, 'utf8');
  let replacedTotal = 0;
  for (const [from, to] of replacements) {
    const parts = code.split(from);
    if (parts.length > 1) {
      replacedTotal += parts.length - 1;
      code = parts.join(to);
    }
  }
  console.log(`[${path.basename(target)}] Replaced ${replacedTotal} mojibake occurrences.`);
  fs.writeFileSync(target, code, 'utf8');
}
