const fs = require('fs');
const b64 = fs.readFileSync('assets/shift-desk-mark.png').toString('base64');
fs.writeFileSync('worker/src/logo-base64.js', `export const LOGO_PNG_BASE64 = '${b64}';\n`);
console.log('Saved logo-base64.js, length:', b64.length);
