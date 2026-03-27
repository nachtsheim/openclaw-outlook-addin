const fs = require('fs');
const path = require('path');
const assetsDir = path.join(__dirname, 'assets');

// Minimal valid 1x1 transparent PNG
const tinyPng = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');

const sizes = [16, 25, 32, 48, 64, 80, 128];
for (const size of sizes) {
  const outPath = path.join(assetsDir, `icon-${size}.png`);
  fs.writeFileSync(outPath, tinyPng);
  console.log(`Created ${outPath}`);
}
console.log('Done');
