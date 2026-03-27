const sharp = require('sharp');
const path = require('path');

const sizes = [16, 25, 32, 48, 64, 80, 128];
const assetsDir = path.join(__dirname, 'assets');

// Use the actual OpenClaw favicon.svg (lobster logo)
const fs = require('fs');
function createSvg() {
  return fs.readFileSync(path.join(__dirname, 'assets', 'openclaw-logo.svg'), 'utf8');
}

async function generate() {
  for (const size of sizes) {
    const svg = createSvg();
    const outPath = path.join(assetsDir, `icon-${size}.png`);
    await sharp(Buffer.from(svg)).resize(size, size).png().toFile(outPath);
    console.log(`✅ icon-${size}.png`);
  }
  console.log('Done!');
}

generate().catch(e => console.error(e));
