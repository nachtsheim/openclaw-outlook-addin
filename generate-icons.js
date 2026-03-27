const sharp = require('sharp');
const path = require('path');

const sizes = [16, 25, 32, 48, 64, 80, 128];
const assetsDir = path.join(__dirname, 'assets');

// OpenClaw icon - blue rounded square with white chat-bubble + claw
function createSvg(size) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 128 128">
    <defs>
      <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" style="stop-color:#0078d4"/>
        <stop offset="100%" style="stop-color:#005a9e"/>
      </linearGradient>
    </defs>
    <rect width="128" height="128" rx="24" fill="url(#bg)"/>
    <!-- Chat bubble -->
    <rect x="24" y="28" width="80" height="56" rx="12" fill="white" opacity="0.95"/>
    <polygon points="36,84 48,84 36,98" fill="white" opacity="0.95"/>
    <!-- Claw symbol inside bubble -->
    <path d="M48,48 C44,40 50,34 56,38 L58,42 M80,48 C84,40 78,34 72,38 L70,42 M54,46 C52,54 58,62 64,62 C70,62 76,54 74,46" 
          stroke="#0078d4" stroke-width="4.5" stroke-linecap="round" fill="none"/>
    <circle cx="56" cy="52" r="2.5" fill="#0078d4"/>
    <circle cx="72" cy="52" r="2.5" fill="#0078d4"/>
  </svg>`;
}

async function generate() {
  for (const size of sizes) {
    const svg = createSvg(size);
    const outPath = path.join(assetsDir, `icon-${size}.png`);
    await sharp(Buffer.from(svg)).resize(size, size).png().toFile(outPath);
    console.log(`✅ icon-${size}.png`);
  }
  console.log('Done!');
}

generate().catch(e => console.error(e));
