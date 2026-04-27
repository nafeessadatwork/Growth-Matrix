import path from "node:path";
import { fileURLToPath } from "node:url";
import fs from "node:fs";
import { Jimp } from "jimp";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const input = path.resolve(__dirname, "..", "src", "assets", "mindx-logo.png");
const output = path.resolve(__dirname, "..", "src", "assets", "mindx-logo-white.png");

if (!fs.existsSync(input)) {
  throw new Error(`Input not found: ${input}`);
}

const img = await Jimp.read(input);

// Heuristic: turn the gray "MIND" letters to white while preserving warm/orange/red mark + X.
// We treat "gray" as pixels where channels are close and brightness is mid-range.
img.scan(0, 0, img.bitmap.width, img.bitmap.height, function (x, y, idx) {
  const r = this.bitmap.data[idx + 0];
  const g = this.bitmap.data[idx + 1];
  const b = this.bitmap.data[idx + 2];
  const a = this.bitmap.data[idx + 3];
  if (a < 8) return;

  const brightness = (r + g + b) / 3;

  // The "MIND" text in this asset is a bluish-gray. Detect it by being moderately dark,
  // with blue slightly higher than red/green, and avoid touching the warm/orange mark.
  const isBluishText =
    brightness > 40 &&
    brightness < 170 &&
    b > r + 8 &&
    b > g + 6 &&
    r < 140 &&
    g < 160 &&
    b < 200;

  if (!isBluishText) return;

  // Push to white, keep alpha.
  this.bitmap.data[idx + 0] = 255;
  this.bitmap.data[idx + 1] = 255;
  this.bitmap.data[idx + 2] = 255;
});

const buf = await img.getBuffer("image/png");
fs.writeFileSync(output, buf);
console.log(`Wrote ${output}`);

