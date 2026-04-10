/**
 * Generate PNG icons from SVG sources for Office Add-in manifest.
 * Uses a canvas-free approach — creates minimal valid PNGs directly.
 * Run: node scripts/generate-icons.js
 */

const fs = require("fs");
const path = require("path");

// Minimal PNG generator for solid-colored icons
// Creates a simple icon with the Streamline brand colors

function createPNG(width, height) {
  // Create raw pixel data (RGBA)
  const pixels = Buffer.alloc(width * height * 4);

  const bgColor = [31, 78, 121, 255]; // #1F4E79
  const barWhite = [255, 255, 255, 230];
  const barBlue = [91, 127, 199, 255]; // #5B7FC7
  const barGray = [255, 255, 255, 153];

  // Fill background
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      // Rounded corners (approximate)
      const r = Math.min(width, height) * 0.125;
      const inCorner = (x < r && y < r && Math.hypot(x - r, y - r) > r) ||
                       (x >= width - r && y < r && Math.hypot(x - (width - r), y - r) > r) ||
                       (x < r && y >= height - r && Math.hypot(x - r, y - (height - r)) > r) ||
                       (x >= width - r && y >= height - r && Math.hypot(x - (width - r), y - (height - r)) > r);

      if (inCorner) {
        pixels[i] = 0; pixels[i+1] = 0; pixels[i+2] = 0; pixels[i+3] = 0;
      } else {
        pixels[i] = bgColor[0]; pixels[i+1] = bgColor[1];
        pixels[i+2] = bgColor[2]; pixels[i+3] = bgColor[3];
      }
    }
  }

  // Draw 3 horizontal bars (Gantt chart representation)
  const barMargin = Math.floor(width * 0.1875);
  const barH = Math.max(Math.floor(height * 0.125), 1);
  const gap = Math.floor(height * 0.25);

  drawBar(pixels, width, barMargin, Math.floor(height * 0.1875), Math.floor(width * 0.5), barH, barWhite);
  drawBar(pixels, width, Math.floor(width * 0.3125), Math.floor(height * 0.1875) + gap, Math.floor(width * 0.4375), barH, barBlue);
  drawBar(pixels, width, barMargin, Math.floor(height * 0.1875) + gap * 2, Math.floor(width * 0.625), barH, barGray);

  return encodePNG(pixels, width, height);
}

function drawBar(pixels, stride, x, y, w, h, color) {
  for (let dy = 0; dy < h; dy++) {
    for (let dx = 0; dx < w; dx++) {
      const i = ((y + dy) * stride + (x + dx)) * 4;
      if (i >= 0 && i < pixels.length - 3) {
        // Alpha blend
        const a = color[3] / 255;
        pixels[i] = Math.round(pixels[i] * (1 - a) + color[0] * a);
        pixels[i+1] = Math.round(pixels[i+1] * (1 - a) + color[1] * a);
        pixels[i+2] = Math.round(pixels[i+2] * (1 - a) + color[2] * a);
        pixels[i+3] = 255;
      }
    }
  }
}

function encodePNG(pixels, width, height) {
  // Minimal PNG encoder (uncompressed)
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type (RGBA)
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace
  const ihdrChunk = makeChunk("IHDR", ihdr);

  // IDAT - raw pixel data with zlib wrapper
  // Each row has a filter byte (0 = None) + RGBA pixels
  const rawData = Buffer.alloc(height * (1 + width * 4));
  for (let y = 0; y < height; y++) {
    rawData[y * (1 + width * 4)] = 0; // filter: None
    pixels.copy(rawData, y * (1 + width * 4) + 1, y * width * 4, (y + 1) * width * 4);
  }

  // Wrap in zlib (deflate with stored blocks)
  const zlibData = zlibStore(rawData);
  const idatChunk = makeChunk("IDAT", zlibData);

  // IEND
  const iendChunk = makeChunk("IEND", Buffer.alloc(0));

  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

function makeChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const typeB = Buffer.from(type, "ascii");
  const crcInput = Buffer.concat([typeB, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(crcInput));
  return Buffer.concat([length, typeB, data, crc]);
}

function zlibStore(data) {
  // zlib header + stored deflate blocks + adler32
  const header = Buffer.from([0x78, 0x01]); // zlib header (deflate, no compression)

  // Split into 65535-byte blocks
  const blocks = [];
  let offset = 0;
  while (offset < data.length) {
    const remaining = data.length - offset;
    const blockSize = Math.min(remaining, 65535);
    const isLast = offset + blockSize >= data.length;

    const blockHeader = Buffer.alloc(5);
    blockHeader[0] = isLast ? 0x01 : 0x00;
    blockHeader.writeUInt16LE(blockSize, 1);
    blockHeader.writeUInt16LE(blockSize ^ 0xFFFF, 3);

    blocks.push(blockHeader);
    blocks.push(data.slice(offset, offset + blockSize));
    offset += blockSize;
  }

  const adler = adler32(data);
  const adlerBuf = Buffer.alloc(4);
  adlerBuf.writeUInt32BE(adler);

  return Buffer.concat([header, ...blocks, adlerBuf]);
}

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function adler32(buf) {
  let a = 1, b = 0;
  for (let i = 0; i < buf.length; i++) {
    a = (a + buf[i]) % 65521;
    b = (b + a) % 65521;
  }
  return ((b << 16) | a) >>> 0;
}

// Generate icons
const sizes = [16, 32, 64, 80, 128];
const assetsDir = path.join(__dirname, "..", "assets");

for (const size of sizes) {
  const png = createPNG(size, size);
  const filePath = path.join(assetsDir, `icon-${size}.png`);
  fs.writeFileSync(filePath, png);
  console.log(`Generated: icon-${size}.png (${png.length} bytes)`);
}

console.log("\nDone. PNG icons generated in assets/");
