const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

const sizes = [256, 48, 32, 16];
const pngPath = path.join(__dirname, 'app-icon.png');
const icoPath = path.join(__dirname, 'app-icon.ico');
const parentIcoPath = path.join(__dirname, '..', 'app-icon.ico');

async function buildIco() {
  const pngBuffers = [];
  for (const s of sizes) {
    const buf = await sharp(pngPath).resize(s, s).png().toBuffer();
    pngBuffers.push({ size: s, data: buf });
  }

  const count = pngBuffers.length;
  const headerSize = 6;
  const entrySize = 16;
  let dataOffset = headerSize + entrySize * count;

  const header = Buffer.alloc(headerSize);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(count, 4);

  const entries = [];
  const images = [];

  for (const { size, data } of pngBuffers) {
    const entry = Buffer.alloc(entrySize);
    entry.writeUInt8(size < 256 ? size : 0, 0);
    entry.writeUInt8(size < 256 ? size : 0, 1);
    entry.writeUInt8(0, 2);
    entry.writeUInt8(0, 3);
    entry.writeUInt16LE(1, 4);
    entry.writeUInt16LE(32, 6);
    entry.writeUInt32LE(data.length, 8);
    entry.writeUInt32LE(dataOffset, 12);
    dataOffset += data.length;
    entries.push(entry);
    images.push(data);
  }

  const ico = Buffer.concat([header, ...entries, ...images]);
  fs.writeFileSync(icoPath, ico);
  fs.writeFileSync(parentIcoPath, ico);
  console.log(`ICO created (${sizes.join(', ')}px): ${icoPath}`);
}

buildIco().catch(console.error);
