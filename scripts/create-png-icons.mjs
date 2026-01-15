/**
 * 创建简单的 PNG 图标文件
 * 使用纯色方块作为占位图标
 * 运行: node scripts/create-png-icons.mjs
 */

import { writeFileSync, mkdirSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const assetsDir = join(__dirname, '..', 'public', 'assets');

// 确保目录存在
if (!existsSync(assetsDir)) {
  mkdirSync(assetsDir, { recursive: true });
}

// 创建简单的 PNG 文件（1x1 像素，橙色）
// PNG 文件头 + IHDR + IDAT + IEND
function createSimplePNG(width, height) {
  // 这是一个最小的有效 PNG 结构
  // 使用橙色 (#E67E22) 作为填充色
  
  const signature = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
  
  // IHDR chunk
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);   // width
  ihdrData.writeUInt32BE(height, 4);  // height
  ihdrData.writeUInt8(8, 8);          // bit depth
  ihdrData.writeUInt8(2, 9);          // color type (RGB)
  ihdrData.writeUInt8(0, 10);         // compression
  ihdrData.writeUInt8(0, 11);         // filter
  ihdrData.writeUInt8(0, 12);         // interlace
  
  const ihdrCrc = crc32(Buffer.concat([Buffer.from('IHDR'), ihdrData]));
  const ihdr = Buffer.concat([
    Buffer.from([0, 0, 0, 13]),  // length
    Buffer.from('IHDR'),
    ihdrData,
    ihdrCrc
  ]);
  
  // IDAT chunk - 简化的图像数据
  // 创建橙色填充的图像
  const rawData = [];
  for (let y = 0; y < height; y++) {
    rawData.push(0); // filter byte
    for (let x = 0; x < width; x++) {
      // 橙色 RGB: 230, 126, 34
      rawData.push(230, 126, 34);
    }
  }
  
  // 使用 zlib 压缩
  const zlib = await import('zlib');
  const compressed = zlib.deflateSync(Buffer.from(rawData));
  
  const idatCrc = crc32(Buffer.concat([Buffer.from('IDAT'), compressed]));
  const idatLen = Buffer.alloc(4);
  idatLen.writeUInt32BE(compressed.length, 0);
  const idat = Buffer.concat([
    idatLen,
    Buffer.from('IDAT'),
    compressed,
    idatCrc
  ]);
  
  // IEND chunk
  const iendCrc = crc32(Buffer.from('IEND'));
  const iend = Buffer.concat([
    Buffer.from([0, 0, 0, 0]),
    Buffer.from('IEND'),
    iendCrc
  ]);
  
  return Buffer.concat([signature, ihdr, idat, iend]);
}

// CRC32 计算
function crc32(data) {
  let crc = 0xFFFFFFFF;
  const table = makeCrcTable();
  
  for (let i = 0; i < data.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ data[i]) & 0xFF];
  }
  
  crc = (crc ^ 0xFFFFFFFF) >>> 0;
  const buf = Buffer.alloc(4);
  buf.writeUInt32BE(crc, 0);
  return buf;
}

function makeCrcTable() {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c;
  }
  return table;
}

// 图标尺寸
const sizes = [16, 32, 80, 128];

async function main() {
  for (const size of sizes) {
    const png = await createSimplePNG(size, size);
    const filename = join(assetsDir, `icon-${size}.png`);
    writeFileSync(filename, png);
    console.log(`Created: icon-${size}.png`);
  }
  console.log('Done! Icons created in public/assets/');
}

main().catch(console.error);
