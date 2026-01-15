/**
 * 生成简单的 PNG 图标
 * 运行: node scripts/make-icons.cjs
 */

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const assetsDir = path.join(__dirname, '..', 'public', 'assets');

// 确保目录存在
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

// CRC32 表
const crcTable = (function() {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c;
  }
  return table;
})();

function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ data[i]) & 0xFF];
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function createPNG(width, height) {
  // PNG signature
  const signature = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
  
  // IHDR chunk
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);
  ihdrData.writeUInt32BE(height, 4);
  ihdrData.writeUInt8(8, 8);   // bit depth
  ihdrData.writeUInt8(6, 9);   // color type (RGBA)
  ihdrData.writeUInt8(0, 10);  // compression
  ihdrData.writeUInt8(0, 11);  // filter
  ihdrData.writeUInt8(0, 12);  // interlace
  
  const ihdrType = Buffer.from('IHDR');
  const ihdrCrc = Buffer.alloc(4);
  ihdrCrc.writeUInt32BE(crc32(Buffer.concat([ihdrType, ihdrData])), 0);
  
  const ihdrLen = Buffer.alloc(4);
  ihdrLen.writeUInt32BE(13, 0);
  
  const ihdr = Buffer.concat([ihdrLen, ihdrType, ihdrData, ihdrCrc]);
  
  // 创建图像数据 - 橙色背景 + 蓝色指针
  const rawData = [];
  const scale = width / 16;
  
  for (let y = 0; y < height; y++) {
    rawData.push(0); // filter byte
    for (let x = 0; x < width; x++) {
      // 归一化坐标
      const nx = x / scale;
      const ny = y / scale;
      
      // 检查是否在橙色矩形边框内
      const inRectBorder = (
        (nx >= 1 && nx <= 11 && ny >= 2 && ny <= 10) &&
        (nx <= 2 || nx >= 10 || ny <= 3 || ny >= 9)
      );
      
      // 检查是否在蓝色指针内 (简化的三角形检测)
      const inPointer = (
        nx >= 9 && nx <= 14 && ny >= 7 && ny <= 15 &&
        (nx - 9) <= (ny - 7) * 0.7
      );
      
      if (inPointer) {
        // 蓝色 #4A7DC4
        rawData.push(74, 125, 196, 255);
      } else if (inRectBorder) {
        // 橙色 #E67E22
        rawData.push(230, 126, 34, 255);
      } else {
        // 透明
        rawData.push(0, 0, 0, 0);
      }
    }
  }
  
  // 压缩图像数据
  const compressed = zlib.deflateSync(Buffer.from(rawData), { level: 9 });
  
  const idatType = Buffer.from('IDAT');
  const idatCrc = Buffer.alloc(4);
  idatCrc.writeUInt32BE(crc32(Buffer.concat([idatType, compressed])), 0);
  
  const idatLen = Buffer.alloc(4);
  idatLen.writeUInt32BE(compressed.length, 0);
  
  const idat = Buffer.concat([idatLen, idatType, compressed, idatCrc]);
  
  // IEND chunk
  const iendType = Buffer.from('IEND');
  const iendCrc = Buffer.alloc(4);
  iendCrc.writeUInt32BE(crc32(iendType), 0);
  const iend = Buffer.concat([Buffer.from([0, 0, 0, 0]), iendType, iendCrc]);
  
  return Buffer.concat([signature, ihdr, idat, iend]);
}

// 生成各尺寸图标
const sizes = [16, 32, 64, 80, 128];

for (const size of sizes) {
  const png = createPNG(size, size);
  const filename = path.join(assetsDir, `icon-${size}.png`);
  fs.writeFileSync(filename, png);
  console.log(`Created: icon-${size}.png (${png.length} bytes)`);
}

console.log('\\nAll icons created in public/assets/');
