/**
 * 生成 PNG 图标的脚本
 * 需要先安装 canvas: npm install canvas
 * 运行: node scripts/generate-icons.mjs
 */

import { createCanvas } from 'canvas';
import { writeFileSync, mkdirSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const assetsDir = join(__dirname, '..', 'public', 'assets');

// 确保目录存在
if (!existsSync(assetsDir)) {
  mkdirSync(assetsDir, { recursive: true });
}

// 图标尺寸
const sizes = [16, 32, 80, 128];

// 绘制图标
function drawIcon(ctx, size) {
  const scale = size / 16;
  
  // 清除背景（透明）
  ctx.clearRect(0, 0, size, size);
  
  // 橙色虚线矩形框
  ctx.strokeStyle = '#E67E22';
  ctx.lineWidth = 1.5 * scale;
  ctx.setLineDash([2 * scale, 1 * scale]);
  ctx.beginPath();
  ctx.roundRect(1 * scale, 2 * scale, 10 * scale, 8 * scale, 1 * scale);
  ctx.stroke();
  
  // 蓝色鼠标指针
  ctx.fillStyle = '#4A7DC4';
  ctx.setLineDash([]);
  ctx.beginPath();
  ctx.moveTo(9 * scale, 7 * scale);
  ctx.lineTo(9 * scale, 14 * scale);
  ctx.lineTo(11 * scale, 12 * scale);
  ctx.lineTo(13 * scale, 15 * scale);
  ctx.lineTo(14 * scale, 14 * scale);
  ctx.lineTo(12 * scale, 11 * scale);
  ctx.lineTo(14 * scale, 10 * scale);
  ctx.closePath();
  ctx.fill();
}

// 生成各尺寸图标
for (const size of sizes) {
  const canvas = createCanvas(size, size);
  const ctx = canvas.getContext('2d');
  
  drawIcon(ctx, size);
  
  const buffer = canvas.toBuffer('image/png');
  const filename = join(assetsDir, `icon-${size}.png`);
  writeFileSync(filename, buffer);
  console.log(`Generated: ${filename}`);
}

console.log('All icons generated successfully!');
