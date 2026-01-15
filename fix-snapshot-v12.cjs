/**
 * 修复 captureSizeSnapshot 调用，传入目标尺寸
 * 版本升级到 v12
 */
const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'src/taskpane-new.ts');
let content = fs.readFileSync(filePath, 'utf8');

// 1. 更新版本号到 v12
content = content.replace(/v11/g, 'v12');

// 2. 修改导入，添加 cmToPoints
// 检查是否已经导入了 cmToPoints
if (!content.includes('cmToPoints') || content.includes('// import { cmToPoints }')) {
  // 如果没有导入或被注释了，确保导入
  content = content.replace(
    /import \{ setStatus, hasOfficeContext(.*?) \} from "\.\/utils";/,
    'import { setStatus, hasOfficeContext, cmToPoints$1 } from "./utils";'
  );
}

// 3. 修改 onWidthChanged 中的 captureSizeSnapshot 调用
// 原来: await captureSizeSnapshot();
// 改为: await captureSizeSnapshot({ targetWidthPt: cmToPoints(next), targetHeightPt: cmToPoints(Number(targetHeightCm.value)) });
content = content.replace(
  /if \(getResizeScope\(\) === "new"\) \{\s*await captureSizeSnapshot\(\);\s*pasteLog\.debug\("保存宽度时捕获尺寸快照"\);/g,
  `if (getResizeScope() === "new") {
      const heightVal = Number(targetHeightCm.value);
      await captureSizeSnapshot({
        targetWidthPt: cmToPoints(next),
        targetHeightPt: Number.isFinite(heightVal) && heightVal > 0 ? cmToPoints(heightVal) : undefined,
      });
      pasteLog.debug("保存宽度时捕获尺寸快照（含目标尺寸）");`
);

// 4. 修改 onHeightChanged 中的 captureSizeSnapshot 调用
content = content.replace(
  /if \(getResizeScope\(\) === "new"\) \{\s*await captureSizeSnapshot\(\);\s*pasteLog\.debug\("保存高度时捕获尺寸快照"\);/g,
  `if (getResizeScope() === "new") {
      const widthVal = Number(targetWidthCm.value);
      await captureSizeSnapshot({
        targetWidthPt: Number.isFinite(widthVal) && widthVal > 0 ? cmToPoints(widthVal) : undefined,
        targetHeightPt: cmToPoints(next),
      });
      pasteLog.debug("保存高度时捕获尺寸快照（含目标尺寸）");`
);

// 5. 修改 scopeNew change 事件中的 captureSizeSnapshot 调用
content = content.replace(
  /await captureBaseline\(\); \/\/ 捕获基线快照\s*await captureSizeSnapshot\(\); \/\/ 捕获尺寸快照（用于内联图片判断）/g,
  `await captureBaseline(); // 捕获基线快照
        // 捕获尺寸快照（含目标尺寸，用于内联图片判断）
        const widthCm = Number((document.getElementById("targetWidthCm") as HTMLInputElement)?.value);
        const heightCm = Number((document.getElementById("targetHeightCm") as HTMLInputElement)?.value);
        await captureSizeSnapshot({
          targetWidthPt: Number.isFinite(widthCm) && widthCm > 0 ? cmToPoints(widthCm) : undefined,
          targetHeightPt: Number.isFinite(heightCm) && heightCm > 0 ? cmToPoints(heightCm) : undefined,
        });`
);

// 6. 修改 btnApply click 事件中的 captureSizeSnapshot 调用
content = content.replace(
  /await captureBaseline\(\);\s*await captureSizeSnapshot\(\);\s*pasteLog\.info\("内存表已重置，基线和尺寸快照已更新"\);/g,
  `await captureBaseline();
        await captureSizeSnapshot({
          targetWidthPt: Number.isFinite(width) && width > 0 ? cmToPoints(width) : undefined,
          targetHeightPt: Number.isFinite(height) && height > 0 ? cmToPoints(height) : undefined,
        });
        pasteLog.info("内存表已重置，基线和尺寸快照已更新（含目标尺寸）");`
);

fs.writeFileSync(filePath, content, 'utf8');
console.log('已修改 taskpane-new.ts，版本升级到 v12，captureSizeSnapshot 调用已添加目标尺寸参数');
