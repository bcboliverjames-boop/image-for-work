/**
 * Excel 粘贴检测器 v3
 * 修复：等比例缩放 + 快速路径（直接调整选中的图片）
 */

import { ResizeSettings } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { saveSetting, getStringSetting } from "../settings";

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[excelPasteDetector]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[excelPasteDetector]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[excelPasteDetector]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[excelPasteDetector]", msg, data ?? ""),
  };
}

// ============ 尺寸快照 ============

interface ExcelSizeSnapshot {
  widths: number[];
  heights: number[];
  timestamp: number;
}

const EXCEL_SIZE_SNAPSHOT_KEY = "pasteWidth_excelSizeSnapshot";

function saveExcelSizeSnapshot(snapshot: ExcelSizeSnapshot): void {
  saveSetting(EXCEL_SIZE_SNAPSHOT_KEY, JSON.stringify(snapshot));
}

function getExcelSizeSnapshot(): ExcelSizeSnapshot | null {
  const raw = getStringSetting(EXCEL_SIZE_SNAPSHOT_KEY, "");
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed.widths) && Array.isArray(parsed.heights)) {
      return parsed as ExcelSizeSnapshot;
    }
    return null;
  } catch {
    return null;
  }
}

function isSizeInExcelSnapshot(width: number, height: number, snapshot: ExcelSizeSnapshot, tolerance: number = 1): boolean {
  const w = Math.round(width);
  const h = Math.round(height);
  const widthInSnapshot = snapshot.widths.some(sw => Math.abs(sw - w) <= tolerance);
  const heightInSnapshot = snapshot.heights.some(sh => Math.abs(sh - h) <= tolerance);
  return widthInSnapshot && heightInSnapshot;
}

function isImageShape(shape: Excel.Shape): boolean {
  const typeString = String(shape.type || "").toLowerCase();
  return typeString.includes("image") || typeString.includes("picture");
}

// ============ Shape ID 内存表 ============

const processedShapeIds = new Set<string>();
let baselineShapeIds = new Set<string>();

function recordShapeId(id: string): void {
  processedShapeIds.add(id);
  baselineShapeIds.add(id);
}

function hasShapeId(id: string): boolean {
  return processedShapeIds.has(id);
}

function clearShapeIds(): void {
  processedShapeIds.clear();
  baselineShapeIds.clear();
}

// ============ 状态管理 ============

interface ExcelPasteState {
  lastShapeCount: number;
  processing: boolean;
  fallbackTimer: number | null;
}

const state: ExcelPasteState = {
  lastShapeCount: 0,
  processing: false,
  fallbackTimer: null,
};

export async function initExcelPasteBaseline(): Promise<void> {
  const log = getLog();
  try {
    clearShapeIds();
    
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await ctx.sync();
      
      state.lastShapeCount = shapes.items?.length ?? 0;
      
      const widthSet = new Set<number>();
      const heightSet = new Set<number>();
      
      for (const shape of shapes.items || []) {
        shape.load(["id", "type", "width", "height"]);
      }
      await ctx.sync();
      
      for (const shape of shapes.items || []) {
        if (shape.id) {
          recordShapeId(shape.id);
        }
        if (isImageShape(shape)) {
          const w = Math.round(shape.width);
          const h = Math.round(shape.height);
          if (w > 0) widthSet.add(w);
          if (h > 0) heightSet.add(h);
        }
      }
      
      const snapshot: ExcelSizeSnapshot = {
        widths: Array.from(widthSet),
        heights: Array.from(heightSet),
        timestamp: Date.now(),
      };
      saveExcelSizeSnapshot(snapshot);
      baselineShapeIds = new Set(processedShapeIds);
      
      log.info("initExcelPasteBaseline", { 
        shapeCount: state.lastShapeCount,
        registeredIds: processedShapeIds.size,
        snapshotWidths: snapshot.widths.length,
        snapshotHeights: snapshot.heights.length,
      });
    });
  } catch (e) {
    log.error("initExcelPasteBaseline error", e);
  }
}

async function resizeNewExcelImagesByIdDiff(settings: ResizeSettings, shapeCount: number): Promise<boolean> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  const sizeSnapshot = getExcelSizeSnapshot();

  return await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const shapes = sheet.shapes;
    shapes.load("items");
    await ctx.sync();

    const items = shapes.items || [];
    for (const s of items) {
      s.load(["id", "type", "width", "height", "name"]);
    }

    const active = ctx.workbook.getActiveShapeOrNullObject();
    active.load(["isNullObject", "id"]);

    await ctx.sync();

    const currentIds = new Set<string>();
    for (const s of items) {
      if (s.id) currentIds.add(s.id);
    }

    const newIds: string[] = [];
    for (const id of currentIds) {
      if (!baselineShapeIds.has(id)) newIds.push(id);
    }

    if (newIds.length === 0) {
      baselineShapeIds = currentIds;
      state.lastShapeCount = shapeCount;
      return false;
    }

    const byId = new Map<string, Excel.Shape>();
    for (const s of items) {
      if (s.id) byId.set(s.id, s);
    }

    const orderedCandidates: Excel.Shape[] = [];
    const activeId = !active.isNullObject && active.id ? active.id : null;
    if (activeId && newIds.includes(activeId)) {
      const s = byId.get(activeId);
      if (s) orderedCandidates.push(s);
    }

    for (const id of newIds) {
      if (activeId && id === activeId) continue;
      const s = byId.get(id);
      if (s) orderedCandidates.push(s);
    }

    let anyResized = false;

    for (const shape of orderedCandidates) {
      if (!shape.id) continue;
      if (!isImageShape(shape)) continue;

      const w = shape.width;
      const h = shape.height;

      if (sizeSnapshot && isSizeInExcelSnapshot(w, h, sizeSnapshot)) {
        log.debug("resizeNewExcelImagesByIdDiff: size in snapshot, skip", { id: shape.id, width: w, height: h });
        recordShapeId(shape.id);
        continue;
      }

      if (
        (!settings.applyWidth || isWithinEpsilon(w, tw)) &&
        (!settings.applyHeight || isWithinEpsilon(h, th))
      ) {
        recordShapeId(shape.id);
        continue;
      }

      let expectedW = w;
      let expectedH = h;
      const aspect = w > 0 ? h / w : 0;

      if (settings.lockAspectRatio && settings.applyWidth && !settings.applyHeight && aspect > 0) {
        expectedW = tw;
        expectedH = tw * aspect;
        shape.width = expectedW;
        shape.height = expectedH;
      } else if (settings.lockAspectRatio && !settings.applyWidth && settings.applyHeight && aspect > 0) {
        expectedH = th;
        expectedW = th / aspect;
        shape.height = expectedH;
        shape.width = expectedW;
      } else {
        if (settings.applyWidth) {
          expectedW = tw;
          shape.width = expectedW;
        }
        if (settings.applyHeight) {
          expectedH = th;
          shape.height = expectedH;
        }
      }

      await ctx.sync();

      let verified = true;
      try {
        shape.load(["width", "height"]);
        await ctx.sync();
        const okW = !settings.applyWidth || isWithinEpsilon(shape.width, expectedW);
        const okH = !settings.applyHeight || isWithinEpsilon(shape.height, expectedH);
        verified = okW && okH;
      } catch {
        verified = false;
      }

      if (!verified) {
        log.debug("resizeNewExcelImagesByIdDiff: verify failed", {
          id: shape.id,
          width: shape.width,
          height: shape.height,
          expectedW,
          expectedH,
          applyWidth: settings.applyWidth,
          applyHeight: settings.applyHeight,
        });
        continue;
      }

      recordShapeId(shape.id);
      anyResized = true;
      log.info("resizeNewExcelImagesByIdDiff: resized", { id: shape.id, name: shape.name });
    }

    baselineShapeIds = currentIds;
    state.lastShapeCount = shapeCount;
    return anyResized;
  });
}

export async function updateExcelSizeSnapshot(targetWidthPt: number, targetHeightPt: number): Promise<void> {
  const log = getLog();
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const shapes = sheet.shapes;
      shapes.load("items");
      await ctx.sync();
      
      const widthSet = new Set<number>();
      const heightSet = new Set<number>();
      
      if (targetWidthPt > 0) widthSet.add(Math.round(targetWidthPt));
      if (targetHeightPt > 0) heightSet.add(Math.round(targetHeightPt));
      
      for (const shape of shapes.items || []) {
        shape.load(["type", "width", "height"]);
      }
      await ctx.sync();
      
      for (const shape of shapes.items || []) {
        if (isImageShape(shape)) {
          const w = Math.round(shape.width);
          const h = Math.round(shape.height);
          if (w > 0) widthSet.add(w);
          if (h > 0) heightSet.add(h);
        }
      }
      
      const snapshot: ExcelSizeSnapshot = {
        widths: Array.from(widthSet),
        heights: Array.from(heightSet),
        timestamp: Date.now(),
      };
      saveExcelSizeSnapshot(snapshot);
      log.info("updateExcelSizeSnapshot", { widths: snapshot.widths.length, heights: snapshot.heights.length });
    });
  } catch (e) {
    log.warn("updateExcelSizeSnapshot error", e);
  }
}

export async function checkExcelCountChange(): Promise<{
  shapeIncreased: boolean;
  shapeCount: number;
}> {
  return await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const shapes = sheet.shapes;
    shapes.load("items");
    await ctx.sync();
    
    const sc = shapes.items?.length ?? 0;
    if (sc < state.lastShapeCount) state.lastShapeCount = sc;
    
    const shapeIncreased = sc > state.lastShapeCount;
    return { shapeIncreased, shapeCount: sc };
  });
}

/**
 * 快速路径：直接调整当前选中的图片
 * Excel 粘贴后图片会自动选中，利用这个特性
 * 复用 adjustSelectedExcelShape 函数
 */
async function resizeSelectedExcelImage(settings: ResizeSettings): Promise<boolean> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  
  return await Excel.run(async (ctx) => {
    // 先检查是否有选中的 shape
    let shape: Excel.Shape;
    try {
      shape = ctx.workbook.getActiveShapeOrNullObject();
      shape.load(["isNullObject", "id", "type", "width", "height", "name"]);
      await ctx.sync();
    } catch (e) {
      log.debug("resizeSelectedExcelImage: getActiveShapeOrNullObject failed", { error: String(e) });
      return false;
    }
    
    if (shape.isNullObject) {
      log.debug("resizeSelectedExcelImage: no active shape (isNullObject=true)");
      return false;
    }
    
    log.debug("resizeSelectedExcelImage: found active shape", { 
      id: shape.id, 
      type: shape.type, 
      name: shape.name,
      width: shape.width,
      height: shape.height 
    });
    
    if (!isImageShape(shape)) {
      log.debug("resizeSelectedExcelImage: not an image shape", { type: shape.type });
      return false;
    }
    
    const shapeId = shape.id;
    const w = shape.width;
    const h = shape.height;
    
    // 检查是否已处理过
    if (shapeId && hasShapeId(shapeId)) {
      log.debug("resizeSelectedExcelImage: already processed", { id: shapeId });
      return false;
    }
    
    // 检查尺寸是否已经正确
    if (isWithinEpsilon(w, tw)) {
      log.debug("resizeSelectedExcelImage: size already correct", { width: w, target: tw });
      if (shapeId) recordShapeId(shapeId);
      return false;
    }
    
    // 缩放（为避免 Excel 不联动，等比时显式写入 width+height）
    let expectedW = w;
    let expectedH = h;
    const aspect = w > 0 ? h / w : 0;

    if (settings.lockAspectRatio && settings.applyWidth && !settings.applyHeight && aspect > 0) {
      expectedW = tw;
      expectedH = tw * aspect;
      shape.width = expectedW;
      shape.height = expectedH;
    } else if (settings.lockAspectRatio && !settings.applyWidth && settings.applyHeight && aspect > 0) {
      expectedH = th;
      expectedW = th / aspect;
      shape.height = expectedH;
      shape.width = expectedW;
    } else {
      if (settings.applyWidth) {
        expectedW = tw;
        shape.width = expectedW;
      }
      if (settings.applyHeight) {
        expectedH = th;
        shape.height = expectedH;
      }
    }

    await ctx.sync();

    try {
      shape.load(["width", "height"]);
      await ctx.sync();
      const okW = !settings.applyWidth || isWithinEpsilon(shape.width, expectedW);
      const okH = !settings.applyHeight || isWithinEpsilon(shape.height, expectedH);
      if (!okW || !okH) {
        log.debug("resizeSelectedExcelImage: verify failed", {
          id: shapeId,
          width: shape.width,
          height: shape.height,
          expectedW,
          expectedH,
          applyWidth: settings.applyWidth,
          applyHeight: settings.applyHeight,
        });
        return false;
      }
    } catch {
      return false;
    }

    if (shapeId) recordShapeId(shapeId);
    
    log.info("resizeSelectedExcelImage: resized", { id: shapeId, name: shape.name, oldWidth: w, newWidth: tw });
    return true;
  });
}

/**
 * 调整最后一个图片的尺寸（备用方案）
 */
export async function resizeLastExcelImage(
  settings: ResizeSettings,
  shapeCount: number
): Promise<boolean> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  const sizeSnapshot = getExcelSizeSnapshot();
  
  return await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const shapes = sheet.shapes;
    shapes.load("items");
    await ctx.sync();
    
    const items = shapes.items || [];
    if (items.length === 0) return false;
    
    for (let i = items.length - 1; i >= 0; i--) {
      const shape = items[i];
      shape.load(["id", "type", "width", "height", "name"]);
    }
    await ctx.sync();
    
    for (let i = items.length - 1; i >= 0; i--) {
      const shape = items[i];
      if (!isImageShape(shape)) continue;
      
      const shapeId = shape.id;
      const w = shape.width;
      const h = shape.height;
      
      log.debug("check shape", { index: i, id: shapeId, width: w, height: h, targetWidth: tw, type: shape.type });
      
      if (shapeId && hasShapeId(shapeId)) {
        log.debug("shape ID in registry, skip", { index: i, id: shapeId });
        continue;
      }
      
      if (sizeSnapshot && isSizeInExcelSnapshot(w, h, sizeSnapshot)) {
        log.debug("shape size in snapshot, skip", { index: i, width: w, height: h });
        continue;
      }
      
      if (!isWithinEpsilon(w, tw)) {
        // 缩放（为避免 Excel 不联动，等比时显式写入 width+height）
        let expectedW = w;
        let expectedH = h;
        const aspect = w > 0 ? h / w : 0;

        if (settings.lockAspectRatio && settings.applyWidth && !settings.applyHeight && aspect > 0) {
          expectedW = tw;
          expectedH = tw * aspect;
          shape.width = expectedW;
          shape.height = expectedH;
        } else if (settings.lockAspectRatio && !settings.applyWidth && settings.applyHeight && aspect > 0) {
          expectedH = th;
          expectedW = th / aspect;
          shape.height = expectedH;
          shape.width = expectedW;
        } else {
          if (settings.applyWidth) {
            expectedW = tw;
            shape.width = expectedW;
          }
          if (settings.applyHeight) {
            expectedH = th;
            shape.height = expectedH;
          }
        }

        await ctx.sync();

        try {
          shape.load(["width", "height"]);
          await ctx.sync();
          const okW = !settings.applyWidth || isWithinEpsilon(shape.width, expectedW);
          const okH = !settings.applyHeight || isWithinEpsilon(shape.height, expectedH);
          if (!okW || !okH) {
            log.debug("resizeLastExcelImage: verify failed", {
              index: i,
              id: shapeId,
              width: shape.width,
              height: shape.height,
              expectedW,
              expectedH,
              applyWidth: settings.applyWidth,
              applyHeight: settings.applyHeight,
            });
            return false;
          }
        } catch {
          return false;
        }

        if (shapeId) {
          recordShapeId(shapeId);
          log.debug("recorded new shape ID", { id: shapeId });
        }
        state.lastShapeCount = shapeCount;
        log.info("resized Excel shape", { index: i, id: shapeId, name: shape.name });
        return true;
      }
    }
    
    state.lastShapeCount = shapeCount;
    return false;
  });
}

/**
 * 处理 shape 数量增加
 * 优先使用快速路径（直接调整选中的图片）
 */
export function handleExcelCountIncrease(
  settings: ResizeSettings,
  shapeCount: number,
  onResult: (resized: boolean, method: string) => void
): void {
  if (state.processing) return;
  state.processing = true;
  if (state.fallbackTimer !== null) {
    window.clearTimeout(state.fallbackTimer);
    state.fallbackTimer = null;
  }
  
  const log = getLog();
  
  void (async () => {
    try {
      // 快速路径：直接调整选中的图片
      const resized = await resizeSelectedExcelImage(settings);
      if (resized) {
        state.lastShapeCount = shapeCount;
        state.processing = false;
        onResult(true, "selected-shape");
        return;
      }

      const resizedDiff = await resizeNewExcelImagesByIdDiff(settings, shapeCount);
      if (resizedDiff) {
        state.processing = false;
        onResult(true, "diff-ids");
        return;
      }
      
      // 备用方案：遍历找最后一个新图片
      const resized2 = await resizeLastExcelImage(settings, shapeCount);
      if (resized2) {
        state.processing = false;
        onResult(true, "last-shape");
        return;
      }
      
      // 延迟重试
      state.fallbackTimer = window.setTimeout(async () => {
        state.fallbackTimer = null;
        try {
          const resized3 = await resizeSelectedExcelImage(settings);
          if (resized3) {
            state.lastShapeCount = shapeCount;
            onResult(true, "delayed-selected");
          } else {
            const resizedDiff2 = await resizeNewExcelImagesByIdDiff(settings, shapeCount);
            if (resizedDiff2) {
              onResult(true, "delayed-diff-ids");
              return;
            }
            const resized4 = await resizeLastExcelImage(settings, shapeCount);
            onResult(resized4, "delayed-fallback");
          }
        } catch {
          onResult(false, "error");
        } finally {
          state.processing = false;
        }
      }, 300);
    } catch (e) {
      log.error("handleExcelCountIncrease error", e);
      state.processing = false;
      onResult(false, "error");
    }
  })();
}

export function cancelExcelPendingDetection(): void {
  if (state.fallbackTimer !== null) {
    window.clearTimeout(state.fallbackTimer);
    state.fallbackTimer = null;
  }
  state.processing = false;
}
