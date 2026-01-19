/**
 * PowerPoint 粘贴检测器 v3
 * 修复：等比例缩放 + 快速路径
 */

import { ResizeSettings } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { saveSetting, getStringSetting } from "../settings";
import { isReferenceBoxName } from "../reference-box/types";

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[pptPasteDetector]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[pptPasteDetector]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[pptPasteDetector]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[pptPasteDetector]", msg, data ?? ""),
  };
}

// ============ 尺寸快照 ============

interface PptSizeSnapshot {
  widths: number[];
  heights: number[];
  timestamp: number;
}

const PPT_SIZE_SNAPSHOT_KEY = "pasteWidth_pptSizeSnapshot";

function savePptSizeSnapshot(snapshot: PptSizeSnapshot): void {
  saveSetting(PPT_SIZE_SNAPSHOT_KEY, JSON.stringify(snapshot));
}

function getPptSizeSnapshot(): PptSizeSnapshot | null {
  const raw = getStringSetting(PPT_SIZE_SNAPSHOT_KEY, "");
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed.widths) && Array.isArray(parsed.heights)) {
      return parsed as PptSizeSnapshot;
    }
    return null;
  } catch {
    return null;
  }
}

function isSizeInPptSnapshot(width: number, height: number, snapshot: PptSizeSnapshot, tolerance: number = 1): boolean {
  const w = Math.round(width);
  const h = Math.round(height);
  const widthInSnapshot = snapshot.widths.some(sw => Math.abs(sw - w) <= tolerance);
  const heightInSnapshot = snapshot.heights.some(sh => Math.abs(sh - h) <= tolerance);
  return widthInSnapshot && heightInSnapshot;
}

function isImageShape(shape: PowerPoint.Shape): boolean {
  const shapeType = (shape as any).type;
  if (shapeType) {
    const typeString = String(shapeType).toLowerCase();
    if (typeString.includes("image") || typeString.includes("picture")) {
      return true;
    }
    if (
      typeString.includes("text") ||
      typeString.includes("chart") ||
      typeString.includes("table") ||
      typeString.includes("smartart") ||
      typeString.includes("media") ||
      typeString.includes("video") ||
      typeString.includes("audio") ||
      typeString.includes("group") ||
      typeString.includes("line") ||
      typeString.includes("connector")
    ) {
      return false;
    }
  }
  
  const shapeName = (shape as any).name || "";
  if (typeof shapeName === "string") {
    const nameLower = shapeName.toLowerCase();
    if (nameLower.includes("picture") || nameLower.includes("image")) {
      return true;
    }
  }
  
  return false;
}

// ============ Shape ID 内存表 ============

const baselineShapeIds = new Set<string>();
const resizedShapeIds = new Set<string>();

function recordBaselineShapeId(id: string): void {
  baselineShapeIds.add(id);
}

function recordShapeId(id: string): void {
  resizedShapeIds.add(id);
  baselineShapeIds.add(id);
}

function hasShapeId(id: string): boolean {
  return baselineShapeIds.has(id) || resizedShapeIds.has(id);
}

function clearShapeIds(): void {
  baselineShapeIds.clear();
  resizedShapeIds.clear();
}

async function resizeNewPptImagesByIdDiff(settings: ResizeSettings, shapeCount: number): Promise<boolean> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  const sizeSnapshot = getPptSizeSnapshot();

  return await PowerPoint.run(async (ctx) => {
    const currentSlideId = await getSelectedSlideIdOrNull(ctx);
    if (state.lastSlideId !== null && currentSlideId !== null && currentSlideId !== state.lastSlideId) {
      await resetBaselineForCurrentSlide(ctx);
      return false;
    }

    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return false;

    const slide = slides.items[0];
    const shapes = slide.shapes;
    shapes.load("items");
    await ctx.sync();

    const items = shapes.items || [];
    for (const s of items) {
      s.load(["id", "width", "height", "name"]);
    }
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
      baselineShapeIds.clear();
      for (const id of currentIds) baselineShapeIds.add(id);
      state.lastShapeCount = shapeCount;
      return false;
    }

    const byId = new Map<string, PowerPoint.Shape>();
    for (const s of items) {
      if (s.id) byId.set(s.id, s);
    }

    let anyResized = false;

    for (const id of newIds) {
      const shape = byId.get(id);
      if (!shape) continue;

      if (isReferenceBoxName(shape.name)) {
        log.debug("resizeNewPptImagesByIdDiff: skip reference box", { id, name: shape.name });
        recordShapeId(id);
        continue;
      }

      if (!isImageShape(shape)) {
        recordShapeId(id);
        continue;
      }

      const w = shape.width;
      const h = shape.height;

      const matchesTarget =
        (!settings.applyWidth || isWithinEpsilon(w, tw)) &&
        (!settings.applyHeight || isWithinEpsilon(h, th));
      if (matchesTarget) {
        recordShapeId(id);
        continue;
      }

      // 这里不把 sizeSnapshot 作为“新旧判断”，只作为 fallback 保护：如果新图尺寸刚好命中快照，也仍然允许处理。
      // 因此：不在 diff-path 中用 snapshot skip。

      let expectedW: number | null = null;
      let expectedH: number | null = null;
      let origLeft: number | null = null;
      let origTop: number | null = null;
      try {
        (shape as any).load(["left", "top"]);
        await ctx.sync();
        const l = Number((shape as any).left);
        const t = Number((shape as any).top);
        if (Number.isFinite(l)) origLeft = l;
        if (Number.isFinite(t)) origTop = t;
      } catch {
        // ignore
      }

      try {
        (shape as any).lockAspectRatio = settings.lockAspectRatio;
      } catch {}

      if (settings.lockAspectRatio && settings.applyWidth && !settings.applyHeight) {
        const scale = w > 0 ? tw / w : 0;
        const newW = tw;
        const newH = scale > 0 ? h * scale : h;
        expectedW = newW;
        expectedH = newH;
        if (Number.isFinite(newW) && newW > 0) shape.width = newW;
        if (Number.isFinite(newH) && newH > 0) shape.height = newH;
      } else if (settings.lockAspectRatio && !settings.applyWidth && settings.applyHeight) {
        const scale = h > 0 ? th / h : 0;
        const newW = scale > 0 ? w * scale : w;
        const newH = th;
        expectedW = newW;
        expectedH = newH;
        if (Number.isFinite(newW) && newW > 0) shape.width = newW;
        if (Number.isFinite(newH) && newH > 0) shape.height = newH;
      } else {
        expectedW = settings.applyWidth ? tw : null;
        expectedH = settings.applyHeight ? th : null;
        if (settings.applyWidth) shape.width = tw;
        if (settings.applyHeight) shape.height = th;
      }

      if (origLeft !== null) (shape as any).left = Math.max(0, origLeft);
      if (origTop !== null) (shape as any).top = Math.max(0, origTop);

      await ctx.sync();

      let verified = true;
      try {
        shape.load(["width", "height"]);
        await ctx.sync();
        const okWidth = expectedW === null || isWithinEpsilon(shape.width, expectedW);
        const okHeight = expectedH === null || isWithinEpsilon(shape.height, expectedH);
        verified = okWidth && okHeight;
      } catch {
        verified = false;
      }

      if (!verified) {
        log.debug("resizeNewPptImagesByIdDiff: verify failed", {
          id,
          width: shape.width,
          height: shape.height,
          expectedWidth: expectedW,
          expectedHeight: expectedH,
        });
        continue;
      }

      recordShapeId(id);
      anyResized = true;
      log.info("resizeNewPptImagesByIdDiff: resized", { id, name: shape.name, width: shape.width, height: shape.height, hasSnapshot: Boolean(sizeSnapshot) });
    }

    baselineShapeIds.clear();
    for (const id of currentIds) baselineShapeIds.add(id);
    state.lastShapeCount = shapeCount;
    return anyResized;
  });
}

// ============ 状态管理 ============

interface PptPasteState {
  lastShapeCount: number;
  processing: boolean;
  fallbackTimer: number | null;
  lastSlideId: string | null;
}

const state: PptPasteState = {
  lastShapeCount: 0,
  processing: false,
  fallbackTimer: null,
  lastSlideId: null,
};

async function getSelectedSlideIdOrNull(ctx: PowerPoint.RequestContext): Promise<string | null> {
  try {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return null;
    const slide = slides.items[0];
    (slide as any).load(["id"]);
    await ctx.sync();
    const id = (slide as any).id;
    return typeof id === "string" && id ? id : null;
  } catch {
    return null;
  }
}

async function resetBaselineForCurrentSlide(ctx: PowerPoint.RequestContext): Promise<void> {
  clearShapeIds();
  const slides = ctx.presentation.getSelectedSlides();
  slides.load("items");
  await ctx.sync();
  if (slides.items.length === 0) {
    state.lastShapeCount = 0;
    state.lastSlideId = null;
    return;
  }
  const slide = slides.items[0];
  const slideId = await getSelectedSlideIdOrNull(ctx);
  state.lastSlideId = slideId;

  const shapes = slide.shapes;
  shapes.load("items");
  await ctx.sync();
  state.lastShapeCount = shapes.items?.length ?? 0;

  for (const shape of shapes.items || []) {
    shape.load(["id"]);
  }
  await ctx.sync();

  for (const shape of shapes.items || []) {
    if (shape.id) recordBaselineShapeId(shape.id);
  }
}

export async function initPptPasteBaseline(): Promise<void> {
  const log = getLog();
  try {
    clearShapeIds();
    
    await PowerPoint.run(async (ctx) => {
      await resetBaselineForCurrentSlide(ctx);
      log.info("initPptPasteBaseline", {
        shapeCount: state.lastShapeCount,
        registeredIds: baselineShapeIds.size,
        slideId: state.lastSlideId,
      });
    });
  } catch (e) {
    log.error("initPptPasteBaseline error", e);
  }
}

export async function updatePptSizeSnapshot(targetWidthPt: number, targetHeightPt: number): Promise<void> {
  const log = getLog();
  try {
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.getSelectedSlides();
      slides.load("items");
      await ctx.sync();
      
      if (slides.items.length === 0) return;
      
      const slide = slides.items[0];
      const shapes = slide.shapes;
      shapes.load("items");
      await ctx.sync();
      
      const widthSet = new Set<number>();
      const heightSet = new Set<number>();
      
      if (targetWidthPt > 0) widthSet.add(Math.round(targetWidthPt));
      if (targetHeightPt > 0) heightSet.add(Math.round(targetHeightPt));
      
      for (const shape of shapes.items || []) {
        shape.load(["width", "height", "name"]);
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
      
      const snapshot: PptSizeSnapshot = {
        widths: Array.from(widthSet),
        heights: Array.from(heightSet),
        timestamp: Date.now(),
      };
      savePptSizeSnapshot(snapshot);
      log.info("updatePptSizeSnapshot", { widths: snapshot.widths.length, heights: snapshot.heights.length });
    });
  } catch (e) {
    log.warn("updatePptSizeSnapshot error", e);
  }
}

export async function checkPptCountChange(): Promise<{
  shapeIncreased: boolean;
  shapeCount: number;
}> {
  return await PowerPoint.run(async (ctx) => {
    const currentSlideId = await getSelectedSlideIdOrNull(ctx);
    if (state.lastSlideId !== null && currentSlideId !== null && currentSlideId !== state.lastSlideId) {
      await resetBaselineForCurrentSlide(ctx);
      return { shapeIncreased: false, shapeCount: state.lastShapeCount };
    }

    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items");
    await ctx.sync();
    
    if (slides.items.length === 0) {
      return { shapeIncreased: false, shapeCount: 0 };
    }
    
    const slide = slides.items[0];
    const shapes = slide.shapes;
    shapes.load("items");
    await ctx.sync();
    
    const sc = shapes.items?.length ?? 0;
    if (sc < state.lastShapeCount) state.lastShapeCount = sc;
    
    const shapeIncreased = sc > state.lastShapeCount;
    return { shapeIncreased, shapeCount: sc };
  });
}

/**
 * 调整最后一个图片的尺寸
 * 等比例缩放：锁定宽高比，只设置宽度
 */
export async function resizeLastPptImage(
  settings: ResizeSettings,
  shapeCount: number
): Promise<boolean> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  const sizeSnapshot = getPptSizeSnapshot();

  return await PowerPoint.run(async (ctx) => {
    const currentSlideId = await getSelectedSlideIdOrNull(ctx);
    if (state.lastSlideId !== null && currentSlideId !== null && currentSlideId !== state.lastSlideId) {
      await resetBaselineForCurrentSlide(ctx);
      return false;
    }

    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items");
    await ctx.sync();
    
    if (slides.items.length === 0) return false;
    
    const slide = slides.items[0];
    const shapes = slide.shapes;
    shapes.load("items");
    await ctx.sync();
    
    const items = shapes.items || [];
    if (items.length === 0) return false;
    
    for (let i = items.length - 1; i >= 0; i--) {
      const shape = items[i];
      shape.load(["id", "width", "height", "name"]);
    }
    await ctx.sync();
    
    for (let i = items.length - 1; i >= 0; i--) {
      const shape = items[i];
      
      if (!isImageShape(shape)) {
        log.debug("skip non-image shape", { index: i, name: shape.name });
        continue;
      }

      if (isReferenceBoxName(shape.name)) {
        log.debug("skip reference box", { index: i, name: shape.name });
        continue;
      }
      
      const shapeId = shape.id;
      const w = shape.width;
      const h = shape.height;
      
      log.debug("check image shape", { index: i, id: shapeId, width: w, height: h, targetWidth: tw, name: shape.name });
      
      if (shapeId && hasShapeId(shapeId)) {
        log.debug("shape ID in registry, skip", { index: i, id: shapeId });
        continue;
      }

      if (sizeSnapshot && isSizeInPptSnapshot(w, h, sizeSnapshot)) {
        log.debug("shape size in snapshot, skip", { index: i, id: shapeId, width: w, height: h });
        if (shapeId) recordShapeId(shapeId);
        continue;
      }

      const matchesTarget =
        (!settings.applyWidth || isWithinEpsilon(w, tw)) &&
        (!settings.applyHeight || isWithinEpsilon(h, th));
      if (matchesTarget) {
        if (shapeId) recordShapeId(shapeId);
        continue;
      }
      
      if (
        (settings.applyWidth && !isWithinEpsilon(w, tw)) ||
        (settings.applyHeight && !isWithinEpsilon(h, th))
      ) {
        let expectedW: number | null = null;
        let expectedH: number | null = null;
        let origLeft: number | null = null;
        let origTop: number | null = null;
        try {
          (shape as any).load(["left", "top"]);
          await ctx.sync();
          const l = Number((shape as any).left);
          const t = Number((shape as any).top);
          if (Number.isFinite(l)) origLeft = l;
          if (Number.isFinite(t)) origTop = t;
        } catch {
          // ignore
        }

        try {
          (shape as any).lockAspectRatio = settings.lockAspectRatio;
        } catch {}

        if (settings.lockAspectRatio && settings.applyWidth && !settings.applyHeight) {
          const scale = tw / w;
          const newW = tw;
          const newH = h * scale;
          expectedW = newW;
          expectedH = newH;
          if (Number.isFinite(newW) && newW > 0) shape.width = newW;
          if (Number.isFinite(newH) && newH > 0) shape.height = newH;
        } else if (settings.lockAspectRatio && !settings.applyWidth && settings.applyHeight) {
          const scale = th / h;
          const newW = w * scale;
          const newH = th;
          expectedW = newW;
          expectedH = newH;
          if (Number.isFinite(newW) && newW > 0) shape.width = newW;
          if (Number.isFinite(newH) && newH > 0) shape.height = newH;
        } else {
          expectedW = settings.applyWidth ? tw : null;
          expectedH = settings.applyHeight ? th : null;
          if (settings.applyWidth) shape.width = tw;
          if (settings.applyHeight) shape.height = th;
        }

        if (origLeft !== null) (shape as any).left = Math.max(0, origLeft);
        if (origTop !== null) (shape as any).top = Math.max(0, origTop);

        await ctx.sync();

        let verified = true;
        try {
          shape.load(["width", "height"]);
          await ctx.sync();
          const okWidth =
            expectedW === null || isWithinEpsilon(shape.width, expectedW);
          const okHeight =
            expectedH === null || isWithinEpsilon(shape.height, expectedH);
          verified = okWidth && okHeight;
        } catch {
          verified = false;
        }

        state.lastShapeCount = shapeCount;
        if (!verified) {
          log.debug("resize verify failed", {
            index: i,
            id: shapeId,
            width: shape.width,
            height: shape.height,
            targetWidth: tw,
            targetHeight: th,
            expectedWidth: expectedW,
            expectedHeight: expectedH,
            applyWidth: settings.applyWidth,
            applyHeight: settings.applyHeight,
          });
          return false;
        }

        if (shapeId) {
          recordShapeId(shapeId);
          log.debug("recorded resized shape ID", { id: shapeId });
        }
        log.info("resized PPT shape", { index: i, id: shapeId, name: shape.name });
        return true;
      }
    }
    
    state.lastShapeCount = shapeCount;
    return false;
  });
}

export function handlePptCountIncrease(
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
      const resizedDiff = await resizeNewPptImagesByIdDiff(settings, shapeCount);
      if (resizedDiff) {
        state.processing = false;
        onResult(true, "diff-ids");
        return;
      }

      const resized = await resizeLastPptImage(settings, shapeCount);
      if (resized) {
        state.processing = false;
        onResult(true, "fast-path");
        return;
      }
      
      state.fallbackTimer = window.setTimeout(async () => {
        state.fallbackTimer = null;
        try {
          const resized2 = await resizeNewPptImagesByIdDiff(settings, shapeCount);
          if (resized2) {
            onResult(true, "delayed-diff-ids");
            return;
          }

          const resized3 = await resizeLastPptImage(settings, shapeCount);
          onResult(resized3, "delayed-fallback");
        } catch {
          onResult(false, "error");
        } finally {
          state.processing = false;
        }
      }, 300);
    } catch (e) {
      log.error("handlePptCountIncrease error", e);
      state.processing = false;
      onResult(false, "error");
    }
  })();
}

export function cancelPptPendingDetection(): void {
  if (state.fallbackTimer !== null) {
    window.clearTimeout(state.fallbackTimer);
    state.fallbackTimer = null;
  }
  state.processing = false;
}
