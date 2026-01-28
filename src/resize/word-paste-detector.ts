import { ResizeSettings } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { isImageShape, applyInlinePictureSize, applyShapeSize } from "./word";
import { getResizeScope } from "../settings";
import { recordShapeId, getRegistry, hasShapeId } from "../resize-scope/registry";
import { getSizeSnapshot, isSizeInSnapshot } from "../resize-scope/size-snapshot";

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[pasteDetector]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[pasteDetector]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[pasteDetector]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[pasteDetector]", msg, data ?? ""),
  };
}

interface PasteState {
  lastInlineCount: number;
  lastShapeCount: number;
  fallbackTimer: number | null;
  processing: boolean;
}

const state: PasteState = {
  lastInlineCount: 0,
  lastShapeCount: 0,
  fallbackTimer: null,
  processing: false,
};

export async function initPasteBaseline(): Promise<void> {
  await Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    let shapes: any = null;
    try {
      shapes = (ctx.document.body as any).shapes;
    } catch {
      shapes = null;
    }
    let ic = 0, sc = 0;
    try {
      const icr = (pics as any).getCount();
      const scr = shapes ? (shapes as any).getCount() : null;
      await ctx.sync();
      ic = (icr as any)?.value ?? 0;
      sc = shapes ? ((scr as any)?.value ?? 0) : 0;
    } catch {
      pics.load("items");
      if (shapes) shapes.load("items");
      await ctx.sync();
      ic = pics.items?.length ?? 0;
      sc = shapes ? (shapes.items?.length ?? 0) : 0;
    }
    state.lastInlineCount = ic;
    state.lastShapeCount = sc;
  });
}

export async function checkCountChange(): Promise<{
  inlineIncreased: boolean;
  shapeIncreased: boolean;
  inlineCount: number;
  shapeCount: number;
}> {
  return await Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    let shapes: any = null;
    try {
      shapes = (ctx.document.body as any).shapes;
    } catch {
      shapes = null;
    }
    let ic = 0, sc = 0;
    try {
      const icr = (pics as any).getCount();
      const scr = shapes ? (shapes as any).getCount() : null;
      await ctx.sync();
      ic = (icr as any)?.value ?? 0;
      sc = shapes ? ((scr as any)?.value ?? 0) : 0;
    } catch {
      pics.load("items");
      if (shapes) shapes.load("items");
      await ctx.sync();
      ic = pics.items?.length ?? 0;
      sc = shapes ? (shapes.items?.length ?? 0) : 0;
    }
    if (ic < state.lastInlineCount) state.lastInlineCount = ic;
    if (sc < state.lastShapeCount) state.lastShapeCount = sc;
    const inlineIncreased = ic > state.lastInlineCount;
    const shapeIncreased = shapes ? sc > state.lastShapeCount : false;
    return { inlineIncreased, shapeIncreased, inlineCount: ic, shapeCount: sc };
  });
}

export function updateBaseline(inlineCount: number, shapeCount: number): void {
  state.lastInlineCount = inlineCount;
  state.lastShapeCount = shapeCount;
}

/**
 * 快速路径 v20：使用 load("items") + 索引访问最后一张图片
 * 修复：Word JS API 没有 getItemAt 方法
 */
export async function resizeLastImage(
  settings: ResizeSettings,
  inlineIncreased: boolean,
  shapeIncreased: boolean,
  inlineCount: number,
  shapeCount: number
): Promise<"inline" | "shape" | "none"> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const scope = getResizeScope() ?? "new";
  const registry = getRegistry();
  const sizeSnapshot = getSizeSnapshot();
  
  return await Word.run(async (ctx) => {
    if (inlineIncreased && inlineCount > 0) {
      try {
        const pics = ctx.document.body.inlinePictures;
        pics.load("items");
        await ctx.sync();
        const items = pics.items || [];
        if (items.length > 0) {
          const lastIndex = items.length - 1;
          const lastPic = items[lastIndex];
          (lastPic as any).load(["width", "height"]);
          await ctx.sync();
          const w = Number((lastPic as any).width);
          const h = Number((lastPic as any).height);
          log.debug("fast-path: check InlinePicture", { index: lastIndex, width: w, height: h, targetWidth: tw, scope });
          if (scope === "new" && sizeSnapshot !== null) {
            if (isSizeInSnapshot(w, h, sizeSnapshot)) {
              log.debug("fast-path: InlinePicture size in snapshot, skip", { width: w, height: h });
              state.lastInlineCount = inlineCount;
              state.lastShapeCount = shapeCount;
              return "none";
            }
            log.debug("fast-path: InlinePicture NOT in snapshot, process", { width: w, height: h });
          }
          if (!isWithinEpsilon(w, tw)) {
            log.info("fast-path: resize InlinePicture", { index: lastIndex });
            const ok = await applyInlinePictureSize(ctx, lastPic, settings);
            if (ok) {
              state.lastInlineCount = inlineCount;
              state.lastShapeCount = shapeCount;
              return "inline";
            }
          } else {
            log.debug("fast-path: InlinePicture already matches target");
          }
        }
      } catch (e) {
        const errMsg = e instanceof Error ? e.message : String(e);
        log.debug("fast-path: get InlinePicture failed", { error: errMsg });
      }
    }
    if (shapeIncreased && shapeCount > 0) {
      try {
        let shapes: any = null;
        try {
          shapes = (ctx.document.body as any).shapes;
        } catch {
          shapes = null;
        }
        if (!shapes) {
          log.debug("fast-path: shapes not supported, skip shape resize");
          return "none";
        }
        shapes.load("items");
        await ctx.sync();
        const items = shapes.items || [];
        if (items.length > 0) {
          const lastIndex = items.length - 1;
          const lastShape = items[lastIndex];
          (lastShape as any).load(["type", "width", "height", "id"]);
          await ctx.sync();
          const shapeId = (lastShape as any).id as number;
          const w = Number((lastShape as any).width);
          const h = Number((lastShape as any).height);
          log.debug("fast-path: check Shape", { index: lastIndex, type: (lastShape as any).type, width: w, height: h, targetWidth: tw, scope, shapeId });
          if (isImageShape(lastShape)) {
            if (scope === "new" && registry !== null && shapeId) {
              recordShapeId(shapeId);
              log.debug("fast-path: recorded Shape ID", { shapeId });
            }
            if (scope === "new" && sizeSnapshot !== null && registry === null) {
              if (isSizeInSnapshot(w, h, sizeSnapshot)) {
                log.debug("fast-path: Shape size in snapshot, skip", { width: w, height: h, shapeId });
                state.lastInlineCount = inlineCount;
                state.lastShapeCount = shapeCount;
                return "none";
              }
              log.debug("fast-path: Shape NOT in snapshot, process", { width: w, height: h, shapeId });
            }
            if (!isWithinEpsilon(w, tw)) {
              log.info("fast-path: resize Shape", { index: lastIndex, shapeId });
              const ok = await applyShapeSize(ctx, lastShape, settings);
              if (ok) {
                state.lastInlineCount = inlineCount;
                state.lastShapeCount = shapeCount;
                return "shape";
              }
            } else {
              log.debug("fast-path: Shape already matches target");
            }
          }
        }
      } catch (e) {
        const errMsg = e instanceof Error ? e.message : String(e);
        log.debug("fast-path: get Shape failed", { error: errMsg });
      }
    }
    log.debug("fast-path: no image to resize");
    return "none";
  });
}

export async function resizeByGlobalScan(settings: ResizeSettings): Promise<"inline" | "shape" | "none"> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  const scope = getResizeScope() ?? "new";
  const registry = getRegistry();
  const sizeSnapshot = getSizeSnapshot();
  return await Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    let shapes: any = null;
    try {
      shapes = (ctx.document.body as any).shapes;
    } catch {
      shapes = null;
    }
    let ic = 0, sc = 0;
    try {
      const icr = (pics as any).getCount();
      const scr = shapes ? (shapes as any).getCount() : null;
      await ctx.sync();
      ic = (icr as any)?.value ?? 0;
      sc = shapes ? ((scr as any)?.value ?? 0) : 0;
    } catch {
      pics.load("items");
      if (shapes) shapes.load("items");
      await ctx.sync();
      ic = pics.items?.length ?? 0;
      sc = shapes ? (shapes.items?.length ?? 0) : 0;
    }
    log.debug("global-scan start", { inlineCount: ic, shapeCount: sc, scope, hasRegistry: registry !== null, hasSizeSnapshot: sizeSnapshot !== null });
    if (ic > 0) {
      pics.load("items"); await ctx.sync();
      for (const p of pics.items || []) (p as any).load(["width", "height"]);
      await ctx.sync();
      const items = pics.items || [];
      for (let i = items.length - 1; i >= 0; i--) {
        const p = items[i];
        const w = Number((p as any).width);
        const h = Number((p as any).height);
        if (scope === "new" && sizeSnapshot !== null) {
          if (isSizeInSnapshot(w, h, sizeSnapshot)) {
            log.debug("global-scan: InlinePicture in snapshot, skip", { index: i, width: w, height: h });
            continue;
          }
          log.debug("global-scan: InlinePicture NOT in snapshot, check", { index: i, width: w, height: h });
        }
        if (!isWithinEpsilon(w, tw)) {
          if (settings.applyWidth && settings.applyHeight) {
            try { (p as any).lockAspectRatio = false; } catch {}
          } else {
            try { (p as any).lockAspectRatio = settings.lockAspectRatio; } catch {}
          }
          if (settings.applyWidth) (p as any).width = tw;
          if (settings.applyHeight) (p as any).height = th;
          await ctx.sync();
          state.lastInlineCount = ic; state.lastShapeCount = sc;
          log.info("global-scan: resized InlinePicture", { index: i });
          return "inline";
        }
      }
    }
    if (shapes && sc > 0) {
      shapes.load("items"); await ctx.sync();
      for (const s of shapes.items || []) (s as any).load(["type", "width", "height", "id"]);
      await ctx.sync();
      const items = shapes.items || [];
      for (let i = items.length - 1; i >= 0; i--) {
        const s = items[i];
        if (!isImageShape(s)) continue;
        const shapeId = (s as any).id as number;
        const w = Number((s as any).width);
        const h = Number((s as any).height);
        if (scope === "new") {
          if (registry !== null) {
            if (!hasShapeId(shapeId)) {
              log.debug("global-scan: Shape not in registry, skip", { index: i, shapeId });
              continue;
            }
            log.debug("global-scan: Shape in registry, check", { index: i, shapeId });
          } else if (sizeSnapshot !== null) {
            if (isSizeInSnapshot(w, h, sizeSnapshot)) {
              log.debug("global-scan: Shape in snapshot, skip", { index: i, width: w, height: h, shapeId });
              continue;
            }
            log.debug("global-scan: Shape NOT in snapshot, check", { index: i, width: w, height: h, shapeId });
          }
        }
        if (!isWithinEpsilon(w, tw)) {
          if (settings.applyWidth && settings.applyHeight) {
            try { (s as any).lockAspectRatio = false; } catch {}
          } else {
            try { (s as any).lockAspectRatio = settings.lockAspectRatio; } catch {}
          }
          if (settings.applyWidth) (s as any).width = tw;
          if (settings.applyHeight) (s as any).height = th;
          await ctx.sync();
          state.lastInlineCount = ic; state.lastShapeCount = sc;
          log.info("global-scan: resized Shape", { index: i, shapeId });
          return "shape";
        }
      }
    }
    state.lastInlineCount = ic; state.lastShapeCount = sc;
    log.debug("global-scan: no image to resize");
    return "none";
  });
}

export function handleCountIncrease(
  settings: ResizeSettings,
  inlineIncreased: boolean,
  shapeIncreased: boolean,
  inlineCount: number,
  shapeCount: number,
  onResult: (result: "inline" | "shape" | "none", method: string) => void
): void {
  if (state.processing) return;
  state.processing = true;
  if (state.fallbackTimer !== null) { window.clearTimeout(state.fallbackTimer); state.fallbackTimer = null; }
  const log = getLog();
  void (async () => {
    try {
      const r = await resizeLastImage(settings, inlineIncreased, shapeIncreased, inlineCount, shapeCount);
      if (r !== "none") { state.processing = false; onResult(r, "fast-path"); return; }
      log.debug("fast-path failed, run global-scan");
      const fr = await resizeByGlobalScan(settings);
      if (fr !== "none") { state.processing = false; onResult(fr, "immediate-fallback"); return; }
      state.fallbackTimer = window.setTimeout(async () => {
        state.fallbackTimer = null;
        try { const fr2 = await resizeByGlobalScan(settings); onResult(fr2, "delayed-fallback"); }
        catch { onResult("none", "error"); }
        finally { state.processing = false; }
      }, 300);
    } catch (e) { log.error("handleCountIncrease error", e); state.processing = false; onResult("none", "error"); }
  })();
}

export function cancelPendingDetection(): void {
  if (state.fallbackTimer !== null) { window.clearTimeout(state.fallbackTimer); state.fallbackTimer = null; }
  state.processing = false;
}
