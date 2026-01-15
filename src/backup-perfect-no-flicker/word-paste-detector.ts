import { ResizeSettings } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { isImageShape, applyInlinePictureSize, applyShapeSize } from "./word";

// 动态获取日志函数
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
    const shapes = ctx.document.body.shapes;
    let ic = 0, sc = 0;
    try {
      const icr = (pics as any).getCount();
      const scr = (shapes as any).getCount();
      await ctx.sync();
      ic = (icr as any)?.value ?? 0;
      sc = (scr as any)?.value ?? 0;
    } catch {
      pics.load("items"); shapes.load("items");
      await ctx.sync();
      ic = pics.items?.length ?? 0;
      sc = shapes.items?.length ?? 0;
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
    const shapes = ctx.document.body.shapes;
    let ic = 0, sc = 0;
    try {
      const icr = (pics as any).getCount();
      const scr = (shapes as any).getCount();
      await ctx.sync();
      ic = (icr as any)?.value ?? 0;
      sc = (scr as any)?.value ?? 0;
    } catch {
      pics.load("items"); shapes.load("items");
      await ctx.sync();
      ic = pics.items?.length ?? 0;
      sc = shapes.items?.length ?? 0;
    }
    if (ic < state.lastInlineCount) state.lastInlineCount = ic;
    if (sc < state.lastShapeCount) state.lastShapeCount = sc;
    
    const inlineIncreased = ic > state.lastInlineCount;
    const shapeIncreased = sc > state.lastShapeCount;
    
    return { inlineIncreased, shapeIncreased, inlineCount: ic, shapeCount: sc };
  });
}

export function updateBaseline(inlineCount: number, shapeCount: number): void {
  state.lastInlineCount = inlineCount;
  state.lastShapeCount = shapeCount;
}

/**
 * 快速路径 v3：使用 load("items") + 索引访问最后一张图片
 * 比 getItemAt() 更可靠
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
  const th = cmToPoints(settings.targetHeightCm);
  
  return await Word.run(async (ctx) => {
    // 优先处理 InlinePicture（内联图片）
    if (inlineIncreased && inlineCount > 0) {
      try {
        const pics = ctx.document.body.inlinePictures;
        pics.load("items");
        await ctx.sync();
        
        if (pics.items && pics.items.length > 0) {
          const lastPic = pics.items[pics.items.length - 1];
          (lastPic as any).load("width");
          await ctx.sync();
          
          const w = Number((lastPic as any).width);
          log.debug("fast-path: check InlinePicture", { index: pics.items.length - 1, width: w, targetWidth: tw });
          
          if (!isWithinEpsilon(w, tw)) {
            log.info("fast-path: resize InlinePicture at index " + (pics.items.length - 1));
            // 直接设置尺寸，不调用 applyInlinePictureSize 以减少 API 调用
            try { (lastPic as any).lockAspectRatio = settings.lockAspectRatio; } catch {}
            if (settings.applyWidth) (lastPic as any).width = tw;
            if (settings.applyHeight) (lastPic as any).height = th;
            await ctx.sync();
            
            state.lastInlineCount = inlineCount;
            state.lastShapeCount = shapeCount;
            return "inline";
          } else {
            log.debug("fast-path: InlinePicture width matches target, skip");
          }
        }
      } catch (e) {
        log.debug("fast-path: get InlinePicture failed", e);
      }
    }
    
    // 处理 Shape（浮动图片）
    if (shapeIncreased && shapeCount > 0) {
      try {
        const shapes = ctx.document.body.shapes;
        shapes.load("items");
        await ctx.sync();
        
        if (shapes.items && shapes.items.length > 0) {
          const lastShape = shapes.items[shapes.items.length - 1];
          (lastShape as any).load(["type", "width"]);
          await ctx.sync();
          
          log.debug("fast-path: check Shape", { 
            index: shapes.items.length - 1,
            type: (lastShape as any).type, 
            width: (lastShape as any).width,
            targetWidth: tw
          });
          
          if (isImageShape(lastShape)) {
            const w = Number((lastShape as any).width);
            if (!isWithinEpsilon(w, tw)) {
              log.info("fast-path: resize Shape at index " + (shapes.items.length - 1));
              try { (lastShape as any).lockAspectRatio = settings.lockAspectRatio; } catch {}
              if (settings.applyWidth) (lastShape as any).width = tw;
              if (settings.applyHeight) (lastShape as any).height = th;
              await ctx.sync();
              
              state.lastInlineCount = inlineCount;
              state.lastShapeCount = shapeCount;
              return "shape";
            } else {
              log.debug("fast-path: Shape width matches target, skip");
            }
          }
        }
      } catch (e) {
        log.debug("fast-path: get Shape failed", e);
      }
    }
    
    log.debug("fast-path: no image to resize, fallback to global scan");
    return "none";
  });
}

export async function resizeByGlobalScan(settings: ResizeSettings): Promise<"inline" | "shape" | "none"> {
  const log = getLog();
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  
  return await Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    const shapes = ctx.document.body.shapes;
    let ic = 0, sc = 0;
    
    try {
      const icr = (pics as any).getCount();
      const scr = (shapes as any).getCount();
      await ctx.sync();
      ic = (icr as any)?.value ?? 0;
      sc = (scr as any)?.value ?? 0;
    } catch {
      pics.load("items"); shapes.load("items");
      await ctx.sync();
      ic = pics.items?.length ?? 0;
      sc = shapes.items?.length ?? 0;
    }
    
    log.debug("global-scan start", { inlineCount: ic, shapeCount: sc });
    
    if (ic > 0) {
      pics.load("items"); await ctx.sync();
      for (const p of pics.items || []) (p as any).load("width");
      await ctx.sync();
      for (const p of pics.items || []) {
        if (!isWithinEpsilon((p as any).width, tw)) {
          try { (p as any).lockAspectRatio = settings.lockAspectRatio; } catch {}
          if (settings.applyWidth) (p as any).width = tw;
          if (settings.applyHeight) (p as any).height = th;
          await ctx.sync();
          state.lastInlineCount = ic; state.lastShapeCount = sc;
          log.info("global-scan: resized InlinePicture");
          return "inline";
        }
      }
    }
    
    if (sc > 0) {
      shapes.load("items"); await ctx.sync();
      for (const s of shapes.items || []) (s as any).load(["type","width"]);
      await ctx.sync();
      for (const s of shapes.items || []) {
        if (!isImageShape(s)) continue;
        if (!isWithinEpsilon((s as any).width, tw)) {
          try { (s as any).lockAspectRatio = settings.lockAspectRatio; } catch {}
          if (settings.applyWidth) (s as any).width = tw;
          if (settings.applyHeight) (s as any).height = th;
          await ctx.sync();
          state.lastInlineCount = ic; state.lastShapeCount = sc;
          log.info("global-scan: resized Shape");
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
      // 快速路径：获取最后一张图片并调整
      const r = await resizeLastImage(settings, inlineIncreased, shapeIncreased, inlineCount, shapeCount);
      if (r !== "none") { 
        state.processing = false; 
        onResult(r, "fast-path"); 
        return; 
      }
      
      // 快速路径失败，立即执行全局扫描
      log.debug("fast-path failed, run global-scan");
      const fr = await resizeByGlobalScan(settings);
      if (fr !== "none") { 
        state.processing = false; 
        onResult(fr, "immediate-fallback"); 
        return; 
      }
      
      // 立即兜底也失败，300ms 后再次兜底
      state.fallbackTimer = window.setTimeout(async () => {
        state.fallbackTimer = null;
        try { 
          const fr2 = await resizeByGlobalScan(settings); 
          onResult(fr2, "delayed-fallback"); 
        }
        catch { onResult("none", "error"); }
        finally { state.processing = false; }
      }, 300);
    } catch (e) { 
      log.error("handleCountIncrease error", e);
      state.processing = false; 
      onResult("none", "error"); 
    }
  })();
}

export function cancelPendingDetection(): void {
  if (state.fallbackTimer !== null) { window.clearTimeout(state.fallbackTimer); state.fallbackTimer = null; }
  state.processing = false;
}