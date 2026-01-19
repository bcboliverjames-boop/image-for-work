/**
 * Word 事件驱动粘贴检测器
 * 
 * 核心思路：
 * - 不做持续轮询，只在 SelectionChanged 事件后启动短暂检测窗口
 * - 检测序列：100ms → 200ms → 500ms（最多 3 次）
 * - 检测到图片增加立即调整，然后停止
 * - 完全消除闪烁问题
 */

import { ResizeSettings } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { isImageShape } from "./word";

// 检测时间配置
const CHECK_DELAYS = [100, 200, 500];  // 第 1/2/3 次检测的延迟

// 状态
interface EventDetectorState {
  // 基线计数
  baselineInline: number;
  baselineShape: number;
  
  // 检测序列状态
  checkTimer: number | null;
  checkIndex: number;
  
  // 防止并发
  checking: boolean;
}

const state: EventDetectorState = {
  baselineInline: 0,
  baselineShape: 0,
  checkTimer: null,
  checkIndex: 0,
  checking: false,
};

/**
 * 初始化基线（记录当前图片数量）
 */
export async function initEventBaseline(): Promise<void> {
  await Word.run(async (ctx) => {
    const { inlineCount, shapeCount } = await getImageCounts(ctx);
    state.baselineInline = inlineCount;
    state.baselineShape = shapeCount;
  });
}

/**
 * 获取当前图片数量
 */
async function getImageCounts(ctx: Word.RequestContext): Promise<{
  inlineCount: number;
  shapeCount: number;
}> {
  const pics = ctx.document.body.inlinePictures;
  const shapes = ctx.document.body.shapes;
  
  let inlineCount = 0;
  let shapeCount = 0;
  
  try {
    const icr = (pics as any).getCount();
    const scr = (shapes as any).getCount();
    await ctx.sync();
    inlineCount = (icr as any)?.value ?? 0;
    shapeCount = (scr as any)?.value ?? 0;
  } catch {
    pics.load("items");
    shapes.load("items");
    await ctx.sync();
    inlineCount = pics.items?.length ?? 0;
    shapeCount = shapes.items?.length ?? 0;
  }
  
  return { inlineCount, shapeCount };
}

/**
 * SelectionChanged 事件触发时调用
 * 启动检测序列
 */
export function onSelectionChangedForPaste(
  settings: ResizeSettings,
  onResult: (result: "inline" | "shape" | "none", checkIndex: number) => void
): void {
  // 取消之前的检测序列
  cancelCheckSequence();
  
  // 重置检测索引
  state.checkIndex = 0;
  
  // 启动新的检测序列
  scheduleNextCheck(settings, onResult);
}

/**
 * 调度下一次检测
 */
function scheduleNextCheck(
  settings: ResizeSettings,
  onResult: (result: "inline" | "shape" | "none", checkIndex: number) => void
): void {
  if (state.checkIndex >= CHECK_DELAYS.length) {
    // 检测序列结束，没有发现新图片
    return;
  }
  
  const delay = CHECK_DELAYS[state.checkIndex];
  
  state.checkTimer = window.setTimeout(() => {
    state.checkTimer = null;
    void runCheck(settings, onResult);
  }, delay);
}

/**
 * 执行一次检测
 */
async function runCheck(
  settings: ResizeSettings,
  onResult: (result: "inline" | "shape" | "none", checkIndex: number) => void
): Promise<void> {
  if (state.checking) return;
  state.checking = true;
  
  const currentCheckIndex = state.checkIndex;
  state.checkIndex++;
  
  try {
    const result = await checkAndResize(settings);
    
    if (result !== "none") {
      // 成功调整，停止检测序列
      onResult(result, currentCheckIndex);
      state.checking = false;
      return;
    }
    
    // 没有发现新图片，继续下一次检测
    state.checking = false;
    scheduleNextCheck(settings, onResult);
    
  } catch (e) {
    state.checking = false;
    // 出错时继续下一次检测
    scheduleNextCheck(settings, onResult);
  }
}

/**
 * 检测并调整新图片
 */
async function checkAndResize(settings: ResizeSettings): Promise<"inline" | "shape" | "none"> {
  const tw = cmToPoints(settings.targetWidthCm);
  const th = cmToPoints(settings.targetHeightCm);
  
  return await Word.run(async (ctx) => {
    const { inlineCount, shapeCount } = await getImageCounts(ctx);
    
    // 处理删除情况：如果数量减少，更新基线
    if (inlineCount < state.baselineInline) {
      state.baselineInline = inlineCount;
    }
    if (shapeCount < state.baselineShape) {
      state.baselineShape = shapeCount;
    }
    
    // 检测内联图片增加
    if (inlineCount > state.baselineInline) {
      const pics = ctx.document.body.inlinePictures;
      try {
        const lastPic = (pics as any).getItemAt(inlineCount - 1);
        (lastPic as any).load(["width", "height", "lockAspectRatio"]);
        await ctx.sync();
        
        if (!isWithinEpsilon((lastPic as any).width, tw)) {
          // 需要调整
          try {
            (lastPic as any).lockAspectRatio = settings.lockAspectRatio;
          } catch {}
          
          if (settings.applyWidth) (lastPic as any).width = tw;
          if (settings.applyHeight) (lastPic as any).height = th;
          await ctx.sync();
          
          // 更新基线
          state.baselineInline = inlineCount;
          state.baselineShape = shapeCount;
          return "inline";
        }
      } catch {}
    }
    
    // 检测浮动图片增加
    if (shapeCount > state.baselineShape) {
      const shapes = ctx.document.body.shapes;
      try {
        const lastShape = (shapes as any).getItemAt(shapeCount - 1);
        (lastShape as any).load(["type", "width", "height"]);
        await ctx.sync();
        
        if (isImageShape(lastShape) && !isWithinEpsilon((lastShape as any).width, tw)) {
          // 需要调整
          try {
            (lastShape as any).lockAspectRatio = settings.lockAspectRatio;
          } catch {}
          
          if (settings.applyWidth) (lastShape as any).width = tw;
          if (settings.applyHeight) (lastShape as any).height = th;
          await ctx.sync();
          
          // 更新基线
          state.baselineInline = inlineCount;
          state.baselineShape = shapeCount;
          return "shape";
        }
      } catch {}
    }
    
    // 没有发现需要调整的图片，但更新基线
    state.baselineInline = inlineCount;
    state.baselineShape = shapeCount;
    return "none";
  });
}

/**
 * 取消检测序列
 */
export function cancelCheckSequence(): void {
  if (state.checkTimer !== null) {
    window.clearTimeout(state.checkTimer);
    state.checkTimer = null;
  }
  state.checking = false;
}

/**
 * 获取当前状态（用于调试）
 */
export function getDetectorState(): Readonly<EventDetectorState> {
  return { ...state };
}
