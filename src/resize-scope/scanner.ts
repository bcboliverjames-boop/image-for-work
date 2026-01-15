/**
 * 缩放生效范围 - 全面扫描模块
 * 
 * 根据生效范围设置执行全面扫描：
 * - 全部图片模式：扫描所有图片
 * - 新图片模式：优先使用内存表，兜底使用基线快照
 */

import { ResizeSettings } from "../types";
import { ScanResult, ResizeScope } from "./types";
import { getRegistry, getShapeIdCount } from "./registry";
import { getBaseline } from "./baseline";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { isImageShape } from "../resize/word";

/**
 * 执行全面扫描
 * @param settings 尺寸设置
 * @param scope 生效范围
 * @returns 扫描结果
 */
export async function scanAllImages(
  settings: ResizeSettings,
  scope: ResizeScope
): Promise<ScanResult> {
  // 全部图片模式或未设置时，执行全量扫描
  if (scope === "all" || scope === null) {
    return await fullScan(settings);
  }

  // 新图片模式
  return await scopedScan(settings);
}

/**
 * 全量扫描（现有逻辑）
 * 从后往前扫描所有图片
 */
async function fullScan(settings: ResizeSettings): Promise<ScanResult> {
  const result: ScanResult = {
    scannedCount: 0,
    adjustedCount: 0,
    skippedCount: 0,
    mode: "all",
  };

  // 检查是否在 Office 环境中
  if (typeof Word === "undefined" || !Word.run) {
    return result;
  }

  await Word.run(async (ctx) => {
    const tw = cmToPoints(settings.targetWidthCm);
    const th = cmToPoints(settings.targetHeightCm);

    // 扫描内联图片（从后往前）
    const pics = ctx.document.body.inlinePictures;
    pics.load("items");
    await ctx.sync();

    const inlineItems = pics.items || [];
    for (const p of inlineItems) {
      (p as any).load(["width", "height"]);
    }
    await ctx.sync();

    for (let i = inlineItems.length - 1; i >= 0; i--) {
      result.scannedCount++;
      const p = inlineItems[i];
      const w = Number((p as any).width);
      if (!isWithinEpsilon(w, tw)) {
        await resizeInlinePicture(p, settings, tw, th);
        result.adjustedCount++;
      } else {
        result.skippedCount++;
      }
    }
    await ctx.sync();

    // 扫描浮动图片（从后往前）
    const shapes = ctx.document.body.shapes;
    shapes.load("items");
    await ctx.sync();

    const shapeItems = shapes.items || [];
    for (const s of shapeItems) {
      (s as any).load(["type", "width", "height", "id"]);
    }
    await ctx.sync();

    for (let i = shapeItems.length - 1; i >= 0; i--) {
      const s = shapeItems[i];
      if (!isImageShape(s)) continue;
      result.scannedCount++;
      const w = Number((s as any).width);
      if (!isWithinEpsilon(w, tw)) {
        await resizeShape(s, settings, tw, th);
        result.adjustedCount++;
      } else {
        result.skippedCount++;
      }
    }
    await ctx.sync();
  });

  return result;
}

/**
 * 范围扫描（新图片模式）
 * 优先使用内存表，兜底使用基线快照
 */
async function scopedScan(settings: ResizeSettings): Promise<ScanResult> {
  const registry = getRegistry();
  const baseline = getBaseline();

  // 无数据时回退到全量扫描
  if ((!registry || getShapeIdCount() === 0) && !baseline) {
    const result = await fullScan(settings);
    result.mode = "fallback-full";
    return result;
  }

  const result: ScanResult = {
    scannedCount: 0,
    adjustedCount: 0,
    skippedCount: 0,
    mode: registry && getShapeIdCount() > 0 ? "new-registry" : "new-baseline",
  };

  // 检查是否在 Office 环境中
  if (typeof Word === "undefined" || !Word.run) {
    return result;
  }

  await Word.run(async (ctx) => {
    const tw = cmToPoints(settings.targetWidthCm);
    const th = cmToPoints(settings.targetHeightCm);

    // 内联图片：始终使用基线兜底（因为 InlinePicture 没有 ID）
    const pics = ctx.document.body.inlinePictures;
    pics.load("items");
    await ctx.sync();

    const inlineItems = pics.items || [];
    const inlineBaseline = baseline?.inlinePictureCount ?? 0;

    // 从后往前扫描，只扫描索引 >= 基线的
    for (let i = inlineItems.length - 1; i >= inlineBaseline; i--) {
      (inlineItems[i] as any).load(["width", "height"]);
    }
    await ctx.sync();

    for (let i = inlineItems.length - 1; i >= inlineBaseline; i--) {
      result.scannedCount++;
      const p = inlineItems[i];
      const w = Number((p as any).width);
      if (!isWithinEpsilon(w, tw)) {
        await resizeInlinePicture(p, settings, tw, th);
        result.adjustedCount++;
      } else {
        result.skippedCount++;
      }
    }
    await ctx.sync();

    // 浮动图片：优先用内存表，空则用基线兜底
    const shapes = ctx.document.body.shapes;
    shapes.load("items");
    await ctx.sync();

    const shapeItems = shapes.items || [];
    for (const s of shapeItems) {
      (s as any).load(["type", "width", "height", "id"]);
    }
    await ctx.sync();

    if (registry && getShapeIdCount() > 0) {
      // 方案1：使用内存表精确识别
      for (let i = shapeItems.length - 1; i >= 0; i--) {
        const s = shapeItems[i];
        if (!isImageShape(s)) continue;
        const shapeId = (s as any).id as number;
        if (!registry.shapeIds.has(shapeId)) continue;

        result.scannedCount++;
        const w = Number((s as any).width);
        if (!isWithinEpsilon(w, tw)) {
          await resizeShape(s, settings, tw, th);
          result.adjustedCount++;
        } else {
          result.skippedCount++;
        }
      }
    } else if (baseline) {
      // 方案2兜底：使用基线数量
      const shapeBaseline = baseline.shapeCount;
      for (let i = shapeItems.length - 1; i >= shapeBaseline; i--) {
        const s = shapeItems[i];
        if (!isImageShape(s)) continue;

        result.scannedCount++;
        const w = Number((s as any).width);
        if (!isWithinEpsilon(w, tw)) {
          await resizeShape(s, settings, tw, th);
          result.adjustedCount++;
        } else {
          result.skippedCount++;
        }
      }
    }
    await ctx.sync();
  });

  return result;
}

/**
 * 辅助函数：调整内联图片尺寸
 */
async function resizeInlinePicture(
  pic: any,
  settings: ResizeSettings,
  tw: number,
  th: number
): Promise<void> {
  // 如果同时设置宽度和高度，必须先解锁宽高比
  if (settings.applyWidth && settings.applyHeight) {
    try {
      (pic as any).lockAspectRatio = false;
    } catch {
      // 忽略
    }
  }
  if (settings.applyWidth) (pic as any).width = tw;
  if (settings.applyHeight) (pic as any).height = th;
}

/**
 * 辅助函数：调整浮动图片尺寸
 */
async function resizeShape(
  shape: any,
  settings: ResizeSettings,
  tw: number,
  th: number
): Promise<void> {
  // 如果同时设置宽度和高度，必须先解锁宽高比
  if (settings.applyWidth && settings.applyHeight) {
    try {
      (shape as any).lockAspectRatio = false;
    } catch {
      // 忽略
    }
  }
  if (settings.applyWidth) (shape as any).width = tw;
  if (settings.applyHeight) (shape as any).height = th;
}
