/**
 * 尺寸快照模块
 * 
 * 记录保存尺寸时文档中所有图片的宽度和高度（去重）
 * 用于判断内联图片是否为"新图片"
 */

import { saveSetting, getStringSetting } from "../settings";
import { CONFIG } from "../types";

const SIZE_SNAPSHOT_KEY = CONFIG.keys.sizeSnapshot || "pasteWidth_sizeSnapshot";

/**
 * 尺寸快照数据结构
 */
export interface SizeSnapshot {
  widths: number[];   // 去重后的宽度列表（单位：points）
  heights: number[];  // 去重后的高度列表（单位：points）
  timestamp: number;  // 创建时间戳
}

/**
 * 保存尺寸快照
 */
export async function saveSizeSnapshot(snapshot: SizeSnapshot): Promise<void> {
  await saveSetting(SIZE_SNAPSHOT_KEY, JSON.stringify(snapshot));
}

/**
 * 读取尺寸快照
 */
export function getSizeSnapshot(): SizeSnapshot | null {
  const raw = getStringSetting(SIZE_SNAPSHOT_KEY, "");
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (
      Array.isArray(parsed.widths) &&
      Array.isArray(parsed.heights) &&
      typeof parsed.timestamp === "number"
    ) {
      return parsed as SizeSnapshot;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * 清除尺寸快照
 */
export async function clearSizeSnapshot(): Promise<void> {
  await saveSetting(SIZE_SNAPSHOT_KEY, "");
}

/**
 * 捕获尺寸快照的选项
 */
export interface CaptureSizeSnapshotOptions {
  targetWidthPt?: number;   // 目标宽度（points），会被加入快照
  targetHeightPt?: number;  // 目标高度（points），会被加入快照
}

/**
 * 捕获当前文档中所有图片的尺寸并保存为快照
 * 需要在 Word 环境中调用
 * 
 * @param options 可选参数，包含目标尺寸（会被加入快照，避免已调整的图片被误判为新图片）
 */
export async function captureSizeSnapshot(options?: CaptureSizeSnapshotOptions): Promise<SizeSnapshot> {
  const { targetWidthPt, targetHeightPt } = options || {};
  
  // 检查是否在 Office 环境中
  if (typeof Word === "undefined" || !Word.run) {
    const widths: number[] = [];
    const heights: number[] = [];
    
    // 即使不在 Office 环境，也要加入目标尺寸
    if (targetWidthPt && Number.isFinite(targetWidthPt) && targetWidthPt > 0) {
      widths.push(Math.round(targetWidthPt));
    }
    if (targetHeightPt && Number.isFinite(targetHeightPt) && targetHeightPt > 0) {
      heights.push(Math.round(targetHeightPt));
    }
    
    const snapshot: SizeSnapshot = {
      widths,
      heights,
      timestamp: Date.now(),
    };
    await saveSizeSnapshot(snapshot);
    return snapshot;
  }

  return await Word.run(async (ctx) => {
    const widthSet = new Set<number>();
    const heightSet = new Set<number>();

    // 首先加入目标尺寸（避免已调整的图片被误判为新图片）
    if (targetWidthPt && Number.isFinite(targetWidthPt) && targetWidthPt > 0) {
      widthSet.add(Math.round(targetWidthPt));
    }
    if (targetHeightPt && Number.isFinite(targetHeightPt) && targetHeightPt > 0) {
      heightSet.add(Math.round(targetHeightPt));
    }

    // 收集内联图片尺寸
    const pics = ctx.document.body.inlinePictures;
    pics.load("items");
    await ctx.sync();

    for (const pic of pics.items || []) {
      (pic as any).load(["width", "height"]);
    }
    await ctx.sync();

    for (const pic of pics.items || []) {
      const w = Math.round(Number((pic as any).width));
      const h = Math.round(Number((pic as any).height));
      if (Number.isFinite(w) && w > 0) widthSet.add(w);
      if (Number.isFinite(h) && h > 0) heightSet.add(h);
    }

    // 收集浮动图片尺寸
    const shapes = ctx.document.body.shapes;
    shapes.load("items");
    await ctx.sync();

    for (const shape of shapes.items || []) {
      (shape as any).load(["type", "width", "height"]);
    }
    await ctx.sync();

    for (const shape of shapes.items || []) {
      // 只处理图片类型的 Shape
      const shapeType = (shape as any).type;
      if (shapeType === "Image" || shapeType === 1) {
        const w = Math.round(Number((shape as any).width));
        const h = Math.round(Number((shape as any).height));
        if (Number.isFinite(w) && w > 0) widthSet.add(w);
        if (Number.isFinite(h) && h > 0) heightSet.add(h);
      }
    }

    const snapshot: SizeSnapshot = {
      widths: Array.from(widthSet),
      heights: Array.from(heightSet),
      timestamp: Date.now(),
    };

    await saveSizeSnapshot(snapshot);
    return snapshot;
  });
}

/**
 * 检查尺寸是否在快照中（判断是否为"老图片"）
 * @param width 图片宽度（points）
 * @param height 图片高度（points）
 * @param snapshot 尺寸快照
 * @param tolerance 容差（默认 1 point）
 * @returns true 表示尺寸在快照中（老图片），false 表示不在（新图片）
 */
export function isSizeInSnapshot(
  width: number,
  height: number,
  snapshot: SizeSnapshot,
  tolerance: number = 1
): boolean {
  const w = Math.round(width);
  const h = Math.round(height);

  // 检查宽度是否在快照中
  const widthInSnapshot = snapshot.widths.some(
    (sw) => Math.abs(sw - w) <= tolerance
  );

  // 检查高度是否在快照中
  const heightInSnapshot = snapshot.heights.some(
    (sh) => Math.abs(sh - h) <= tolerance
  );

  // 宽度和高度都在快照中，才认为是老图片
  return widthInSnapshot && heightInSnapshot;
}

/**
 * 检查尺寸快照是否存在
 */
export function hasSizeSnapshot(): boolean {
  return getSizeSnapshot() !== null;
}
