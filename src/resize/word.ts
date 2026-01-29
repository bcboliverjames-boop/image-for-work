/**
 * Word 图片调整模块
 */

import { CONFIG, ResizeSettings, AdjustResult, WordState, createWordState } from "../types";
import { cmToPoints, isWithinEpsilon, setStatus, hasOfficeContext, safeAsync } from "../utils";
import { getFeatureSettings, getResizeSettings, isEnabled } from "../settings";

// Word 全局状态
export const wordState: WordState = createWordState();

/**
 * 应用内联图片尺寸
 */
export async function applyInlinePictureSize(
  context: Word.RequestContext,
  pic: any,
  settings: ResizeSettings
): Promise<boolean> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);

  (pic as any).load(["width", "height", "lockAspectRatio"]);
  await context.sync();

  let changed = false;

  // 如果同时设置宽度和高度，必须先解锁宽高比
  if (settings.applyWidth && settings.applyHeight) {
    try {
      (pic as any).lockAspectRatio = false;
      await context.sync();
      changed = true;
    } catch {
      // 忽略
    }
  }

  // 设置宽度
  if (settings.applyWidth) {
    try {
      const w = Number((pic as any).width);
      if (!Number.isFinite(w) || Math.abs(w - targetWidthPts) > CONFIG.word.sizeEpsilonPts) {
        (pic as any).width = targetWidthPts;
        changed = true;
      }
    } catch {
      // 忽略
    }
  }

  // 设置高度
  if (settings.applyHeight) {
    try {
      const h = Number((pic as any).height);
      if (!Number.isFinite(h) || Math.abs(h - targetHeightPts) > CONFIG.word.sizeEpsilonPts) {
        (pic as any).height = targetHeightPts;
        changed = true;
      }
    } catch {
      // 忽略
    }
  }

  if (changed) {
    await context.sync();

    // 验证修改是否生效
    try {
      (pic as any).load(["width", "height"]);
      await context.sync();
      const okWidth = !settings.applyWidth || isWithinEpsilon((pic as any).width, targetWidthPts);
      const okHeight = !settings.applyHeight || isWithinEpsilon((pic as any).height, targetHeightPts);
      if (!okWidth || !okHeight) return false;
    } catch {
      return false;
    }
  }

  return changed;
}

/**
 * 应用浮动图片（Shape）尺寸
 */
export async function applyShapeSize(
  context: Word.RequestContext,
  shape: any,
  settings: ResizeSettings
): Promise<boolean> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);

  let changed = false;

  try {
    (shape as any).load(["width", "height", "lockAspectRatio"]);
    await context.sync();
  } catch {
    // 忽略
  }

  // 如果同时设置宽度和高度，必须先解锁宽高比
  if (settings.applyWidth && settings.applyHeight) {
    try {
      (shape as any).lockAspectRatio = false;
      await context.sync();
      changed = true;
    } catch {
      // 忽略
    }
  } else {
    // 否则按照设置来处理宽高比锁定
    try {
      const currentLock = (shape as any).lockAspectRatio;
      if (typeof currentLock === "boolean" && currentLock !== settings.lockAspectRatio) {
        (shape as any).lockAspectRatio = settings.lockAspectRatio;
        changed = true;
      }
    } catch {
      // 忽略
    }
  }

  // 设置宽度
  if (settings.applyWidth) {
    try {
      const w = Number((shape as any).width);
      if (!Number.isFinite(w) || Math.abs(w - targetWidthPts) > CONFIG.word.sizeEpsilonPts) {
        (shape as any).width = targetWidthPts;
        changed = true;
      }
    } catch {
      // 忽略
    }
  }

  // 设置高度
  if (settings.applyHeight) {
    try {
      const h = Number((shape as any).height);
      if (!Number.isFinite(h) || Math.abs(h - targetHeightPts) > CONFIG.word.sizeEpsilonPts) {
        (shape as any).height = targetHeightPts;
        changed = true;
      }
    } catch {
      // 忽略
    }
  }

  if (changed) {
    try {
      await context.sync();

      // 验证修改
      (shape as any).load(["width", "height"]);
      await context.sync();
      const okWidth = !settings.applyWidth || isWithinEpsilon((shape as any).width, targetWidthPts);
      const okHeight = !settings.applyHeight || isWithinEpsilon((shape as any).height, targetHeightPts);
      if (!okWidth || !okHeight) return false;
    } catch {
      return false;
    }
  }

  return changed;
}

/**
 * 判断 Shape 是否为图片类型
 */
export function isImageShape(shape: any): boolean {
  const typeString = String((shape as any)?.type || "").toLowerCase();
  return typeString.includes("picture") || typeString.includes("image");
}

/**
 * 调整选中的 Word 对象尺寸
 */
export async function adjustSelectedWordObject(
  context: Word.RequestContext,
  settings: ResizeSettings
): Promise<AdjustResult> {
  const range = context.document.getSelection();
  const inlinePics = range.inlinePictures;
  inlinePics.load("items");

  // 先只 sync inlinePictures，避免 shapes 不支持导致整个 sync 失败
  await context.sync();

  let selectionShapes: any = null;
  try {
    selectionShapes = (range as any).shapes;
    if (selectionShapes?.load) selectionShapes.load("items");
    await context.sync();
  } catch {
    selectionShapes = null;
  }

  // 优先处理浮动图片
  if (selectionShapes?.items?.length >= 1) {
    const shape = selectionShapes.items[0];
    try {
      (shape as any).load("type");
      await context.sync();

      if (isImageShape(shape)) {
        const ok = await applyShapeSize(context, shape, settings);
        return ok ? "word-selection-shape" : "none";
      }
    } catch {
      // ignore shapes path if not supported
    }
  }

  // 处理内联图片
  if (inlinePics.items?.length >= 1) {
    const pic = inlinePics.items[0];
    const ok = await applyInlinePictureSize(context, pic, settings);
    return ok ? "word-inline" : "none";
  }

  return "none";
}
