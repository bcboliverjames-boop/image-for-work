/**
 * PowerPoint 图片调整模块
 */

import { ResizeSettings, AdjustResult } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";
import { isReferenceBoxName } from "../reference-box/types";

function isPptImageShape(shape: PowerPoint.Shape): boolean {
  const shapeType = (shape as any).type;
  if (shapeType) {
    const typeString = String(shapeType).toLowerCase();
    if (typeString.includes("image") || typeString.includes("picture")) return true;
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
    if (nameLower.includes("picture") || nameLower.includes("image")) return true;
  }

  return false;
}

/**
 * 调整 PowerPoint 中选中的图片尺寸
 */
export async function adjustSelectedPptShape(
  settings: ResizeSettings
): Promise<AdjustResult> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);

  const found = await PowerPoint.run(
    async (context: PowerPoint.RequestContext) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (shapes.items.length < 1) return false;

      const shape = shapes.items[0];
      shape.load(["width", "height", "id", "name"]);
      await context.sync();

      if (isReferenceBoxName((shape as any).name)) return false;
      if (!isPptImageShape(shape)) return false;

      let origLeft: number | null = null;
      let origTop: number | null = null;
      try {
        (shape as any).load(["left", "top"]);
        await context.sync();
        const l = Number((shape as any).left);
        const t = Number((shape as any).top);
        if (Number.isFinite(l)) origLeft = l;
        if (Number.isFinite(t)) origTop = t;
      } catch {
        // ignore
      }

      // 如果尺寸已经正确，跳过
      if (
        settings.applyWidth &&
        !settings.applyHeight &&
        isWithinEpsilon(shape.width, targetWidthPts)
      ) {
        return true;
      }

      // 设置宽高比锁定
      try {
        (shape as any).lockAspectRatio = settings.lockAspectRatio;
      } catch {
        // 忽略
      }

      // 应用尺寸
      let expectedW: number | null = null;
      let expectedH: number | null = null;
      if (settings.lockAspectRatio && settings.applyWidth && !settings.applyHeight) {
        const ow = shape.width;
        const oh = shape.height;
        const scale = targetWidthPts / ow;
        const newW = targetWidthPts;
        const newH = oh * scale;
        expectedW = newW;
        expectedH = newH;
        if (Number.isFinite(newW) && newW > 0) shape.width = newW;
        if (Number.isFinite(newH) && newH > 0) shape.height = newH;
      } else if (settings.lockAspectRatio && !settings.applyWidth && settings.applyHeight) {
        const ow = shape.width;
        const oh = shape.height;
        const scale = targetHeightPts / oh;
        const newW = ow * scale;
        const newH = targetHeightPts;
        expectedW = newW;
        expectedH = newH;
        if (Number.isFinite(newW) && newW > 0) shape.width = newW;
        if (Number.isFinite(newH) && newH > 0) shape.height = newH;
      } else {
        expectedW = settings.applyWidth ? targetWidthPts : null;
        expectedH = settings.applyHeight ? targetHeightPts : null;
        if (settings.applyWidth) shape.width = targetWidthPts;
        if (settings.applyHeight) shape.height = targetHeightPts;
      }

      if (origLeft !== null) (shape as any).left = Math.max(0, origLeft);
      if (origTop !== null) (shape as any).top = Math.max(0, origTop);

      await context.sync();

      try {
        shape.load(["width", "height"]);
        await context.sync();
        const okWidth = expectedW === null || isWithinEpsilon(shape.width, expectedW);
        const okHeight = expectedH === null || isWithinEpsilon(shape.height, expectedH);
        if (!okWidth || !okHeight) return false;
      } catch {
        return false;
      }
      return true;
    }
  );

  return found ? "ppt-shape" : "none";
}
