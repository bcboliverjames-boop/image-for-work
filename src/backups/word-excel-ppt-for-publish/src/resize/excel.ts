/**
 * Excel 图片调整模块
 */

import { ResizeSettings, AdjustResult } from "../types";
import { cmToPoints, isWithinEpsilon } from "../utils";

/**
 * 调整 Excel 中选中的图片尺寸
 */
export async function adjustSelectedExcelShape(
  settings: ResizeSettings
): Promise<AdjustResult> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);

  const found = await Excel.run(async (context: Excel.RequestContext) => {
    let shape: Excel.Shape;
    try {
      shape = context.workbook.getActiveShapeOrNullObject();
      shape.load(["isNullObject", "type", "width", "height", "name", "id"]);
      await context.sync();
    } catch (e) {
      console.log("[adjustSelectedExcelShape] getActiveShapeOrNullObject failed", e);
      return false;
    }

    if (shape.isNullObject) {
      console.log("[adjustSelectedExcelShape] no active shape");
      return false;
    }

    console.log("[adjustSelectedExcelShape] found shape", {
      id: shape.id,
      name: shape.name,
      type: shape.type,
      width: shape.width,
      height: shape.height,
    });

    // 检查是否为图片类型
    const typeString = String(shape.type || "").toLowerCase();
    const looksLikeImage =
      typeString.includes("image") || typeString.includes("picture");
    if (!looksLikeImage && typeof shape.type === "string") {
      console.log("[adjustSelectedExcelShape] not an image", { type: shape.type });
      return false;
    }

    // 如果尺寸已经正确，跳过
    if (
      settings.applyWidth &&
      !settings.applyHeight &&
      isWithinEpsilon(shape.width, targetWidthPts)
    ) {
      console.log("[adjustSelectedExcelShape] size already correct");
      return true;
    }

    // 设置宽高比锁定（等比例缩放）
    // 应用尺寸
    const oldWidth = shape.width;
    const oldHeight = shape.height;

    const aspect = oldWidth > 0 ? oldHeight / oldWidth : undefined;

    const wantsAspect = Boolean(settings.lockAspectRatio);

    const applyWidthOnly = settings.applyWidth && !settings.applyHeight;
    const applyHeightOnly = !settings.applyWidth && settings.applyHeight;

    let expectedWidth = oldWidth;
    let expectedHeight = oldHeight;

    if (wantsAspect && aspect && applyWidthOnly) {
      expectedWidth = targetWidthPts;
      expectedHeight = targetWidthPts * aspect;
      shape.width = expectedWidth;
      shape.height = expectedHeight;
    } else if (wantsAspect && aspect && applyHeightOnly) {
      expectedHeight = targetHeightPts;
      expectedWidth = targetHeightPts / aspect;
      shape.height = expectedHeight;
      shape.width = expectedWidth;
    } else {
      if (settings.applyWidth) {
        expectedWidth = targetWidthPts;
        shape.width = expectedWidth;
      }
      if (settings.applyHeight) {
        expectedHeight = targetHeightPts;
        shape.height = expectedHeight;
      }
    }

    await context.sync();

    try {
      shape.load(["width", "height"]);
      await context.sync();

      const shouldVerifyWidth =
        settings.applyWidth || (wantsAspect && Boolean(aspect) && applyHeightOnly);
      const shouldVerifyHeight =
        settings.applyHeight || (wantsAspect && Boolean(aspect) && applyWidthOnly);

      const okWidth = !shouldVerifyWidth || isWithinEpsilon(shape.width, expectedWidth);
      const okHeight = !shouldVerifyHeight || isWithinEpsilon(shape.height, expectedHeight);
      if (!okWidth || !okHeight) {
        console.log("[adjustSelectedExcelShape] verify failed", {
          width: shape.width,
          height: shape.height,
          targetWidthPts,
          targetHeightPts,
          expectedWidth,
          expectedHeight,
          applyWidth: settings.applyWidth,
          applyHeight: settings.applyHeight,
        });
        return false;
      }
    } catch (e) {
      console.log("[adjustSelectedExcelShape] verify error", e);
      return false;
    }
    
    console.log("[adjustSelectedExcelShape] resized", {
      oldWidth,
      oldHeight,
      newWidth: targetWidthPts,
      applyHeight: settings.applyHeight,
      lockAspectRatio: settings.lockAspectRatio,
    });
    
    return true;
  });

  return found ? "excel-shape" : "none";
}
