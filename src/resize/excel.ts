/**
 * Excel 图片调整模块
 */

import { ResizeSettings, AdjustResult } from "../types";
import { cmToPoints } from "../utils";

/**
 * 调整 Excel 中选中的图片尺寸
 */
export async function adjustSelectedExcelShape(
  settings: ResizeSettings
): Promise<AdjustResult> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);

  const found = await Excel.run(async (context: Excel.RequestContext) => {
    const shape = context.workbook.getActiveShapeOrNullObject();
    shape.load(["isNullObject", "type", "width", "height"]);
    await context.sync();

    if (shape.isNullObject) return false;

    // 检查是否为图片类型
    const typeString = String(shape.type || "").toLowerCase();
    const looksLikeImage =
      typeString.includes("image") || typeString.includes("picture");
    if (!looksLikeImage && typeof shape.type === "string") return false;

    // 如果尺寸已经正确，跳过
    if (
      settings.applyWidth &&
      !settings.applyHeight &&
      Math.abs(shape.width - targetWidthPts) < 1
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
    if (settings.applyWidth) shape.width = targetWidthPts;
    if (settings.applyHeight) shape.height = targetHeightPts;

    await context.sync();
    return true;
  });

  return found ? "excel-shape" : "none";
}
