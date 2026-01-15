/**
 * PowerPoint 图片调整模块
 */

import { ResizeSettings, AdjustResult } from "../types";
import { cmToPoints } from "../utils";

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
      shape.load(["width", "height"]);
      await context.sync();

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
    }
  );

  return found ? "ppt-shape" : "none";
}
