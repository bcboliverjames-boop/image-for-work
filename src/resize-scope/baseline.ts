/**
 * 缩放生效范围 - 基线快照管理模块
 * 
 * 基线快照用于记录设置保存时的图片数量，作为内存表的兜底方案
 * 持久化存储到 roamingSettings 或 localStorage
 */

import { BaselineSnapshot } from "./types";
import { saveSetting, getStringSetting } from "../settings";
import { CONFIG } from "../types";

const BASELINE_KEY = CONFIG.keys.baselineSnapshot;

/**
 * 保存基线快照
 * @param snapshot 基线快照数据
 */
export async function saveBaseline(snapshot: BaselineSnapshot): Promise<void> {
  await saveSetting(BASELINE_KEY, JSON.stringify(snapshot));
}

/**
 * 读取基线快照
 * @returns 基线快照数据，如果不存在或解析失败则返回 null
 */
export function getBaseline(): BaselineSnapshot | null {
  const raw = getStringSetting(BASELINE_KEY, "");
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    // 验证数据结构
    if (
      typeof parsed.inlinePictureCount === "number" &&
      typeof parsed.shapeCount === "number" &&
      typeof parsed.timestamp === "number"
    ) {
      return parsed as BaselineSnapshot;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * 清除基线快照
 */
export async function clearBaseline(): Promise<void> {
  await saveSetting(BASELINE_KEY, "");
}

/**
 * 捕获当前图片数量并保存为基线快照
 * 需要在 Word 环境中调用
 * @returns 创建的基线快照
 */
export async function captureBaseline(): Promise<BaselineSnapshot> {
  // 检查是否在 Office 环境中
  if (typeof Word === "undefined" || !Word.run) {
    // 非 Office 环境，返回空基线
    const snapshot: BaselineSnapshot = {
      inlinePictureCount: 0,
      shapeCount: 0,
      timestamp: Date.now(),
    };
    await saveBaseline(snapshot);
    return snapshot;
  }

  return await Word.run(async (ctx) => {
    const pics = ctx.document.body.inlinePictures;
    const shapes = ctx.document.body.shapes;

    let inlineCount = 0;
    let shapeCount = 0;

    try {
      // 尝试使用 getCount() 方法
      const icr = (pics as any).getCount();
      const scr = (shapes as any).getCount();
      await ctx.sync();
      inlineCount = (icr as any)?.value ?? 0;
      shapeCount = (scr as any)?.value ?? 0;
    } catch {
      // 回退到 items.length
      pics.load("items");
      shapes.load("items");
      await ctx.sync();
      inlineCount = pics.items?.length ?? 0;
      shapeCount = shapes.items?.length ?? 0;
    }

    const snapshot: BaselineSnapshot = {
      inlinePictureCount: inlineCount,
      shapeCount: shapeCount,
      timestamp: Date.now(),
    };

    await saveBaseline(snapshot);
    return snapshot;
  });
}

/**
 * 检查基线快照是否存在
 */
export function hasBaseline(): boolean {
  return getBaseline() !== null;
}
