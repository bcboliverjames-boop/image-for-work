/**
 * 参考框模块类型定义
 * 参考框是一个可视化辅助工具，用于直观地预览和设置图片的目标尺寸
 */

/**
 * 参考框状态
 */
export interface ReferenceBoxState {
  /** 参考框是否存在 */
  exists: boolean;
  /** 形状名称（用于后续操作） */
  shapeName: string | null;
  /** 当前宽度（厘米） */
  widthCm: number;
  /** 当前高度（厘米） */
  heightCm: number;
}

/**
 * 参考框尺寸
 */
export interface ReferenceBoxSize {
  /** 宽度（厘米） */
  widthCm: number;
  /** 高度（厘米） */
  heightCm: number;
}

/**
 * 参考框配置常量
 */
export const REFERENCE_BOX_CONFIG = {
  /** 形状名称前缀，用于识别参考框 */
  namePrefix: "OfficePasteWidthRefBox_",
  /** 默认宽度（厘米） */
  defaultWidthCm: 14.0,
  /** 默认高度（厘米） */
  defaultHeightCm: 10.0,
  /** 轮询间隔（毫秒） */
  pollIntervalMs: 200,
  /** 填充颜色（橙黄色，便于识别） */
  fillColor: "#FFCC00",
  /** 填充透明度（85% 透明） */
  fillTransparency: 0.85,
  /** 边框颜色（橙色） */
  lineColor: "#FF6600",
  /** 边框粗细（磅） */
  lineWeight: 1.5,
} as const;

/**
 * 厘米转磅的转换系数
 * 1 英寸 = 72 磅，1 英寸 = 2.54 厘米
 * 所以 1 厘米 = 72 / 2.54 ≈ 28.3465 磅
 */
export const CM_TO_POINTS = 72.0 / 2.54;

/**
 * 厘米转换为磅值
 */
export function cmToPoints(cm: number): number {
  return cm * CM_TO_POINTS;
}

/**
 * 磅值转换为厘米
 */
export function pointsToCm(points: number): number {
  return points / CM_TO_POINTS;
}

/**
 * 生成唯一的参考框名称
 */
export function generateReferenceBoxName(): string {
  const uuid = crypto.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  return `${REFERENCE_BOX_CONFIG.namePrefix}${uuid}`;
}

/**
 * 检查形状名称是否为参考框
 */
export function isReferenceBoxName(name: string | null | undefined): boolean {
  return typeof name === "string" && name.startsWith(REFERENCE_BOX_CONFIG.namePrefix);
}
