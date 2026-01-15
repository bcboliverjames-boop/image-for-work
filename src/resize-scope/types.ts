/**
 * 缩放生效范围 - 类型定义
 */

/**
 * 生效范围类型
 * - "all": 全部图片模式，对所有图片应用缩放
 * - "new": 新图片模式，仅对设置保存后新增的图片应用缩放
 * - null: 未设置，首次保存时需要用户选择
 */
export type ResizeScope = "all" | "new" | null;

/**
 * 新图片内存表（方案1）
 * 仅存在于内存中，不持久化
 * 用于精确记录设置保存后新增的浮动图片 ID
 */
export interface NewImageRegistry {
  /** 新增浮动图片的 ID 集合（Shape.id） */
  shapeIds: Set<number>;
  /** 设置保存时间戳 */
  settingsTimestamp: number;
}

/**
 * 基线快照（方案2 - 兜底）
 * 持久化存储，用于内存表为空时的兜底方案
 */
export interface BaselineSnapshot {
  /** 内联图片数量 */
  inlinePictureCount: number;
  /** 浮动图片数量 */
  shapeCount: number;
  /** 记录时间戳 */
  timestamp: number;
}

/**
 * 全面扫描结果
 */
export interface ScanResult {
  /** 扫描的图片总数 */
  scannedCount: number;
  /** 调整的图片数量 */
  adjustedCount: number;
  /** 跳过的图片数量（尺寸已匹配） */
  skippedCount: number;
  /** 扫描模式 */
  mode: "all" | "new-registry" | "new-baseline" | "fallback-full";
}
