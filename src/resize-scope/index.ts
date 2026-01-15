/**
 * 缩放生效范围模块 - 统一导出入口
 */

// 类型定义
export type { ResizeScope, NewImageRegistry, BaselineSnapshot, ScanResult } from "./types";

// 内存表管理
export {
  createRegistry,
  clearRegistry,
  getRegistry,
  recordShapeId,
  hasShapeId,
  getShapeIdCount,
} from "./registry";

// 基线快照管理
export {
  saveBaseline,
  getBaseline,
  clearBaseline,
  captureBaseline,
} from "./baseline";

// 尺寸快照管理（用于内联图片的新图片判断）
export {
  saveSizeSnapshot,
  getSizeSnapshot,
  clearSizeSnapshot,
  captureSizeSnapshot,
  isSizeInSnapshot,
  hasSizeSnapshot,
} from "./size-snapshot";
export type { SizeSnapshot, CaptureSizeSnapshotOptions } from "./size-snapshot";

// 全面扫描
export { scanAllImages } from "./scanner";
