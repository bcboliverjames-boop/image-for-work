/**
 * 参考框模块入口
 * 导出所有公共 API 和类型
 */

// 导出类型
export type { ReferenceBoxState, ReferenceBoxSize } from "./types";

// 导出配置和工具函数
export {
  REFERENCE_BOX_CONFIG,
  cmToPoints,
  pointsToCm,
  isReferenceBoxName,
} from "./types";

// 导出核心 API
export {
  insertReferenceBox,
  getReferenceBoxSize,
  removeReferenceBox,
  findReferenceBox,
  getReferenceBoxState,
} from "./core";
