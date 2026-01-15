/**
 * 缩放生效范围 - 内存表管理模块
 * 
 * 内存表用于精确记录设置保存后新增的浮动图片 ID（Shape.id）
 * 仅存在于内存中，不持久化
 */

import { NewImageRegistry } from "./types";

/** 内存表实例（模块级单例） */
let registry: NewImageRegistry | null = null;

/**
 * 创建新的内存表
 * 在用户选择"新图片"模式或保存尺寸设置时调用
 */
export function createRegistry(): void {
  registry = {
    shapeIds: new Set<number>(),
    settingsTimestamp: Date.now(),
  };
}

/**
 * 清空内存表
 * 在用户切换到"全部图片"模式时调用
 */
export function clearRegistry(): void {
  registry = null;
}

/**
 * 获取内存表实例
 * @returns 内存表实例，如果未创建则返回 null
 */
export function getRegistry(): NewImageRegistry | null {
  return registry;
}

/**
 * 检查内存表是否存在
 */
export function hasRegistry(): boolean {
  return registry !== null;
}

/**
 * 记录新增浮动图片 ID
 * 在粘贴检测到新浮动图片时调用
 * @param id Shape.id 值
 */
export function recordShapeId(id: number): void {
  if (registry) {
    registry.shapeIds.add(id);
  }
}

/**
 * 检查 Shape ID 是否在内存表中
 * @param id Shape.id 值
 * @returns 是否存在
 */
export function hasShapeId(id: number): boolean {
  return registry?.shapeIds.has(id) ?? false;
}

/**
 * 获取内存表中的 Shape ID 数量
 * @returns ID 数量，如果内存表不存在则返回 0
 */
export function getShapeIdCount(): number {
  return registry?.shapeIds.size ?? 0;
}

/**
 * 获取所有记录的 Shape ID
 * @returns Shape ID 集合的副本，如果内存表不存在则返回空集合
 */
export function getAllShapeIds(): Set<number> {
  return registry ? new Set(registry.shapeIds) : new Set();
}

/**
 * 重置内存表（清空 ID 集合但保留时间戳）
 * 在保存尺寸设置时调用
 */
export function resetRegistry(): void {
  if (registry) {
    registry.shapeIds.clear();
    registry.settingsTimestamp = Date.now();
  }
}
