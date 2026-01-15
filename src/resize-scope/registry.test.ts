/**
 * 内存表管理模块 - 单元测试
 */

import { describe, it, expect, beforeEach } from "vitest";
import {
  createRegistry,
  clearRegistry,
  getRegistry,
  hasRegistry,
  recordShapeId,
  hasShapeId,
  getShapeIdCount,
  getAllShapeIds,
  resetRegistry,
} from "./registry";

describe("内存表管理模块", () => {
  // 每个测试前清空内存表
  beforeEach(() => {
    clearRegistry();
  });

  describe("createRegistry", () => {
    it("应该创建空的内存表", () => {
      createRegistry();
      expect(hasRegistry()).toBe(true);
      expect(getShapeIdCount()).toBe(0);
    });

    it("应该设置时间戳", () => {
      const before = Date.now();
      createRegistry();
      const after = Date.now();
      const registry = getRegistry();
      expect(registry).not.toBeNull();
      expect(registry!.settingsTimestamp).toBeGreaterThanOrEqual(before);
      expect(registry!.settingsTimestamp).toBeLessThanOrEqual(after);
    });
  });

  describe("clearRegistry", () => {
    it("应该清空内存表", () => {
      createRegistry();
      recordShapeId(123);
      clearRegistry();
      expect(hasRegistry()).toBe(false);
      expect(getRegistry()).toBeNull();
    });
  });

  describe("recordShapeId", () => {
    it("应该记录 Shape ID", () => {
      createRegistry();
      recordShapeId(100);
      recordShapeId(200);
      expect(getShapeIdCount()).toBe(2);
      expect(hasShapeId(100)).toBe(true);
      expect(hasShapeId(200)).toBe(true);
    });

    it("应该忽略重复的 ID", () => {
      createRegistry();
      recordShapeId(100);
      recordShapeId(100);
      expect(getShapeIdCount()).toBe(1);
    });

    it("内存表不存在时不应报错", () => {
      expect(() => recordShapeId(100)).not.toThrow();
      expect(hasShapeId(100)).toBe(false);
    });
  });

  describe("hasShapeId", () => {
    it("应该正确检查 ID 是否存在", () => {
      createRegistry();
      recordShapeId(100);
      expect(hasShapeId(100)).toBe(true);
      expect(hasShapeId(999)).toBe(false);
    });

    it("内存表不存在时应返回 false", () => {
      expect(hasShapeId(100)).toBe(false);
    });
  });

  describe("getShapeIdCount", () => {
    it("应该返回正确的数量", () => {
      createRegistry();
      expect(getShapeIdCount()).toBe(0);
      recordShapeId(1);
      expect(getShapeIdCount()).toBe(1);
      recordShapeId(2);
      recordShapeId(3);
      expect(getShapeIdCount()).toBe(3);
    });

    it("内存表不存在时应返回 0", () => {
      expect(getShapeIdCount()).toBe(0);
    });
  });

  describe("getAllShapeIds", () => {
    it("应该返回所有 ID 的副本", () => {
      createRegistry();
      recordShapeId(1);
      recordShapeId(2);
      const ids = getAllShapeIds();
      expect(ids.size).toBe(2);
      expect(ids.has(1)).toBe(true);
      expect(ids.has(2)).toBe(true);
    });

    it("返回的应该是副本，修改不影响原数据", () => {
      createRegistry();
      recordShapeId(1);
      const ids = getAllShapeIds();
      ids.add(999);
      expect(hasShapeId(999)).toBe(false);
    });

    it("内存表不存在时应返回空集合", () => {
      const ids = getAllShapeIds();
      expect(ids.size).toBe(0);
    });
  });

  describe("resetRegistry", () => {
    it("应该清空 ID 集合但保留内存表", () => {
      createRegistry();
      recordShapeId(1);
      recordShapeId(2);
      resetRegistry();
      expect(hasRegistry()).toBe(true);
      expect(getShapeIdCount()).toBe(0);
    });

    it("应该更新时间戳", () => {
      createRegistry();
      const registry1 = getRegistry();
      const ts1 = registry1!.settingsTimestamp;
      
      // 等待一小段时间确保时间戳不同
      const start = Date.now();
      while (Date.now() - start < 5) { /* 等待 */ }
      
      resetRegistry();
      const registry2 = getRegistry();
      expect(registry2!.settingsTimestamp).toBeGreaterThanOrEqual(ts1);
    });

    it("内存表不存在时不应报错", () => {
      expect(() => resetRegistry()).not.toThrow();
    });
  });

  describe("模式切换场景", () => {
    it("切换到新图片模式 -> 记录 ID -> 切换到全部图片模式", () => {
      // 切换到新图片模式
      createRegistry();
      expect(hasRegistry()).toBe(true);
      
      // 记录一些 ID
      recordShapeId(100);
      recordShapeId(200);
      expect(getShapeIdCount()).toBe(2);
      
      // 切换到全部图片模式
      clearRegistry();
      expect(hasRegistry()).toBe(false);
      expect(getShapeIdCount()).toBe(0);
    });

    it("保存尺寸设置时重置内存表", () => {
      createRegistry();
      recordShapeId(100);
      recordShapeId(200);
      
      // 保存设置时重置
      resetRegistry();
      expect(hasRegistry()).toBe(true);
      expect(getShapeIdCount()).toBe(0);
    });
  });
});
