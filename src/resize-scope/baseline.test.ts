/**
 * 基线快照管理模块 - 单元测试
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { BaselineSnapshot } from "./types";

// Mock localStorage
const mockStorage: Record<string, string> = {};
const mockLocalStorage = {
  getItem: vi.fn((key: string) => mockStorage[key] ?? null),
  setItem: vi.fn((key: string, value: string) => {
    mockStorage[key] = value;
  }),
  removeItem: vi.fn((key: string) => {
    delete mockStorage[key];
  }),
  clear: vi.fn(() => {
    Object.keys(mockStorage).forEach((key) => delete mockStorage[key]);
  }),
};

// 设置全局 mock
vi.stubGlobal("window", { localStorage: mockLocalStorage });
vi.stubGlobal("Office", undefined);
vi.stubGlobal("Word", undefined);

// 动态导入模块（在 mock 设置后）
const { saveBaseline, getBaseline, clearBaseline, hasBaseline } = await import(
  "./baseline"
);

describe("基线快照管理模块", () => {
  beforeEach(() => {
    // 清空 mock storage
    Object.keys(mockStorage).forEach((key) => delete mockStorage[key]);
    vi.clearAllMocks();
  });

  describe("saveBaseline", () => {
    it("应该保存基线快照到 localStorage", async () => {
      const snapshot: BaselineSnapshot = {
        inlinePictureCount: 5,
        shapeCount: 3,
        timestamp: 1704067200000,
      };

      await saveBaseline(snapshot);

      expect(mockLocalStorage.setItem).toHaveBeenCalledWith(
        "opw_baselineSnapshot",
        JSON.stringify(snapshot)
      );
    });
  });

  describe("getBaseline", () => {
    it("应该读取并解析基线快照", () => {
      const snapshot: BaselineSnapshot = {
        inlinePictureCount: 5,
        shapeCount: 3,
        timestamp: 1704067200000,
      };
      mockStorage["opw_baselineSnapshot"] = JSON.stringify(snapshot);

      const result = getBaseline();

      expect(result).toEqual(snapshot);
    });

    it("存储为空时应返回 null", () => {
      const result = getBaseline();
      expect(result).toBeNull();
    });

    it("JSON 解析失败时应返回 null", () => {
      mockStorage["opw_baselineSnapshot"] = "invalid json";

      const result = getBaseline();

      expect(result).toBeNull();
    });

    it("数据结构不完整时应返回 null", () => {
      mockStorage["opw_baselineSnapshot"] = JSON.stringify({
        inlinePictureCount: 5,
        // 缺少 shapeCount 和 timestamp
      });

      const result = getBaseline();

      expect(result).toBeNull();
    });

    it("数据类型错误时应返回 null", () => {
      mockStorage["opw_baselineSnapshot"] = JSON.stringify({
        inlinePictureCount: "5", // 应该是 number
        shapeCount: 3,
        timestamp: 1704067200000,
      });

      const result = getBaseline();

      expect(result).toBeNull();
    });
  });

  describe("clearBaseline", () => {
    it("应该清除基线快照", async () => {
      mockStorage["opw_baselineSnapshot"] = JSON.stringify({
        inlinePictureCount: 5,
        shapeCount: 3,
        timestamp: 1704067200000,
      });

      await clearBaseline();

      expect(mockLocalStorage.setItem).toHaveBeenCalledWith(
        "opw_baselineSnapshot",
        ""
      );
    });
  });

  describe("hasBaseline", () => {
    it("存在有效基线时应返回 true", () => {
      mockStorage["opw_baselineSnapshot"] = JSON.stringify({
        inlinePictureCount: 5,
        shapeCount: 3,
        timestamp: 1704067200000,
      });

      expect(hasBaseline()).toBe(true);
    });

    it("不存在基线时应返回 false", () => {
      expect(hasBaseline()).toBe(false);
    });

    it("基线数据无效时应返回 false", () => {
      mockStorage["opw_baselineSnapshot"] = "invalid";

      expect(hasBaseline()).toBe(false);
    });
  });

  describe("基线快照数据完整性", () => {
    it("保存后读取应该得到相同的数据", async () => {
      const snapshot: BaselineSnapshot = {
        inlinePictureCount: 10,
        shapeCount: 5,
        timestamp: Date.now(),
      };

      await saveBaseline(snapshot);
      const result = getBaseline();

      expect(result).toEqual(snapshot);
    });

    it("应该正确处理边界值", async () => {
      const snapshot: BaselineSnapshot = {
        inlinePictureCount: 0,
        shapeCount: 0,
        timestamp: 0,
      };

      await saveBaseline(snapshot);
      const result = getBaseline();

      expect(result).toEqual(snapshot);
    });

    it("应该正确处理大数值", async () => {
      const snapshot: BaselineSnapshot = {
        inlinePictureCount: 999999,
        shapeCount: 888888,
        timestamp: Number.MAX_SAFE_INTEGER,
      };

      await saveBaseline(snapshot);
      const result = getBaseline();

      expect(result).toEqual(snapshot);
    });
  });
});
