/**
 * 生效范围设置 - 属性测试
 * 
 * Property 1: 生效范围设置的正确性
 * 对于任意生效范围设置操作，设置值应正确保存并在下次读取时恢复。
 * 验证: 需求 1.2, 1.3, 1.5, 8.1, 8.2
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import * as fc from "fast-check";
import { ResizeScope } from "./types";

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

// 动态导入模块
const { getResizeScope, saveResizeScope } = await import("../settings");

// 生效范围值的 Arbitrary
const resizeScopeArb = fc.oneof(
  fc.constant("all" as ResizeScope),
  fc.constant("new" as ResizeScope),
  fc.constant(null as ResizeScope)
);

describe("生效范围设置 - 属性测试", () => {
  beforeEach(() => {
    Object.keys(mockStorage).forEach((key) => delete mockStorage[key]);
    vi.clearAllMocks();
  });

  /**
   * Property 1: 生效范围设置的正确性
   * 对于任意生效范围设置操作，设置值应正确保存并在下次读取时恢复。
   * 验证: 需求 1.2, 1.3, 1.5, 8.1, 8.2
   */
  it("Property 1: 保存后读取应返回相同的值", async () => {
    await fc.assert(
      fc.asyncProperty(resizeScopeArb, async (scope) => {
        // 保存设置
        await saveResizeScope(scope);
        
        // 读取设置
        const result = getResizeScope();
        
        // 验证
        expect(result).toBe(scope);
      }),
      { numRuns: 100 }
    );
  });

  it("Property 1: 多次保存应保留最后一次的值", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(resizeScopeArb, { minLength: 1, maxLength: 10 }),
        async (scopes) => {
          // 依次保存多个值
          for (const scope of scopes) {
            await saveResizeScope(scope);
          }
          
          // 读取应返回最后一个值
          const result = getResizeScope();
          const lastScope = scopes[scopes.length - 1];
          
          expect(result).toBe(lastScope);
        }
      ),
      { numRuns: 100 }
    );
  });

  it("Property 1: 无效值应返回 null", () => {
    // 测试各种无效值
    const invalidValues = ["invalid", "ALL", "NEW", "true", "false", "123", ""];
    
    for (const value of invalidValues) {
      mockStorage["opw_resizeScope"] = value;
      const result = getResizeScope();
      expect(result).toBe(null);
    }
  });

  it("Property 1: 存储为空时应返回 null", () => {
    const result = getResizeScope();
    expect(result).toBe(null);
  });
});

describe("生效范围设置 - 单元测试", () => {
  beforeEach(() => {
    Object.keys(mockStorage).forEach((key) => delete mockStorage[key]);
    vi.clearAllMocks();
  });

  it("应该正确保存 'all' 值", async () => {
    await saveResizeScope("all");
    expect(mockLocalStorage.setItem).toHaveBeenCalledWith(
      "opw_resizeScope",
      "all"
    );
  });

  it("应该正确保存 'new' 值", async () => {
    await saveResizeScope("new");
    expect(mockLocalStorage.setItem).toHaveBeenCalledWith(
      "opw_resizeScope",
      "new"
    );
  });

  it("应该正确保存 null 值（转为空字符串）", async () => {
    await saveResizeScope(null);
    expect(mockLocalStorage.setItem).toHaveBeenCalledWith(
      "opw_resizeScope",
      ""
    );
  });

  it("应该正确读取 'all' 值", () => {
    mockStorage["opw_resizeScope"] = "all";
    expect(getResizeScope()).toBe("all");
  });

  it("应该正确读取 'new' 值", () => {
    mockStorage["opw_resizeScope"] = "new";
    expect(getResizeScope()).toBe("new");
  });

  it("空字符串应返回 null", () => {
    mockStorage["opw_resizeScope"] = "";
    expect(getResizeScope()).toBe(null);
  });
});
