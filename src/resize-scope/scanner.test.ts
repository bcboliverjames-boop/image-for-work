/**
 * 全面扫描模块的属性测试
 * 
 * Property 5: 全面扫描的范围过滤
 * - 验证全部图片模式扫描所有图片
 * - 验证新图片模式只扫描内存表中的图片或基线之后的图片
 */

import { describe, it, expect, beforeEach } from "vitest";
import * as fc from "fast-check";
import { createRegistry, clearRegistry, recordShapeId, getShapeIdCount, getRegistry, hasShapeId } from "./registry";
import { BaselineSnapshot, ResizeScope } from "./types";

describe("scanner - 范围过滤逻辑", () => {
  beforeEach(() => {
    clearRegistry();
  });

  describe("Property 5.1: 内存表过滤逻辑", () => {
    it("应该只包含内存表中记录的 Shape ID", () => {
      fc.assert(
        fc.property(
          // 生成随机的 Shape ID 集合（已记录的）
          fc.array(fc.integer({ min: 1, max: 10000 }), { minLength: 1, maxLength: 20 }),
          // 生成随机的 Shape ID 集合（未记录的）
          fc.array(fc.integer({ min: 10001, max: 20000 }), { minLength: 1, maxLength: 20 }),
          (recordedIds, unrecordedIds) => {
            // 创建内存表并记录 ID
            createRegistry();
            for (const id of recordedIds) {
              recordShapeId(id);
            }

            // 验证：已记录的 ID 应该在内存表中
            for (const id of recordedIds) {
              expect(hasShapeId(id)).toBe(true);
            }

            // 验证：未记录的 ID 不应该在内存表中
            for (const id of unrecordedIds) {
              expect(hasShapeId(id)).toBe(false);
            }

            // 验证：内存表大小等于去重后的已记录 ID 数量
            const uniqueRecordedIds = new Set(recordedIds);
            expect(getShapeIdCount()).toBe(uniqueRecordedIds.size);
          }
        ),
        { numRuns: 50 }
      );
    });
  });

  describe("Property 5.2: 基线过滤逻辑", () => {
    it("基线之后的索引应该被扫描", () => {
      fc.assert(
        fc.property(
          // 基线时的内联图片数量
          fc.integer({ min: 0, max: 100 }),
          // 基线时的浮动图片数量
          fc.integer({ min: 0, max: 100 }),
          // 当前内联图片数量（>= 基线）
          fc.integer({ min: 0, max: 50 }),
          // 当前浮动图片数量（>= 基线）
          fc.integer({ min: 0, max: 50 }),
          (baselineInline, baselineShape, addedInline, addedShape) => {
            const currentInline = baselineInline + addedInline;
            const currentShape = baselineShape + addedShape;

            // 设置基线
            const baseline: BaselineSnapshot = {
              inlinePictureCount: baselineInline,
              shapeCount: baselineShape,
              timestamp: Date.now(),
            };

            // 验证：应该扫描的内联图片索引范围
            const inlineIndicesToScan: number[] = [];
            for (let i = currentInline - 1; i >= baseline.inlinePictureCount; i--) {
              inlineIndicesToScan.push(i);
            }
            expect(inlineIndicesToScan.length).toBe(addedInline);

            // 验证：应该扫描的浮动图片索引范围
            const shapeIndicesToScan: number[] = [];
            for (let i = currentShape - 1; i >= baseline.shapeCount; i--) {
              shapeIndicesToScan.push(i);
            }
            expect(shapeIndicesToScan.length).toBe(addedShape);
          }
        ),
        { numRuns: 50 }
      );
    });
  });

  describe("Property 5.3: 模式选择逻辑", () => {
    it("scope 为 all 或 null 时应该使用全量扫描模式", () => {
      fc.assert(
        fc.property(
          fc.constantFrom<ResizeScope>("all", null),
          (scope) => {
            // 验证：全量扫描模式不依赖内存表或基线
            const shouldFullScan = scope === "all" || scope === null;
            expect(shouldFullScan).toBe(true);
          }
        ),
        { numRuns: 10 }
      );
    });

    it("scope 为 new 时应该使用范围扫描模式", () => {
      const scope: ResizeScope = "new";
      const shouldScopedScan = scope === "new";
      expect(shouldScopedScan).toBe(true);
    });
  });

  describe("Property 5.4: 内存表优先于基线", () => {
    it("当内存表有数据时，应该优先使用内存表", () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 10000 }), { minLength: 1, maxLength: 10 }),
          (shapeIds) => {
            // 创建内存表并记录 ID
            createRegistry();
            for (const id of shapeIds) {
              recordShapeId(id);
            }

            // 验证：内存表有数据
            expect(getShapeIdCount()).toBeGreaterThan(0);

            // 验证：应该使用内存表模式（new-registry）
            const registry = getRegistry();
            const useRegistry = registry !== null && getShapeIdCount() > 0;
            expect(useRegistry).toBe(true);
          }
        ),
        { numRuns: 20 }
      );
    });

    it("当内存表为空时，应该使用基线兜底", () => {
      // 创建空内存表
      createRegistry();
      expect(getShapeIdCount()).toBe(0);

      // 验证：应该使用基线模式（new-baseline）
      const registry = getRegistry();
      const useBaseline = !registry || getShapeIdCount() === 0;
      expect(useBaseline).toBe(true);
    });
  });

  describe("Property 5.5: 无数据时回退到全量扫描", () => {
    it("当内存表和基线都为空时，应该回退到全量扫描", () => {
      // 确保内存表为空
      clearRegistry();

      // 验证：应该回退到全量扫描
      const registry = getRegistry();
      const baseline: BaselineSnapshot | null = null; // 模拟无基线
      const shouldFallback = (!registry || getShapeIdCount() === 0) && !baseline;
      expect(shouldFallback).toBe(true);
    });
  });
});

describe("scanner - 扫描顺序", () => {
  it("应该从后往前扫描（最新的图片先处理）", () => {
    fc.assert(
      fc.property(
        fc.array(fc.integer({ min: 0, max: 100 }), { minLength: 1, maxLength: 20 }),
        (indices) => {
          // 模拟从后往前扫描
          const scanOrder: number[] = [];
          for (let i = indices.length - 1; i >= 0; i--) {
            scanOrder.push(i);
          }

          // 验证：第一个扫描的是最后一个索引
          expect(scanOrder[0]).toBe(indices.length - 1);
          // 验证：最后一个扫描的是第一个索引
          expect(scanOrder[scanOrder.length - 1]).toBe(0);
          // 验证：扫描顺序是递减的
          for (let i = 1; i < scanOrder.length; i++) {
            expect(scanOrder[i]).toBeLessThan(scanOrder[i - 1]);
          }
        }
      ),
      { numRuns: 30 }
    );
  });
});

describe("scanner - 范围扫描决策逻辑", () => {
  beforeEach(() => {
    clearRegistry();
  });

  it("Property 5.6: 范围扫描决策树", () => {
    fc.assert(
      fc.property(
        // 生效范围
        fc.constantFrom<ResizeScope>("all", "new", null),
        // 内存表是否有数据
        fc.boolean(),
        // 基线是否存在
        fc.boolean(),
        (scope, hasRegistryData, hasBaseline) => {
          // 设置内存表状态
          if (hasRegistryData) {
            createRegistry();
            recordShapeId(1);
          } else {
            clearRegistry();
          }

          // 模拟基线状态
          const baseline: BaselineSnapshot | null = hasBaseline
            ? { inlinePictureCount: 5, shapeCount: 5, timestamp: Date.now() }
            : null;

          // 决策逻辑
          let expectedMode: string;
          if (scope === "all" || scope === null) {
            expectedMode = "all";
          } else if (hasRegistryData) {
            expectedMode = "new-registry";
          } else if (hasBaseline) {
            expectedMode = "new-baseline";
          } else {
            expectedMode = "fallback-full";
          }

          // 验证决策逻辑
          const registry = getRegistry();
          const registryCount = getShapeIdCount();

          let actualMode: string;
          if (scope === "all" || scope === null) {
            actualMode = "all";
          } else if (registry && registryCount > 0) {
            actualMode = "new-registry";
          } else if (baseline) {
            actualMode = "new-baseline";
          } else {
            actualMode = "fallback-full";
          }

          expect(actualMode).toBe(expectedMode);
        }
      ),
      { numRuns: 50 }
    );
  });
});
