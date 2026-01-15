/**
 * 粘贴检测与内存表集成的属性测试
 * 
 * Property 3: 新图片 ID 记录的正确性
 * - 验证粘贴检测时正确记录 Shape ID
 * - 验证只在"新图片"模式下记录
 */

import { describe, it, expect, beforeEach } from "vitest";
import * as fc from "fast-check";
import {
  createRegistry,
  clearRegistry,
  recordShapeId,
  hasShapeId,
  getShapeIdCount,
  getRegistry,
} from "./registry";
import { ResizeScope } from "./types";

describe("粘贴检测 - 新图片 ID 记录", () => {
  beforeEach(() => {
    clearRegistry();
  });

  describe("Property 3.1: Shape ID 记录的正确性", () => {
    it("记录的 Shape ID 应该可以被查询到", () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 100000 }), { minLength: 1, maxLength: 50 }),
          (shapeIds) => {
            createRegistry();

            // 模拟粘贴检测时记录 ID
            for (const id of shapeIds) {
              recordShapeId(id);
            }

            // 验证所有记录的 ID 都可以查询到
            for (const id of shapeIds) {
              expect(hasShapeId(id)).toBe(true);
            }

            // 验证数量正确（去重后）
            const uniqueIds = new Set(shapeIds);
            expect(getShapeIdCount()).toBe(uniqueIds.size);
          }
        ),
        { numRuns: 30 }
      );
    });

    it("未记录的 Shape ID 不应该被查询到", () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 50000 }), { minLength: 1, maxLength: 20 }),
          fc.array(fc.integer({ min: 50001, max: 100000 }), { minLength: 1, maxLength: 20 }),
          (recordedIds, unrecordedIds) => {
            createRegistry();

            // 只记录第一组 ID
            for (const id of recordedIds) {
              recordShapeId(id);
            }

            // 验证第二组 ID 不在内存表中
            for (const id of unrecordedIds) {
              expect(hasShapeId(id)).toBe(false);
            }
          }
        ),
        { numRuns: 30 }
      );
    });
  });

  describe("Property 3.2: 模式条件检查", () => {
    it("只有在'新图片'模式且内存表存在时才应该记录", () => {
      fc.assert(
        fc.property(
          fc.constantFrom<ResizeScope>("all", "new", null),
          fc.integer({ min: 1, max: 100000 }),
          fc.boolean(),
          (scope, shapeId, registryExists) => {
            // 设置内存表状态
            if (registryExists) {
              createRegistry();
            } else {
              clearRegistry();
            }

            // 模拟粘贴检测的条件判断
            const shouldRecord = scope === "new" && getRegistry() !== null;

            if (shouldRecord) {
              recordShapeId(shapeId);
              expect(hasShapeId(shapeId)).toBe(true);
            } else {
              // 不应该记录，但如果内存表不存在，recordShapeId 会静默失败
              if (getRegistry() !== null) {
                // 内存表存在但模式不是 "new"，不记录
                expect(hasShapeId(shapeId)).toBe(false);
              }
            }
          }
        ),
        { numRuns: 30 }
      );
    });
  });

  describe("Property 3.3: 重复 ID 处理", () => {
    it("重复记录同一个 ID 不应该增加计数", () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 1, max: 100000 }),
          fc.integer({ min: 2, max: 10 }),
          (shapeId, repeatCount) => {
            createRegistry();

            // 重复记录同一个 ID
            for (let i = 0; i < repeatCount; i++) {
              recordShapeId(shapeId);
            }

            // 验证只记录了一次
            expect(getShapeIdCount()).toBe(1);
            expect(hasShapeId(shapeId)).toBe(true);
          }
        ),
        { numRuns: 20 }
      );
    });
  });

  describe("Property 3.4: 内存表生命周期", () => {
    it("清空内存表后所有 ID 都应该被移除", () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 100000 }), { minLength: 1, maxLength: 30 }),
          (shapeIds) => {
            createRegistry();

            // 记录一些 ID
            for (const id of shapeIds) {
              recordShapeId(id);
            }
            expect(getShapeIdCount()).toBeGreaterThan(0);

            // 清空内存表
            clearRegistry();

            // 验证所有 ID 都被移除
            expect(getRegistry()).toBeNull();
            for (const id of shapeIds) {
              expect(hasShapeId(id)).toBe(false);
            }
          }
        ),
        { numRuns: 20 }
      );
    });

    it("重新创建内存表后应该是空的", () => {
      fc.assert(
        fc.property(
          fc.array(fc.integer({ min: 1, max: 100000 }), { minLength: 1, maxLength: 20 }),
          (shapeIds) => {
            // 第一次创建并记录
            createRegistry();
            for (const id of shapeIds) {
              recordShapeId(id);
            }
            const countBefore = getShapeIdCount();

            // 清空并重新创建
            clearRegistry();
            createRegistry();

            // 验证新内存表是空的
            expect(getShapeIdCount()).toBe(0);
            expect(countBefore).toBeGreaterThan(0);
          }
        ),
        { numRuns: 20 }
      );
    });
  });
});
