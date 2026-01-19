/**
 * 工具函数模块
 */

import { CONFIG } from "./types";

/**
 * 厘米转换为磅值
 */
export function cmToPoints(cm: number): number {
  return cm * 28.3464566929;
}

/**
 * 判断值是否在目标值的误差范围内
 */
export function isWithinEpsilon(
  value: unknown,
  target: number,
  epsilon = CONFIG.word.sizeEpsilonPts
): boolean {
  const v = Number(value);
  return Number.isFinite(v) && Math.abs(v - target) <= epsilon;
}

/**
 * 安全执行异步函数，出错时返回默认值
 */
export async function safeAsync<T>(
  fn: () => Promise<T>,
  fallback: T
): Promise<T> {
  try {
    return await fn();
  } catch {
    return fallback;
  }
}

/**
 * 安全执行同步函数，出错时返回默认值
 */
export function safeSync<T>(fn: () => T, fallback: T): T {
  try {
    return fn();
  } catch {
    return fallback;
  }
}

/**
 * 更新状态栏显示
 */
export function setStatus(message: string): void {
  const el = document.getElementById("status");
  if (el) el.textContent = message;
}

/**
 * 检查是否有有效的 Office 上下文
 */
export function hasOfficeContext(): boolean {
  const anyOffice = Office as unknown as { context?: unknown };
  const ctx = anyOffice.context as { document?: unknown } | undefined;
  return Boolean(ctx?.document);
}

/**
 * 防抖函数
 */
export function debounce<T extends (...args: unknown[]) => void>(
  fn: T,
  delayMs: number
): (...args: Parameters<T>) => void {
  let timer: number | null = null;
  return (...args: Parameters<T>) => {
    if (timer !== null) {
      window.clearTimeout(timer);
    }
    timer = window.setTimeout(() => {
      timer = null;
      fn(...args);
    }, delayMs);
  };
}

/**
 * 节流函数
 */
export function throttle<T extends (...args: unknown[]) => void>(
  fn: T,
  intervalMs: number
): (...args: Parameters<T>) => void {
  let lastCall = 0;
  return (...args: Parameters<T>) => {
    const now = Date.now();
    if (now - lastCall >= intervalMs) {
      lastCall = now;
      fn(...args);
    }
  };
}
