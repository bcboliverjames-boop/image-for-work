/**
 * 设置管理模块
 */

import { CONFIG, FeatureSettings, ResizeSettings } from "./types";
import { ResizeScope } from "./resize-scope/types";

/**
 * 获取数字类型设置
 */
export function getNumberSetting(key: string, fallback: number): number {
  let raw: unknown;

  // 优先从 roamingSettings 读取
  try {
    const rs = (Office as any)?.context?.roamingSettings;
    raw = rs?.get ? rs.get(key) : undefined;
  } catch {
    raw = undefined;
  }

  // 回退到 localStorage
  if (raw === undefined) {
    try {
      raw = window.localStorage.getItem(key) ?? undefined;
    } catch {
      raw = undefined;
    }
  }

  const num = typeof raw === "number" ? raw : Number(raw);
  return Number.isFinite(num) ? num : fallback;
}

/**
 * 获取布尔类型设置
 */
export function getBoolSetting(key: string, fallback: boolean): boolean {
  let raw: unknown;

  try {
    const rs = (Office as any)?.context?.roamingSettings;
    raw = rs?.get ? rs.get(key) : undefined;
  } catch {
    raw = undefined;
  }

  if (raw === undefined) {
    try {
      raw = window.localStorage.getItem(key) ?? undefined;
    } catch {
      raw = undefined;
    }
  }

  if (typeof raw === "boolean") return raw;
  if (raw === "true") return true;
  if (raw === "false") return false;
  return fallback;
}

/**
 * 获取字符串类型设置
 */
export function getStringSetting(key: string, fallback: string): string {
  let raw: unknown;

  try {
    const rs = (Office as any)?.context?.roamingSettings;
    raw = rs?.get ? rs.get(key) : undefined;
  } catch {
    raw = undefined;
  }

  if (raw === undefined) {
    try {
      raw = window.localStorage.getItem(key) ?? undefined;
    } catch {
      raw = undefined;
    }
  }

  if (typeof raw === "string") return raw;
  return fallback;
}

/**
 * 保存设置
 */
export function saveSetting(key: string, value: unknown): Promise<void> {
  try {
    const rs = (Office as any)?.context?.roamingSettings;
    if (rs?.set && rs?.saveAsync) {
      rs.set(key, value);
      return new Promise((resolve, reject) => {
        rs.saveAsync((result: Office.AsyncResult<void>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(result.error);
        });
      });
    }
  } catch {
    // 忽略并回退
  }

  try {
    window.localStorage.setItem(key, String(value));
  } catch {
    // 忽略
  }

  return Promise.resolve();
}

/**
 * 获取功能设置（粘贴模式/选择模式）
 */
export function getFeatureSettings(): FeatureSettings {
  const resizeOnPaste = getBoolSetting(
    CONFIG.keys.resizeOnPaste,
    CONFIG.defaults.resizeOnPaste
  );
  const resizeOnSelection = getBoolSetting(
    CONFIG.keys.resizeOnSelection,
    CONFIG.defaults.resizeOnSelection
  );

  // 互斥逻辑：两个都开或都关时，默认使用粘贴模式
  if (resizeOnPaste && resizeOnSelection) {
    return { resizeOnPaste: true, resizeOnSelection: false };
  }
  if (!resizeOnPaste && !resizeOnSelection) {
    return { resizeOnPaste: true, resizeOnSelection: false };
  }

  return { resizeOnPaste, resizeOnSelection };
}

/**
 * 获取尺寸调整设置
 */
export function getResizeSettings(): ResizeSettings {
  const targetWidthCm = getNumberSetting(
    CONFIG.keys.targetWidthCm,
    CONFIG.defaults.targetWidthCm
  );
  const setHeightEnabled = getBoolSetting(
    CONFIG.keys.setHeightEnabled,
    CONFIG.defaults.setHeightEnabled
  );
  const targetHeightCm = getNumberSetting(
    CONFIG.keys.targetHeightCm,
    CONFIG.defaults.targetHeightCm
  );

  let applyWidth = Number.isFinite(targetWidthCm) && targetWidthCm > 0;
  const applyHeight =
    setHeightEnabled && Number.isFinite(targetHeightCm) && targetHeightCm > 0;

  // 如果宽高都没设置，回退到默认宽度
  const effectiveTargetWidthCm =
    !applyWidth && !applyHeight ? CONFIG.defaults.targetWidthCm : targetWidthCm;
  if (!applyWidth && !applyHeight) applyWidth = true;

  // 锁定宽高比：只设置宽度或只设置高度时锁定
  const lockAspectRatio =
    (applyWidth && !applyHeight) || (!applyWidth && applyHeight);

  return {
    targetWidthCm: effectiveTargetWidthCm,
    setHeightEnabled,
    targetHeightCm,
    applyWidth,
    applyHeight,
    lockAspectRatio,
  };
}

/**
 * 检查是否启用
 */
export function isEnabled(): boolean {
  return getBoolSetting(CONFIG.keys.enabled, CONFIG.defaults.enabled);
}

/**
 * 获取生效范围设置
 * @returns "all" | "new" | null
 */
export function getResizeScope(): ResizeScope {
  const raw = getStringSetting(CONFIG.keys.resizeScope, "");
  if (raw === "all" || raw === "new") return raw;
  return null;
}

/**
 * 保存生效范围设置
 * @param scope "all" | "new" | null
 */
export async function saveResizeScope(scope: ResizeScope): Promise<void> {
  await saveSetting(CONFIG.keys.resizeScope, scope ?? "");
}
