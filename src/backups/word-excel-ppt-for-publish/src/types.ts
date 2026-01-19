/**
 * Office Add-in 类型定义
 */

// 调整结果类型
export type AdjustResult =
  | "word-inline"
  | "word-selection-shape"
  | "word-body-last-inline"
  | "word-body-last-picture"
  | "ppt-shape"
  | "excel-shape"
  | "none";

// 功能设置
export interface FeatureSettings {
  resizeOnPaste: boolean;
  resizeOnSelection: boolean;
}

// 尺寸设置
export interface ResizeSettings {
  targetWidthCm: number;
  setHeightEnabled: boolean;
  targetHeightCm: number;
  applyWidth: boolean;
  applyHeight: boolean;
  lockAspectRatio: boolean;
}

// 配置常量
export const CONFIG = {
  // 设置键名
  keys: {
    enabled: "opw_enabled",
    resizeOnPaste: "opw_resizeOnPaste",
    resizeOnSelection: "opw_resizeOnSelection",
    targetWidthCm: "opw_targetWidthCm",
    setHeightEnabled: "opw_setHeightEnabled",
    targetHeightCm: "opw_targetHeightCm",
    resizeScope: "opw_resizeScope",
    baselineSnapshot: "opw_baselineSnapshot",
    sizeSnapshot: "opw_sizeSnapshot",
  },

  // 默认值
  defaults: {
    enabled: true,
    resizeOnPaste: true,
    resizeOnSelection: false,
    targetWidthCm: 15,
    setHeightEnabled: false,
    targetHeightCm: 10,
    resizeScope: null as null,
  },

  // Word 相关时间配置
  word: {
    selection: {
      debounceMs: 100,
      throttleMs: 120,
      verifyQuietMs: 200,
    },
    watcher: {
      fastIntervalMs: 80,
      slowIntervalMs: 150,
      selectedIntervalMs: 500,
      pasteFastIntervalMs: 150,
      pasteSlowIntervalMs: 300,
    },
    burst: {
      intervalMs: 120,
      pasteIntervalMs: 120,
      noneStreakMax: 12,
      pasteNoneStreakMax: 20,
    },
    suppress: {
      pictureMs: 2000,
      pasteSelectionMs: 800,
    },
    statusErrorThrottleMs: 2000,
    sizeEpsilonPts: 2,
  },

  // 其他应用节流
  otherSelectionThrottleMs: 800,
} as const;

// Word 状态管理
export interface WordState {
  lastAdjustAt: number;
  lastBodyShapeCount: number;
  lastBodyInlinePictureCount: number;
  pollingInFlight: boolean;
  bodyInlineIds: Set<string>;
  bodyShapeIds: Set<string>;
  suppressPollingUntil: number;
  suppressSelectionUntil: number;

  // 粘贴突发检测
  pasteBurstTimer: number | null;
  pasteBurstUntil: number;
  pasteBurstNoneStreak: number;
  pasteAnchorRange: Word.Range | null;

  // 粘贴确认
  pasteConfirmSeq: number;
  pasteConfirmTimers: number[];

  // 计数监视器
  countWatcherTimer: number | null;
  countWatcherInFlight: boolean;
  watcherInlineCount: number;
  watcherShapeCount: number;
  watcherFastUntil: number;

  // 选择状态
  pictureSelected: boolean;
  lastWatcherErrorAt: number;
  selectionDebounceTimer: number | null;
  selectionVerifyTimer: number | null;
  lastSelectionEventAt: number;
  pasteSelectionProbeInFlight: boolean;
  pasteLastSelectionAutoResizeAt: number;
}

// 创建初始 Word 状态
export function createWordState(): WordState {
  return {
    lastAdjustAt: 0,
    lastBodyShapeCount: 0,
    lastBodyInlinePictureCount: 0,
    pollingInFlight: false,
    bodyInlineIds: new Set<string>(),
    bodyShapeIds: new Set<string>(),
    suppressPollingUntil: 0,
    suppressSelectionUntil: 0,

    pasteBurstTimer: null,
    pasteBurstUntil: 0,
    pasteBurstNoneStreak: 0,
    pasteAnchorRange: null,

    pasteConfirmSeq: 0,
    pasteConfirmTimers: [],

    countWatcherTimer: null,
    countWatcherInFlight: false,
    watcherInlineCount: 0,
    watcherShapeCount: 0,
    watcherFastUntil: 0,

    pictureSelected: false,
    lastWatcherErrorAt: 0,
    selectionDebounceTimer: null,
    selectionVerifyTimer: null,
    lastSelectionEventAt: 0,
    pasteSelectionProbeInFlight: false,
    pasteLastSelectionAutoResizeAt: 0,
  };
}
