/**
 * Office Paste Width - 任务窗格入口
 * 方案 A：轮询检测 + 快速路径 + 兜底扫描（功能完美但有闪烁）
 * 
 * 备份版本 - 保留原有轮询方案
 */

import { CONFIG, AdjustResult, createWordState, WordState } from "./types";
import { setStatus, hasOfficeContext } from "./utils";
import {
  getBoolSetting,
  getNumberSetting,
  saveSetting,
  getFeatureSettings,
  getResizeSettings,
  isEnabled,
} from "./settings";
import { adjustSelectedWordObject } from "./resize/word";
import { adjustSelectedExcelShape } from "./resize/excel";
import { adjustSelectedPptShape } from "./resize/powerpoint";
import {
  initPasteBaseline,
  checkCountChange,
  handleCountIncrease,
} from "./resize/word-paste-detector";

const wordState: WordState = createWordState();

async function adjustSelectedObject(): Promise<AdjustResult> {
  const host = Office.context.host;
  const settings = getResizeSettings();
  if (host === Office.HostType.Word) {
    return await Word.run(async (context) => {
      return await adjustSelectedWordObject(context, settings);
    });
  }
  if (host === Office.HostType.Excel) {
    return await adjustSelectedExcelShape(settings);
  }
  if (host === Office.HostType.PowerPoint) {
    return await adjustSelectedPptShape(settings);
  }
  return "none";
}

async function wordCountWatcherTick(): Promise<void> {
  if (!hasOfficeContext()) return;
  if (Office.context.host !== Office.HostType.Word) return;
  if (wordState.countWatcherInFlight) return;
  if (Date.now() < wordState.suppressPollingUntil) return;
  if (!isEnabled()) return;
  const { resizeOnPaste } = getFeatureSettings();
  if (!resizeOnPaste) return;

  try {
    wordState.countWatcherInFlight = true;
    const { inlineIncreased, shapeIncreased, inlineCount, shapeCount } = await checkCountChange();

    if (inlineIncreased || shapeIncreased) {
      const settings = getResizeSettings();
      handleCountIncrease(settings, inlineCount, shapeCount, (result, isFallback) => {
        if (result !== "none") {
          wordState.lastAdjustAt = Date.now();
          const method = isFallback ? "（兜底）" : "（快速）";
          setStatus(`已启用。目标宽度: ${settings.targetWidthCm} cm\n已调整: ${result} ${method}\n时间: ${new Date().toLocaleTimeString()}`);
        }
      });
      setWordWatcherFastWindow(2000);
    }
  } catch (e) {
    const now = Date.now();
    if (now - wordState.lastWatcherErrorAt > CONFIG.word.statusErrorThrottleMs) {
      wordState.lastWatcherErrorAt = now;
      setStatus(`监视器错误: ${String(e)}`);
    }
  } finally {
    wordState.countWatcherInFlight = false;
  }
}

function startWordCountWatcher(): void {
  if (!hasOfficeContext()) return;
  if (Office.context.host !== Office.HostType.Word) return;
  if (wordState.countWatcherTimer !== null) return;
  scheduleNextWatcherTick(0);
}

function stopWordCountWatcher(): void {
  if (wordState.countWatcherTimer !== null) {
    window.clearTimeout(wordState.countWatcherTimer);
    wordState.countWatcherTimer = null;
  }
}

function scheduleNextWatcherTick(delayMs: number): void {
  if (Office.context.host !== Office.HostType.Word) return;
  if (wordState.countWatcherTimer !== null) {
    window.clearTimeout(wordState.countWatcherTimer);
    wordState.countWatcherTimer = null;
  }
  wordState.countWatcherTimer = window.setTimeout(() => {
    wordState.countWatcherTimer = null;
    void (async () => {
      await wordCountWatcherTick();
      const now = Date.now();
      const nextDelay = now < wordState.watcherFastUntil ? 50 : 150;
      scheduleNextWatcherTick(nextDelay);
    })();
  }, Math.max(0, delayMs));
}

function setWordWatcherFastWindow(durationMs: number): void {
  const now = Date.now();
  wordState.watcherFastUntil = Math.max(wordState.watcherFastUntil, now + durationMs);
  scheduleNextWatcherTick(0);
}

async function onSelectionChanged(): Promise<void> {
  if (!isEnabled()) return;
  const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();

  if (Office.context.host === Office.HostType.Word) {
    if (!resizeOnSelection && resizeOnPaste) {
      // 粘贴模式：SelectionChanged 触发快速轮询窗口
      setWordWatcherFastWindow(2000);
      return;
    }

    // 选择模式
    if (wordState.selectionDebounceTimer !== null) {
      window.clearTimeout(wordState.selectionDebounceTimer);
    }
    wordState.selectionDebounceTimer = window.setTimeout(() => {
      wordState.selectionDebounceTimer = null;
      void processSelectionChanged();
    }, CONFIG.word.selection.debounceMs);
    return;
  }

  await processSelectionChanged();
}

async function processSelectionChanged(): Promise<void> {
  if (!isEnabled()) return;
  const { resizeOnSelection } = getFeatureSettings();
  if (!resizeOnSelection) return;

  const now = Date.now();
  const threshold = Office.context.host === Office.HostType.Word
    ? CONFIG.word.selection.throttleMs
    : CONFIG.otherSelectionThrottleMs;
  if (now - wordState.lastAdjustAt < threshold) return;
  wordState.lastAdjustAt = now;

  const settings = getResizeSettings();
  try {
    const result = await adjustSelectedObject();
    setStatus(`已启用。目标宽度: ${settings.targetWidthCm} cm\n已调整: ${result}\n时间: ${new Date().toLocaleTimeString()}`);
  } catch (e) {
    setStatus(`错误: ${String(e)}`);
  }
}

function bindUi(): void {
  const elements = {
    enabled: document.getElementById("enabled") as HTMLInputElement | null,
    resizeOnPaste: document.getElementById("resizeOnPaste") as HTMLInputElement | null,
    resizeOnSelection: document.getElementById("resizeOnSelection") as HTMLInputElement | null,
    targetWidthCm: document.getElementById("targetWidthCm") as HTMLInputElement | null,
    setHeightEnabled: document.getElementById("setHeightEnabled") as HTMLInputElement | null,
    targetHeightCm: document.getElementById("targetHeightCm") as HTMLInputElement | null,
    lockAspectRatio: document.getElementById("lockAspectRatio") as HTMLInputElement | null,
  };

  if (Object.values(elements).some((el) => !el)) return;

  const { enabled, resizeOnPaste, resizeOnSelection, targetWidthCm, setHeightEnabled, targetHeightCm, lockAspectRatio } = 
    elements as { [K in keyof typeof elements]: HTMLInputElement };

  enabled.checked = getBoolSetting(CONFIG.keys.enabled, CONFIG.defaults.enabled);
  const feature = getFeatureSettings();
  resizeOnPaste.checked = feature.resizeOnPaste;
  resizeOnSelection.checked = feature.resizeOnSelection;
  targetWidthCm.value = String(getNumberSetting(CONFIG.keys.targetWidthCm, CONFIG.defaults.targetWidthCm));
  setHeightEnabled.checked = getBoolSetting(CONFIG.keys.setHeightEnabled, CONFIG.defaults.setHeightEnabled);
  targetHeightCm.value = String(getNumberSetting(CONFIG.keys.targetHeightCm, CONFIG.defaults.targetHeightCm));

  const updateLockAspectRatioUi = () => {
    const width = Number(targetWidthCm.value);
    const applyWidth = Number.isFinite(width) && width > 0;
    const height = Number(targetHeightCm.value);
    const applyHeight = setHeightEnabled.checked && Number.isFinite(height) && height > 0;
    lockAspectRatio.checked = (applyWidth && !applyHeight) || (!applyWidth && applyHeight);
  };
  updateLockAspectRatioUi();

  enabled.addEventListener("change", async () => {
    await saveSetting(CONFIG.keys.enabled, enabled.checked);
    setStatus(`已保存。启用: ${enabled.checked}`);
    if (Office.context.host === Office.HostType.Word) {
      const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
      if (enabled.checked && resizeOnPaste && !resizeOnSelection) {
        await initPasteBaseline();
        startWordCountWatcher();
      } else {
        stopWordCountWatcher();
      }
    }
  });

  const applyMode = async (mode: "paste" | "selection") => {
    const nextPaste = mode === "paste";
    const nextSelection = mode === "selection";
    resizeOnPaste.checked = nextPaste;
    resizeOnSelection.checked = nextSelection;
    await saveSetting(CONFIG.keys.resizeOnPaste, nextPaste);
    await saveSetting(CONFIG.keys.resizeOnSelection, nextSelection);
    setStatus(`已保存。粘贴模式: ${nextPaste}，选择模式: ${nextSelection}`);
    if (Office.context.host === Office.HostType.Word) {
      if (nextPaste && isEnabled()) {
        await initPasteBaseline();
        startWordCountWatcher();
      } else {
        stopWordCountWatcher();
      }
    }
  };

  resizeOnPaste.addEventListener("change", async () => {
    if (!resizeOnPaste.checked) { resizeOnPaste.checked = true; return; }
    await applyMode("paste");
  });

  resizeOnSelection.addEventListener("change", async () => {
    if (!resizeOnSelection.checked) { resizeOnSelection.checked = true; return; }
    await applyMode("selection");
  });

  const onWidthChanged = async () => {
    let next = Number(targetWidthCm.value);
    if (!Number.isFinite(next) || next < 0) return;
    if (!setHeightEnabled.checked && next <= 0) {
      next = CONFIG.defaults.targetWidthCm;
      targetWidthCm.value = String(next);
    }
    await saveSetting(CONFIG.keys.targetWidthCm, next);
    setStatus(`已保存。目标宽度: ${next} cm`);
    updateLockAspectRatioUi();
  };

  targetWidthCm.addEventListener("input", () => void onWidthChanged());
  targetWidthCm.addEventListener("change", () => void onWidthChanged());

  setHeightEnabled.addEventListener("change", async () => {
    await saveSetting(CONFIG.keys.setHeightEnabled, setHeightEnabled.checked);
    setStatus(`已保存。设置高度: ${setHeightEnabled.checked}`);
    updateLockAspectRatioUi();
  });

  const onHeightChanged = async () => {
    const next = Number(targetHeightCm.value);
    if (!Number.isFinite(next) || next <= 0) return;
    await saveSetting(CONFIG.keys.targetHeightCm, next);
    setStatus(`已保存。目标高度: ${next} cm`);
    updateLockAspectRatioUi();
  };

  targetHeightCm.addEventListener("input", () => void onHeightChanged());
  targetHeightCm.addEventListener("change", () => void onHeightChanged());

  lockAspectRatio.disabled = true;
}

Office.onReady(async () => {
  if (!hasOfficeContext()) {
    setStatus("此页面必须在 Word/Excel/PowerPoint 任务窗格中打开。");
    return;
  }

  bindUi();

  if (Office.context.host === Office.HostType.Word) {
    await initPasteBaseline();
    const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
    if (isEnabled() && resizeOnPaste && !resizeOnSelection) {
      startWordCountWatcher();
    }
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => void onSelectionChanged(),
    (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("就绪。粘贴图片后将自动调整尺寸。");
      } else {
        setStatus(`注册监听失败: ${result.error?.message ?? String(result.error)}`);
      }
    }
  );
});
