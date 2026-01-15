/**
 * Office Paste Width - Task Pane Entry
 * v7: Event-driven + 3-layer fallback (no polling)
 *
 * Logic:
 * - SelectionChanged triggers image count check
 * - 3-layer: fast path -> immediate fallback -> 300ms delayed fallback
 * - On success: update baseline
 */

import { CONFIG, AdjustResult, createWordState, WordState } from "./types";
import { setStatus, hasOfficeContext, cmToPoints, isWithinEpsilon } from "./utils";
import {
  getBoolSetting,
  getNumberSetting,
  saveSetting,
  getFeatureSettings,
  getResizeSettings,
  isEnabled,
} from "./settings";
import { adjustSelectedWordObject, isImageShape } from "./resize/word";
import { adjustSelectedExcelShape } from "./resize/excel";
import { adjustSelectedPptShape } from "./resize/powerpoint";
import {
  initPasteBaseline,
  checkCountChange,
  handleCountIncrease,
  cancelPendingDetection,
  updateBaseline,
} from "./resize/word-paste-detector";

import "./logger";

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[pasteMode]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[pasteMode]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[pasteMode]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[pasteMode]", msg, data ?? ""),
  };
}

const pasteLog = {
  debug: (msg: string, data?: any) => getLog().debug(msg, data),
  info: (msg: string, data?: any) => getLog().info(msg, data),
  warn: (msg: string, data?: any) => getLog().warn(msg, data),
  error: (msg: string, data?: any) => getLog().error(msg, data),
};

pasteLog.info("taskpane-new module loaded v7");
console.log("=== taskpane-new v7 loaded ===");

const wordState: WordState = createWordState();
let pasteDebounceTimer: number | null = null;

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

async function checkAndHandlePaste(): Promise<void> {
  try {
    const { inlineIncreased, shapeIncreased, inlineCount, shapeCount } = await checkCountChange();
    pasteLog.info("check result", { inlineIncreased, shapeIncreased, inlineCount, shapeCount });
    
    if (inlineIncreased || shapeIncreased) {
      pasteLog.info("new image detected");
      const settings = getResizeSettings();
      handleCountIncrease(settings, inlineIncreased, shapeIncreased, inlineCount, shapeCount, (result, method) => {
        if (result !== "none") {
          wordState.lastAdjustAt = Date.now();
          pasteLog.info("resize success: " + result + " (" + method + ")");
          updateBaseline(inlineCount, shapeCount);
          setStatus("Target: " + settings.targetWidthCm + "cm, Adjusted: " + result + " (" + method + ")");
        }
      });
    } else {
      pasteLog.debug("no new image");
      updateBaseline(inlineCount, shapeCount);
    }
  } catch (e) {
    pasteLog.error("check error", e);
  }
}

async function onSelectionChanged(): Promise<void> {
  if (!isEnabled()) return;
  const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();

  if (Office.context.host === Office.HostType.Word) {
    if (!resizeOnSelection && resizeOnPaste) {
      if (pasteDebounceTimer !== null) {
        window.clearTimeout(pasteDebounceTimer);
      }
      pasteDebounceTimer = window.setTimeout(() => {
        pasteDebounceTimer = null;
        void checkAndHandlePaste();
      }, 0); // 无防抖，直接响应
      return;
    }

    if (resizeOnSelection) {
      if (wordState.selectionDebounceTimer !== null) {
        window.clearTimeout(wordState.selectionDebounceTimer);
      }
      wordState.selectionDebounceTimer = window.setTimeout(() => {
        wordState.selectionDebounceTimer = null;
        void processSelectionChanged();
      }, CONFIG.word.selection.debounceMs);
    }
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
    setStatus("Target: " + settings.targetWidthCm + "cm, Adjusted: " + result);
  } catch (e) {
    setStatus("Error: " + String(e));
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
    setStatus("Saved. Enabled: " + enabled.checked);
    if (Office.context.host === Office.HostType.Word) {
      if (enabled.checked) {
        await initPasteBaseline();
      } else {
        cancelPendingDetection();
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
    setStatus("Saved. Paste: " + nextPaste + ", Selection: " + nextSelection);
    if (Office.context.host === Office.HostType.Word) {
      if (nextPaste && isEnabled()) {
        await initPasteBaseline();
      } else {
        cancelPendingDetection();
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
    setStatus("Saved. Width: " + next + " cm");
    updateLockAspectRatioUi();
  };

  targetWidthCm.addEventListener("input", () => void onWidthChanged());
  targetWidthCm.addEventListener("change", () => void onWidthChanged());

  setHeightEnabled.addEventListener("change", async () => {
    await saveSetting(CONFIG.keys.setHeightEnabled, setHeightEnabled.checked);
    setStatus("Saved. Set height: " + setHeightEnabled.checked);
    updateLockAspectRatioUi();
  });

  const onHeightChanged = async () => {
    const next = Number(targetHeightCm.value);
    if (!Number.isFinite(next) || next <= 0) return;
    await saveSetting(CONFIG.keys.targetHeightCm, next);
    setStatus("Saved. Height: " + next + " cm");
    updateLockAspectRatioUi();
  };

  targetHeightCm.addEventListener("input", () => void onHeightChanged());
  targetHeightCm.addEventListener("change", () => void onHeightChanged());

  lockAspectRatio.disabled = true;
}

Office.onReady(async () => {
  pasteLog.info("=== Office.onReady v7 ===");
  console.log("=== [taskpane-new] Office.onReady v7 ===");
  
  if (!hasOfficeContext()) {
    setStatus("Must open in Office task pane.");
    return;
  }

  bindUi();
  pasteLog.info("UI bindUi done");

  if (Office.context.host === Office.HostType.Word) {
    const { resizeOnPaste } = getFeatureSettings();
    pasteLog.info("Word init", { resizeOnPaste, enabled: isEnabled() });
    if (isEnabled() && resizeOnPaste) {
      await initPasteBaseline();
      pasteLog.info("baseline init done");
    }
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => void onSelectionChanged(),
    (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        pasteLog.info("SelectionChanged registered");
        setStatus("Ready. Auto-resize on paste.");
      } else {
        pasteLog.error("SelectionChanged failed", result.error);
        setStatus("Registration failed: " + (result.error?.message ?? String(result.error)));
      }
    }
  );
});
