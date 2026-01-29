/**
 * Office Paste Width - Taskpane UI v26
 * 修复: 参考框删除使用前缀匹配, 保存后更新 UI 显示
 */

import { setStatus, hasOfficeContext, cmToPoints, debounce } from "./utils";
import { getBoolSetting, getNumberSetting, saveSetting } from "./settings";
import { adjustSelectedWordObject } from "./resize/word";
import { adjustSelectedExcelShape } from "./resize/excel";
import { adjustSelectedPptShape } from "./resize/powerpoint";
import { initPasteBaseline, checkCountChange, handleCountIncrease, resizeByGlobalScan } from "./resize/word-paste-detector";
import {
  initExcelPasteBaseline,
  checkExcelCountChange,
  handleExcelCountIncrease,
  cancelExcelPendingDetection,
  updateExcelSizeSnapshot,
} from "./resize/excel-paste-detector";
import {
  initPptPasteBaseline,
  checkPptCountChange,
  handlePptCountIncrease,
  cancelPptPendingDetection,
  updatePptSizeSnapshot,
} from "./resize/ppt-paste-detector";
import { insertReferenceBox, removeReferenceBox, getReferenceBoxState } from "./reference-box";
import { captureSizeSnapshot, createRegistry } from "./resize-scope";
import "./logger";

const BUILD_TAG = "v39";

interface PendingSettings {
  targetWidthCm: number;
  targetHeightCm: number;
  lockHeight: boolean;
  lockAspectRatio: boolean;
}

interface UIState {
  pendingSettings: PendingSettings | null;
  refBoxActive: boolean;
  refBoxShapeName: string | null;
  refBoxPollingInterval: number | null;
  refBoxUnsupported: boolean;
  webUnsupported: boolean;
}

const uiState: UIState = {
  pendingSettings: null,
  refBoxActive: false,
  refBoxShapeName: null,
  refBoxPollingInterval: null,
  refBoxUnsupported: false,
  webUnsupported: false,
};

let selectionAdjustInFlight = false;

let excelActiveShapePollTimer: number | null = null;
let excelLastActiveShapeId: string | null = null;

let excelPasteCountPollTimer: number | null = null;
let excelPastePendingShapeCount: number | null = null;

let wordPasteCountPollTimer: number | null = null;
let wordPastePendingKey: string | null = null;

let excelConfirmSeq = 0;
let excelConfirmTimers: number[] = [];

let pptConfirmSeq = 0;
let pptConfirmTimers: number[] = [];

function cancelExcelConfirmRetries(): void {
  excelConfirmSeq += 1;
  for (const t of excelConfirmTimers) window.clearTimeout(t);
  excelConfirmTimers = [];
}

function scheduleExcelConfirmRetries(): void {
  cancelExcelConfirmRetries();
  const seq = excelConfirmSeq;
  const delays = [100, 250, 600];
  for (const delay of delays) {
    const timer = window.setTimeout(() => {
      void (async () => {
        if (seq !== excelConfirmSeq) return;
        await handleSelectionChange("retry");
      })();
    }, delay);
    excelConfirmTimers.push(timer);
  }
}

function cancelPptConfirmRetries(): void {
  pptConfirmSeq += 1;
  for (const t of pptConfirmTimers) window.clearTimeout(t);
  pptConfirmTimers = [];
}

function schedulePptConfirmRetries(): void {
  cancelPptConfirmRetries();
  const seq = pptConfirmSeq;
  const delays = [100, 250, 600];
  for (const delay of delays) {
    const timer = window.setTimeout(() => {
      void (async () => {
        if (seq !== pptConfirmSeq) return;
        await handleSelectionChange("retry");
      })();
    }, delay);
    pptConfirmTimers.push(timer);
  }
}

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[taskpane]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[taskpane]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[taskpane]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[taskpane]", msg, data ?? ""),
  };
}

function isOfficeOnlinePlatform(platform: any): boolean {
  try {
    const officeAny = Office as any;
    if (officeAny?.PlatformType?.OfficeOnline && platform === officeAny.PlatformType.OfficeOnline) return true;
  } catch {
    // ignore
  }
  return String(platform ?? "").toLowerCase() === "officeonline";
}

function isMacPlatform(): boolean {
  try {
    const p = (Office as any)?.context?.diagnostics?.platform ?? (Office as any)?.context?.platform;
    return String(p ?? "").toLowerCase() === "mac";
  } catch {
    return false;
  }
}

function toLoggableError(e: unknown): any {
  if (e && typeof e === "object") {
    const anyErr = e as any;
    return {
      name: anyErr?.name,
      message: anyErr?.message ?? String(e),
      code: anyErr?.code,
      debugInfo: anyErr?.debugInfo,
      httpStatusCode: anyErr?.httpStatusCode,
    };
  }
  return { message: String(e) };
}

function createCardSection(title?: string): HTMLDivElement {
  const section = document.createElement("div");
  section.className = "card-section";
  if (title) {
    const titleEl = document.createElement("div");
    titleEl.className = "card-section-title";
    titleEl.textContent = title;
    section.appendChild(titleEl);
  }
  return section;
}

function createToggleSwitch(
  id: string,
  label: string,
  checked: boolean,
  onChange: (checked: boolean) => void,
  disabled: boolean = false
): HTMLDivElement {
  const container = document.createElement("div");
  container.className = "toggle-container";
  const labelEl = document.createElement("span");
  labelEl.className = "toggle-label";
  labelEl.textContent = label;
  const toggle = document.createElement("div");
  toggle.className = `toggle-switch${checked ? " active" : ""}`;
  toggle.id = id;
  toggle.setAttribute("role", "switch");
  toggle.setAttribute("aria-checked", String(checked));
  toggle.tabIndex = 0;
  if (disabled) {
    toggle.setAttribute("aria-disabled", "true");
    toggle.style.pointerEvents = "none";
    toggle.style.opacity = "0.6";
  }
  const knob = document.createElement("div");
  knob.className = "toggle-switch-knob";
  toggle.appendChild(knob);
  toggle.addEventListener("click", () => {
    if (disabled) return;
    const newState = !toggle.classList.contains("active");
    toggle.classList.toggle("active", newState);
    toggle.setAttribute("aria-checked", String(newState));
    onChange(newState);
  });
  container.appendChild(labelEl);
  container.appendChild(toggle);
  return container;
}

function createRadioButton(
  name: string,
  value: string,
  label: string,
  checked: boolean,
  onChange: () => void,
  disabled: boolean = false
): HTMLDivElement {
  const container = document.createElement("div");
  const input = document.createElement("input");
  input.type = "radio";
  input.name = name;
  input.value = value;
  input.id = `${name}-${value}`;
  input.className = "modern-radio";
  input.checked = checked;
  input.disabled = disabled;
  const labelEl = document.createElement("label");
  labelEl.className = "modern-radio-label";
  labelEl.htmlFor = input.id;
  const circle = document.createElement("span");
  circle.className = "modern-radio-circle";
  const dot = document.createElement("span");
  dot.className = "modern-radio-dot";
  circle.appendChild(dot);
  const text = document.createElement("span");
  text.textContent = label;
  labelEl.appendChild(circle);
  labelEl.appendChild(text);
  input.addEventListener("change", () => {
    if (disabled) return;
    if (input.checked) onChange();
  });
  container.appendChild(input);
  container.appendChild(labelEl);
  return container;
}

function createInput(
  id: string,
  label: string,
  value: number,
  onChange: (value: number) => void,
  unit: string = "cm",
  disabled: boolean = false
): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "input-row";
  const labelEl = document.createElement("span");
  labelEl.className = "input-label";
  labelEl.textContent = label;
  const input = document.createElement("input");
  input.type = "number";
  input.id = id;
  input.className = "modern-input";
  input.value = String(value);
  input.min = "0";
  input.max = "100";
  input.step = "0.5";
  input.disabled = disabled;
  const unitEl = document.createElement("span");
  unitEl.className = "input-unit";
  unitEl.textContent = unit;
  const handleChange = () => {
    if (disabled) return;
    const val = parseFloat(input.value);
    if (!isNaN(val) && val >= 0) onChange(val);
  };
  input.addEventListener("change", handleChange);
  input.addEventListener("input", handleChange);
  row.appendChild(labelEl);
  row.appendChild(input);
  row.appendChild(unitEl);
  return row;
}

function createButton(
  text: string,
  onClick: () => void,
  variant: "primary" | "secondary" = "primary",
  fullWidth: boolean = false,
  disabled: boolean = false
): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `modern-button modern-button-${variant}${fullWidth ? " modern-button-full" : ""}`;
  button.textContent = text;
  button.disabled = disabled;
  button.addEventListener("click", () => {
    if (disabled) return;
    onClick();
  });
  return button;
}

function getCurrentSettings() {
  return {
    enabled: getBoolSetting("enabled", true),
    resizeOnPaste: getBoolSetting("resizeOnPaste", true),
    resizeOnSelection: getBoolSetting("resizeOnSelection", false),
    targetWidthCm: getNumberSetting("targetWidthCm", 15),
    targetHeightCm: getNumberSetting("targetHeightCm", 10),
    lockHeight: getBoolSetting("setHeightEnabled", false),
    lockAspectRatio: getBoolSetting("lockAspectRatio", true),
  };
}

function getResizeSettings() {
  const settings = getCurrentSettings();

  const applyWidth = Number.isFinite(settings.targetWidthCm) && settings.targetWidthCm > 0;
  const applyHeight =
    settings.lockHeight && Number.isFinite(settings.targetHeightCm) && settings.targetHeightCm > 0;

  // 如果宽高都没设置，回退到默认宽度（避免什么都不做）
  const effectiveTargetWidthCm = !applyWidth && !applyHeight ? 15 : settings.targetWidthCm;
  const effectiveApplyWidth = !applyWidth && !applyHeight ? true : applyWidth;

  // 等比：只设置一边时才等比
  const lockAspectRatio =
    (effectiveApplyWidth && !applyHeight) || (!effectiveApplyWidth && applyHeight);

  return {
    targetWidthCm: effectiveTargetWidthCm,
    targetHeightCm: settings.targetHeightCm,
    setHeightEnabled: settings.lockHeight,
    applyWidth: effectiveApplyWidth,
    applyHeight,
    lockAspectRatio,
  };
}

function startRefBoxPolling(): void {
  const log = getLog();
  if (uiState.refBoxPollingInterval !== null) return;
  log.info("startRefBoxPolling: start");
  let consecutiveFailures = 0;
  uiState.refBoxPollingInterval = window.setInterval(async () => {
    try {
      const state = await getReferenceBoxState();
      if (state && state.widthCm && state.heightCm) {
        const widthInput = document.getElementById("widthInput") as HTMLInputElement;
        const heightInput = document.getElementById("heightInput") as HTMLInputElement;
        if (widthInput) widthInput.value = state.widthCm.toFixed(1);
        if (heightInput) heightInput.value = state.heightCm.toFixed(1);
        uiState.pendingSettings = {
          targetWidthCm: state.widthCm,
          targetHeightCm: state.heightCm,
          lockHeight: getBoolSetting("setHeightEnabled", false),
          lockAspectRatio: !getBoolSetting("setHeightEnabled", false),
        };
      }
      consecutiveFailures = 0;
    } catch (e) {
      consecutiveFailures += 1;
      log.warn("refBoxPolling error", toLoggableError(e));
      if (consecutiveFailures >= 3) {
        log.warn("refBoxPolling: too many failures, stop polling");
        stopRefBoxPolling();
      }
    }
  }, 300);
}

function stopRefBoxPolling(): void {
  const log = getLog();
  if (uiState.refBoxPollingInterval !== null) {
    log.info("stopRefBoxPolling: stop");
    window.clearInterval(uiState.refBoxPollingInterval);
    uiState.refBoxPollingInterval = null;
  }
}

function startExcelActiveShapePolling(): void {
  const log = getLog();
  if (excelActiveShapePollTimer !== null) return;
  excelLastActiveShapeId = null;
  log.info("startExcelActiveShapePolling: start");
  excelActiveShapePollTimer = window.setInterval(async () => {
    if (!hasOfficeContext()) return;
    if (Office.context.host !== Office.HostType.Excel) return;

    const settings = getCurrentSettings();
    if (!settings.enabled) return;
    if (!settings.resizeOnSelection && !settings.resizeOnPaste) return;
    if (selectionAdjustInFlight) return;

    try {
      const info = await Excel.run(async (ctx) => {
        const shape = ctx.workbook.getActiveShapeOrNullObject();
        shape.load(["isNullObject", "id", "type"]);
        await ctx.sync();
        if (shape.isNullObject) return { id: null as string | null, isImage: false };
        const typeString = String((shape as any).type ?? "").toLowerCase();
        const isImage = typeString.includes("image") || typeString.includes("picture");
        return { id: shape.id as string, isImage };
      });

      if (!info?.id || !info.isImage) {
        excelLastActiveShapeId = null;
        return;
      }

      if (info.id === excelLastActiveShapeId) return;
      excelLastActiveShapeId = info.id;

      selectionAdjustInFlight = true;
      log.info("excelActiveShapePoll: new active image", { id: info.id });
      const result = await adjustSelectedExcelShape(getResizeSettings());
      log.info("excelActiveShapePoll: resize attempted", { result });
      if (result !== "none") {
        setStatus(`Resized ${result} image (excel active shape)`);
        syncUIWithSavedSettings();
      }
    } catch (e) {
      log.warn("excelActiveShapePoll error", toLoggableError(e));
    } finally {
      selectionAdjustInFlight = false;
    }
  }, 300);
}

function stopExcelActiveShapePolling(): void {
  const log = getLog();
  if (excelActiveShapePollTimer !== null) {
    log.info("stopExcelActiveShapePolling: stop");
    window.clearInterval(excelActiveShapePollTimer);
    excelActiveShapePollTimer = null;
  }
  excelLastActiveShapeId = null;
}

function startExcelPasteCountPolling(): void {
  const log = getLog();
  if (excelPasteCountPollTimer !== null) return;
  log.info("startExcelPasteCountPolling: start");
  excelPasteCountPollTimer = window.setInterval(async () => {
    if (!hasOfficeContext()) return;
    if (Office.context.host !== Office.HostType.Excel) return;

    const settings = getCurrentSettings();
    if (!settings.enabled) return;
    if (!settings.resizeOnPaste) return;
    if (selectionAdjustInFlight) return;

    try {
      const change = await checkExcelCountChange();
      if (!change.shapeIncreased) return;

      if (excelPastePendingShapeCount === change.shapeCount) return;
      excelPastePendingShapeCount = change.shapeCount;

      log.info("excelPasteCountPoll: paste detected", {
        shapeCount: change.shapeCount,
      });

      handleExcelCountIncrease(getResizeSettings(), change.shapeCount, (resized, method) => {
        log.info("excelPasteCountPoll: paste resize done", { resized, method });
        excelPastePendingShapeCount = null;
        if (resized) {
          setStatus(`Resized excel-shape image (paste ${method})`);
          syncUIWithSavedSettings();
        }
      });
    } catch (e) {
      log.warn("excelPasteCountPoll error", toLoggableError(e));
      excelPastePendingShapeCount = null;
    }
  }, 300);
}

function stopExcelPasteCountPolling(): void {
  const log = getLog();
  if (excelPasteCountPollTimer !== null) {
    log.info("stopExcelPasteCountPolling: stop");
    window.clearInterval(excelPasteCountPollTimer);
    excelPasteCountPollTimer = null;
  }
  excelPastePendingShapeCount = null;
}

function startWordPasteCountPolling(): void {
  const log = getLog();
  if (wordPasteCountPollTimer !== null) return;
  wordPastePendingKey = null;
  let lastGlobalScanAt = 0;
  let lastSelectionAttemptAt = 0;
  let lastHeartbeatAt = 0;
  log.info("startWordPasteCountPolling: start");
  wordPasteCountPollTimer = window.setInterval(async () => {
    if (!hasOfficeContext()) return;
    if (Office.context.host !== Office.HostType.Word) return;

    const settings = getCurrentSettings();
    if (!settings.enabled) return;
    if (!settings.resizeOnPaste) return;
    if (selectionAdjustInFlight) return;

    try {
      const change = await checkCountChange();
      const now = Date.now();
      if (now - lastHeartbeatAt > 6000) {
        lastHeartbeatAt = now;
        log.info("wordPasteCountPoll: heartbeat", {
          inlineCount: change.inlineCount,
          shapeCount: change.shapeCount,
        });
      }

      if (change.inlineIncreased || change.shapeIncreased) {
        const key = `${change.inlineCount}/${change.shapeCount}`;
        if (wordPastePendingKey === key) return;
        wordPastePendingKey = key;

        log.info("wordPasteCountPoll: paste detected", {
          inlineIncreased: change.inlineIncreased,
          shapeIncreased: change.shapeIncreased,
          inlineCount: change.inlineCount,
          shapeCount: change.shapeCount,
        });

        handleCountIncrease(
          getResizeSettings(),
          change.inlineIncreased,
          change.shapeIncreased,
          change.inlineCount,
          change.shapeCount,
          (resizeResult, method) => {
            log.info("wordPasteCountPoll: paste resize done", { result: resizeResult, method });
            wordPastePendingKey = null;
            if (resizeResult !== "none") {
              setStatus(`Resized ${resizeResult} image (${method})`);
              syncUIWithSavedSettings();
            }
          }
        );
        return;
      }

      if (now - lastSelectionAttemptAt > 800) {
        lastSelectionAttemptAt = now;
        selectionAdjustInFlight = true;
        try {
          await Word.run(async (ctx) => {
            const res = await adjustSelectedWordObject(ctx, getResizeSettings());
            if (res !== "none") {
              log.info("wordPasteCountPoll: selection resized", { result: res });
              setStatus(`Resized ${res} image (selection)`);
              syncUIWithSavedSettings();
            }
          });
        } catch (e) {
          log.warn("wordPasteCountPoll: selection error", toLoggableError(e));
        } finally {
          selectionAdjustInFlight = false;
        }
      }

      if (isMacPlatform() && now - lastGlobalScanAt > 1200) {
        lastGlobalScanAt = now;
        selectionAdjustInFlight = true;
        try {
          const result = await resizeByGlobalScan(getResizeSettings());
          if (result !== "none") {
            log.info("wordPasteCountPoll: global-scan resized", { result });
            setStatus(`Resized ${result} image (global-scan)`);
            syncUIWithSavedSettings();
          }
        } catch (e) {
          log.warn("wordPasteCountPoll: global-scan error", toLoggableError(e));
        } finally {
          selectionAdjustInFlight = false;
        }
      }
    } catch (e) {
      log.warn("wordPasteCountPoll error", toLoggableError(e));
      wordPastePendingKey = null;
    }
  }, 300);
}

function stopWordPasteCountPolling(): void {
  const log = getLog();
  if (wordPasteCountPollTimer !== null) {
    log.info("stopWordPasteCountPolling: stop");
    window.clearInterval(wordPasteCountPollTimer);
    wordPasteCountPollTimer = null;
  }
  wordPastePendingKey = null;
}

const LOCK_ICON_LOCKED = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path></svg>`;
const LOCK_ICON_UNLOCKED = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 9.9-1"></path></svg>`;
const MOUSE_CURSOR_ICON = `<svg class="mouse-cursor-svg" width="24" height="24" viewBox="0 0 16 16"><rect x="1" y="1" width="10" height="8" fill="none" stroke="#E67E22" stroke-width="1.2" stroke-dasharray="2,1" rx="0.5"/><polygon points="9,7 9,14 11,12 13,15 14,14 12,11 14,10" fill="#4A7DC4"/></svg>`;

function syncUIWithSavedSettings(): void {
  const settings = getCurrentSettings();
  const widthInput = document.getElementById("widthInput") as HTMLInputElement | null;
  const heightInput = document.getElementById("heightInput") as HTMLInputElement | null;
  const lockBtn = document.getElementById("lockHeightBtn") as HTMLButtonElement | null;

  if (widthInput) widthInput.value = settings.targetWidthCm.toFixed(1);
  if (heightInput) heightInput.value = settings.targetHeightCm.toFixed(1);

  if (lockBtn) {
    lockBtn.classList.toggle("active", settings.lockHeight);
    lockBtn.innerHTML = settings.lockHeight ? LOCK_ICON_LOCKED : LOCK_ICON_UNLOCKED;
    lockBtn.title = settings.lockHeight ? "Height locked" : "Height unlocked";
  }
}

async function handleToggleRefBox(): Promise<void> {
  const log = getLog();
  const btn = document.getElementById("refBoxBtn") as HTMLButtonElement;
  if (uiState.refBoxActive) {
    log.info("handleToggleRefBox: close");
    stopRefBoxPolling();
    try {
      if (uiState.refBoxShapeName) await removeReferenceBox(uiState.refBoxShapeName);
    } catch (e) { log.warn("removeReferenceBox error", e); }
    uiState.refBoxActive = false;
    uiState.refBoxShapeName = null;
    if (btn) {
      btn.classList.remove("active");
      btn.innerHTML = `<span class="mouse-icon">${MOUSE_CURSOR_ICON}</span> Draggable Size Setting`;
    }
  } else {
    log.info("handleToggleRefBox: open");
    try {
      const shapeName = await insertReferenceBox();
      if (shapeName) {
        uiState.refBoxUnsupported = false;
        uiState.refBoxActive = true;
        uiState.refBoxShapeName = shapeName;
        startRefBoxPolling();
        if (btn) {
          btn.classList.add("active");
          btn.textContent = "Click to End Dragging";
        }
        setStatus("Reference box inserted (tip: click a cell near the screen center, then insert)");
      } else {
        log.warn("insertReferenceBox returned null");
        uiState.refBoxUnsupported = true;
        if (btn) {
          btn.disabled = true;
          (btn as any).style.opacity = "0.6";
          btn.title = "Reference box is not supported on this platform";
        }
        buildUI();
        setStatus("Failed to insert reference box");
      }
    } catch (e) {
      log.error("insertReferenceBox error", e);
      setStatus("Failed to insert reference box");
    }
  }
}

async function applySettings(): Promise<void> {
  const log = getLog();
  log.info("applySettings: save");
  stopRefBoxPolling();
  if (uiState.refBoxActive) {
    try {
      if (uiState.refBoxShapeName) await removeReferenceBox(uiState.refBoxShapeName);
    } catch (e) { log.warn("removeReferenceBox error", e); }
    uiState.refBoxActive = false;
    uiState.refBoxShapeName = null;
    const btn = document.getElementById("refBoxBtn") as HTMLButtonElement;
    if (btn) {
      btn.classList.remove("active");
      btn.innerHTML = `<span class="mouse-icon">${MOUSE_CURSOR_ICON}</span> Draggable Size Setting`;
    }
  }
  const widthInput = document.getElementById("widthInput") as HTMLInputElement;
  const heightInput = document.getElementById("heightInput") as HTMLInputElement;
  let savedWidth = 0, savedHeight = 0;
  if (widthInput) {
    const width = parseFloat(widthInput.value);
    if (!isNaN(width) && width >= 0) {
      await saveSetting("targetWidthCm", width);
      savedWidth = width;
    }
  }
  if (heightInput) {
    const height = parseFloat(heightInput.value);
    if (!isNaN(height) && height >= 0) {
      await saveSetting("targetHeightCm", height);
      savedHeight = height;
    }
  }
  log.info("applySettings: saved values", { width: savedWidth, height: savedHeight });
  uiState.pendingSettings = null;
  
  // 保存设置后重置内存表并更新尺寸快照
  if (hasOfficeContext() && Office.context.host === Office.HostType.Word) {
    try {
      // 重置内存表，清空已记录的 Shape ID
      createRegistry();
      log.info("applySettings: registry reset");
      
      const settings = getCurrentSettings();
      await captureSizeSnapshot({
        targetWidthPt: cmToPoints(settings.targetWidthCm),
        targetHeightPt: cmToPoints(settings.targetHeightCm),
      });
      log.info("applySettings: size snapshot updated");
    } catch (e) {
      log.warn("applySettings: update size snapshot failed", toLoggableError(e));
    }
  }

  if (hasOfficeContext() && Office.context.host === Office.HostType.PowerPoint) {
    try {
      const after = getCurrentSettings();
      if (after.resizeOnPaste) {
        await updatePptSizeSnapshot(
          cmToPoints(after.targetWidthCm),
          cmToPoints(after.targetHeightCm)
        );
        cancelPptPendingDetection();
        await initPptPasteBaseline();
        log.info("applySettings: ppt paste baseline reset");
      }
    } catch (e) {
      log.warn("applySettings: ppt paste baseline reset failed", toLoggableError(e));
    }
  }

  if (hasOfficeContext() && Office.context.host === Office.HostType.Excel) {
    try {
      const after = getCurrentSettings();
      if (after.resizeOnPaste) {
        await updateExcelSizeSnapshot(
          cmToPoints(after.targetWidthCm),
          cmToPoints(after.targetHeightCm)
        );
        cancelExcelPendingDetection();
        await initExcelPasteBaseline();
        log.info("applySettings: excel paste baseline reset");
        startExcelPasteCountPolling();
      }
    } catch (e) {
      log.warn("applySettings: excel paste baseline reset failed", toLoggableError(e));
    }
  }
  
  // 刷新 UI 显示保存后的值
  if (widthInput) widthInput.value = savedWidth.toFixed(1);
  if (heightInput) heightInput.value = savedHeight.toFixed(1);

  try {
    const after = getCurrentSettings();
    log.info("applySettings: readback", {
      targetWidthCm: after.targetWidthCm,
      targetHeightCm: after.targetHeightCm,
      lockHeight: after.lockHeight,
      resizeOnPaste: after.resizeOnPaste,
      resizeOnSelection: after.resizeOnSelection,
    });
  } catch (e) {
    log.warn("applySettings: readback failed", toLoggableError(e));
  }

  setStatus("Settings saved");
  log.info("applySettings: done");

  if (hasOfficeContext() && Office.context.host === Office.HostType.Word) {
    const after = getCurrentSettings();
    if (after.enabled && after.resizeOnPaste) startWordPasteCountPolling();
    else stopWordPasteCountPolling();
  }
}

async function handleSelectionChange(source: "event" | "retry" = "event"): Promise<void> {
  const log = getLog();
  const settings = getCurrentSettings();
  if (!settings.enabled) return;
  const host = Office.context.host;

  // 粘贴模式：Word 使用粘贴探测；Excel/PPT 使用“选中即缩放”（粘贴后通常会自动选中新图片）
  if (settings.resizeOnPaste && host === Office.HostType.Word) {
    if (wordPasteCountPollTimer !== null) return;
    try {
      const result = await checkCountChange();
      if (result.inlineIncreased || result.shapeIncreased) {
        log.info("handleSelectionChange: image count increased", {
          inlineIncreased: result.inlineIncreased,
          shapeIncreased: result.shapeIncreased,
          inlineCount: result.inlineCount,
          shapeCount: result.shapeCount,
        });
        handleCountIncrease(
          getResizeSettings(),
          result.inlineIncreased,
          result.shapeIncreased,
          result.inlineCount,
          result.shapeCount,
          (resizeResult, method) => {
            log.info("handleSelectionChange: resize done", { result: resizeResult, method });
              if (resizeResult !== "none") { setStatus(`Resized ${resizeResult} image (${method})`); syncUIWithSavedSettings(); }
          }
        );
      }
    } catch (e) { log.error("handleSelectionChange paste mode error", toLoggableError(e)); }
  } else if (settings.resizeOnPaste && host === Office.HostType.Excel) {
    try {
      if (selectionAdjustInFlight) return;
      selectionAdjustInFlight = true;
      const change = await checkExcelCountChange();
      if (change.shapeIncreased) {
        log.info("handleSelectionChange: excel paste detected", {
          shapeCount: change.shapeCount,
          source,
        });
        handleExcelCountIncrease(getResizeSettings(), change.shapeCount, (resized, method) => {
          log.info("handleSelectionChange: excel paste resize done", { resized, method, source });
          if (resized) {
            setStatus(`Resized excel-shape image (paste ${method})`);
            syncUIWithSavedSettings();
          }
        });
      } else {
        log.info("handleSelectionChange: excel paste check no change", { source });
      }
    } catch (e) {
      log.error("handleSelectionChange excel paste error", toLoggableError(e));
    } finally {
      selectionAdjustInFlight = false;
    }
  } else if (settings.resizeOnPaste && host === Office.HostType.PowerPoint) {
    try {
      if (selectionAdjustInFlight) return;
      selectionAdjustInFlight = true;
      const change = await checkPptCountChange();
      if (change.shapeIncreased) {
        log.info("handleSelectionChange: ppt paste detected", {
          shapeCount: change.shapeCount,
          source,
        });
        handlePptCountIncrease(getResizeSettings(), change.shapeCount, (resized, method) => {
          log.info("handleSelectionChange: ppt paste resize done", { resized, method, source });
          if (resized) {
            setStatus(`Resized ppt-shape image (paste ${method})`);
            syncUIWithSavedSettings();
          }
        });
      } else {
        log.info("handleSelectionChange: ppt paste check no change", { source });
      }
    } catch (e) {
      log.error("handleSelectionChange ppt paste error", toLoggableError(e));
    } finally {
      selectionAdjustInFlight = false;
    }
  }

  if (settings.resizeOnSelection) {
    if (selectionAdjustInFlight) return;
    try {
      selectionAdjustInFlight = true;
      log.info("handleSelectionChange: selection mode triggered", { host, source });
      if (host === Office.HostType.Word) {
        await Word.run(async (ctx) => { await adjustSelectedWordObject(ctx, getResizeSettings()); });
      } else if (host === Office.HostType.Excel) {
        if (source === "event") {
          const rs = getResizeSettings();
          log.info("handleSelectionChange: Excel resize settings", {
            targetWidthCm: rs.targetWidthCm,
            targetHeightCm: rs.targetHeightCm,
            applyWidth: rs.applyWidth,
            applyHeight: rs.applyHeight,
            setHeightEnabled: rs.setHeightEnabled,
            lockAspectRatio: rs.lockAspectRatio,
          });
        }
        const result = await adjustSelectedExcelShape(getResizeSettings());
        log.info("handleSelectionChange: Excel selection resize attempted", { result, source });
        if (result === "none") {
          if (source === "event") scheduleExcelConfirmRetries();
        } else {
          setStatus(`Resized ${result} image (selection)`);
          syncUIWithSavedSettings();
          cancelExcelConfirmRetries();
        }
      } else if (host === Office.HostType.PowerPoint) {
        if (source === "event") {
          const rs = getResizeSettings();
          log.info("handleSelectionChange: PowerPoint resize settings", {
            targetWidthCm: rs.targetWidthCm,
            targetHeightCm: rs.targetHeightCm,
            applyWidth: rs.applyWidth,
            applyHeight: rs.applyHeight,
            setHeightEnabled: rs.setHeightEnabled,
            lockAspectRatio: rs.lockAspectRatio,
          });
        }
        const result = await adjustSelectedPptShape(getResizeSettings());
        log.info("handleSelectionChange: PowerPoint selection resize attempted", { result, source });
        if (result === "none") {
          if (source === "event") schedulePptConfirmRetries();
        } else {
          setStatus(`Resized ${result} image (selection)`);
          syncUIWithSavedSettings();
          cancelPptConfirmRetries();
        }
      }
    } catch (e) { log.error("handleSelectionChange selection mode error", toLoggableError(e)); }
    finally { selectionAdjustInFlight = false; }
  }
}

const debouncedSelectionHandler = debounce(() => {
  void handleSelectionChange("event");
}, 80);

async function registerSelectionChangedEvent(): Promise<void> {
  const log = getLog();
  if (!hasOfficeContext()) {
    log.warn("registerSelectionChangedEvent: No Office context");
    return;
  }
  const host = Office.context.host;
  
  if (host === Office.HostType.Word) {
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => { debouncedSelectionHandler(); },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            log.info("registerSelectionChangedEvent: Word SelectionChanged registered");
          } else {
            log.error("registerSelectionChangedEvent: Word registration failed", result.error);
          }
        }
      );
    } catch (e) { 
      log.error("registerSelectionChangedEvent Word error", e); 
    }
  } else if (host === Office.HostType.Excel) {
    try {
      await Excel.run(async (ctx) => {
        ctx.workbook.onSelectionChanged.add(() => { debouncedSelectionHandler(); });
        await ctx.sync();
        log.info("registerSelectionChangedEvent: Excel SelectionChanged registered");
      });
    } catch (e) { log.error("registerSelectionChangedEvent Excel error", e); }

    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => { debouncedSelectionHandler(); },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            log.info("registerSelectionChangedEvent: Excel DocumentSelectionChanged registered");
          } else {
            log.error("registerSelectionChangedEvent: Excel DocumentSelectionChanged failed", result.error);
          }
        }
      );
    } catch (e) {
      log.error("registerSelectionChangedEvent Excel DocumentSelectionChanged error", e);
    }
  } else if (host === Office.HostType.PowerPoint) {
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        () => { debouncedSelectionHandler(); },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            log.info("registerSelectionChangedEvent: PowerPoint SelectionChanged registered");
          } else {
            log.error("registerSelectionChangedEvent: PowerPoint registration failed", result.error);
          }
        }
      );
    } catch (e) {
      log.error("registerSelectionChangedEvent PowerPoint error", e);
    }
  }
}

function buildUI(): void {
  const log = getLog();
  const container = document.getElementById("ui-container");
  if (!container) { log.error("buildUI: ui-container not found"); return; }
  container.innerHTML = "";
  const settings = getCurrentSettings();

  const disabled = uiState.webUnsupported;
  const refBoxDisabled =
    disabled ||
    (uiState.refBoxUnsupported && hasOfficeContext() && Office.context.host === Office.HostType.Word);
  if (uiState.webUnsupported) {
    const warnSection = createCardSection("Web version limitation");
    const warn = document.createElement("div");
    warn.style.color = "#b45309";
    warn.style.fontSize = "12px";
    warn.style.lineHeight = "1.4";
    const hostName = hasOfficeContext() ? String((Office as any)?.context?.host ?? "Office") : "Office";
    warn.textContent = `${hostName} on the web is not supported yet. Please use the desktop app (Windows/Mac).`;
    warnSection.appendChild(warn);
    container.appendChild(warnSection);
  }

  const enableSection = createCardSection();
  enableSection.appendChild(createToggleSwitch("enableToggle", "Enable Auto Resize", settings.enabled, (checked) => {
    saveSetting("enabled", checked);
    setStatus(checked ? "Enabled" : "Disabled");
  }, disabled));
  container.appendChild(enableSection);

  const modeSection = createCardSection("Mode Selection");
  const radioContainer = document.createElement("div");
  radioContainer.className = "radio-container";
  radioContainer.appendChild(createRadioButton("mode", "paste", "Resize on Paste", settings.resizeOnPaste, () => {
    saveSetting("resizeOnPaste", true);
    saveSetting("resizeOnSelection", false);
    if (hasOfficeContext() && Office.context.host === Office.HostType.PowerPoint) {
      void (async () => {
        try {
          cancelPptPendingDetection();
          await initPptPasteBaseline();
        } catch {
          // ignore
        }
      })();
    }
    if (hasOfficeContext() && Office.context.host === Office.HostType.Excel) {
      void (async () => {
        try {
          cancelExcelPendingDetection();
          await initExcelPasteBaseline();
        } catch {
          // ignore
        }
      })();
      stopExcelActiveShapePolling();
      startExcelPasteCountPolling();
    }
    setStatus("Mode: Paste");
  }, disabled));
  radioContainer.appendChild(createRadioButton("mode", "selection", "Resize on Selection", settings.resizeOnSelection && !settings.resizeOnPaste, () => {
    saveSetting("resizeOnPaste", false);
    saveSetting("resizeOnSelection", true);
    if (hasOfficeContext() && Office.context.host === Office.HostType.PowerPoint) {
      cancelPptPendingDetection();
    }
    if (hasOfficeContext() && Office.context.host === Office.HostType.Excel) {
      cancelExcelPendingDetection();
      stopExcelPasteCountPolling();
      startExcelActiveShapePolling();
    }
    setStatus("Mode: Selection");
  }, disabled));
  modeSection.appendChild(radioContainer);
  container.appendChild(modeSection);

  const sizeSection = createCardSection("Size Settings");

  const refBoxBtn = document.createElement("button");
  refBoxBtn.type = "button";
  refBoxBtn.id = "refBoxBtn";
  refBoxBtn.className = "draggable-size-btn";
  refBoxBtn.innerHTML = `<span class="mouse-icon">${MOUSE_CURSOR_ICON}</span> Draggable Size Setting`;
  if (refBoxDisabled) {
    refBoxBtn.disabled = true;
    (refBoxBtn as any).style.opacity = "0.6";
    if (uiState.refBoxUnsupported) refBoxBtn.title = "Reference box is not supported on this platform";
  } else {
    refBoxBtn.addEventListener("click", () => void handleToggleRefBox());
  }
  sizeSection.appendChild(refBoxBtn);

  sizeSection.appendChild(createInput("widthInput", "Width", settings.targetWidthCm, (value) => {
    if (!uiState.pendingSettings) {
      uiState.pendingSettings = { targetWidthCm: value, targetHeightCm: settings.targetHeightCm, lockHeight: settings.lockHeight, lockAspectRatio: !settings.lockHeight };
    } else { uiState.pendingSettings.targetWidthCm = value; }
  }, "cm", disabled));

  const heightRow = createInput("heightInput", "Height", settings.targetHeightCm, (value) => {
    if (!uiState.pendingSettings) {
      uiState.pendingSettings = { targetWidthCm: settings.targetWidthCm, targetHeightCm: value, lockHeight: settings.lockHeight, lockAspectRatio: !settings.lockHeight };
    } else { uiState.pendingSettings.targetHeightCm = value; }
  }, "cm", disabled);

  const lockBtn = document.createElement("button");
  lockBtn.type = "button";
  lockBtn.id = "lockHeightBtn";
  lockBtn.className = `lock-icon-btn${settings.lockHeight ? " active" : ""}`;
  lockBtn.innerHTML = settings.lockHeight ? LOCK_ICON_LOCKED : LOCK_ICON_UNLOCKED;
  lockBtn.title = settings.lockHeight ? "Height locked" : "Height unlocked";
  lockBtn.disabled = disabled;
  lockBtn.addEventListener("click", () => {
    if (disabled) return;
    const newLockState = !lockBtn.classList.contains("active");
    lockBtn.classList.toggle("active", newLockState);
    lockBtn.innerHTML = newLockState ? LOCK_ICON_LOCKED : LOCK_ICON_UNLOCKED;
    lockBtn.title = newLockState ? "Height locked" : "Height unlocked";
    saveSetting("setHeightEnabled", newLockState);
  });
  heightRow.appendChild(lockBtn);
  sizeSection.appendChild(heightRow);

  sizeSection.appendChild(createButton("Save", () => void applySettings(), "primary", true, disabled));
  container.appendChild(sizeSection);

  log.info("buildUI: UI built", { build: BUILD_TAG });
}

async function initialize(): Promise<void> {
  const log = getLog();
  log.info("initialize: start", { build: BUILD_TAG });
  buildUI();
  if (!hasOfficeContext()) {
    log.warn("initialize: No Office context, running in browser mode");
    setStatus("Browser mode (no Office)");
    return;
  }
  const host = Office.context.host;
  log.info("initialize: Office host", { host });
  if (uiState.webUnsupported) {
    setStatus(`Web not supported (${BUILD_TAG})`);
    log.warn("initialize: webUnsupported - skip Office handlers", { host });
    return;
  }
  if (host === Office.HostType.Word) {
    try { await initPasteBaseline(); log.info("initialize: paste baseline initialized"); } catch (e) { log.error("initialize: initPasteBaseline error", toLoggableError(e)); }
    try {
      // 捕获当前所有图片的尺寸快照（包含目标尺寸）
      const settings = getCurrentSettings();
      await captureSizeSnapshot({
        targetWidthPt: cmToPoints(settings.targetWidthCm),
        targetHeightPt: cmToPoints(settings.targetHeightCm),
      });
      log.info("initialize: size snapshot captured");
    } catch (e) { log.error("initialize: captureSizeSnapshot error", toLoggableError(e)); }
    // 创建内存表（粘贴模式只有新图片模式）
    createRegistry();
    log.info("initialize: registry created");
  }
  if (host === Office.HostType.PowerPoint) {
    try {
      const settings = getCurrentSettings();
      if (settings.resizeOnPaste) {
        await initPptPasteBaseline();
        log.info("initialize: ppt paste baseline initialized");
      }
    } catch (e) {
      log.warn("initialize: initPptPasteBaseline error", toLoggableError(e));
    }
  }
  if (host === Office.HostType.Excel) {
    try {
      const settings = getCurrentSettings();
      if (settings.resizeOnPaste) {
        await initExcelPasteBaseline();
        log.info("initialize: excel paste baseline initialized");
      }
    } catch (e) {
      log.warn("initialize: initExcelPasteBaseline error", toLoggableError(e));
    }
  }
  await registerSelectionChangedEvent();
  if (host === Office.HostType.Word) {
    const settings = getCurrentSettings();
    if (settings.enabled && settings.resizeOnPaste) startWordPasteCountPolling();
    else stopWordPasteCountPolling();
  } else {
    stopWordPasteCountPolling();
  }
  if (host === Office.HostType.Excel) {
    const settings = getCurrentSettings();
    if (settings.resizeOnSelection) {
      stopExcelPasteCountPolling();
      startExcelActiveShapePolling();
    } else if (settings.resizeOnPaste) {
      stopExcelActiveShapePolling();
      startExcelPasteCountPolling();
    } else {
      stopExcelActiveShapePolling();
      stopExcelPasteCountPolling();
    }
  } else stopExcelActiveShapePolling();
  setStatus(`Ready (${BUILD_TAG})`);
  log.info("initialize: done", { build: BUILD_TAG });
}

;(Office as any).onReady(async (info: any) => {
  const log = getLog();
  log.info("Office.onReady", { host: info.host, platform: info.platform });
  uiState.webUnsupported = isOfficeOnlinePlatform((info as any)?.platform);
  try { await initialize(); } catch (e) { log.error("Office.onReady error", e); setStatus("Initialization failed"); }
});

export { getCurrentSettings, getResizeSettings, handleSelectionChange, applySettings, handleToggleRefBox, startRefBoxPolling, stopRefBoxPolling, buildUI, initialize };
