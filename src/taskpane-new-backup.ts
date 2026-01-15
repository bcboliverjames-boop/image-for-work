/**
 * Office Paste Width - Taskpane UI v24
 * 修复: 参考框删除使用前缀匹配, 保存后更新 UI 显示
 */

import { setStatus, hasOfficeContext, cmToPoints } from "./utils";
import { getBoolSetting, getNumberSetting, saveSetting } from "./settings";
import { adjustSelectedWordObject } from "./resize/word";
import { adjustSelectedExcelShape } from "./resize/excel";
import { adjustSelectedPptShape } from "./resize/powerpoint";
import { initPasteBaseline, checkCountChange, handleCountIncrease } from "./resize/word-paste-detector";
import { insertReferenceBox, removeReferenceBox, getReferenceBoxState } from "./reference-box";
import { captureSizeSnapshot, createRegistry } from "./resize-scope";
import "./logger";

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
}

const uiState: UIState = {
  pendingSettings: null,
  refBoxActive: false,
  refBoxShapeName: null,
  refBoxPollingInterval: null,
};

function getLog() {
  return (window as any)._pasteLog || {
    debug: (msg: string, data?: any) => console.debug("[taskpane]", msg, data ?? ""),
    info: (msg: string, data?: any) => console.info("[taskpane]", msg, data ?? ""),
    warn: (msg: string, data?: any) => console.warn("[taskpane]", msg, data ?? ""),
    error: (msg: string, data?: any) => console.error("[taskpane]", msg, data ?? ""),
  };
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

function createToggleSwitch(id: string, label: string, checked: boolean, onChange: (checked: boolean) => void): HTMLDivElement {
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
  const knob = document.createElement("div");
  knob.className = "toggle-switch-knob";
  toggle.appendChild(knob);
  toggle.addEventListener("click", () => {
    const newState = !toggle.classList.contains("active");
    toggle.classList.toggle("active", newState);
    toggle.setAttribute("aria-checked", String(newState));
    onChange(newState);
  });
  container.appendChild(labelEl);
  container.appendChild(toggle);
  return container;
}

function createRadioButton(name: string, value: string, label: string, checked: boolean, onChange: () => void): HTMLDivElement {
  const container = document.createElement("div");
  const input = document.createElement("input");
  input.type = "radio";
  input.name = name;
  input.value = value;
  input.id = `${name}-${value}`;
  input.className = "modern-radio";
  input.checked = checked;
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
  input.addEventListener("change", () => { if (input.checked) onChange(); });
  container.appendChild(input);
  container.appendChild(labelEl);
  return container;
}

function createInput(id: string, label: string, value: number, onChange: (value: number) => void, unit: string = "cm"): HTMLDivElement {
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
  input.min = "0.1";
  input.max = "100";
  input.step = "0.5";
  const unitEl = document.createElement("span");
  unitEl.className = "input-unit";
  unitEl.textContent = unit;
  const handleChange = () => {
    const val = parseFloat(input.value);
    if (!isNaN(val) && val > 0) onChange(val);
  };
  input.addEventListener("change", handleChange);
  input.addEventListener("input", handleChange);
  row.appendChild(labelEl);
  row.appendChild(input);
  row.appendChild(unitEl);
  return row;
}

function createButton(text: string, onClick: () => void, variant: "primary" | "secondary" = "primary", fullWidth: boolean = false): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `modern-button modern-button-${variant}${fullWidth ? " modern-button-full" : ""}`;
  button.textContent = text;
  button.addEventListener("click", onClick);
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
  return {
    targetWidthCm: settings.targetWidthCm,
    targetHeightCm: settings.targetHeightCm,
    setHeightEnabled: settings.lockHeight,
    applyWidth: true,
    applyHeight: settings.lockHeight,
    lockAspectRatio: !settings.lockHeight,
  };
}

function startRefBoxPolling(): void {
  const log = getLog();
  if (uiState.refBoxPollingInterval !== null) return;
  log.info("startRefBoxPolling: start");
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
    } catch (e) {
      log.warn("refBoxPolling error", e);
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

const LOCK_ICON_LOCKED = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path></svg>`;
const LOCK_ICON_UNLOCKED = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 9.9-1"></path></svg>`;
const MOUSE_CURSOR_ICON = `<svg class="mouse-cursor-svg" width="24" height="24" viewBox="0 0 16 16"><rect x="1" y="1" width="10" height="8" fill="none" stroke="#E67E22" stroke-width="1.2" stroke-dasharray="2,1" rx="0.5"/><polygon points="9,7 9,14 11,12 13,15 14,14 12,11 14,10" fill="#4A7DC4"/></svg>`;

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
        uiState.refBoxActive = true;
        uiState.refBoxShapeName = shapeName;
        startRefBoxPolling();
        if (btn) {
          btn.classList.add("active");
          btn.textContent = "Click to End Dragging";
        }
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
  if (widthInput) {
    const width = parseFloat(widthInput.value);
    if (!isNaN(width) && width > 0) saveSetting("targetWidthCm", width);
  }
  if (heightInput) {
    const height = parseFloat(heightInput.value);
    if (!isNaN(height) && height > 0) saveSetting("targetHeightCm", height);
  }
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
      log.warn("applySettings: update size snapshot failed", e);
    }
  }
  
  setStatus("Settings saved");
  log.info("applySettings: done");
}

async function handleSelectionChange(): Promise<void> {
  const log = getLog();
  const settings = getCurrentSettings();
  if (!settings.enabled) return;

  if (settings.resizeOnPaste) {
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
            if (resizeResult !== "none") setStatus(`Resized ${resizeResult} image (${method})`);
          }
        );
      }
    } catch (e) { log.error("handleSelectionChange paste mode error", e); }
  }

  if (settings.resizeOnSelection) {
    try {
      const host = Office.context.host;
      if (host === Office.HostType.Word) {
        await Word.run(async (ctx) => { await adjustSelectedWordObject(ctx, getResizeSettings()); });
      } else if (host === Office.HostType.Excel) {
        await adjustSelectedExcelShape(getResizeSettings());
      } else if (host === Office.HostType.PowerPoint) {
        await adjustSelectedPptShape(getResizeSettings());
      }
    } catch (e) { log.error("handleSelectionChange selection mode error", e); }
  }
}

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
        () => { void handleSelectionChange(); },
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
        ctx.workbook.onSelectionChanged.add(handleSelectionChange);
        await ctx.sync();
        log.info("registerSelectionChangedEvent: Excel SelectionChanged registered");
      });
    } catch (e) { log.error("registerSelectionChangedEvent Excel error", e); }
  }
}

function buildUI(): void {
  const log = getLog();
  const container = document.getElementById("ui-container");
  if (!container) { log.error("buildUI: ui-container not found"); return; }
  container.innerHTML = "";
  const settings = getCurrentSettings();

  const enableSection = createCardSection();
  enableSection.appendChild(createToggleSwitch("enableToggle", "Enable Auto Resize", settings.enabled, (checked) => {
    saveSetting("enabled", checked);
    setStatus(checked ? "Enabled" : "Disabled");
  }));
  container.appendChild(enableSection);

  const modeSection = createCardSection("Mode Selection");
  const radioContainer = document.createElement("div");
  radioContainer.className = "radio-container";
  radioContainer.appendChild(createRadioButton("mode", "paste", "Resize on Paste", settings.resizeOnPaste, () => {
    saveSetting("resizeOnPaste", true);
    saveSetting("resizeOnSelection", false);
    setStatus("Mode: Paste");
  }));
  radioContainer.appendChild(createRadioButton("mode", "selection", "Resize on Selection", settings.resizeOnSelection && !settings.resizeOnPaste, () => {
    saveSetting("resizeOnPaste", false);
    saveSetting("resizeOnSelection", true);
    setStatus("Mode: Selection");
  }));
  modeSection.appendChild(radioContainer);
  container.appendChild(modeSection);

  const sizeSection = createCardSection("Size Settings");

  const refBoxBtn = document.createElement("button");
  refBoxBtn.type = "button";
  refBoxBtn.id = "refBoxBtn";
  refBoxBtn.className = "draggable-size-btn";
  refBoxBtn.innerHTML = `<span class="mouse-icon">${MOUSE_CURSOR_ICON}</span> Draggable Size Setting`;
  refBoxBtn.addEventListener("click", () => void handleToggleRefBox());
  sizeSection.appendChild(refBoxBtn);

  sizeSection.appendChild(createInput("widthInput", "Width", settings.targetWidthCm, (value) => {
    if (!uiState.pendingSettings) {
      uiState.pendingSettings = { targetWidthCm: value, targetHeightCm: settings.targetHeightCm, lockHeight: settings.lockHeight, lockAspectRatio: !settings.lockHeight };
    } else { uiState.pendingSettings.targetWidthCm = value; }
  }));

  const heightRow = createInput("heightInput", "Height", settings.targetHeightCm, (value) => {
    if (!uiState.pendingSettings) {
      uiState.pendingSettings = { targetWidthCm: settings.targetWidthCm, targetHeightCm: value, lockHeight: settings.lockHeight, lockAspectRatio: !settings.lockHeight };
    } else { uiState.pendingSettings.targetHeightCm = value; }
  });

  const lockBtn = document.createElement("button");
  lockBtn.type = "button";
  lockBtn.id = "lockHeightBtn";
  lockBtn.className = `lock-icon-btn${settings.lockHeight ? " active" : ""}`;
  lockBtn.innerHTML = settings.lockHeight ? LOCK_ICON_LOCKED : LOCK_ICON_UNLOCKED;
  lockBtn.title = settings.lockHeight ? "Height locked" : "Height unlocked";
  lockBtn.addEventListener("click", () => {
    const newLockState = !lockBtn.classList.contains("active");
    lockBtn.classList.toggle("active", newLockState);
    lockBtn.innerHTML = newLockState ? LOCK_ICON_LOCKED : LOCK_ICON_UNLOCKED;
    lockBtn.title = newLockState ? "Height locked" : "Height unlocked";
    saveSetting("setHeightEnabled", newLockState);
  });
  heightRow.appendChild(lockBtn);
  sizeSection.appendChild(heightRow);

  sizeSection.appendChild(createButton("Save", () => void applySettings(), "primary", true));
  container.appendChild(sizeSection);

  const statusSection = createCardSection("Status");
  const statusContainer = document.createElement("div");
  statusContainer.className = "status-container";
  statusContainer.id = "status";
  statusSection.appendChild(statusContainer);
  container.appendChild(statusSection);

  log.info("buildUI: UI built v23");
}

async function initialize(): Promise<void> {
  const log = getLog();
  log.info("initialize: start v23");
  buildUI();
  if (!hasOfficeContext()) {
    log.warn("initialize: No Office context, running in browser mode");
    setStatus("Browser mode (no Office)");
    return;
  }
  const host = Office.context.host;
  log.info("initialize: Office host", { host });
  if (host === Office.HostType.Word) {
    try { await initPasteBaseline(); log.info("initialize: paste baseline initialized"); } catch (e) { log.error("initialize: initPasteBaseline error", e); }
    try {
      // 捕获当前所有图片的尺寸快照（包含目标尺寸）
      const settings = getCurrentSettings();
      await captureSizeSnapshot({
        targetWidthPt: cmToPoints(settings.targetWidthCm),
        targetHeightPt: cmToPoints(settings.targetHeightCm),
      });
      log.info("initialize: size snapshot captured");
    } catch (e) { log.error("initialize: captureSizeSnapshot error", e); }
    // 创建内存表（粘贴模式只有新图片模式）
    createRegistry();
    log.info("initialize: registry created");
  }
  await registerSelectionChangedEvent();
  setStatus("Ready (v23)");
  log.info("initialize: done v23");
}

Office.onReady(async (info) => {
  const log = getLog();
  log.info("Office.onReady", { host: info.host, platform: info.platform });
  try { await initialize(); } catch (e) { log.error("Office.onReady error", e); setStatus("Initialization failed"); }
});

export { getCurrentSettings, getResizeSettings, handleSelectionChange, applySettings, handleToggleRefBox, startRefBoxPolling, stopRefBoxPolling, buildUI, initialize };

