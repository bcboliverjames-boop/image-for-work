const fs = require('fs');
const path = require('path');

const content = `/**
 * Office Paste Width - Taskpane UI v12
 * Compact modern SaaS-style UI with segmented card layout
 */

import { setStatus, hasOfficeContext } from "./utils";
import {
  getBoolSetting,
  getNumberSetting,
  saveSetting,
  getResizeScope,
  saveResizeScope,
  getResizeSettings,
} from "./settings";
import { adjustSelectedWordObject } from "./resize/word";
import { adjustSelectedExcelShape } from "./resize/excel";
import { adjustSelectedPptShape } from "./resize/powerpoint";
import {
  initPasteBaseline,
  checkCountChange,
  handleCountIncrease,
} from "./resize/word-paste-detector";
import {
  insertReferenceBox,
  removeReferenceBox,
  getReferenceBoxState,
} from "./reference-box";
import {
  captureBaseline,
} from "./resize-scope";
import "./logger";

// Types
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
}

// State
const uiState: UIState = {
  pendingSettings: null,
  refBoxActive: false,
  refBoxShapeName: null,
};

// UI Component Helpers
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
  onChange: (checked: boolean) => void
): HTMLDivElement {
  const container = document.createElement("div");
  container.className = "toggle-container";
  const labelEl = document.createElement("span");
  labelEl.className = "toggle-label";
  labelEl.textContent = label;
  const toggle = document.createElement("div");
  toggle.className = \`toggle-switch\${checked ? " active" : ""}\`;
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
  toggle.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      toggle.click();
    }
  });
  container.appendChild(labelEl);
  container.appendChild(toggle);
  return container;
}

function createCheckbox(
  id: string,
  label: string,
  checked: boolean,
  onChange: (checked: boolean) => void,
  disabled: boolean = false
): HTMLDivElement {
  const container = document.createElement("div");
  container.className = "checkbox-container";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.id = id;
  input.className = "modern-checkbox";
  input.checked = checked;
  input.disabled = disabled;
  const labelEl = document.createElement("label");
  labelEl.className = "modern-checkbox-label";
  labelEl.htmlFor = id;
  const box = document.createElement("span");
  box.className = "modern-checkbox-box";
  const checkmark = document.createElement("span");
  checkmark.className = "modern-checkbox-checkmark";
  checkmark.textContent = "✓";
  box.appendChild(checkmark);
  const text = document.createElement("span");
  text.textContent = label;
  labelEl.appendChild(box);
  labelEl.appendChild(text);
  input.addEventListener("change", () => onChange(input.checked));
  container.appendChild(input);
  container.appendChild(labelEl);
  return container;
}

function createRadioButton(
  name: string,
  value: string,
  label: string,
  checked: boolean,
  onChange: (value: string) => void
): HTMLDivElement {
  const container = document.createElement("div");
  const input = document.createElement("input");
  input.type = "radio";
  input.name = name;
  input.value = value;
  input.id = \`\${name}-\${value}\`;
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
  input.addEventListener("change", () => { if (input.checked) onChange(value); });
  container.appendChild(input);
  container.appendChild(labelEl);
  return container;
}

function createInput(
  id: string,
  label: string,
  value: number,
  onChange: (value: number) => void,
  unit: string = "cm"
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

function createButton(
  text: string,
  onClick: () => void,
  variant: "primary" | "secondary" | "toggle" = "primary",
  fullWidth: boolean = false
): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = \`modern-button modern-button-\${variant}\${fullWidth ? " modern-button-full" : ""}\`;
  button.textContent = text;
  button.addEventListener("click", onClick);
  return button;
}

// Settings Management
function getCurrentSettings() {
  return {
    enabled: getBoolSetting("enabled", true),
    resizeOnPaste: getBoolSetting("resizeOnPaste", true),
    resizeOnSelection: getBoolSetting("resizeOnSelection", false),
    targetWidthCm: getNumberSetting("targetWidthCm", 15),
    targetHeightCm: getNumberSetting("targetHeightCm", 10),
    lockHeight: getBoolSetting("setHeightEnabled", false),
    lockAspectRatio: getBoolSetting("lockAspectRatio", true),
    resizeScope: getResizeScope(),
  };
}

function applySettings() {
  if (!uiState.pendingSettings) return;
  const { targetWidthCm, targetHeightCm, lockHeight, lockAspectRatio } = uiState.pendingSettings;
  saveSetting("targetWidthCm", targetWidthCm);
  saveSetting("targetHeightCm", targetHeightCm);
  saveSetting("setHeightEnabled", lockHeight);
  saveSetting("lockAspectRatio", lockAspectRatio);
  saveResizeScope("new");
  setStatus("Settings saved");
  captureBaseline().catch(console.error);
}

// UI Building
function buildUI() {
  const container = document.getElementById("ui-container");
  if (!container) { console.error("[Taskpane] ui-container not found"); return; }
  container.innerHTML = "";
  const settings = getCurrentSettings();
  uiState.pendingSettings = {
    targetWidthCm: settings.targetWidthCm,
    targetHeightCm: settings.targetHeightCm,
    lockHeight: settings.lockHeight,
    lockAspectRatio: settings.lockAspectRatio,
  };

  // Section 1: Enable Control
  const enableSection = createCardSection();
  enableSection.appendChild(createToggleSwitch("enabled", "Enabled (auto-resize)", settings.enabled, (checked) => {
    saveSetting("enabled", checked);
    setStatus(checked ? "Auto-resize enabled" : "Auto-resize disabled");
  }));
  container.appendChild(enableSection);

  // Section 2: Mode Selection (Radio buttons - mutually exclusive)
  const modeSection = createCardSection("Mode Selection");
  const radioContainer = document.createElement("div");
  radioContainer.className = "radio-container";
  const currentMode = settings.resizeOnPaste ? "paste" : "selection";
  const pasteRadio = createRadioButton("mode", "paste", "Resize on paste", currentMode === "paste", (value) => {
    saveSetting("resizeOnPaste", value === "paste");
    saveSetting("resizeOnSelection", value === "selection");
  });
  const selectionRadio = createRadioButton("mode", "selection", "Resize on selection", currentMode === "selection", (value) => {
    saveSetting("resizeOnPaste", value === "paste");
    saveSetting("resizeOnSelection", value === "selection");
  });
  radioContainer.appendChild(pasteRadio);
  radioContainer.appendChild(selectionRadio);
  modeSection.appendChild(radioContainer);
  container.appendChild(modeSection);

  // Section 3: Target Size Setting
  const sizeSection = createCardSection("Target Size Setting");
  
  const refBoxBtn = createButton("Draggable Size Setting", handleToggleRefBox, "toggle", true);
  refBoxBtn.id = "btnInsertRefBox";
  updateRefBoxButtonState(refBoxBtn);
  sizeSection.appendChild(refBoxBtn);

  const widthRow = createInput("targetWidthCm", "Width:", settings.targetWidthCm, (value) => {
    if (uiState.pendingSettings) {
      uiState.pendingSettings.targetWidthCm = value;
      updateLockAspectRatioDisplay();
    }
  });
  sizeSection.appendChild(widthRow);

  const heightRow = document.createElement("div");
  heightRow.className = "input-row";
  const heightLabel = document.createElement("span");
  heightLabel.className = "input-label";
  heightLabel.textContent = "Height:";
  const heightInput = document.createElement("input");
  heightInput.type = "number";
  heightInput.id = "targetHeightCm";
  heightInput.className = "modern-input";
  heightInput.value = String(settings.targetHeightCm);
  heightInput.min = "0.1";
  heightInput.max = "100";
  heightInput.step = "0.5";
  heightInput.addEventListener("change", () => {
    const val = parseFloat(heightInput.value);
    if (!isNaN(val) && val > 0 && uiState.pendingSettings) {
      uiState.pendingSettings.targetHeightCm = val;
      updateLockAspectRatioDisplay();
    }
  });
  heightInput.addEventListener("input", () => {
    const val = parseFloat(heightInput.value);
    if (!isNaN(val) && val > 0 && uiState.pendingSettings) {
      uiState.pendingSettings.targetHeightCm = val;
      updateLockAspectRatioDisplay();
    }
  });
  const heightUnit = document.createElement("span");
  heightUnit.className = "input-unit";
  heightUnit.textContent = "cm";
  
  const lockHeightBtn = document.createElement("button");
  lockHeightBtn.type = "button";
  lockHeightBtn.id = "lockHeightBtn";
  lockHeightBtn.className = \`lock-icon-btn\${settings.lockHeight ? " active" : ""}\`;
  lockHeightBtn.innerHTML = settings.lockHeight ? "🔒" : "🔓";
  lockHeightBtn.title = settings.lockHeight ? "Height locked" : "Height unlocked";
  lockHeightBtn.addEventListener("click", () => {
    if (uiState.pendingSettings) {
      uiState.pendingSettings.lockHeight = !uiState.pendingSettings.lockHeight;
      lockHeightBtn.classList.toggle("active", uiState.pendingSettings.lockHeight);
      lockHeightBtn.innerHTML = uiState.pendingSettings.lockHeight ? "🔒" : "🔓";
      lockHeightBtn.title = uiState.pendingSettings.lockHeight ? "Height locked" : "Height unlocked";
    }
  });
  
  heightRow.appendChild(heightLabel);
  heightRow.appendChild(heightInput);
  heightRow.appendChild(heightUnit);
  heightRow.appendChild(lockHeightBtn);
  sizeSection.appendChild(heightRow);

  const aspectRatioRow = document.createElement("div");
  aspectRatioRow.className = "aspect-ratio-display";
  aspectRatioRow.id = "aspectRatioDisplay";
  const aspectIcon = document.createElement("span");
  aspectIcon.className = "aspect-icon";
  aspectIcon.textContent = settings.lockAspectRatio ? "🔗" : "⛓️‍💥";
  const aspectText = document.createElement("span");
  aspectText.className = "aspect-text";
  aspectText.textContent = settings.lockAspectRatio ? "Aspect ratio locked" : "Aspect ratio unlocked";
  aspectRatioRow.appendChild(aspectIcon);
  aspectRatioRow.appendChild(aspectText);
  sizeSection.appendChild(aspectRatioRow);

  const saveBtn = createButton("Save", handleSave, "primary", true);
  saveBtn.id = "btnSave";
  sizeSection.appendChild(saveBtn);
  container.appendChild(sizeSection);

  const statusSection = createCardSection("System Message");
  const statusContainer = document.createElement("div");
  statusContainer.id = "status";
  statusContainer.className = "status-container";
  statusSection.appendChild(statusContainer);
  container.appendChild(statusSection);
  setStatus("Ready");
}

function handleSave() {
  applySettings();
}

async function handleToggleRefBox() {
  const btn = document.getElementById("btnInsertRefBox") as HTMLButtonElement;
  try {
    if (uiState.refBoxActive && uiState.refBoxShapeName) {
      await removeReferenceBox(uiState.refBoxShapeName);
      uiState.refBoxActive = false;
      uiState.refBoxShapeName = null;
      setStatus("Reference box removed");
    } else {
      setStatus("Inserting reference box...");
      const shapeName = await insertReferenceBox();
      uiState.refBoxActive = true;
      uiState.refBoxShapeName = shapeName;
      setStatus("Reference box inserted. Drag to resize, then click Save.");
    }
    updateRefBoxButtonState(btn);
  } catch (error) {
    setStatus(\`Error: \${error}\`);
  }
}

function updateRefBoxButtonState(btn: HTMLButtonElement) {
  if (uiState.refBoxActive) {
    btn.classList.add("active");
    btn.textContent = "✓ Draggable Size Setting";
  } else {
    btn.classList.remove("active");
    btn.textContent = "Draggable Size Setting";
  }
}

function updateLockAspectRatioDisplay() {
  const display = document.getElementById("aspectRatioDisplay");
  if (!display || !uiState.pendingSettings) return;
  
  const { targetWidthCm, targetHeightCm } = uiState.pendingSettings;
  const isLocked = targetWidthCm > 0 && targetHeightCm > 0;
  uiState.pendingSettings.lockAspectRatio = isLocked;
  
  const icon = display.querySelector(".aspect-icon");
  const text = display.querySelector(".aspect-text");
  if (icon) icon.textContent = isLocked ? "🔗" : "⛓️‍💥";
  if (text) text.textContent = isLocked ? "Aspect ratio locked" : "Aspect ratio unlocked";
}

async function handleRefBoxUpdate() {
  try {
    const state = await getReferenceBoxState();
    if (state.exists) {
      uiState.refBoxActive = true;
      uiState.refBoxShapeName = state.shapeName;
      const widthCm = Math.round(state.widthCm * 10) / 10;
      const heightCm = Math.round(state.heightCm * 10) / 10;
      if (uiState.pendingSettings) {
        uiState.pendingSettings.targetWidthCm = widthCm;
        uiState.pendingSettings.targetHeightCm = heightCm;
      }
      const widthInput = document.getElementById("targetWidthCm") as HTMLInputElement;
      const heightInput = document.getElementById("targetHeightCm") as HTMLInputElement;
      if (widthInput) widthInput.value = String(widthCm);
      if (heightInput) heightInput.value = String(heightCm);
      updateLockAspectRatioDisplay();
    } else {
      uiState.refBoxActive = false;
      uiState.refBoxShapeName = null;
    }
    const btn = document.getElementById("btnInsertRefBox") as HTMLButtonElement;
    if (btn) updateRefBoxButtonState(btn);
  } catch (error) {
    console.error("[Taskpane] Error getting reference box size:", error);
  }
}

let pollingInterval: number | null = null;

function startPolling() {
  if (pollingInterval) return;
  pollingInterval = window.setInterval(async () => {
    if (!hasOfficeContext()) return;
    const settings = getCurrentSettings();
    if (!settings.enabled) return;
    try {
      if (settings.resizeOnPaste) {
        const countResult = await checkCountChange();
        if (countResult.inlineIncreased || countResult.shapeIncreased) {
          const resizeSettings = getResizeSettings();
          handleCountIncrease(resizeSettings, countResult.inlineIncreased, countResult.shapeIncreased, countResult.inlineCount, countResult.shapeCount, (result, method) => {
            if (result !== "none") setStatus(\`Resized pasted image (\${method})\`);
          });
        }
      }
      await handleRefBoxUpdate();
    } catch (error) { /* ignore */ }
  }, 500);
}

async function handleSelectionChange() {
  const settings = getCurrentSettings();
  if (!settings.enabled || !settings.resizeOnSelection) return;
  try {
    const hostType = Office.context.host;
    const resizeSettings = getResizeSettings();
    if (hostType === Office.HostType.Word) {
      await Word.run(async (context) => { await adjustSelectedWordObject(context, resizeSettings); });
    } else if (hostType === Office.HostType.Excel) {
      await adjustSelectedExcelShape(resizeSettings);
    } else if (hostType === Office.HostType.PowerPoint) {
      await adjustSelectedPptShape(resizeSettings);
    }
  } catch (error) { /* ignore */ }
}

async function initialize() {
  console.info("[Taskpane] Initializing v12...");
  buildUI();
  if (typeof Office !== "undefined" && Office.onReady) {
    await Office.onReady();
    console.info("[Taskpane] Office ready");
    try { await initPasteBaseline(); } catch (error) { console.warn("[Taskpane] Failed to init paste baseline:", error); }
    startPolling();
    try {
      const hostType = Office.context.host;
      if (hostType === Office.HostType.Word) {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChange);
      }
    } catch (error) { console.warn("[Taskpane] Failed to register selection handler:", error); }
  }
  console.info("[Taskpane] Initialized");
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initialize);
} else {
  initialize();
}

export { createCardSection, createToggleSwitch, createCheckbox, createRadioButton, createInput, createButton, getCurrentSettings, applySettings, uiState };

export function resetUIState() {
  uiState.pendingSettings = null;
  uiState.refBoxActive = false;
  uiState.refBoxShapeName = null;
}
`;

const targetPath = path.join(__dirname, 'src', 'taskpane-new.ts');
fs.writeFileSync(targetPath, content, 'utf8');
console.log('Written to:', targetPath);
console.log('File size:', fs.statSync(targetPath).size, 'bytes');

// Verify
const written = fs.readFileSync(targetPath, 'utf8');
if (written.includes('v12')) {
  console.log('SUCCESS: v12 found in file');
} else {
  console.log('ERROR: v12 NOT found');
}
