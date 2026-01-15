/**
 * Taskpane UI Redesign - Property-Based Tests
 * Feature: taskpane-ui-redesign
 * 
 * These tests verify the correctness properties defined in the design document.
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import * as fc from "fast-check";

// Mock Office.js and settings
vi.mock("./settings", () => ({
  getBoolSetting: vi.fn((_key: string, defaultValue: boolean) => defaultValue),
  getNumberSetting: vi.fn((_key: string, defaultValue: number) => defaultValue),
  saveSetting: vi.fn(),
  getResizeScope: vi.fn(() => "new" as const),
  saveResizeScope: vi.fn(),
  getResizeSettings: vi.fn(() => ({
    targetWidthCm: 15,
    targetHeightCm: 10,
    applyWidth: true,
    applyHeight: false,
    lockAspectRatio: true,
  })),
}));

vi.mock("./utils", () => ({
  setStatus: vi.fn(),
  hasOfficeContext: vi.fn(() => false),
}));

vi.mock("./resize/word", () => ({
  adjustSelectedWordObject: vi.fn(),
}));

vi.mock("./resize/excel", () => ({
  adjustSelectedExcelShape: vi.fn(),
}));

vi.mock("./resize/powerpoint", () => ({
  adjustSelectedPptShape: vi.fn(),
}));

vi.mock("./resize/word-paste-detector", () => ({
  initPasteBaseline: vi.fn(),
  checkCountChange: vi.fn(),
  handleCountIncrease: vi.fn(),
}));

vi.mock("./reference-box", () => ({
  insertReferenceBox: vi.fn(),
  getReferenceBoxState: vi.fn(() => ({ exists: false, shapeName: null, widthCm: 14, heightCm: 10 })),
}));

vi.mock("./resize-scope", () => ({
  captureBaseline: vi.fn(() => Promise.resolve()),
  scanAllImages: vi.fn(() => Promise.resolve({ scannedCount: 0, adjustedCount: 0, skippedCount: 0, mode: "all" })),
}));

vi.mock("./logger", () => ({}));

// Import after mocks
const taskpaneModule = await import("./taskpane-new");
const {
  createToggleSwitch,
  createCheckbox,
  createInput,
  showScopeDialog,
  applyPendingSettings,
  uiState,
  resetUIState,
} = taskpaneModule;
import { saveSetting, saveResizeScope } from "./settings";

describe("Taskpane UI Redesign - Property Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    document.body.innerHTML = "";
    // Reset UI state using exported function
    resetUIState();
  });

  /**
   * Property 1: Toggle Switch State Synchronization
   * For any toggle switch action (on or off), the enabled state in settings
   * SHALL immediately reflect the toggle's new position.
   * 
   * Feature: taskpane-ui-redesign, Property 1: Toggle Switch State Synchronization
   * Validates: Requirements 2.3
   */
  describe("Property 1: Toggle Switch State Synchronization", () => {
    it("should synchronize toggle state with settings for any sequence of toggles", () => {
      fc.assert(
        fc.property(
          fc.array(fc.boolean(), { minLength: 1, maxLength: 20 }),
          (toggleSequence) => {
            vi.clearAllMocks();
            let currentState = false;
            const toggle = createToggleSwitch("test-toggle", "Test", currentState, (checked) => {
              saveSetting("enabled", checked);
            });
            document.body.appendChild(toggle);
            const toggleEl = toggle.querySelector(".toggle-switch") as HTMLElement;

            // Apply each toggle in sequence
            for (const targetState of toggleSequence) {
              // Only click if we need to change state
              if ((toggleEl.classList.contains("active")) !== targetState) {
                toggleEl.click();
              }
              currentState = targetState;
            }

            // Verify final state matches
            const finalState = toggleEl.classList.contains("active");
            expect(finalState).toBe(currentState);
            
            // Verify saveSetting was called with correct final value
            const calls = vi.mocked(saveSetting).mock.calls;
            if (calls.length > 0) {
              const lastCall = calls[calls.length - 1];
              expect(lastCall[0]).toBe("enabled");
              expect(lastCall[1]).toBe(currentState);
            }
          }
        ),
        { numRuns: 100 }
      );
    });
  });


  /**
   * Property 2: Checkbox Setting Synchronization
   * For any checkbox toggle action in the Mode Selection section, the corresponding
   * setting (resizeOnPaste or resizeOnSelection) SHALL immediately update to match
   * the checkbox state.
   * 
   * Feature: taskpane-ui-redesign, Property 2: Checkbox Setting Synchronization
   * Validates: Requirements 3.4
   */
  describe("Property 2: Checkbox Setting Synchronization", () => {
    it("should synchronize checkbox state with settings for any sequence of toggles", () => {
      fc.assert(
        fc.property(
          fc.array(fc.boolean(), { minLength: 1, maxLength: 20 }),
          fc.constantFrom("resizeOnPaste", "resizeOnSelection"),
          (toggleSequence, settingKey) => {
            vi.clearAllMocks();
            let currentState = false;
            const checkbox = createCheckbox(settingKey, "Test", currentState, (checked) => {
              saveSetting(settingKey, checked);
            });
            document.body.appendChild(checkbox);
            const input = checkbox.querySelector("input") as HTMLInputElement;

            // Apply each toggle in sequence
            for (const targetState of toggleSequence) {
              if (input.checked !== targetState) {
                input.checked = targetState;
                input.dispatchEvent(new Event("change"));
              }
              currentState = targetState;
            }

            // Verify final state matches
            expect(input.checked).toBe(currentState);
            
            // Verify saveSetting was called with correct final value
            const calls = vi.mocked(saveSetting).mock.calls;
            if (calls.length > 0) {
              const lastCall = calls[calls.length - 1];
              expect(lastCall[0]).toBe(settingKey);
              expect(lastCall[1]).toBe(currentState);
            }
          }
        ),
        { numRuns: 100 }
      );
    });
  });

  /**
   * Property 3: Pending Settings Isolation
   * For any modification to width or height input values, the actual applied settings
   * SHALL remain unchanged until the Save button is clicked and a scope is selected.
   * 
   * Feature: taskpane-ui-redesign, Property 3: Pending Settings Isolation
   * Validates: Requirements 4.7
   */
  describe("Property 3: Pending Settings Isolation", () => {
    it("should not apply settings until save is clicked", () => {
      fc.assert(
        fc.property(
          fc.double({ min: 0.1, max: 100, noNaN: true }),
          fc.double({ min: 0.1, max: 100, noNaN: true }),
          (newWidth, newHeight) => {
            vi.clearAllMocks();
            
            // Initialize pending settings
            uiState.pendingSettings = {
              targetWidthCm: 15,
              targetHeightCm: 10,
              lockHeight: false,
              lockAspectRatio: true,
            };

            // Create inputs
            const widthInput = createInput("targetWidthCm", "Width:", 15, (value) => {
              if (uiState.pendingSettings) uiState.pendingSettings.targetWidthCm = value;
            });
            const heightInput = createInput("targetHeightCm", "Height:", 10, (value) => {
              if (uiState.pendingSettings) uiState.pendingSettings.targetHeightCm = value;
            });
            document.body.appendChild(widthInput);
            document.body.appendChild(heightInput);

            // Modify inputs
            const wInput = widthInput.querySelector("input") as HTMLInputElement;
            const hInput = heightInput.querySelector("input") as HTMLInputElement;
            wInput.value = String(newWidth);
            wInput.dispatchEvent(new Event("change"));
            hInput.value = String(newHeight);
            hInput.dispatchEvent(new Event("change"));

            // Verify saveSetting was NOT called (settings not applied)
            expect(vi.mocked(saveSetting)).not.toHaveBeenCalledWith("targetWidthCm", expect.anything());
            expect(vi.mocked(saveSetting)).not.toHaveBeenCalledWith("targetHeightCm", expect.anything());

            // Verify pending settings were updated
            expect(uiState.pendingSettings?.targetWidthCm).toBe(newWidth);
            expect(uiState.pendingSettings?.targetHeightCm).toBe(newHeight);
          }
        ),
        { numRuns: 100 }
      );
    });
  });

  /**
   * Property 4: Scope Selection Propagation
   * For any scope option selected in the save dialog ("All images" or "New images only"),
   * both the Resize_Scope setting and the Resize Scope section display SHALL update
   * to reflect the selected option.
   * 
   * Feature: taskpane-ui-redesign, Property 4: Scope Selection Propagation
   * Validates: Requirements 5.4, 7.4
   */
  describe("Property 4: Scope Selection Propagation", () => {
    it("should update scope setting when option is selected", () => {
      fc.assert(
        fc.property(
          fc.constantFrom("all", "new") as fc.Arbitrary<"all" | "new">,
          (selectedScope) => {
            vi.clearAllMocks();
            
            // Initialize pending settings
            uiState.pendingSettings = {
              targetWidthCm: 15,
              targetHeightCm: 10,
              lockHeight: false,
              lockAspectRatio: true,
            };

            // Apply pending settings with selected scope
            applyPendingSettings(selectedScope);

            // Verify saveResizeScope was called with correct scope
            expect(vi.mocked(saveResizeScope)).toHaveBeenCalledWith(selectedScope);
          }
        ),
        { numRuns: 100 }
      );
    });
  });


  /**
   * Property 5: Save Dialog Completion
   * For any scope option selection in the save dialog, the dialog SHALL close
   * and the pending settings SHALL be applied to the actual settings.
   * 
   * Feature: taskpane-ui-redesign, Property 5: Save Dialog Completion
   * Validates: Requirements 7.5
   */
  describe("Property 5: Save Dialog Completion", () => {
    it("should close dialog and apply settings when option is selected", () => {
      fc.assert(
        fc.property(
          fc.constantFrom("all", "new") as fc.Arbitrary<"all" | "new">,
          fc.double({ min: 0.1, max: 100, noNaN: true }),
          fc.double({ min: 0.1, max: 100, noNaN: true }),
          (selectedScope, width, height) => {
            vi.clearAllMocks();
            
            // Initialize pending settings
            uiState.pendingSettings = {
              targetWidthCm: width,
              targetHeightCm: height,
              lockHeight: false,
              lockAspectRatio: true,
            };

            let dialogClosed = false;
            let selectedOption: "all" | "new" | null = null;

            // Show dialog
            showScopeDialog(
              () => { selectedOption = "all"; },
              () => { selectedOption = "new"; },
              () => { dialogClosed = true; }
            );

            // Verify dialog is open
            expect(uiState.isScopeDialogOpen).toBe(true);
            const overlay = document.getElementById("scope-dialog-overlay");
            expect(overlay).not.toBeNull();

            // Click the appropriate option
            const options = overlay!.querySelectorAll(".scope-dialog-option");
            const optionIndex = selectedScope === "all" ? 0 : 1;
            (options[optionIndex] as HTMLElement).click();

            // Verify dialog closed
            expect(uiState.isScopeDialogOpen).toBe(false);
            expect(document.getElementById("scope-dialog-overlay")).toBeNull();

            // Verify correct option was selected
            expect(selectedOption).toBe(selectedScope);
          }
        ),
        { numRuns: 100 }
      );
    });
  });

  /**
   * Property 6: Cancel Preserves State
   * For any cancel action in the save dialog, the pending settings SHALL be discarded
   * and the actual settings SHALL remain unchanged from their state before the Save
   * button was clicked.
   * 
   * Feature: taskpane-ui-redesign, Property 6: Cancel Preserves State
   * Validates: Requirements 7.6
   */
  describe("Property 6: Cancel Preserves State", () => {
    it("should preserve original state when dialog is cancelled", () => {
      fc.assert(
        fc.property(
          fc.double({ min: 0.1, max: 100, noNaN: true }),
          fc.double({ min: 0.1, max: 100, noNaN: true }),
          (pendingWidth, pendingHeight) => {
            vi.clearAllMocks();
            
            // Initialize pending settings with modified values
            uiState.pendingSettings = {
              targetWidthCm: pendingWidth,
              targetHeightCm: pendingHeight,
              lockHeight: false,
              lockAspectRatio: true,
            };

            let cancelCalled = false;

            // Show dialog
            showScopeDialog(
              () => {},
              () => {},
              () => { cancelCalled = true; }
            );

            // Verify dialog is open
            expect(uiState.isScopeDialogOpen).toBe(true);

            // Click cancel
            const cancelBtn = document.querySelector(".scope-dialog-cancel") as HTMLElement;
            cancelBtn.click();

            // Verify dialog closed
            expect(uiState.isScopeDialogOpen).toBe(false);
            expect(cancelCalled).toBe(true);

            // Verify saveSetting was NOT called (settings not applied)
            expect(vi.mocked(saveSetting)).not.toHaveBeenCalledWith("targetWidthCm", expect.anything());
            expect(vi.mocked(saveSetting)).not.toHaveBeenCalledWith("targetHeightCm", expect.anything());
            expect(vi.mocked(saveResizeScope)).not.toHaveBeenCalled();
          }
        ),
        { numRuns: 100 }
      );
    });
  });
});

/**
 * Integration Tests for Taskpane UI
 * These tests verify the complete workflows work correctly
 */
describe("Taskpane UI Integration Tests", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    document.body.innerHTML = "";
    resetUIState();
  });

  /**
   * Test full save workflow: modify → save → select scope → verify
   */
  describe("Full Save Workflow", () => {
    it("should apply settings when save workflow is completed", () => {
      // Initialize pending settings
      uiState.pendingSettings = {
        targetWidthCm: 20,
        targetHeightCm: 15,
        lockHeight: true,
        lockAspectRatio: false,
      };

      let scopeSelected: "all" | "new" | null = null;

      // Show dialog (simulating Save button click)
      showScopeDialog(
        () => { scopeSelected = "all"; applyPendingSettings("all"); },
        () => { scopeSelected = "new"; applyPendingSettings("new"); },
        () => {}
      );

      // Select "All images" option
      const overlay = document.getElementById("scope-dialog-overlay");
      const allOption = overlay!.querySelector(".scope-dialog-option") as HTMLElement;
      allOption.click();

      // Verify settings were applied
      expect(scopeSelected).toBe("all");
      expect(vi.mocked(saveSetting)).toHaveBeenCalledWith("targetWidthCm", 20);
      expect(vi.mocked(saveSetting)).toHaveBeenCalledWith("targetHeightCm", 15);
      expect(vi.mocked(saveSetting)).toHaveBeenCalledWith("setHeightEnabled", true);
      expect(vi.mocked(saveSetting)).toHaveBeenCalledWith("lockAspectRatio", false);
      expect(vi.mocked(saveResizeScope)).toHaveBeenCalledWith("all");
    });
  });

  /**
   * Test cancel workflow: modify → save → cancel → verify unchanged
   */
  describe("Cancel Workflow", () => {
    it("should not apply settings when cancel is clicked", () => {
      // Initialize pending settings with modified values
      uiState.pendingSettings = {
        targetWidthCm: 25,
        targetHeightCm: 18,
        lockHeight: false,
        lockAspectRatio: true,
      };

      let cancelCalled = false;

      // Show dialog
      showScopeDialog(
        () => {},
        () => {},
        () => { cancelCalled = true; }
      );

      // Click cancel
      const cancelBtn = document.querySelector(".scope-dialog-cancel") as HTMLElement;
      cancelBtn.click();

      // Verify cancel was called and settings were NOT applied
      expect(cancelCalled).toBe(true);
      expect(vi.mocked(saveSetting)).not.toHaveBeenCalled();
      expect(vi.mocked(saveResizeScope)).not.toHaveBeenCalled();
    });
  });

  /**
   * Test initial load state
   */
  describe("Initial Load State", () => {
    it("should create toggle switch with correct initial state", () => {
      const toggle = createToggleSwitch("test", "Test Label", true, () => {});
      document.body.appendChild(toggle);

      const toggleEl = toggle.querySelector(".toggle-switch") as HTMLElement;
      expect(toggleEl.classList.contains("active")).toBe(true);
      expect(toggleEl.getAttribute("aria-checked")).toBe("true");
    });

    it("should create checkbox with correct initial state", () => {
      const checkbox = createCheckbox("test", "Test Label", true, () => {});
      document.body.appendChild(checkbox);

      const input = checkbox.querySelector("input") as HTMLInputElement;
      expect(input.checked).toBe(true);
    });

    it("should create input with correct initial value", () => {
      const inputRow = createInput("test", "Test:", 15.5, () => {});
      document.body.appendChild(inputRow);

      const input = inputRow.querySelector("input") as HTMLInputElement;
      expect(input.value).toBe("15.5");
    });
  });
});
