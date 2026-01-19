const SETTINGS_KEY_ENABLED = "opw_enabled";
const SETTINGS_KEY_RESIZE_ON_PASTE = "opw_resizeOnPaste";
const SETTINGS_KEY_RESIZE_ON_SELECTION = "opw_resizeOnSelection";
const SETTINGS_KEY_TARGET_WIDTH_CM = "opw_targetWidthCm";
const SETTINGS_KEY_SET_HEIGHT_ENABLED = "opw_setHeightEnabled";
const SETTINGS_KEY_TARGET_HEIGHT_CM = "opw_targetHeightCm";

const DEFAULT_ENABLED = true;
const DEFAULT_RESIZE_ON_PASTE = true;
const DEFAULT_RESIZE_ON_SELECTION = false;
const DEFAULT_TARGET_WIDTH_CM = 15;
const DEFAULT_SET_HEIGHT_ENABLED = false;
const DEFAULT_TARGET_HEIGHT_CM = 10;

let lastAdjustAt = 0;
let lastWordBodyShapeCount = 0;
let wordPollingInFlight = false;
let lastWordBodyInlinePictureCount = 0;
let wordBodyInlineIds = new Set<string>();
let wordBodyShapeIds = new Set<string>();
let wordSuppressPollingUntil = 0;
let wordSuppressSelectionUntil = 0;

let wordPasteBurstTimer: number | null = null;
let wordPasteBurstUntil = 0;
let wordPasteBurstNoneStreak = 0;
let wordPasteAnchorRange: Word.Range | null = null;

let wordPasteConfirmSeq = 0;
let wordPasteConfirmTimers: number[] = [];

let wordCountWatcherTimer: number | null = null;
let wordCountWatcherInFlight = false;
let wordWatcherInlineCount = 0;
let wordWatcherShapeCount = 0;
let wordWatcherFastUntil = 0;
let wordPictureSelected = false;
let lastWordWatcherErrorAt = 0;
let wordSelectionDebounceTimer: number | null = null;
let wordSelectionVerifyTimer: number | null = null;
let wordLastSelectionEventAt = 0;
let wordPasteSelectionProbeInFlight = false;

let wordPasteLastSelectionAutoResizeAt = 0;

const WORD_SELECTION_VERIFY_QUIET_MS = 200;

const WORD_SELECTION_DEBOUNCE_MS = 100;
const WORD_SELECTION_THROTTLE_MS = 120;
const OTHER_SELECTION_THROTTLE_MS = 800;

const WORD_WATCHER_FAST_INTERVAL_MS = 80;
const WORD_WATCHER_SLOW_INTERVAL_MS = 150;
const WORD_WATCHER_SELECTED_INTERVAL_MS = 500;

const WORD_WATCHER_PASTE_FAST_INTERVAL_MS = 150;
const WORD_WATCHER_PASTE_SLOW_INTERVAL_MS = 300;

const WORD_BURST_INTERVAL_MS = 120;
const WORD_PASTE_BURST_INTERVAL_MS = 120;
const WORD_BURST_NONE_STREAK_MAX = 12;
const WORD_PASTE_BURST_NONE_STREAK_MAX = 20;

const WORD_STATUS_ERROR_THROTTLE_MS = 2000;

const WORD_PICTURE_SUPPRESS_MS = 2000;

const WORD_PASTE_SELECTION_SUPPRESS_MS = 800;

const WORD_SIZE_EPSILON_PTS = 2;

function isWithinEpsilonPts(value: unknown, target: number, epsilon = WORD_SIZE_EPSILON_PTS): boolean {
  const v = Number(value);
  return Number.isFinite(v) && Math.abs(v - target) <= epsilon;
}

function cmToPoints(cm: number): number {
  return cm * 28.3464566929;
}

function getFeatureSettings(): {
  resizeOnPaste: boolean;
  resizeOnSelection: boolean;
} {
  const resizeOnPaste = getBoolSetting(SETTINGS_KEY_RESIZE_ON_PASTE, DEFAULT_RESIZE_ON_PASTE);
  const resizeOnSelection = getBoolSetting(
    SETTINGS_KEY_RESIZE_ON_SELECTION,
    DEFAULT_RESIZE_ON_SELECTION
  );

  if (resizeOnPaste && resizeOnSelection) {
    return { resizeOnPaste: true, resizeOnSelection: false };
  }
  if (!resizeOnPaste && !resizeOnSelection) {
    return { resizeOnPaste: true, resizeOnSelection: false };
  }

  return { resizeOnPaste, resizeOnSelection };
}

function getResizeSettings(): {
  targetWidthCm: number;
  setHeightEnabled: boolean;
  targetHeightCm: number;
  applyWidth: boolean;
  applyHeight: boolean;
  lockAspectRatio: boolean;
} {
  const targetWidthCm = getNumberSetting(SETTINGS_KEY_TARGET_WIDTH_CM, DEFAULT_TARGET_WIDTH_CM);
  const setHeightEnabled = getBoolSetting(SETTINGS_KEY_SET_HEIGHT_ENABLED, DEFAULT_SET_HEIGHT_ENABLED);
  const targetHeightCm = getNumberSetting(SETTINGS_KEY_TARGET_HEIGHT_CM, DEFAULT_TARGET_HEIGHT_CM);

  let applyWidth = Number.isFinite(targetWidthCm) && targetWidthCm > 0;
  const applyHeight = setHeightEnabled && Number.isFinite(targetHeightCm) && targetHeightCm > 0;

  // If user hasn't enabled height and also cleared width (0/empty), fall back to default width
  // so selection-based resize still works.
  const effectiveTargetWidthCm = !applyWidth && !applyHeight ? DEFAULT_TARGET_WIDTH_CM : targetWidthCm;
  if (!applyWidth && !applyHeight) applyWidth = true;

  const lockAspectRatio = (applyWidth && !applyHeight) || (!applyWidth && applyHeight);

  return {
    targetWidthCm: effectiveTargetWidthCm,
    setHeightEnabled,
    targetHeightCm,
    applyWidth,
    applyHeight,
    lockAspectRatio
  };
}

function getNumberSetting(key: string, fallback: number): number {
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

  const num = typeof raw === "number" ? raw : Number(raw);
  return Number.isFinite(num) ? num : fallback;
}

function getBoolSetting(key: string, fallback: boolean): boolean {
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

function saveSetting(key: string, value: unknown): Promise<void> {
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
    // ignore and fallback
  }

  try {
    window.localStorage.setItem(key, String(value));
  } catch {
    // ignore
  }

  return Promise.resolve();
}

function setStatus(message: string): void {
  const el = document.getElementById("status");
  if (el) el.textContent = message;
}

function hasOfficeContext(): boolean {
  const anyOffice = Office as unknown as { context?: unknown };
  const ctx = anyOffice.context as
    | { document?: unknown }
    | undefined;
  return Boolean(ctx?.document);
}

async function adjustLatestWordImageIfAdded(settings: {
  targetWidthCm: number;
  setHeightEnabled: boolean;
  targetHeightCm: number;
  applyWidth: boolean;
  applyHeight: boolean;
  lockAspectRatio: boolean;
}): Promise<"word-body-last-inline" | "word-body-last-picture" | "none"> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);
  const applyWidth = settings.applyWidth;
  const applyHeight = settings.applyHeight;
  const lockAspectRatio = settings.lockAspectRatio;

  const result = await Word.run(async (context: Word.RequestContext) => {
    const inlinePics = context.document.body.inlinePictures;
    const shapes = context.document.body.shapes;

    let inlineItems: any[] = [];
    let shapeItems: any[] = [];
    let inlineCount = 0;
    let shapeCount = 0;

    try {
      const inlineCountResult = (inlinePics as any).getCount();
      const shapeCountResult = (shapes as any).getCount();
      await context.sync();
      const inlineValue = (inlineCountResult as any)?.value;
      const shapeValue = (shapeCountResult as any)?.value;
      if (typeof inlineValue !== "number" || typeof shapeValue !== "number") {
        throw new Error("getCount did not return a number");
      }
      inlineCount = inlineValue;
      shapeCount = shapeValue;
    } catch {
      inlinePics.load("items");
      shapes.load("items");
      await context.sync();
      inlineItems = inlinePics.items || [];
      shapeItems = shapes.items || [];
      inlineCount = inlineItems.length;
      shapeCount = shapeItems.length;
    }

    // If all images were deleted, reflect it immediately so the next paste can be detected.
    if (inlineCount === 0) lastWordBodyInlinePictureCount = 0;
    if (shapeCount === 0) lastWordBodyShapeCount = 0;

    // If counts went down, clamp baselines immediately to avoid false "increased" detection
    // and avoid resizing existing images on a decrease/change tick.
    if (inlineCount < lastWordBodyInlinePictureCount) lastWordBodyInlinePictureCount = inlineCount;
    if (shapeCount < lastWordBodyShapeCount) lastWordBodyShapeCount = shapeCount;

    const prevInlineCount = lastWordBodyInlinePictureCount;
    const inlineIncreased = inlineCount > prevInlineCount;

    if (inlineIncreased && inlineCount >= 1) {
      try {
        const paraCandidates: any[] = [];
        const selection = context.document.getSelection();
        const anchorOrSelection = wordPasteAnchorRange ?? selection;
        let basePara: any;
        try {
          basePara = (anchorOrSelection as any).paragraphs.getFirst();
        } catch {
          basePara = null;
        }
        if (basePara) {
          paraCandidates.push(basePara);

          let prev: any = basePara;
          let next: any = basePara;
          for (let i = 0; i < 8; i += 1) {
            try {
              prev = (prev as any)?.getPrevious ? (prev as any).getPrevious() : null;
              if (prev) paraCandidates.push(prev);
            } catch {
              // ignore
            }
            try {
              next = (next as any)?.getNext ? (next as any).getNext() : null;
              if (next) paraCandidates.push(next);
            } catch {
              // ignore
            }
          }
        }

        const inlineCollections: any[] = [];
        for (const p of paraCandidates) {
          const c = (p as any).inlinePictures;
          if (c?.load) {
            c.load("items");
            inlineCollections.push(c);
          }
        }

        await context.sync();

        let pic: any | null = null;
        for (const c of inlineCollections) {
          const items = c.items || [];
          if (items.length >= 1) {
            pic = items[items.length - 1];
            break;
          }
        }

        if (!pic) {
          try {
            pic = (inlinePics as any).getItemAt(inlineCount - 1);
            (pic as any).load(["width", "height"]);
            await context.sync();
          } catch {
            pic = null;
          }
        }

        if (pic) {
          const changed = await applyWordInlinePictureSizeInContext(context, pic, settings);
          if (changed) {
            lastWordBodyInlinePictureCount = inlineCount;
            return "word-body-last-inline" as const;
          }
        }

        // Fallback: new image may not be the last item in body order (e.g. pasted near the top).
        // Probe a few indices from both ends to find a candidate that needs resize.
        try {
          const indices: number[] = [];
          for (let i = 0; i < 5; i += 1) {
            const back = inlineCount - 1 - i;
            if (back >= 0) indices.push(back);
          }
          for (let i = 0; i < 5 && i < inlineCount; i += 1) {
            indices.push(i);
          }

          for (const idx of indices) {
            let candidate: any;
            try {
              candidate = (inlinePics as any).getItemAt(idx);
            } catch {
              candidate = null;
            }
            if (!candidate) continue;

            const changed = await applyWordInlinePictureSizeInContext(context, candidate, settings);
            if (changed) {
              lastWordBodyInlinePictureCount = inlineCount;
              return "word-body-last-inline" as const;
            }
          }
        } catch {
          // ignore
        }
      } catch {
        // ignore
      }
    }

    const prevShapeCount = lastWordBodyShapeCount;
    const shapeIncreased = shapeCount > prevShapeCount;

    if (!shapeIncreased || shapeCount < 1) {
      return "none" as const;
    }

    const selection = context.document.getSelection();
    const anchorOrSelection = wordPasteAnchorRange ?? selection;

    try {
      const selShapes = (selection as any).shapes;
      if (selShapes?.load) {
        selShapes.load("items");
        await context.sync();
        if (selShapes.items?.length >= 1) {
          const shape = selShapes.items[0];
          (shape as any).load("type");
          await context.sync();
          const typeString = String((shape as any).type || "").toLowerCase();
          const looksLikeImage = typeString.includes("picture") || typeString.includes("image");
          if (looksLikeImage) {
            const changed = await applyWordShapeSizeInContext(context, selection as any, shape, settings);
            if (changed) {
              lastWordBodyShapeCount = shapeCount;
              return "word-body-last-picture" as const;
            }
          }
        }
      }
    } catch {
      // ignore
    }

    // Fallback: pasted floating picture may be anchored to a nearby paragraph,
    // while the cursor jumps to the next paragraph (selection.shapes is empty).
    // Try scanning shapes within ranges of nearby paragraphs.
    try {
      const paraCandidates: any[] = [];
      let basePara: any;
      try {
        basePara = (anchorOrSelection as any).paragraphs.getFirst();
      } catch {
        basePara = null;
      }

      if (basePara) {
        paraCandidates.push(basePara);

        let prev: any = basePara;
        let next: any = basePara;
        for (let i = 0; i < 8; i += 1) {
          try {
            prev = (prev as any)?.getPrevious ? (prev as any).getPrevious() : null;
            if (prev) paraCandidates.push(prev);
          } catch {
            // ignore
          }
          try {
            next = (next as any)?.getNext ? (next as any).getNext() : null;
            if (next) paraCandidates.push(next);
          } catch {
            // ignore
          }
        }
      }

      const rangeShapeCollections: any[] = [];
      for (const p of paraCandidates) {
        try {
          const r = (p as any)?.getRange ? (p as any).getRange() : null;
          const c = r ? (r as any).shapes : null;
          if (c?.load) {
            c.load("items/type");
            rangeShapeCollections.push(c);
          }
        } catch {
          // ignore
        }
      }

      await context.sync();

      for (const c of rangeShapeCollections) {
        const items = c.items || [];
        for (let i = items.length - 1; i >= 0; i -= 1) {
          const candidate = items[i];
          const typeString = String((candidate as any)?.type || "").toLowerCase();
          if (!(typeString.includes("picture") || typeString.includes("image"))) continue;

          const changed = await applyWordAnyShapeSizeInContext(context, candidate, settings);
          if (changed) {
            lastWordBodyShapeCount = shapeCount;
            return "word-body-last-picture" as const;
          }
        }
      }
    } catch {
      // ignore
    }

    let lastShape: any;
    try {
      lastShape = (shapes as any).getItemAt(shapeCount - 1);
      (lastShape as any).load("type");
      await context.sync();
    } catch {
      shapes.load("items");
      await context.sync();
      shapeItems = shapes.items || [];
      if (shapeItems.length < 1) return "none" as const;
      lastShape = shapeItems[shapeItems.length - 1];
      (lastShape as any).load("type");
      await context.sync();
    }

    try {
      const indices: number[] = [];
      for (let i = 0; i < 5; i += 1) {
        const back = shapeCount - 1 - i;
        if (back >= 0) indices.push(back);
      }
      for (let i = 0; i < 5 && i < shapeCount; i += 1) {
        indices.push(i);
      }

      for (const idx of indices) {
        let candidate: any;
        try {
          candidate = (shapes as any).getItemAt(idx);
          (candidate as any).load("type");
        } catch {
          candidate = null;
        }
        if (!candidate) continue;

        await context.sync();
        const typeString = String((candidate as any).type || "").toLowerCase();
        if (!(typeString.includes("picture") || typeString.includes("image"))) continue;

        const changed = await applyWordAnyShapeSizeInContext(context, candidate, settings);
        if (changed) {
          lastWordBodyShapeCount = shapeCount;
          return "word-body-last-picture" as const;
        }
      }
    } catch {
      // ignore
    }

    try {
      const typeString = String((lastShape as any)?.type || "").toLowerCase();
      if (typeString.includes("picture") || typeString.includes("image")) {
        const changed = await applyWordAnyShapeSizeInContext(context, lastShape, settings);
        if (changed) {
          lastWordBodyShapeCount = shapeCount;
          return "word-body-last-picture" as const;
        }
      }
    } catch {
      // ignore
    }

    return "none" as const;
  });

  return result;
}

async function initWordPollingBaseline(): Promise<void> {
  await Word.run(async (context: Word.RequestContext) => {
    const inlinePics = context.document.body.inlinePictures;
    const shapes = context.document.body.shapes;

    let inlineCount = 0;
    let shapeCount = 0;
    try {
      const inlineCountResult = (inlinePics as any).getCount();
      const shapeCountResult = (shapes as any).getCount();
      await context.sync();
      const inlineValue = (inlineCountResult as any)?.value;
      const shapeValue = (shapeCountResult as any)?.value;
      inlineCount = typeof inlineValue === "number" ? inlineValue : 0;
      shapeCount = typeof shapeValue === "number" ? shapeValue : 0;
    } catch {
      inlinePics.load("items");
      shapes.load("items");
      await context.sync();
      inlineCount = inlinePics.items?.length ?? 0;
      shapeCount = shapes.items?.length ?? 0;
    }

    lastWordBodyInlinePictureCount = inlineCount;
    lastWordBodyShapeCount = shapeCount;

    wordWatcherInlineCount = lastWordBodyInlinePictureCount;
    wordWatcherShapeCount = lastWordBodyShapeCount;

    wordBodyInlineIds = new Set<string>();
    wordBodyShapeIds = new Set<string>();
  });
}

async function getWordCollectionCount(collection: any, context: Word.RequestContext): Promise<number> {
  try {
    if (typeof collection?.getCount === "function") {
      const r = collection.getCount();
      await context.sync();
      const v = (r as any)?.value;
      return typeof v === "number" ? v : 0;
    }
  } catch {
    // ignore and fallback
  }

  try {
    collection.load("items");
    await context.sync();
    return collection.items?.length ?? 0;
  } catch {
    return 0;
  }
}

async function wordCountWatcherTick(): Promise<void> {
  if (!hasOfficeContext()) return;
  if (Office.context.host !== Office.HostType.Word) return;
  if (wordCountWatcherInFlight) return;
  if (wordPictureSelected) return;
  if (Date.now() < wordSuppressPollingUntil) return;

  const enabled = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
  if (!enabled) return;
  const { resizeOnPaste } = getFeatureSettings();
  if (!resizeOnPaste) return;

  try {
    wordCountWatcherInFlight = true;
    const changeInfo = await Word.run(async (context: Word.RequestContext) => {
      const inlinePics = context.document.body.inlinePictures;
      const shapes = context.document.body.shapes;

      const selection = context.document.getSelection();
      const selInlinePics = selection.inlinePictures;
      selInlinePics.load("items");
      const selShapes = (selection as any).shapes;
      if (selShapes?.load) selShapes.load("items");

      const prevInlineCount = wordWatcherInlineCount;
      const prevShapeCount = wordWatcherShapeCount;

      let inlineCount = 0;
      let shapeCount = 0;
      try {
        const inlineCountResult = (inlinePics as any).getCount();
        const shapeCountResult = (shapes as any).getCount();
        await context.sync();
        const inlineValue = (inlineCountResult as any)?.value;
        const shapeValue = (shapeCountResult as any)?.value;
        if (typeof inlineValue !== "number" || typeof shapeValue !== "number") {
          throw new Error("getCount did not return a number");
        }
        inlineCount = inlineValue;
        shapeCount = shapeValue;
      } catch {
        inlinePics.load("items");
        shapes.load("items");
        await context.sync();
        inlineCount = inlinePics.items?.length ?? 0;
        shapeCount = shapes.items?.length ?? 0;
      }

      const changedAny =
        inlineCount !== prevInlineCount || shapeCount !== prevShapeCount;
      const increased =
        inlineCount > prevInlineCount || shapeCount > prevShapeCount;

      try {
        const selHasInline = (selInlinePics.items?.length ?? 0) >= 1;
        const selHasShape = (selShapes?.items?.length ?? 0) >= 1;
        // Preserve the pre-paste insertion point: on increase ticks, Word may have already
        // moved the caret to a different paragraph (e.g. line-end paste causing newline).
        // Only refresh anchor when not increased (or when anchor is uninitialized).
        if (!selHasInline && !selHasShape && (!increased || !wordPasteAnchorRange)) {
          if (wordPasteAnchorRange) {
            try {
              (context.trackedObjects as any).remove(wordPasteAnchorRange as any);
            } catch {
              // ignore
            }
          }
          const anchor = (selection as any).getRange ? (selection as any).getRange("Start") : selection;
          try {
            (context.trackedObjects as any).add(anchor as any);
            wordPasteAnchorRange = anchor;
          } catch {
            // ignore
          }
        }
      } catch {
        // ignore
      }

      // Keep paste baseline in sync even when counts go down (e.g. user deletes all images).
      if (inlineCount < lastWordBodyInlinePictureCount) lastWordBodyInlinePictureCount = inlineCount;
      if (shapeCount < lastWordBodyShapeCount) lastWordBodyShapeCount = shapeCount;

      wordWatcherInlineCount = inlineCount;
      wordWatcherShapeCount = shapeCount;
      return { changedAny, increased };
    });

    if (changeInfo?.changedAny) {
      // Only start high-frequency burst when counts increase.
      if (changeInfo.increased) {
        startWordPasteBurst(6000);
        setWordWatcherFastWindow(4000);
      } else {
        setWordWatcherFastWindow(1500);
      }
      scheduleWordPasteConfirmRetries();
    }
  } catch (e) {
    const now = Date.now();
    if (now - lastWordWatcherErrorAt > WORD_STATUS_ERROR_THROTTLE_MS) {
      lastWordWatcherErrorAt = now;
      setStatus(`Enabled. Word watcher error: ${String(e)}`);
    }
  } finally {
    wordCountWatcherInFlight = false;
  }
}

function setWordWatcherFastWindow(durationMs: number): void {
  const now = Date.now();
  wordWatcherFastUntil = Math.max(wordWatcherFastUntil, now + durationMs);
  scheduleNextWordCountWatcherTick(0);
}

function startWordCountWatcher(): void {
  if (!hasOfficeContext()) return;
  if (Office.context.host !== Office.HostType.Word) return;
  if (wordCountWatcherTimer !== null) return;
  scheduleNextWordCountWatcherTick(0);
}

function stopWordCountWatcher(): void {
  if (wordCountWatcherTimer !== null) {
    window.clearTimeout(wordCountWatcherTimer);
    wordCountWatcherTimer = null;
  }
}

function scheduleNextWordCountWatcherTick(delayMs: number): void {
  if (Office.context.host !== Office.HostType.Word) return;
  if (wordCountWatcherTimer !== null) {
    window.clearTimeout(wordCountWatcherTimer);
    wordCountWatcherTimer = null;
  }

  wordCountWatcherTimer = window.setTimeout(() => {
    wordCountWatcherTimer = null;

    // If picture is selected, do not run Word.run; just reschedule.
    if (wordPictureSelected) {
      scheduleNextWordCountWatcherTick(WORD_WATCHER_SELECTED_INTERVAL_MS);
      return;
    }

    void (async () => {
      await wordCountWatcherTick();

      const now = Date.now();
      const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
      const fastInterval =
        resizeOnPaste && !resizeOnSelection
          ? WORD_WATCHER_PASTE_FAST_INTERVAL_MS
          : WORD_WATCHER_FAST_INTERVAL_MS;
      const slowInterval =
        resizeOnPaste && !resizeOnSelection
          ? WORD_WATCHER_PASTE_SLOW_INTERVAL_MS
          : WORD_WATCHER_SLOW_INTERVAL_MS;
      const nextDelay = now < wordWatcherFastUntil ? fastInterval : slowInterval;
      scheduleNextWordCountWatcherTick(nextDelay);
    })();
  }, Math.max(0, delayMs));
}

async function applyWordInlinePictureSizeInContext(
  context: Word.RequestContext,
  pic: any,
  settings: {
    targetWidthCm: number;
    setHeightEnabled: boolean;
    targetHeightCm: number;
    applyWidth: boolean;
    applyHeight: boolean;
    lockAspectRatio: boolean;
  }
): Promise<boolean> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);
  const applyWidth = settings.applyWidth;
  const applyHeight = settings.applyHeight;
  const lockAspectRatio = settings.lockAspectRatio;

  (pic as any).load(["width", "height", "lockAspectRatio"]);
  await context.sync();

  let changed = false;
  try {
    const currentLock = (pic as any).lockAspectRatio;
    if (typeof currentLock === "boolean" && currentLock !== lockAspectRatio) {
      (pic as any).lockAspectRatio = lockAspectRatio;
      changed = true;
    }
  } catch {
    // ignore
  }

  try {
    const w = Number((pic as any).width);
    if (applyWidth && (!Number.isFinite(w) || Math.abs(w - targetWidthPts) > WORD_SIZE_EPSILON_PTS)) {
      (pic as any).width = targetWidthPts;
      changed = true;
    }
  } catch {
    // ignore
  }

  try {
    const h = Number((pic as any).height);
    if (applyHeight && (!Number.isFinite(h) || Math.abs(h - targetHeightPts) > WORD_SIZE_EPSILON_PTS)) {
      (pic as any).height = targetHeightPts;
      changed = true;
    }
  } catch {
    // ignore
  }

  if (changed) {
    await context.sync();

    try {
      (pic as any).load(["width", "height"]);
      await context.sync();
      const okWidth = !applyWidth || isWithinEpsilonPts((pic as any).width, targetWidthPts);
      const okHeight = !applyHeight || isWithinEpsilonPts((pic as any).height, targetHeightPts);
      if (!okWidth || !okHeight) return false;
    } catch {
      return false;
    }
  }
  return changed;
}

async function applyWordAnyShapeSizeInContext(
  context: Word.RequestContext,
  shape: any,
  settings: {
    targetWidthCm: number;
    setHeightEnabled: boolean;
    targetHeightCm: number;
    applyWidth: boolean;
    applyHeight: boolean;
    lockAspectRatio: boolean;
  }
): Promise<boolean> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);
  const applyWidth = settings.applyWidth;
  const applyHeight = settings.applyHeight;
  const lockAspectRatio = settings.lockAspectRatio;

  let changed = false;
  try {
    (shape as any).load(["width", "height"]);
    await context.sync();
  } catch {
    // ignore
  }

  try {
    const currentLock = (shape as any).lockAspectRatio;
    if (typeof currentLock === "boolean" && currentLock !== lockAspectRatio) {
      (shape as any).lockAspectRatio = lockAspectRatio;
      changed = true;
    }
  } catch {
    // ignore
  }

  try {
    const w = Number((shape as any).width);
    if (applyWidth && (!Number.isFinite(w) || Math.abs(w - targetWidthPts) > WORD_SIZE_EPSILON_PTS)) {
      (shape as any).width = targetWidthPts;
      changed = true;
    }
  } catch {
    // ignore
  }
  try {
    const h = Number((shape as any).height);
    if (applyHeight && (!Number.isFinite(h) || Math.abs(h - targetHeightPts) > WORD_SIZE_EPSILON_PTS)) {
      (shape as any).height = targetHeightPts;
      changed = true;
    }
  } catch {
    // ignore
  }

  if (changed) {
    try {
      await context.sync();

      try {
        (shape as any).load(["width", "height"]);
        await context.sync();
        const okWidth = !applyWidth || isWithinEpsilonPts((shape as any).width, targetWidthPts);
        const okHeight = !applyHeight || isWithinEpsilonPts((shape as any).height, targetHeightPts);
        if (!okWidth || !okHeight) return false;
      } catch {
        return false;
      }
    } catch {
      // ignore
      return false;
    }
  }
  return changed;
}

async function applyWordShapeSizeInContext(
  context: Word.RequestContext,
  range: Word.Range,
  shape: any,
  settings: {
    targetWidthCm: number;
    setHeightEnabled: boolean;
    targetHeightCm: number;
    applyWidth: boolean;
    applyHeight: boolean;
    lockAspectRatio: boolean;
  }
): Promise<boolean> {
  const targetWidthPts = cmToPoints(settings.targetWidthCm);
  const targetHeightPts = cmToPoints(settings.targetHeightCm);
  const applyWidth = settings.applyWidth;
  const applyHeight = settings.applyHeight;
  const lockAspectRatio = settings.lockAspectRatio;

  (shape as any).load(["type", "width", "height"]);
  await context.sync();
  try {
    (shape as any).load(["lockAspectRatio"]);
    await context.sync();
  } catch {
    // ignore
  }

  const typeString = String((shape as any).type || "").toLowerCase();
  const looksLikeImage = typeString.includes("picture") || typeString.includes("image");

  let shouldHandle = looksLikeImage;
  if (!shouldHandle) {
    try {
      const ooxmlResult = range.getOoxml();
      await context.sync();
      const xml = String((ooxmlResult as any).value || "");
      const s = xml.toLowerCase();
      shouldHandle = s.includes("<pic:pic") || s.includes("w:drawing") || s.includes("v:imagedata");
    } catch {
      shouldHandle = false;
    }
  }

  if (!shouldHandle) return false;

  let changed = false;
  try {
    const currentLock = (shape as any).lockAspectRatio;
    if (typeof currentLock === "boolean" && currentLock !== lockAspectRatio) {
      (shape as any).lockAspectRatio = lockAspectRatio;
      changed = true;
    }
  } catch {
    // ignore
  }

  try {
    const w = Number((shape as any).width);
    if (applyWidth && (!Number.isFinite(w) || Math.abs(w - targetWidthPts) > WORD_SIZE_EPSILON_PTS)) {
      (shape as any).width = targetWidthPts;
      changed = true;
    }
  } catch {
    // ignore
  }

  try {
    const h = Number((shape as any).height);
    if (applyHeight && (!Number.isFinite(h) || Math.abs(h - targetHeightPts) > WORD_SIZE_EPSILON_PTS)) {
      (shape as any).height = targetHeightPts;
      changed = true;
    }
  } catch {
    // ignore
  }

  if (changed) {
    await context.sync();

    try {
      (shape as any).load(["width", "height"]);
      await context.sync();
      const okWidth = !applyWidth || isWithinEpsilonPts((shape as any).width, targetWidthPts);
      const okHeight = !applyHeight || isWithinEpsilonPts((shape as any).height, targetHeightPts);
      if (!okWidth || !okHeight) return false;
    } catch {
      return false;
    }
  }
  return changed;
}

async function adjustSelectedWordObjectWidthInContext(
  context: Word.RequestContext,
  settings: {
    targetWidthCm: number;
    setHeightEnabled: boolean;
    targetHeightCm: number;
    applyWidth: boolean;
    applyHeight: boolean;
    lockAspectRatio: boolean;
  }
): Promise<"word-inline" | "word-selection-shape" | "none"> {
  const range = context.document.getSelection();

  const inlinePics = range.inlinePictures;
  inlinePics.load("items");

  const selectionShapes = (range as any).shapes;
  if (selectionShapes?.load) selectionShapes.load("items");

  await context.sync();

  if (selectionShapes?.items?.length >= 1) {
    const shape = selectionShapes.items[0];
    const ok = await applyWordShapeSizeInContext(context, range, shape, settings);
    return ok ? ("word-selection-shape" as const) : ("none" as const);
  }

  if (inlinePics.items?.length >= 1) {
    const pic = inlinePics.items[0];
    const ok = await applyWordInlinePictureSizeInContext(context, pic, settings);
    return ok ? ("word-inline" as const) : ("none" as const);
  }

  return "none" as const;
}

async function adjustSelectedObjectWidth(settings: {
  targetWidthCm: number;
  setHeightEnabled: boolean;
  targetHeightCm: number;
  applyWidth: boolean;
  applyHeight: boolean;
  lockAspectRatio: boolean;
}): Promise<"word-inline" | "word-selection-shape" | "word-body-last-inline" | "word-body-last-picture" | "ppt-shape" | "excel-shape" | "none"> {
  const host = Office.context.host;

  if (host === Office.HostType.Word) {
    return await Word.run(async (context: Word.RequestContext) => {
      return await adjustSelectedWordObjectWidthInContext(context, settings);
    });
  }

  if (host === Office.HostType.PowerPoint) {
    const targetWidthPts = cmToPoints(settings.targetWidthCm);
    const targetHeightPts = cmToPoints(settings.targetHeightCm);
    const applyWidth = settings.applyWidth;
    const applyHeight = settings.applyHeight;
    const lockAspectRatio = settings.lockAspectRatio;
    const found = await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      if (shapes.items.length < 1) return false;

      const shape = shapes.items[0];
      shape.load(["width", "height"]);
      await context.sync();

      if (applyWidth && !applyHeight && Math.abs(shape.width - targetWidthPts) < 1) return true;

      try {
        (shape as any).lockAspectRatio = lockAspectRatio;
      } catch {
        // ignore
      }

      if (applyWidth) shape.width = targetWidthPts;
      if (applyHeight) shape.height = targetHeightPts;

      await context.sync();
      return true;
    });
    return found ? "ppt-shape" : "none";
  }

  if (host === Office.HostType.Excel) {
    const targetWidthPts = cmToPoints(settings.targetWidthCm);
    const targetHeightPts = cmToPoints(settings.targetHeightCm);
    const applyWidth = settings.applyWidth;
    const applyHeight = settings.applyHeight;
    const lockAspectRatio = settings.lockAspectRatio;
    const found = await Excel.run(async (context: Excel.RequestContext) => {
      const shape = context.workbook.getActiveShapeOrNullObject();
      shape.load(["isNullObject", "type", "width", "height"]);
      await context.sync();

      if (shape.isNullObject) return false;

      const typeString = String(shape.type || "").toLowerCase();
      const looksLikeImage = typeString.includes("image") || typeString.includes("picture");
      if (!looksLikeImage && typeof shape.type === "string") return false;

      if (applyWidth && !applyHeight && Math.abs(shape.width - targetWidthPts) < 1) return true;

      try {
        (shape as any).lockAspectRatio = lockAspectRatio;
      } catch {
        // ignore
      }

      if (applyWidth) shape.width = targetWidthPts;
      if (applyHeight) shape.height = targetHeightPts;

      await context.sync();
      return true;
    });
    return found ? "excel-shape" : "none";
  }

  return "none";
}

async function onSelectionChanged(): Promise<void> {
  const enabled = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
  if (!enabled) return;

  const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();

  if (Office.context.host === Office.HostType.Word) {
    wordLastSelectionEventAt = Date.now();

    if (!resizeOnSelection) {
      if (resizeOnPaste) {
        if (wordPasteSelectionProbeInFlight) return;

        // Paste-only mode: selecting anything should stop all Word.run polling.
        // We'll resume watcher only after selection is no longer an image.
        stopWordPasteBurst();
        stopWordCountWatcher();
        wordPictureSelected = true;
        wordSuppressPollingUntil = Date.now() + WORD_PASTE_SELECTION_SUPPRESS_MS;

        if (wordSelectionVerifyTimer !== null) {
          window.clearTimeout(wordSelectionVerifyTimer);
        }
        wordSelectionVerifyTimer = window.setTimeout(() => {
          wordSelectionVerifyTimer = null;
          void verifyWordPictureSelectedAfterQuiet();
        }, WORD_SELECTION_VERIFY_QUIET_MS);

        scheduleWordPasteConfirmRetries();
      }
      return;
    }

    if (wordSelectionDebounceTimer !== null) {
      window.clearTimeout(wordSelectionDebounceTimer);
    }
    wordSelectionDebounceTimer = window.setTimeout(() => {
      wordSelectionDebounceTimer = null;
      void processSelectionChanged();
    }, WORD_SELECTION_DEBOUNCE_MS);
    return;
  }

  await processSelectionChanged();
}

async function verifyWordPictureSelectedAfterQuiet(): Promise<void> {
  if (!hasOfficeContext()) return;
  if (Office.context.host !== Office.HostType.Word) return;

  const now = Date.now();
  if (now - wordLastSelectionEventAt < WORD_SELECTION_VERIFY_QUIET_MS) {
    if (wordSelectionVerifyTimer !== null) {
      window.clearTimeout(wordSelectionVerifyTimer);
    }
    wordSelectionVerifyTimer = window.setTimeout(() => {
      wordSelectionVerifyTimer = null;
      void verifyWordPictureSelectedAfterQuiet();
    }, WORD_SELECTION_VERIFY_QUIET_MS);
    return;
  }

  const before = wordPictureSelected;
  const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
  if (!resizeOnSelection && resizeOnPaste) {
    const selectionAt = wordLastSelectionEventAt;
    wordPictureSelected = await probeWordImageSelectedViaOfficeAsync();

    if (wordPictureSelected && wordPasteLastSelectionAutoResizeAt !== selectionAt) {
      wordPasteLastSelectionAutoResizeAt = selectionAt;
      try {
        const resizeSettings = getResizeSettings();
        const result = await Word.run(async (context: Word.RequestContext) => {
          return await adjustSelectedWordObjectWidthInContext(context, resizeSettings);
        });
        if (result === "word-inline" || result === "word-selection-shape") {
          setStatus(
            `Enabled. Target width: ${resizeSettings.targetWidthCm} cm.\nAdjusted: ${result}.\nLast check: ${new Date().toLocaleTimeString()}`
          );
        }
      } catch {
        // ignore
      }
    }

    // Paste-only mode: do not keep probing while an image remains selected.
    // If selection is not an image anymore, resume watcher.
    if (!wordPictureSelected) {
      startWordCountWatcher();
      setWordWatcherFastWindow(1000);
    }
    return;
  }

  await refreshWordPictureSelectedState();

  // If we left picture selection state, allow background watcher/burst to resume.
  if (before && !wordPictureSelected) {
    if (resizeOnPaste && !resizeOnSelection) {
      startWordCountWatcher();
      setWordWatcherFastWindow(1000);
    }
  }
}

function probeWordImageSelectedViaOfficeAsync(): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      const doc: any = (Office as any)?.context?.document;
      const coercionImage = (Office as any)?.CoercionType?.Image ?? "image";
      const coercionOoxml = (Office as any)?.CoercionType?.Ooxml ?? "ooxml";
      const imageFormat = (Office as any)?.ImageFormat?.Png ?? "png";

      if (!doc?.getSelectedDataAsync) {
        resolve(false);
        return;
      }

      doc.getSelectedDataAsync(coercionImage, { imageFormat }, (result: Office.AsyncResult<any>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(true);
          return;
        }

        doc.getSelectedDataAsync(coercionOoxml, (r2: Office.AsyncResult<any>) => {
          if (r2.status !== Office.AsyncResultStatus.Succeeded) {
            resolve(false);
            return;
          }

          const xml = typeof (r2 as any).value === "string" ? String((r2 as any).value) : "";
          const s = xml.toLowerCase();
          resolve(s.includes("<pic:pic") || s.includes("w:drawing") || s.includes("v:imagedata"));
        });
      });
    } catch {
      resolve(false);
    }
  });
}

async function refreshWordPictureSelectedState(): Promise<void> {
  if (!hasOfficeContext()) return;
  if (Office.context.host !== Office.HostType.Word) return;

  if (wordPasteSelectionProbeInFlight) return;

  try {
    wordPasteSelectionProbeInFlight = true;
    const pictureSelected = await Word.run(async (context: Word.RequestContext) => {
      const range = context.document.getSelection();
      const inlinePics = range.inlinePictures;
      inlinePics.load("items");

      const selectionShapes = (range as any).shapes;
      if (selectionShapes?.load) selectionShapes.load("items");

      await context.sync();

      if (selectionShapes?.items?.length >= 1) {
        const shape = selectionShapes.items[0];
        (shape as any).load(["type"]);
        await context.sync();
        const typeString = String((shape as any).type || "").toLowerCase();
        if (typeString.includes("picture") || typeString.includes("image")) return true;
      }

      if (inlinePics.items?.length >= 1) return true;

      return false;
    });

    wordPictureSelected = pictureSelected;
  } catch {
    // ignore
  } finally {
    wordPasteSelectionProbeInFlight = false;
  }
}

async function processSelectionChanged(ignoreSuppress = false, ignoreThrottle = false): Promise<void> {
  const enabled = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
  if (!enabled) return;

  const { resizeOnSelection, resizeOnPaste } = getFeatureSettings();
  if (!resizeOnSelection) return;

  const now = Date.now();
  if (
    !ignoreSuppress &&
    Office.context.host === Office.HostType.Word &&
    resizeOnPaste &&
    !resizeOnSelection &&
    now < wordSuppressSelectionUntil
  )
    return;
  if (!ignoreThrottle) {
    const threshold = Office.context.host === Office.HostType.Word ? WORD_SELECTION_THROTTLE_MS : OTHER_SELECTION_THROTTLE_MS;
    if (now - lastAdjustAt < threshold) return;
  }
  lastAdjustAt = now;

  const resizeSettings = getResizeSettings();

  try {
    const result = await adjustSelectedObjectWidth(resizeSettings);

    if (Office.context.host === Office.HostType.Word) {
      const nextSelected = result === "word-inline" || result === "word-selection-shape";
      if (wordPictureSelected && !nextSelected) {
        wordSuppressSelectionUntil = 0;
        wordSuppressPollingUntil = 0;
        setWordWatcherFastWindow(8000);
      }

      wordPictureSelected = nextSelected;
      if (wordPictureSelected) {
        stopWordPasteBurst();
      }
    }

    if (
      Office.context.host === Office.HostType.Word &&
      (result === "word-inline" || result === "word-selection-shape")
    ) {
      // Only suppress selection/polling when paste watcher is active (paste-only mode).
      // In selection mode, suppressing causes rapid clicks to be ignored (e.g. 1/3/5 ok, 2/4/6 skipped).
      if (resizeOnPaste && !resizeOnSelection) {
        const until = Date.now() + WORD_PICTURE_SUPPRESS_MS;
        wordSuppressSelectionUntil = until;
        wordSuppressPollingUntil = until;
        stopWordPasteBurst();
      } else {
        wordSuppressSelectionUntil = 0;
        wordSuppressPollingUntil = 0;
      }
    }

    if (Office.context.host === Office.HostType.Word && result === "none" && resizeOnPaste && !resizeOnSelection) {
      startWordPasteBurst(5000);
      setWordWatcherFastWindow(8000);
    }

    if (result === "none" && Office.context.host !== Office.HostType.Word) {
      const delays = [30, 80, 150];
      for (const delay of delays) {
        setTimeout(() => {
          void adjustSelectedObjectWidth(resizeSettings).catch((e) => {
            setStatus(`Enabled. Target width: ${resizeSettings.targetWidthCm} cm.\nLast error: ${String(e)}`);
          });
        }, delay);
      }
    }
    setStatus(
      `Enabled. Target width: ${resizeSettings.targetWidthCm} cm.\nAdjusted: ${result}.\nLast check: ${new Date().toLocaleTimeString()}`
    );
  } catch (e) {
    setStatus(`Enabled. Target width: ${resizeSettings.targetWidthCm} cm.\nLast error: ${String(e)}`);
  }
}

async function wordPastePollingTick(force = false): Promise<"word-body-last-inline" | "word-body-last-picture" | "none"> {
  if (!hasOfficeContext()) return "none";
  if (Office.context.host !== Office.HostType.Word) return "none";
  if (wordPollingInFlight) return "none";
  if (!force) {
    if (wordPictureSelected) return "none";
    if (Date.now() < wordSuppressPollingUntil) return "none";
  }

  const enabled = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
  if (!enabled) return "none";
  const { resizeOnPaste } = getFeatureSettings();
  if (!resizeOnPaste) return "none";

  const resizeSettings = getResizeSettings();
  try {
    wordPollingInFlight = true;
    const result = await adjustLatestWordImageIfAdded(resizeSettings);
    if (result !== "none") {
      if (force) cancelWordPasteConfirmRetries();
      lastAdjustAt = Date.now();
      setStatus(
        `Enabled. Target width: ${resizeSettings.targetWidthCm} cm.\nAdjusted: ${result}.\nLast check: ${new Date().toLocaleTimeString()}`
      );
    }
    return result;
  } catch (e) {
    setStatus(`Enabled. Target width: ${resizeSettings.targetWidthCm} cm.\nLast error: ${String(e)}`);
    return "none";
  } finally {
    wordPollingInFlight = false;
  }
}

function startWordPasteBurst(durationMs: number): void {
  if (Office.context.host !== Office.HostType.Word) return;
  if (!hasOfficeContext()) return;

  const now = Date.now();
  wordPasteBurstUntil = Math.max(wordPasteBurstUntil, now + durationMs);
  wordPasteBurstNoneStreak = 0;

  if (wordPasteBurstTimer !== null) return;

  // Run one tick immediately to reduce latency (don't wait for the first interval).
  void (async () => {
    if (Date.now() >= wordPasteBurstUntil) {
      stopWordPasteBurst();
      return;
    }

    const result = await wordPastePollingTick();
    if (result === "none") {
      wordPasteBurstNoneStreak += 1;
      const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
      const noneStreakMax =
        resizeOnPaste && !resizeOnSelection
          ? WORD_PASTE_BURST_NONE_STREAK_MAX
          : WORD_BURST_NONE_STREAK_MAX;
      if (wordPasteBurstNoneStreak >= noneStreakMax) {
        stopWordPasteBurst();
      }
    } else {
      stopWordPasteBurst();
    }
  })();

  const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
  const intervalMs = resizeOnPaste && !resizeOnSelection ? WORD_PASTE_BURST_INTERVAL_MS : WORD_BURST_INTERVAL_MS;
  const noneStreakMax =
    resizeOnPaste && !resizeOnSelection
      ? WORD_PASTE_BURST_NONE_STREAK_MAX
      : WORD_BURST_NONE_STREAK_MAX;

  wordPasteBurstTimer = window.setInterval(() => {
    void (async () => {
      if (Date.now() >= wordPasteBurstUntil) {
        stopWordPasteBurst();
        return;
      }

      const result = await wordPastePollingTick();
      if (result === "none") {
        wordPasteBurstNoneStreak += 1;
        if (wordPasteBurstNoneStreak >= noneStreakMax) {
          stopWordPasteBurst();
        }
      } else {
        stopWordPasteBurst();
      }
    })();
  }, intervalMs);
}

function cancelWordPasteConfirmRetries(): void {
  wordPasteConfirmSeq += 1;
  for (const t of wordPasteConfirmTimers) {
    window.clearTimeout(t);
  }
  wordPasteConfirmTimers = [];
}

function scheduleWordPasteConfirmRetries(): void {
  if (Office.context.host !== Office.HostType.Word) return;
  if (!hasOfficeContext()) return;

  cancelWordPasteConfirmRetries();
  const seq = wordPasteConfirmSeq;
  const delays = [200, 450, 900];

  for (const delay of delays) {
    const timer = window.setTimeout(() => {
      void (async () => {
        if (seq !== wordPasteConfirmSeq) return;
        const result = await wordPastePollingTick(true);
        if (seq !== wordPasteConfirmSeq) return;
        if (result !== "none") {
          cancelWordPasteConfirmRetries();
          stopWordPasteBurst();
        }
      })();
    }, delay);
    wordPasteConfirmTimers.push(timer);
  }
}

function stopWordPasteBurst(): void {
  if (wordPasteBurstTimer !== null) {
    window.clearInterval(wordPasteBurstTimer);
    wordPasteBurstTimer = null;
  }
  wordPasteBurstUntil = 0;
  wordPasteBurstNoneStreak = 0;
}

function bindUi(): void {
  const enabledEl = document.getElementById("enabled") as HTMLInputElement | null;
  const resizeOnPasteEl = document.getElementById("resizeOnPaste") as HTMLInputElement | null;
  const resizeOnSelectionEl = document.getElementById("resizeOnSelection") as HTMLInputElement | null;
  const widthEl = document.getElementById("targetWidthCm") as HTMLInputElement | null;
  const setHeightEnabledEl = document.getElementById("setHeightEnabled") as HTMLInputElement | null;
  const heightEl = document.getElementById("targetHeightCm") as HTMLInputElement | null;
  const lockAspectRatioEl = document.getElementById("lockAspectRatio") as HTMLInputElement | null;

  if (
    !enabledEl ||
    !resizeOnPasteEl ||
    !resizeOnSelectionEl ||
    !widthEl ||
    !setHeightEnabledEl ||
    !heightEl ||
    !lockAspectRatioEl
  ) {
    return;
  }

  enabledEl.checked = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
  const feature = getFeatureSettings();
  resizeOnPasteEl.checked = feature.resizeOnPaste;
  resizeOnSelectionEl.checked = feature.resizeOnSelection;
  widthEl.value = String(getNumberSetting(SETTINGS_KEY_TARGET_WIDTH_CM, DEFAULT_TARGET_WIDTH_CM));
  setHeightEnabledEl.checked = getBoolSetting(SETTINGS_KEY_SET_HEIGHT_ENABLED, DEFAULT_SET_HEIGHT_ENABLED);
  heightEl.value = String(getNumberSetting(SETTINGS_KEY_TARGET_HEIGHT_CM, DEFAULT_TARGET_HEIGHT_CM));

  const updateLockAspectRatioUi = () => {
    const width = Number(widthEl.value);
    const applyWidth = Number.isFinite(width) && width > 0;
    const height = Number(heightEl.value);
    const applyHeight =
      setHeightEnabledEl.checked && Number.isFinite(height) && height > 0;
    lockAspectRatioEl.checked = (applyWidth && !applyHeight) || (!applyWidth && applyHeight);
  };

  updateLockAspectRatioUi();

  enabledEl.addEventListener("change", async () => {
    try {
      await saveSetting(SETTINGS_KEY_ENABLED, enabledEl.checked);
      setStatus(`Saved. Enabled: ${enabledEl.checked}`);
    } catch (e) {
      setStatus(`Failed to save setting: ${String(e)}`);
    }

    if (Office.context.host === Office.HostType.Word) {
      const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
      if (enabledEl.checked && resizeOnPaste && !resizeOnSelection) {
        try {
          await initWordPollingBaseline();
        } catch {}
        startWordCountWatcher();
      } else {
        stopWordPasteBurst();
        stopWordCountWatcher();
      }
    }
  });

  const applyExclusiveMode = async (mode: "paste" | "selection") => {
    const nextPaste = mode === "paste";
    const nextSelection = mode === "selection";

    resizeOnPasteEl.checked = nextPaste;
    resizeOnSelectionEl.checked = nextSelection;

    try {
      await saveSetting(SETTINGS_KEY_RESIZE_ON_PASTE, nextPaste);
      await saveSetting(SETTINGS_KEY_RESIZE_ON_SELECTION, nextSelection);
      setStatus(`Saved. Resize on paste: ${nextPaste}. Resize on selection: ${nextSelection}`);
    } catch (e) {
      setStatus(`Failed to save setting: ${String(e)}`);
    }

    if (Office.context.host === Office.HostType.Word) {
      const enabled = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
      if (nextPaste && enabled) {
        try {
          await initWordPollingBaseline();
        } catch {}
        startWordCountWatcher();
      } else {
        stopWordPasteBurst();
        stopWordCountWatcher();
      }
    }

    if (nextPaste) {
      wordPictureSelected = false;
      wordSuppressSelectionUntil = 0;
      wordSuppressPollingUntil = 0;
    }
    if (nextSelection) {
      stopWordPasteBurst();
      stopWordCountWatcher();
      wordWatcherFastUntil = 0;
      wordSuppressPollingUntil = 0;
    }
  };

  resizeOnPasteEl.addEventListener("change", async () => {
    if (!resizeOnPasteEl.checked) {
      resizeOnPasteEl.checked = true;
      return;
    }
    await applyExclusiveMode("paste");
  });

  resizeOnSelectionEl.addEventListener("change", async () => {
    if (!resizeOnSelectionEl.checked) {
      resizeOnSelectionEl.checked = true;
      return;
    }
    await applyExclusiveMode("selection");
  });

  const onWidthChanged = async () => {
    const next = Number(widthEl.value);
    if (!Number.isFinite(next) || next < 0) return;

    const normalized = !setHeightEnabledEl.checked && next <= 0 ? DEFAULT_TARGET_WIDTH_CM : next;
    if (normalized !== next) widthEl.value = String(normalized);

    try {
      await saveSetting(SETTINGS_KEY_TARGET_WIDTH_CM, normalized);
      setStatus(`Saved. Target width: ${normalized} cm`);
    } catch (e) {
      setStatus(`Failed to save setting: ${String(e)}`);
    }

    updateLockAspectRatioUi();
  };

  widthEl.addEventListener("input", () => {
    void onWidthChanged();
  });

  widthEl.addEventListener("change", () => {
    void onWidthChanged();
  });

  const onHeightEnabledChanged = async () => {
    try {
      await saveSetting(SETTINGS_KEY_SET_HEIGHT_ENABLED, setHeightEnabledEl.checked);
      setStatus(`Saved. Set height: ${setHeightEnabledEl.checked}`);
    } catch (e) {
      setStatus(`Failed to save setting: ${String(e)}`);
    }

    updateLockAspectRatioUi();
  };

  setHeightEnabledEl.addEventListener("change", () => {
    void onHeightEnabledChanged();
  });

  const onHeightChanged = async () => {
    const next = Number(heightEl.value);
    if (!Number.isFinite(next) || next <= 0) return;

    try {
      await saveSetting(SETTINGS_KEY_TARGET_HEIGHT_CM, next);
      setStatus(`Saved. Target height: ${next} cm`);
    } catch (e) {
      setStatus(`Failed to save setting: ${String(e)}`);
    }

    updateLockAspectRatioUi();
  };

  heightEl.addEventListener("input", () => {
    void onHeightChanged();
  });

  heightEl.addEventListener("change", () => {
    void onHeightChanged();
  });

  lockAspectRatioEl.disabled = true;
}

Office.onReady(async () => {
  if (!hasOfficeContext()) {
    setStatus("This page must be opened inside Word/Excel/PowerPoint task pane, not a regular browser.");
    return;
  }

  bindUi();

  if (Office.context.host === Office.HostType.Word) {
    try {
      await initWordPollingBaseline();
    } catch {
      // ignore
    }

    const enabled = getBoolSetting(SETTINGS_KEY_ENABLED, DEFAULT_ENABLED);
    const { resizeOnPaste, resizeOnSelection } = getFeatureSettings();
    if (enabled && resizeOnPaste && !resizeOnSelection) {
      startWordCountWatcher();
    } else {
      stopWordPasteBurst();
      stopWordCountWatcher();
    }
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => {
      void onSelectionChanged();
    },
    (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("Ready. SelectionChanged handler attached. Paste an image and it will be resized if selected.");
      } else {
        setStatus(`Failed to attach handler: ${result.error?.message ?? String(result.error)}`);
      }
    }
  );
});
