/**
 * Logger module - uses the log system defined in taskpane.html
 */

// Re-export from window._pasteLog (defined in taskpane.html)
export const pasteLog = {
  debug: (msg: string, data?: any) => {
    const log = (window as any)._pasteLog;
    if (log) log.debug(msg, data);
    else console.debug("[pasteMode]", msg, data ?? "");
  },
  info: (msg: string, data?: any) => {
    const log = (window as any)._pasteLog;
    if (log) log.info(msg, data);
    else console.info("[pasteMode]", msg, data ?? "");
  },
  warn: (msg: string, data?: any) => {
    const log = (window as any)._pasteLog;
    if (log) log.warn(msg, data);
    else console.warn("[pasteMode]", msg, data ?? "");
  },
  error: (msg: string, data?: any) => {
    const log = (window as any)._pasteLog;
    if (log) log.error(msg, data);
    else console.error("[pasteMode]", msg, data ?? "");
  },
};

export const selectionLog = {
  debug: (msg: string, data?: any) => {
    const logs = (window as any).pasteWidthLogs;
    if (logs?.log) logs.log("debug", "selectionMode", msg, data);
    else console.debug("[selectionMode]", msg, data ?? "");
  },
  info: (msg: string, data?: any) => {
    const logs = (window as any).pasteWidthLogs;
    if (logs?.log) logs.log("info", "selectionMode", msg, data);
    else console.info("[selectionMode]", msg, data ?? "");
  },
  warn: (msg: string, data?: any) => {
    const logs = (window as any).pasteWidthLogs;
    if (logs?.log) logs.log("warn", "selectionMode", msg, data);
    else console.warn("[selectionMode]", msg, data ?? "");
  },
  error: (msg: string, data?: any) => {
    const logs = (window as any).pasteWidthLogs;
    if (logs?.log) logs.log("error", "selectionMode", msg, data);
    else console.error("[selectionMode]", msg, data ?? "");
  },
};

// Re-export functions from window.pasteWidthLogs
export function getLogs() {
  return (window as any).pasteWidthLogs?.getLogs?.() ?? [];
}

export function getLogsAsText() {
  return (window as any).pasteWidthLogs?.getLogsAsText?.() ?? "";
}

export function clearLogs() {
  (window as any).pasteWidthLogs?.clearLogs?.();
}

export function downloadLogs() {
  (window as any).pasteWidthLogs?.downloadLogs?.();
}

export async function sendLogsToServer() {
  return (window as any).pasteWidthLogs?.sendLogsToServer?.() ?? false;
}

export function setLoggingEnabled(enabled: boolean) {
  // No-op, logging is always enabled in taskpane.html
}

console.info("[logger.ts] Module loaded - using taskpane.html log system");
