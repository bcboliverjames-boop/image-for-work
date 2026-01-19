import { getHttpsServerOptions } from "office-addin-dev-certs";
import { createServer } from "vite";

import { execSync, spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

function getPidsByPort(port) {
  if (process.platform !== "win32") {
    return [];
  }

  try {
    const output = execSync(`netstat -ano -p tcp | findstr :${port}`, {
      stdio: ["ignore", "pipe", "ignore"],
      encoding: "utf8"
    });

    const pids = new Set();
    for (const line of output.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed) continue;

      const parts = trimmed.split(/\s+/);
      if (parts.length < 5) continue;

      const localAddress = parts[1];
      const state = parts[3];
      if (!localAddress.endsWith(`:${port}`)) continue;
      if (state !== "LISTENING") continue;

      const pidStr = parts[parts.length - 1];
      const pid = Number(pidStr);
      if (Number.isFinite(pid) && pid > 0) {
        pids.add(pid);
      }
    }
    return [...pids];
  } catch {
    return [];
  }
}

function killPid(pid) {
  if (process.platform !== "win32") {
    return;
  }

  try {
    execSync(`taskkill /PID ${pid} /T /F`, {
      stdio: ["ignore", "ignore", "ignore"],
      encoding: "utf8"
    });
  } catch {
    // ignore
  }
}

function ensurePortFree(port) {
  const pids = getPidsByPort(port);
  for (const pid of pids) {
    if (pid === process.pid) continue;
    if (pid === 4) continue;
    killPid(pid);
  }
}

const https = await getHttpsServerOptions();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const configFile = path.resolve(root, "vite.config.ts");

ensurePortFree(3001);
ensurePortFree(3002);

const logServerProcess = spawn(process.execPath, [path.resolve(__dirname, "log-server.cjs")], {
  cwd: root,
  stdio: "inherit"
});

const server = await createServer({
  root,
  configFile,
  server: {
    host: "::",
    port: 3002,
    https
  }
});

await server.listen();
server.printUrls();

const shutdown = async () => {
  try {
    await server.close();
  } catch {
    // ignore
  }

  try {
    if (logServerProcess && !logServerProcess.killed) {
      logServerProcess.kill("SIGTERM");
    }
  } catch {
    // ignore
  }
};

process.on("SIGINT", async () => {
  await shutdown();
  process.exit(0);
});

process.on("SIGTERM", async () => {
  await shutdown();
  process.exit(0);
});

process.on("exit", () => {
  try {
    if (logServerProcess && !logServerProcess.killed) {
      logServerProcess.kill();
    }
  } catch {
    // ignore
  }
});
