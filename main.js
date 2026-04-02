/**
 * Excel AI Analyzer — Electron Main Process (Stage 3)
 *
 * Responsibilities:
 *   1. Show splash screen immediately
 *   2. Start Flask backend (binary in production, Python in dev)
 *   3. Check if Ollama is already running — if not, launch bundled binary
 *   4. Poll /health until Flask is ready
 *   5. Check if any models are installed
 *   6. If no models: show first-run wizard; else open main app directly
 *   7. Gracefully kill Flask + Ollama on exit
 */

const { app, BrowserWindow, dialog, shell, Menu } = require("electron");
const { spawn } = require("child_process");
const path = require("path");
const http = require("http");
const fs = require("fs");

// ── Configuration ──────────────────────────────────────────────────────────────
const BACKEND_PORT    = 3000;
const OLLAMA_PORT     = 11434;
const HEALTH_URL      = `http://localhost:${BACKEND_PORT}/health`;
const MODELS_URL      = `http://localhost:${BACKEND_PORT}/models`;
const HEALTH_POLL_INTERVAL_MS = 500;
const HEALTH_TIMEOUT_MS       = 60000;

// ── State ──────────────────────────────────────────────────────────────────────
let splashWindow   = null;
let mainWindow     = null;
let firstRunWindow = null;
let backendProcess = null;
let ollamaProcess  = null;   // only set if WE launched Ollama (not pre-existing)

// ── Path helpers ───────────────────────────────────────────────────────────────

/**
 * Resolve a path to a bundled resource.
 * In production (packaged app), resources live in process.resourcesPath.
 * In development, they are relative to __dirname.
 */
function resourcePath(...segments) {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, ...segments);
  }
  return path.join(__dirname, ...segments);
}

// ── Ollama management ──────────────────────────────────────────────────────────

/**
 * Check if Ollama is already listening on port 11434.
 * Returns true if something is already running there.
 */
function isOllamaRunning() {
  return new Promise((resolve) => {
    http.get(`http://localhost:${OLLAMA_PORT}/api/tags`, (res) => {
      resolve(res.statusCode < 500);
    }).on("error", () => resolve(false));
  });
}

/**
 * Get the bundled Ollama binary path.
 * The binary is placed in ollama/ folder, included via extraResources.
 */
function getOllamaBinaryPath() {
  const binaryName = process.platform === "win32" ? "ollama.exe" : "ollama";
  return resourcePath("ollama", binaryName);
}

/**
 * Start the bundled Ollama binary as a background process.
 * Resolves when Ollama is confirmed listening on port 11434.
 */
async function startOllama() {
  const ollamaPath = getOllamaBinaryPath();

  if (!fs.existsSync(ollamaPath)) {
    console.log("[Ollama] Bundled binary not found — skipping (user must have Ollama separately)");
    return;
  }

  console.log(`[Ollama] Starting bundled binary: ${ollamaPath}`);

  ollamaProcess = spawn(ollamaPath, ["serve"], {
    env: {
      ...process.env,
      // Store models in user's home dir so they persist between app updates
      OLLAMA_MODELS: path.join(
        process.env.HOME || process.env.USERPROFILE || "",
        ".ollama", "models"
      ),
    },
  });

  ollamaProcess.stdout.on("data", (d) => process.stdout.write(`[Ollama] ${d}`));
  ollamaProcess.stderr.on("data", (d) => process.stderr.write(`[Ollama:err] ${d}`));
  ollamaProcess.on("exit", (code) => {
    console.log(`[Ollama] Process exited — code=${code}`);
    ollamaProcess = null;
  });

  // Wait for Ollama to be ready (up to 15 seconds)
  const start = Date.now();
  while (Date.now() - start < 15000) {
    if (await isOllamaRunning()) {
      console.log("[Ollama] Ready ✓");
      return;
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  console.log("[Ollama] Did not respond within 15s — continuing anyway");
}

/**
 * Ensure Ollama is running — uses existing if already up, else starts bundled.
 */
async function ensureOllama() {
  const alreadyRunning = await isOllamaRunning();
  if (alreadyRunning) {
    console.log("[Ollama] Already running on port 11434 — using existing instance");
    return;
  }
  await startOllama();
}

/**
 * Check if any models are installed via the Flask /models endpoint.
 */
async function hasModels() {
  return new Promise((resolve) => {
    http.get(MODELS_URL, (res) => {
      let body = "";
      res.on("data", (d) => (body += d));
      res.on("end", () => {
        try {
          const json = JSON.parse(body);
          resolve(Array.isArray(json.models) && json.models.length > 0);
        } catch { resolve(false); }
      });
    }).on("error", () => resolve(false));
  });
}

// ── Python detection ───────────────────────────────────────────────────────────

/**
 * Find a working Python 3 executable on the user's machine.
 * Tries common locations in order and returns the first one that works.
 */
async function findPython() {
  const { execFile } = require("child_process");
  const { promisify } = require("util");
  const execFileAsync = promisify(execFile);

  // Candidates — order matters, most specific first
  const isMac = process.platform === "darwin";
  const isWin = process.platform === "win32";

  const candidates = [];

  if (isWin) {
    candidates.push(
      "py",        // Windows Python launcher (winget/official install)
      "python",
      "python3",
      path.join(process.env.LOCALAPPDATA || "", "Programs", "Python", "Python311", "python.exe"),
      path.join(process.env.LOCALAPPDATA || "", "Programs", "Python", "Python310", "python.exe"),
      path.join(process.env.USERPROFILE || "", "miniconda3", "envs", "excel-ai", "python.exe"),
      path.join(process.env.USERPROFILE || "", "anaconda3", "envs", "excel-ai", "python.exe"),
      path.join(process.env.USERPROFILE || "", "miniconda3", "python.exe"),
      path.join(process.env.USERPROFILE || "", "anaconda3", "python.exe"),
    );
  } else {
    // Mac / Linux
    candidates.push(
      "python3",
      "python",
      "/usr/bin/python3",
      "/usr/local/bin/python3",
      "/opt/homebrew/bin/python3",                          // Homebrew on Apple Silicon
      path.join(process.env.HOME || "", "miniconda3", "envs", "excel-ai", "bin", "python"),
      path.join(process.env.HOME || "", "anaconda3", "envs", "excel-ai", "bin", "python"),
      path.join(process.env.HOME || "", "miniconda3", "bin", "python3"),
      path.join(process.env.HOME || "", "anaconda3", "bin", "python3"),
    );
  }

  for (const candidate of candidates) {
    try {
      const { stdout } = await execFileAsync(candidate, ["--version"]);
      const version = (stdout || "").trim();
      if (version.startsWith("Python 3")) {
        console.log(`[Python] Found: ${candidate} → ${version}`);
        return candidate;
      }
    } catch {
      // Not found or wrong version — keep trying
    }
  }

  return null; // Not found
}

// ── Health polling ─────────────────────────────────────────────────────────────

/**
 * Poll the Flask /health endpoint until it responds or we time out.
 * Returns a promise that resolves when healthy, rejects on timeout.
 */
function waitForServer() {
  return new Promise((resolve, reject) => {
    let settled = false;
    let stderrOutput = "";

    // Collect stderr so we can report it if the process crashes
    if (backendProcess && backendProcess.stderr) {
      backendProcess.stderr.on("data", (d) => { stderrOutput += d.toString(); });
    }

    // If the backend process exits before Flask is ready, fail immediately
    if (backendProcess) {
      backendProcess.on("exit", (code) => {
        if (!settled) {
          settled = true;
          const reason = stderrOutput.trim() || `Backend process exited with code ${code}`;
          reject(new Error(`AI backend crashed on startup: ${reason}`));
        }
      });
    }

    // Poll /health indefinitely — no timeout
    const check = () => {
      if (settled) return;
      http.get(HEALTH_URL, (res) => {
        let body = "";
        res.on("data", (d) => (body += d));
        res.on("end", () => {
          if (settled) return;
          try {
            const json = JSON.parse(body);
            if (json.status === "ok") {
              settled = true;
              console.log("[Health] Flask server is ready ✓");
              resolve();
              return;
            }
          } catch {}
          setTimeout(check, HEALTH_POLL_INTERVAL_MS);
        });
      }).on("error", () => {
        if (!settled) setTimeout(check, HEALTH_POLL_INTERVAL_MS);
      });
    };

    check();
  });
}

// ── Splash window ──────────────────────────────────────────────────────────────

function createSplashWindow() {
  splashWindow = new BrowserWindow({
    width: 480,
    height: 300,
    frame: false,
    resizable: false,
    center: true,
    alwaysOnTop: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  splashWindow.loadFile(path.join(__dirname, "splash.html"));
  splashWindow.on("closed", () => { splashWindow = null; });
}

// ── First-run wizard window ────────────────────────────────────────────────────

function createFirstRunWindow() {
  firstRunWindow = new BrowserWindow({
    width: 620,
    height: 580,
    resizable: false,
    center: true,
    title: "Excel AI Analyzer — Setup",
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  firstRunWindow.loadFile(path.join(__dirname, "first-run.html"));

  // Handle the launch://app protocol triggered by the wizard's buttons
  firstRunWindow.webContents.on("will-navigate", (event, url) => {
    if (url.startsWith("launch://app")) {
      event.preventDefault();
      firstRunWindow.close();
      createMainWindow();
    }
  });

  firstRunWindow.on("closed", () => { firstRunWindow = null; });
}

// ── Main app window ────────────────────────────────────────────────────────────

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    show: false,
    title: "Excel AI Analyzer",
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  mainWindow.loadFile(resourcePath("frontend", "excel-ai-analyzer.html"));

  mainWindow.once("ready-to-show", () => {
    mainWindow.show();
    if (splashWindow && !splashWindow.isDestroyed()) {
      splashWindow.close();
    }
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });

  mainWindow.on("closed", () => { mainWindow = null; });
}

// ── Backend launcher ───────────────────────────────────────────────────────────

/**
 * Returns the path to the PyInstaller-compiled binary.
 * Layout inside the packaged app (set by electron-builder extraResources):
 *   Mac:     <app>/Contents/Resources/backend/dist/server/server
 *   Windows: <app>/resources/backend/dist/server/server.exe
 */
function getCompiledBinaryPath() {
  const binaryName = process.platform === "win32" ? "server.exe" : "server";
  return resourcePath("backend", "dist", "server", binaryName);
}

/**
 * Unified backend starter.
 *
 * PRODUCTION (app.isPackaged = true):
 *   Runs the PyInstaller binary directly — no Python needed on user's machine.
 *
 * DEVELOPMENT (npm start):
 *   Falls back to finding Python and running server.py — convenient for coding.
 */
async function startBackend() {
  let execPath;
  let args = [];
  let label;

  if (app.isPackaged) {
    // ── PRODUCTION: use compiled binary ──────────────────────────────────────
    execPath = getCompiledBinaryPath();
    label = "Binary";

    if (!fs.existsSync(execPath)) {
      dialog.showErrorBox(
        "Backend Binary Not Found",
        `Could not find the compiled backend at:\n${execPath}\n\n` +
        "Please rebuild the app with:\n  npm run build:backend\n  npm run build:mac"
      );
      app.quit();
      return;
    }

    console.log(`[Backend] Production mode — using compiled binary`);
    console.log(`[Backend] Path: ${execPath}`);

  } else {
    // ── DEVELOPMENT: use Python + server.py ──────────────────────────────────
    label = "Python";
    const pythonExec = await findPython();

    if (!pythonExec) {
      dialog.showErrorBox(
        "Python Not Found",
        "Excel AI Analyzer requires Python 3 to run the AI backend.\n\n" +
        "Please install Python 3 from https://www.python.org and try again.\n\n" +
        "If you used Conda, make sure the 'excel-ai' environment is activated " +
        "before launching this app."
      );
      app.quit();
      return;
    }

    execPath = pythonExec;
    args = [resourcePath("backend", "server.py")];
    console.log(`[Backend] Dev mode — using Python: ${pythonExec}`);
  }

  backendProcess = spawn(execPath, args, {
    cwd: resourcePath("backend"),
    env: {
      ...process.env,
      PYTHONUNBUFFERED: "1",
      PYTHONUTF8: "1",        // Force UTF-8 on Windows (fixes charmap encoding errors)
      FLASK_ENV: "production",
    },
  });

  backendProcess.stdout.on("data", (d) => process.stdout.write(`[${label}] ${d}`));
  backendProcess.stderr.on("data", (d) => process.stderr.write(`[${label}:err] ${d}`));

  backendProcess.on("exit", (code, signal) => {
    console.log(`[Backend] Process exited — code=${code} signal=${signal}`);
    backendProcess = null;

    if (mainWindow && !mainWindow.isDestroyed()) {
      dialog.showErrorBox(
        "Backend Stopped",
        "The AI backend stopped unexpectedly.\n\nPlease restart the app."
      );
    }
  });
}

// ── Update splash status ───────────────────────────────────────────────────────

function setSplashStatus(message) {
  if (splashWindow && !splashWindow.isDestroyed()) {
    splashWindow.webContents.executeJavaScript(
      `document.getElementById('status').textContent = ${JSON.stringify(message)}`
    ).catch(() => {});
  }
}

// ── App metadata ───────────────────────────────────────────────────────────────
const APP_NAME    = "Excel AI Analyzer";
const APP_VERSION = app.getVersion();
const APP_YEAR    = new Date().getFullYear();

// ── Mac application menu ───────────────────────────────────────────────────────

function buildAppMenu() {
  const isMac = process.platform === "darwin";

  const template = [
    // ── App menu (Mac only) ──────────────────────────────────
    ...(isMac ? [{
      label: APP_NAME,
      submenu: [
        {
          label: `About ${APP_NAME}`,
          click: showAboutDialog,
        },
        { type: "separator" },
        {
          label: "Check for Model Updates",
          click: () => {
            if (mainWindow) {
              // Open the + Model modal in the frontend
              mainWindow.webContents.executeJavaScript(
                `document.getElementById('addModelBtn')?.click()`
              );
            }
          },
        },
        { type: "separator" },
        { role: "services" },
        { type: "separator" },
        { role: "hide" },
        { role: "hideOthers" },
        { role: "unhide" },
        { type: "separator" },
        { role: "quit", label: `Quit ${APP_NAME}` },
      ],
    }] : []),

    // ── File ─────────────────────────────────────────────────
    {
      label: "File",
      submenu: [
        {
          label: "New Chat",
          accelerator: "CmdOrCtrl+N",
          click: () => {
            if (mainWindow) {
              mainWindow.webContents.executeJavaScript(
                `document.getElementById('newChatBtn')?.click()`
              );
            }
          },
        },
        { type: "separator" },
        isMac ? { role: "close" } : { role: "quit" },
      ],
    },

    // ── Edit ─────────────────────────────────────────────────
    {
      label: "Edit",
      submenu: [
        { role: "undo" },
        { role: "redo" },
        { type: "separator" },
        { role: "cut" },
        { role: "copy" },
        { role: "paste" },
        { role: "selectAll" },
      ],
    },

    // ── View ─────────────────────────────────────────────────
    {
      label: "View",
      submenu: [
        { role: "reload" },
        { type: "separator" },
        { role: "resetZoom" },
        { role: "zoomIn" },
        { role: "zoomOut" },
        { type: "separator" },
        { role: "togglefullscreen" },
        ...(process.env.NODE_ENV === "development"
          ? [{ type: "separator" }, { role: "toggleDevTools" }]
          : []),
      ],
    },

    // ── Window ───────────────────────────────────────────────
    {
      label: "Window",
      submenu: [
        { role: "minimize" },
        { role: "zoom" },
        ...(isMac
          ? [{ type: "separator" }, { role: "front" }]
          : [{ role: "close" }]),
      ],
    },

    // ── Help ─────────────────────────────────────────────────
    {
      label: "Help",
      submenu: [
        {
          label: "Ollama Model Library",
          click: () => shell.openExternal("https://ollama.com/library"),
        },
        { type: "separator" },
        {
          label: `About ${APP_NAME}`,
          click: showAboutDialog,
        },
      ],
    },
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

// ── About dialog ───────────────────────────────────────────────────────────────

function showAboutDialog() {
  dialog.showMessageBox({
    type: "none",
    title: `About ${APP_NAME}`,
    message: APP_NAME,
    detail: [
      `Version ${APP_VERSION}`,
      "",
      "AI-powered spreadsheet analysis that runs 100% locally.",
      "Your data never leaves your machine.",
      "",
      "Powered by Ollama · Flask · Electron",
      "",
      `© ${APP_YEAR} ${APP_NAME}`,
    ].join("\n"),
    buttons: ["OK"],
    defaultId: 0,
  });
}

// ── App lifecycle ──────────────────────────────────────────────────────────────

app.whenReady().then(async () => {
  // Build the application menu immediately
  buildAppMenu();

  // 1. Show splash immediately
  createSplashWindow();
  setSplashStatus("Starting up…");
  await new Promise((r) => setTimeout(r, 200));

  // 2. Start Flask backend (binary in production, Python in dev)
  setSplashStatus("Starting AI backend…");
  await startBackend();

  // 3. Ensure Ollama is running (use existing or launch bundled)
  setSplashStatus("Starting Ollama…");
  await ensureOllama();

  // 4. Wait for Flask to be healthy
  setSplashStatus("Waiting for backend to be ready…");
  try {
    await waitForServer();
  } catch (err) {
    const hint = app.isPackaged
      ? "The AI backend binary failed to start."
      : "Please check that all Python packages are installed:\n  pip install flask flask-cors pandas openpyxl xlrd requests numpy";
    dialog.showErrorBox("Backend Failed to Start",
      `The AI backend did not start correctly.\n\n${err.message}\n\n${hint}`);
    app.quit();
    return;
  }

  // 5. Check if any models are installed
  setSplashStatus("Checking AI models…");
  const modelsInstalled = await hasModels();

  // 6. Show first-run wizard OR go straight to main app
  setSplashStatus("Opening app…");
  if (!modelsInstalled) {
    console.log("[Setup] No models found — showing first-run wizard");
    createFirstRunWindow();
    // Close splash once first-run window is ready
    if (firstRunWindow) {
      firstRunWindow.once("ready-to-show", () => {
        if (splashWindow && !splashWindow.isDestroyed()) splashWindow.close();
        firstRunWindow.show();
      });
      firstRunWindow.webContents.once("did-finish-load", () => {
        if (splashWindow && !splashWindow.isDestroyed()) splashWindow.close();
      });
    }
  } else {
    console.log("[Setup] Models found — opening main app");
    createMainWindow();
  }
});

// ── Quit behaviour ─────────────────────────────────────────────────────────────

// On macOS, keep app running when all windows are closed (standard Mac behaviour)
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  // Re-open main window when dock icon is clicked on Mac and no windows are open
  if (BrowserWindow.getAllWindows().length === 0) {
    createMainWindow();
  }
});

// Kill backend AND Ollama when Electron exits
app.on("before-quit", () => {
  if (backendProcess) {
    console.log("[Cleanup] Killing backend process…");
    backendProcess.kill("SIGTERM");
    backendProcess = null;
  }
  if (ollamaProcess) {
    console.log("[Cleanup] Killing Ollama process…");
    ollamaProcess.kill("SIGTERM");
    ollamaProcess = null;
  }
});

process.on("exit", () => {
  if (backendProcess) backendProcess.kill("SIGTERM");
  if (ollamaProcess)  ollamaProcess.kill("SIGTERM");
});
