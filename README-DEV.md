# Excel AI Analyzer — Electron App (Dev Guide)

## Project Structure

```
excel-ai-electron/
├── main.js                        ← Electron main process
├── preload.js                     ← Security bridge
├── splash.html                    ← Loading screen
├── package.json                   ← Node/Electron config + build scripts
├── build-backend-mac.sh           ← Compiles Python → binary (Mac)
├── build-backend-win.bat          ← Compiles Python → binary (Windows)
├── backend/
│   ├── server.py                  ← Flask backend (source, unchanged)
│   ├── server.spec                ← PyInstaller config
│   └── dist/server/               ← PyInstaller output (created by build script)
│       ├── server                 ← Mac binary (no Python needed)
│       └── server.exe             ← Windows binary (no Python needed)
├── frontend/
│   └── excel-ai-analyzer.html    ← HTML frontend (unchanged)
└── assets/
    ├── icon.icns                  ← Mac icon (add before building installer)
    └── icon.ico                   ← Windows icon (add before building installer)
```

---

## Stage 1 — Development (Python required on your machine)

For daily coding and testing. Fast iteration, no compile step needed.

```bash
cd excel-ai-electron
npm install          # first time only
npm start
```

Electron auto-detects Python, starts server.py, opens the app.

---

## Stage 2 — Build the Compiled Backend (No Python on user's machine)

Run this ONCE on each platform you are building for.
Produces a ~80-120MB self-contained binary with Python baked in.

### Mac

```bash
chmod +x build-backend-mac.sh
./build-backend-mac.sh
```

### Windows

Double-click `build-backend-win.bat`, or from Command Prompt:
```cmd
build-backend-win.bat
```

Both scripts automatically:
1. Detect Python on your machine
2. Install PyInstaller if missing
3. Install all required Python packages
4. Compile server.py → standalone binary
5. Output to `backend/dist/server/`

### Test the binary directly before packaging

```bash
# Mac — should show Flask startup output
./backend/dist/server/server

# Windows
backend\dist\server\server.exe
```

---

## Stage 2 — Build the Distributable Installer

After the binary is built and tested:

```bash
# Mac (.dmg for Apple Silicon + Intel)
npm run build:mac

# Windows (.exe installer)
npm run build:win
```

Output in `dist/`:
- `Excel AI Analyzer-1.0.0-arm64.dmg`   (Mac Apple Silicon)
- `Excel AI Analyzer-1.0.0.dmg`          (Mac Intel)
- `Excel AI Analyzer Setup 1.0.0.exe`    (Windows)

---

## How main.js Chooses Execution Path

```
app.isPackaged?
  YES (installed app) → run backend/dist/server/server binary (no Python needed)
  NO  (npm start)     → find Python, run backend/server.py
```

---

## Icons (Required Before Shipping)

**Mac** — `assets/icon.icns` (1024x1024 source PNG)
Use: https://cloudconvert.com/png-to-icns

**Windows** — `assets/icon.ico` (256x256 multi-resolution)
Use: https://cloudconvert.com/png-to-ico

---

## Troubleshooting

**PyInstaller fails with ModuleNotFoundError**
Add the module name to `hiddenimports` in `backend/server.spec`, rerun build.

**"Backend Binary Not Found" error in app**
The build script hasn't been run yet. Run `./build-backend-mac.sh` first.

**Binary is very large (>200MB)**
Add unused packages to the `excludes` list in `server.spec` and rebuild.

**Binary works in terminal but not in packaged app**
Path issue. Add console.log to main.js to print resolved paths and debug.

---

## Stage 3 Preview

Stage 3 bundles the Ollama binary inside the app — users won't need to
install Ollama separately. Includes a first-run model download wizard.
