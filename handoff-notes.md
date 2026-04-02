# Excel AI Analyzer — Project Handoff Notes
## Session Date: March 9, 2026

---

## What This App Does

A commercial Mac desktop app that lets users analyze Excel/CSV files using local AI.
- 100% offline — data never leaves the machine
- No Python, no terminal required for end users
- Powered by Ollama + Flask + Electron

---

## Project Location

Working directory on Jack's machine:
```
/Users/jacklau/Documents/Programming/2026Programming/TestFolderCVS/TestOllamaCVS_Electron/Stage4/excel-ai-electron/
```

---

## Tech Stack

| Layer | Technology |
|---|---|
| Desktop shell | Electron 31.7.7 |
| Backend | Flask (Python), compiled via PyInstaller |
| AI engine | Ollama (local LLM) |
| Frontend | Single HTML file |
| Build tool | electron-builder 24.13.3 |
| Deployment | Vercel (planned for landing page) |

---

## Project Structure

```
excel-ai-electron/
├── main.js                    ← Electron main process (Stage 4 complete)
├── preload.js                 ← Security bridge
├── splash.html                ← Loading screen
├── first-run.html             ← Setup wizard (Ollama + model download)
├── package.json
├── generate-icns.sh           ← Converts iconset → icon.icns (run on Mac)
├── build-backend-mac.sh       ← Compiles Python → binary
├── build-backend-win.bat      ← Windows equivalent
├── download-ollama.sh         ← Downloads Ollama binary for bundling
├── backend/
│   ├── server.py              ← Flask backend
│   ├── server.spec            ← PyInstaller spec
│   └── dist/server/server     ← Compiled binary (generated)
├── frontend/
│   └── excel-ai-analyzer.html ← Main app UI
└── assets/
    ├── icon.icns              ← Mac icon (generate via generate-icns.sh)
    ├── icon.ico               ← Windows icon
    ├── icon-1024.png          ← Master PNG
    └── icon.iconset/          ← All icon sizes for Mac
```

---

## Backend API Endpoints (server.py)

| Method | Route | Purpose |
|---|---|---|
| POST | /upload | Upload Excel/CSV file |
| POST | /chat | Streaming AI chat response |
| GET | /models | List installed Ollama models |
| GET | /health | Health check (returns ollama: true/false) |
| GET | /preview/<id> | Preview uploaded file |
| GET | /text/<id> | Get file as text |
| DELETE | /files | Delete uploaded files |
| POST | /pull | Stream model download from Ollama |
| POST | /delete-model | Remove an Ollama model |

---

## App Lifecycle (main.js)

```
Launch
  → splash.html shown immediately
  → Flask backend started (binary in prod, Python in dev)
  → Ollama check: port 11434 running?
      YES → use existing (never kill it on quit)
      NO  → launch bundled ollama binary (if present)
  → Poll /health until Flask ready (60s timeout)
  → GET /models
      Has models → open main app
      No models  → open first-run.html wizard
```

---

## First-Run Wizard (first-run.html)

Two-screen flow:

**Screen 1 — Ollama Setup** (shown if Ollama not on port 11434):
- Explains what Ollama is
- "Download Ollama — free → ollama.com" button (opens browser)
- "Check Again" button (re-polls /health for ollama status)
- Auto-skipped if Ollama already running

**Screen 2 — Model Selection**:
- Three model cards: qwen3:1.7b (recommended), qwen3:8b, llama3.2:3b
- Streams download progress from /pull endpoint
- On success → launches main app via `window.location.href = 'launch://app'`

---

## Key Behaviours

- Ollama started by USER (not app) → app detects it, uses it, NEVER kills it on quit
- Ollama started by APP (bundled binary) → app kills it cleanly on quit
- First-run wizard only appears when models folder is empty
- `+ Model` button in main app lets users download additional models anytime

---

## Build Commands

```bash
# Development
npm install
npm start

# Generate Mac icon (one-time, requires macOS)
chmod +x generate-icns.sh && ./generate-icns.sh

# Compile Python backend (run after any server.py changes)
chmod +x build-backend-mac.sh && ./build-backend-mac.sh

# Bundle Ollama (one-time, for zero-knowledge users)
chmod +x download-ollama.sh && ./download-ollama.sh

# Build production DMG
npx electron-builder --mac

# Output
dist/Excel AI Analyzer-1.0.0-arm64.dmg   ← Apple Silicon
dist/Excel AI Analyzer-1.0.0.dmg          ← Intel Mac
```

---

## Current Status — What's Done

| Stage | Feature | Status |
|---|---|---|
| 1 | Electron wraps app, auto-starts Flask | ✅ Done |
| 2 | PyInstaller binary, DMG installer | ✅ Done |
| 3 | Ollama bundling, first-run wizard, + Model button | ✅ Done |
| 4 | App icon, Mac menu, About dialog | ✅ Done |
| 4+ | Ollama detection screen + guided download | ✅ Done |

---

## What's Next (Not Done Yet)

### High Priority
1. **GitHub repo** — create `excel-ai-analyzer` repo, push code, create Release with DMG attached
2. **Landing page** — single HTML page on Vercel with download button linking to GitHub Release
3. **Bundle Ollama** — run `./download-ollama.sh` before final public build so zero-knowledge users are fully supported

### Medium Priority
4. **Code signing** — Apple Developer ID ($99/yr) removes Gatekeeper warning for public users
5. **Windows build** — run `build-backend-win.bat` on a Windows machine, build with `npx electron-builder --win`

### Lower Priority
6. **Auto-updater** — push version updates without users reinstalling
7. **Analytics** — basic download/usage counting

---

## Known Notes

- Without `./download-ollama.sh` + rebuild, users must have Ollama pre-installed. The wizard guides them to ollama.com but doesn't install it silently.
- Unsigned DMG: users must right-click → Open on first launch to bypass Gatekeeper. Fine for pilots, not for public launch.
- `package.json` `name` field must stay lowercase (`excel-ai-analyzer`) — electron-builder rejects spaces. Display name is controlled by `productName: "Excel AI Analyzer"`.
- Ollama latest version as of build: v0.17.7. `download-ollama.sh` auto-fetches latest via GitHub API.

---

## Testing Checklist

```
[ ] npm start → splash → main app (Ollama running)
[ ] Ollama dot green, models in dropdown
[ ] + Model button opens modal, progress bar works
[ ] First-run Screen 1: kill Ollama + hide ~/.ollama/models
[ ] First-run Screen 2: Ollama running, no models
[ ] App quits cleanly (no orphan processes)
[ ] DMG installs, launches, shows correct app name in menu bar
```

---

## Target Market

This is a commercial product. Target customers are finance and legal professionals in Hong Kong and Singapore. Key selling point: complete data privacy — AI runs locally, nothing sent to cloud. No subscription, no API key, no cloud dependency.
