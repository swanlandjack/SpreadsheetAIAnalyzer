/**
 * preload.js — Electron Preload Script
 *
 * This runs in the renderer process (browser window) before any page scripts.
 * It is the ONLY place where Node.js APIs can be selectively exposed to
 * the frontend HTML — keeping the app secure via contextIsolation.
 *
 * For Stage 1, the frontend (excel-ai-analyzer.html) communicates only with
 * the Flask backend on localhost:3000 via standard fetch() calls, so no
 * Node.js APIs need to be bridged yet.
 *
 * This file is kept as a clean foundation for future stages (e.g. exposing
 * file-system access for drag-and-drop, or IPC calls for model management).
 */

const { contextBridge } = require("electron");

// Expose a minimal API surface to the renderer
contextBridge.exposeInMainWorld("electronAPI", {
  // App version — useful to display in the UI
  version: process.env.npm_package_version || "1.0.0",

  // Platform — lets the frontend adapt UI if needed
  platform: process.platform,   // 'darwin' | 'win32' | 'linux'
});
