/*
 * Settings persistence for the Word add-in.
 *
 * Deliberately NOT reading extension/clausesmith_core.py's config file
 * (~/.config/libreoffice-clausesmith/config.json) - a browser sandbox has
 * no filesystem access at all, so that's not a bridgeable gap the way the
 * RedactAI -> ClauseSmith rename's config migration was. Settings here
 * are independent per platform: configure Word separately from
 * LibreOffice. localStorage is the closest analog available - it's
 * per-machine, per-browser-profile, and survives closing Word, same
 * shape as the LibreOffice config file even though it isn't the same
 * file.
 */
import { DEFAULT_CONFIG } from "./ollama-client.js";

const STORAGE_KEY = "clausesmith.config";

export function loadConfig() {
  const cfg = { ...DEFAULT_CONFIG };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) Object.assign(cfg, JSON.parse(raw));
  } catch {
    // corrupt or inaccessible storage - fall back to defaults rather than
    // break the task pane, same policy as load_config()'s bare except.
  }
  return cfg;
}

export function saveConfig(cfg) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
}
