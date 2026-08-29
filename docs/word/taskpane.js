/* global Office, Word */

import { callOllama, listOllamaModels } from "./shared/ollama-client.js";
import { loadConfig, saveConfig } from "./shared/config.js";
import { REWRITE_ACTIONS, buildPrompt } from "./shared/prompts.js";

let cfg = loadConfig();
let busy = false;
// The range an in-flight (or just-completed) suggestion applies to. Word.js
// range objects don't survive past the Word.run() batch they were created
// in unless explicitly .track()ed - this is that pattern's equivalent of
// clausesmith_panel.py's self._target_anchor (a UNO text cursor, which is
// a live document object and doesn't need this dance).
let trackedAnchor = null;

const $ = (id) => document.getElementById(id);

// Declared ABOVE Office.onReady, not next to populateQuickActions where it
// reads more naturally. onReady's callback can run synchronously (it does
// when Office.js is already initialised), and a `const` declared further
// down the module is still in its temporal dead zone at that point - which
// throws, aborts the rest of onReady including wireButtons(), and leaves
// every control in the pane inert with only a console error to show for it.
//
// Button captions are decoupled from the REWRITE_ACTIONS keys on purpose.
// Those keys are shared verbatim with extension/clausesmith_core.py so both
// platforms send byte-identical prompts, but "Improve writing" doesn't fit
// a ~97px grid cell at the default pane width and rendered as
// "Improve writ…". Shortening the key would break prompt parity; shortening
// only the caption doesn't. The full name stays as the tooltip.
const ACTION_LABELS = {
  "Improve writing": "Improve",
};

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;

  populateQuickActions();
  wireButtons();
  registerSelectionTracking();
  checkConnection();
});

// -- connection status --------------------------------------------------
//
// Doubles as the live localhost-fetch check: if WebView2 ever blocks the
// https-taskpane -> http-localhost request as mixed content, this is where
// it surfaces, as a network error from ollama-client.js rather than
// anything Word-specific.
//
// Healthy state is deliberately one terse line (model + endpoint) instead
// of the old two-line "Connected to Ollama at ..." banner. In a pane that
// stays docked all day, standing information should cost one row, and
// which model you're about to hit is worth more than a success sentence
// you've already read. Failures still get the full message.
function setConnection(state, text) {
  $("statusDot").className = `dot dot-${state}`;
  $("statusText").textContent = text;
  $("statusText").title = text; // full text on hover when ellipsised
  $("statusStrip").classList.toggle("is-error", state === "error");
}

async function checkConnection() {
  setConnection("checking", "Checking Ollama…");
  const host = cfg.endpoint.replace(/^https?:\/\//, "");
  try {
    const models = await listOllamaModels(cfg.endpoint);
    if (models.length === 0) {
      setConnection("warn", `${host} — no models installed`);
    } else if (!models.includes(cfg.model)) {
      // Worth flagging up front: the configured model isn't one Ollama has,
      // so every action would fail with a 404 until Settings is corrected.
      setConnection("warn", `${cfg.model} not installed on ${host}`);
    } else {
      setConnection("ok", `${cfg.model} · ${host}`);
    }
  } catch (e) {
    setConnection("error", e.message);
  }
}

// -- quick actions -----------------------------------------------------
// (ACTION_LABELS lives at the top of the file - see the note there.)
function populateQuickActions() {
  const container = $("quickActions");
  for (const key of Object.keys(REWRITE_ACTIONS)) {
    const btn = document.createElement("button");
    btn.textContent = ACTION_LABELS[key] || key;
    btn.title = key;
    btn.addEventListener("click", () => onQuickAction(key));
    container.appendChild(btn);
  }
}

function wireButtons() {
  $("applyCustomBtn").addEventListener("click", onCustom);
  $("replaceBtn").addEventListener("click", onReplace);
  $("insertBelowBtn").addEventListener("click", onInsertBelow);
  $("copyBtn").addEventListener("click", onCopy);
  $("discardBtn").addEventListener("click", onDiscard);

  $("settingsToggle").addEventListener("click", openSettings);
  $("saveSettingsBtn").addEventListener("click", onSaveSettings);
  $("cancelSettingsBtn").addEventListener("click", closeSettings);
  $("fetchModelsBtn").addEventListener("click", onFetchModels);

  // Enter in the instruction box runs Custom - the field is right next to
  // its button, and reaching for the mouse to submit a one-line input is
  // friction in a pane you're using constantly.
  $("customInstruction").addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      onCustom();
    }
  });
}

// -- selection tracking (fills the Original pane automatically) --------
function registerSelectionTracking() {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );
  onSelectionChanged(); // populate with whatever's selected right now
}

async function onSelectionChanged() {
  try {
    const text = await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();
      return range.text;
    });
    $("originalPreview").value = text;
  } catch (e) {
    // Non-fatal - the Original pane just doesn't auto-update this time.
    // Matches _start_selection_tracking's best-effort policy in
    // clausesmith_panel.py.
    console.error("selection tracking failed", e);
  }
}

/**
 * Captures the current selection as a range that survives past this
 * Word.run batch (via .track()), for use by a later Replace/Insert
 * action once the model response comes back. Releases any previously
 * tracked range first, mirroring _capture_selection() being called
 * fresh on every action rather than reusing a stale anchor.
 */
async function captureAnchor() {
  await clearAnchor();
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    range.track();
    return range;
  });
}

/** Untrack and forget the current anchor, if any. Safe to call anytime. */
async function clearAnchor() {
  const anchor = trackedAnchor;
  trackedAnchor = null;
  if (!anchor) return;
  try {
    await Word.run(anchor, async (context) => {
      anchor.untrack();
      await context.sync();
    });
  } catch (e) {
    console.error("releasing tracked range failed", e);
  }
}

// -- actions that populate the Suggested pane ---------------------------
async function onQuickAction(actionKey) {
  await runAction(REWRITE_ACTIONS[actionKey]);
}

async function onCustom() {
  const instruction = $("customInstruction").value.trim();
  if (!instruction) {
    setStatus("Type a custom instruction first.", true);
    return;
  }
  await runAction(instruction);
}

async function runAction(instruction) {
  const selection = $("originalPreview").value;
  if (!selection.trim()) {
    setStatus("Select some text in the document first.", true);
    return;
  }
  if (busy) return;
  busy = true;
  setActionsEnabled(false);
  setStatus("Asking Ollama… (first run after idle is slow - model loads)");

  try {
    trackedAnchor = await captureAnchor();
    const prompt = buildPrompt(instruction, selection);
    const result = await callOllama(cfg, prompt);
    $("suggested").value = result;
    setStatus("Review, then Replace or Insert.");
  } catch (e) {
    // Release the anchor rather than leaving a tracked range attached to a
    // request that never produced anything to commit.
    await clearAnchor();
    // Errors go inline, not through alert(): a modal dialog in a pane
    // that's docked permanently beside the document is the wrong register,
    // and Ollama's messages (timeouts, missing models) are long enough to
    // be worth reading in place rather than dismissing.
    setStatus(e.message, true);
  } finally {
    busy = false;
    setActionsEnabled(true);
  }
}

// -- footer: commit or discard the current suggestion --------------------

/** Shared guard for the two commit actions. Returns the anchor or null. */
function takeAnchorForCommit() {
  if (!$("suggested").value.trim()) {
    setStatus("Nothing to commit yet - run an action first.", true);
    return null;
  }
  if (!trackedAnchor) {
    setStatus("Re-select the text and run an action again.", true);
    return null;
  }
  const anchor = trackedAnchor;
  trackedAnchor = null;
  return anchor;
}

async function onReplace() {
  const anchor = takeAnchorForCommit();
  if (!anchor) return;
  const text = $("suggested").value;
  try {
    await Word.run(anchor, async (context) => {
      anchor.insertText(text, Word.InsertLocation.replace);
      anchor.untrack();
      await context.sync();
    });
    setStatus("Replaced.");
  } catch (e) {
    setStatus(`Couldn't replace: ${e.message}`, true);
  }
}

async function onInsertBelow() {
  const anchor = takeAnchorForCommit();
  if (!anchor) return;
  const text = $("suggested").value;
  try {
    await Word.run(anchor, async (context) => {
      // insertParagraph, not insertText("\n" + ...): a literal newline in
      // insertText does not reliably produce a paragraph break in Word,
      // whereas insertParagraph is the documented way to add one after a
      // range - and "below" here means a new paragraph, not a line break.
      anchor.insertParagraph(text, Word.InsertLocation.after);
      anchor.untrack();
      await context.sync();
    });
    setStatus("Inserted below.");
  } catch (e) {
    setStatus(`Couldn't insert: ${e.message}`, true);
  }
}

async function onCopy() {
  const text = $("suggested").value;
  if (!text.trim()) return;
  try {
    await navigator.clipboard.writeText(text);
    setStatus("Copied to clipboard.");
  } catch (e) {
    setStatus(`Couldn't copy: ${e.message}`, true);
  }
}

async function onDiscard() {
  $("suggested").value = "";
  await clearAnchor();
  setStatus("");
}

function setStatus(text, isError = false) {
  const el = $("statusLabel");
  el.textContent = text;
  el.title = text; // full text on hover when ellipsised
  el.classList.toggle("is-error", Boolean(isError));
}

// Commit buttons are disabled during a run too: the anchor is mid-flight
// and the suggestion on screen belongs to the previous request.
function setActionsEnabled(enabled) {
  document
    .querySelectorAll(
      "#quickActions button, #applyCustomBtn, #replaceBtn, #insertBelowBtn"
    )
    .forEach((btn) => (btn.disabled = !enabled));
}

// -- settings ------------------------------------------------------------
function openSettings() {
  $("endpoint").value = cfg.endpoint;
  $("temperature").value = cfg.temperature;
  $("timeout").value = cfg.timeout;
  $("systemPrompt").value = cfg.systemPrompt;
  const select = $("modelSelect");
  select.innerHTML = "";
  const opt = document.createElement("option");
  opt.value = cfg.model;
  opt.textContent = cfg.model;
  select.appendChild(opt);
  setSettingsStatus("");

  $("mainView").classList.add("hidden");
  $("settingsView").classList.remove("hidden");
}

function closeSettings() {
  $("settingsView").classList.add("hidden");
  $("mainView").classList.remove("hidden");
}

function setSettingsStatus(text, isError = false) {
  const el = $("settingsStatus");
  el.textContent = text;
  el.title = text;
  el.classList.toggle("is-error", Boolean(isError));
}

async function onFetchModels() {
  const endpoint = $("endpoint").value.trim() || cfg.endpoint;
  setSettingsStatus("Fetching…");
  try {
    const models = await listOllamaModels(endpoint);
    const select = $("modelSelect");
    if (models.length === 0) {
      setSettingsStatus("Connected, but no models installed (ollama pull …).", true);
      return;
    }
    select.innerHTML = "";
    for (const name of models) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      select.appendChild(opt);
    }
    if (models.includes(cfg.model)) select.value = cfg.model;
    setSettingsStatus(`${models.length} model(s) found.`);
  } catch (e) {
    setSettingsStatus(e.message, true);
  }
}

function onSaveSettings() {
  const temperature = $("temperature").value.trim();
  const timeout = $("timeout").value.trim();

  // Inline, not alert() - same reasoning as runAction's error path, and it
  // keeps the invalid field on screen next to the complaint about it.
  const temp = parseFloat(temperature);
  if (Number.isNaN(temp) || temp < 0 || temp > 1) {
    setSettingsStatus("Temperature must be a number between 0 and 1.", true);
    return;
  }
  const timeoutNum = parseFloat(timeout);
  if (Number.isNaN(timeoutNum) || timeoutNum <= 0) {
    setSettingsStatus("Timeout must be a positive number of seconds.", true);
    return;
  }

  cfg = {
    endpoint: $("endpoint").value.trim() || cfg.endpoint,
    model: $("modelSelect").value || cfg.model,
    temperature,
    timeout,
    systemPrompt: $("systemPrompt").value,
  };
  saveConfig(cfg);
  closeSettings();
  checkConnection();
}
