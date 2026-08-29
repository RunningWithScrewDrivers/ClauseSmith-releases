/*
 * Ollama HTTP client - a JS port of extension/clausesmith_core.py's
 * call_ollama / list_ollama_models. Same endpoints, same request bodies,
 * same error-message wording, so the two platforms behave identically
 * from the user's point of view. Keep them in sync by hand; there's no
 * shared build step across a Python UNO extension and a JS Office
 * add-in, so this is a deliberate duplication, not an oversight.
 *
 * fetch() has no built-in socket timeout, unlike Python's
 * urllib.request(timeout=...) - AbortController stands in for it below.
 */

export const DEFAULT_CONFIG = {
  endpoint: "http://localhost:11434",
  model: "llama3.1",
  temperature: "0.7",
  timeout: "180",
  systemPrompt:
    "You are a concise writing assistant. Help draft, improve, and " +
    "brainstorm prose. Return only the requested text, with no preamble " +
    "or commentary.",
};

/**
 * Pull Ollama's own message out of an error response body.
 *
 * Ollama answers a request for a model that isn't installed with 404 and
 * a body like {"error": "model 'llama3.1' not found, try pulling it
 * first"} - far more useful than the bare status text.
 */
async function httpErrorDetail(response) {
  let body;
  try {
    body = await response.text();
  } catch {
    return response.statusText;
  }
  try {
    const parsed = JSON.parse(body);
    if (parsed && parsed.error) return parsed.error;
  } catch {
    // not JSON - fall through to the raw body
  }
  return body.trim() || response.statusText;
}

/**
 * POST /api/generate. Returns the generated text, or throws an Error
 * whose .message is meant to be shown to the user directly (mirrors
 * call_ollama's exception messages in clausesmith_core.py).
 */
export async function callOllama(cfg, prompt, systemPrompt) {
  const url = cfg.endpoint.replace(/\/+$/, "") + "/api/generate";
  const payload = {
    model: cfg.model,
    prompt,
    system: systemPrompt !== undefined ? systemPrompt : cfg.systemPrompt,
    options: { temperature: parseFloat(cfg.temperature || "0.7") },
    stream: false,
  };

  const timeoutMs = parseFloat(cfg.timeout || "180") * 1000;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  let response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
  } catch (e) {
    if (e.name === "AbortError") {
      throw new Error(
        `Ollama didn't respond within ${cfg.timeout}s. The connection ` +
          `worked, so this is generation taking longer than the timeout - ` +
          `common on the first request after Ollama's been idle (it has ` +
          `to load the model into memory first), or with a large model on ` +
          `CPU. Raise the timeout in Settings, or run 'ollama ps' to ` +
          `check whether the model is still loading.`
      );
    }
    // fetch() rejects with a generic "Failed to fetch" / "NetworkError" for
    // everything from "nothing listening on that port" to "browser blocked
    // the request as mixed content" - both look identical here, which is
    // why the Windows-desktop fetch spike (see word/README.md) has to be
    // confirmed by hand rather than inferred from an error string.
    throw new Error(
      `Couldn't reach Ollama at ${cfg.endpoint} (${e.message}). Is ` +
        `'ollama serve' running?`
    );
  } finally {
    clearTimeout(timer);
  }

  if (!response.ok) {
    throw new Error(
      `Ollama returned HTTP ${response.status}: ${await httpErrorDetail(response)}`
    );
  }

  const data = await response.json();
  // Checked outside any catch that could re-wrap it - see the 0.5.1 fix to
  // call_ollama in clausesmith_core.py for why this order matters: raising
  // here from inside a try/catch that also handles transport errors is how
  // "model not found" became "Ollama request failed: Ollama error: model
  // not found" in the LibreOffice version.
  if (data.error) {
    throw new Error(`Ollama error: ${data.error}`);
  }
  return data.response || "";
}

/** GET /api/tags -> sorted list of locally-installed model names. */
export async function listOllamaModels(endpoint) {
  const url = endpoint.replace(/\/+$/, "") + "/api/tags";
  let response;
  try {
    response = await fetch(url, { method: "GET" });
  } catch (e) {
    throw new Error(
      `Couldn't reach Ollama at ${endpoint} (${e.message}). Is 'ollama ` +
        `serve' running?`
    );
  }
  if (!response.ok) {
    throw new Error(
      `Ollama returned HTTP ${response.status}: ${await httpErrorDetail(response)}`
    );
  }
  const data = await response.json();
  const names = (data.models || [])
    .map((m) => m.name || m.model)
    .filter(Boolean);
  return names.sort();
}
