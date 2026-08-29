/*
 * Word-for-word port of REWRITE_ACTIONS / NO_PREAMBLE_SUFFIX /
 * build_prompt from extension/clausesmith_core.py. Keep the wording
 * identical across platforms so the same instruction produces the same
 * behavior regardless of which app it's running in.
 */

export const REWRITE_ACTIONS = {
  "Improve writing":
    "Improve the clarity and flow of the following text without " +
    "changing its meaning:",
  "Fix grammar":
    "Correct any grammar, spelling, and punctuation errors in the " +
    "following text, otherwise leaving it unchanged:",
  Shorten:
    "Make the following text more concise without losing the key points:",
  "Make formal":
    "Rewrite the following text in a more formal, professional tone:",
  Summarize: "Summarize the following text:",
};

const NO_PREAMBLE_SUFFIX =
  "\n\nReturn ONLY the rewritten text itself - no preamble, no " +
  'explanation, no introductory phrase like "Here is the rewritten ' +
  'text:", and no quotation marks wrapping it.';

/**
 * A rewrite prompt with the no-preamble instruction reinforced right next
 * to the text, not just in the system prompt - smaller local models don't
 * always follow a system-level instruction reliably, and a stray "Here's
 * the rewritten version:" is a real problem when the output gets
 * inserted straight into a document.
 */
export function buildPrompt(instruction, text) {
  return `${instruction}\n\n${text}${NO_PREAMBLE_SUFFIX}`;
}
