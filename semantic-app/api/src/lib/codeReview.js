const { z } = require("zod");
const { anthropicCreate } = require("./anthropic");
const { REVIEW_TARGET_FILE, fetchFileContent } = require("./github");
const { HttpError } = require("./httpError");

// Ported from Code.js's runCodeReview()/runTargetedCodeReview() + its
// two-layer verification pass, adapted to this app's single reviewed file
// (semantic-app/index.html, not Code.js+Index.html).
//
// Uses the SAME plain-text-generation + manual-JSON-extraction approach
// Code.js used (extractJsonSpan/sanitizeClaudeJson below, ported verbatim),
// NOT Anthropic's newer structured-output API (`output_config.format` /
// `messages.parse()`) — an earlier version of this file used that instead,
// and every route built on it failed live in production (Azure's proxy
// returned a generic "Backend call failure") while the one route NOT using
// it (AI commentary, plain `anthropicCreate`) worked. Rather than keep
// debugging an unverified API combination, this reverts to the exact
// pattern already proven working in this app. Zod schemas below now only
// validate the parsed JSON locally (a safety net), not to enforce the
// API's response shape.

// ---------- JSON-from-prose helpers (ported verbatim from Code.js) ----------

// Finds the JSON array/object in Claude's response text, even when Claude
// prefaces it with a sentence or two despite being told to return ONLY
// JSON. Tracks bracket depth — skipping brackets inside "..." strings so
// they don't throw off the count — and returns the largest complete
// balanced span, since the real JSON payload is always far bigger than an
// incidental bracket in a sentence. Returns null if no balanced span exists.
function extractJsonSpan(text, openChar, closeChar) {
  let best = null;
  let inString = false;
  let escaped = false;
  let depth = 0;
  let start = -1;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) { escaped = false; }
      else if (ch === "\\") { escaped = true; }
      else if (ch === '"') { inString = false; }
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === openChar) {
      if (depth === 0) start = i;
      depth++;
      continue;
    }
    if (ch === closeChar && depth > 0) {
      depth--;
      if (depth === 0 && start !== -1) {
        const span = text.slice(start, i + 1);
        if (!best || span.length > best.length) best = span;
        start = -1;
      }
    }
  }
  return best;
}

// Claude's JSON output sometimes isn't valid JSON as-is: bare backslashes in
// string values, literal newline/tab/CR characters instead of escapes, and
// unescaped literal " characters (especially in "quote" fields that
// reproduce source code verbatim) that prematurely end the JSON string.
// Walks the text tracking whether it's inside a "..." string and fixes all
// three only there.
function sanitizeClaudeJson(raw) {
  let out = "";
  let inString = false;
  let escaped = false;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (!inString) {
      if (ch === '"') inString = true;
      out += ch;
      continue;
    }
    if (escaped) {
      escaped = false;
      if (ch === "u") {
        const hex = raw.substr(i + 1, 4);
        out += /^[0-9a-fA-F]{4}$/.test(hex) ? "\\u" : "\\\\u";
      } else if ('"\\/bfnrt'.indexOf(ch) !== -1) {
        out += "\\" + ch;
      } else {
        out += "\\\\" + ch; // invalid escape sequence — escape the backslash itself
      }
      continue;
    }
    if (ch === "\\") { escaped = true; continue; }
    if (ch === '"') {
      // A real string terminator is followed by a JSON structural character
      // (or end of input) once trailing whitespace is skipped — anything
      // else means this quote is literal content Claude failed to escape.
      let j = i + 1;
      while (j < raw.length && /\s/.test(raw[j])) j++;
      const next = raw[j];
      if (next === undefined || ",}]:".indexOf(next) !== -1) {
        inString = false;
        out += ch;
      } else {
        out += '\\"';
      }
      continue;
    }
    if (ch === "\n") { out += "\\n"; continue; }
    if (ch === "\r") { out += "\\r"; continue; }
    if (ch === "\t") { out += "\\t"; continue; }
    out += ch;
  }
  return out;
}

// Sends a prompt, extracts the JSON array/object from the plain-text
// response, sanitizes it, and validates it against a local Zod schema
// (a safety net over the parsed result — not an API-enforced shape).
async function askForJson(prompt, schema, bracket, effort, maxTokens) {
  const openChar = bracket === "array" ? "[" : "{";
  const closeChar = bracket === "array" ? "]" : "}";
  const response = await anthropicCreate({
    max_tokens: maxTokens,
    output_config: { effort: effort || "high" },
    messages: [{ role: "user", content: prompt }],
  });
  if (response.stop_reason === "max_tokens") {
    throw new HttpError(502, "Claude's response was cut off by the token limit before finishing — try a narrower description.");
  }
  const textBlock = response.content.find((b) => b.type === "text");
  const text = textBlock ? textBlock.text : "";
  const match = extractJsonSpan(text, openChar, closeChar);
  if (!match) throw new HttpError(502, "Could not find a JSON " + bracket + " in Claude's response");
  let parsed;
  try {
    parsed = JSON.parse(sanitizeClaudeJson(match));
  } catch (e) {
    throw new HttpError(502, "Claude returned malformed JSON: " + (e && e.message));
  }
  const result = schema.safeParse(parsed);
  if (!result.success) {
    throw new HttpError(502, "Claude's JSON didn't match the expected shape: " + result.error.issues.map((iss) => iss.message).join("; "));
  }
  return result.data;
}

// ---------- Schemas (local validation only — see header comment) ----------

const FindingSchema = z.object({
  severity: z.enum(["critical", "high", "medium", "low"]),
  location: z.string(),
  title: z.string(),
  description: z.string(),
  quote: z.string(),
  suggestion: z.string(),
});
const FindingsSchema = z.array(FindingSchema).max(15);

const SelectionSchema = z.object({
  matched: z.array(z.string()),
  note: z.string(),
});

const VerdictsSchema = z.array(
  z.object({
    index: z.number(),
    verdict: z.enum(["CONFIRMED", "REFUTED"]),
    reason: z.string(),
  })
);

const PatchSchema = z.object({
  old_string: z.string(),
  new_string: z.string(),
  explanation: z.string(),
});

const REVIEW_FOCUS =
  "Review the code for real bugs and security issues only. Focus on:\n" +
  "- Logic bugs: wrong calculations, off-by-one errors, missing return statements\n" +
  "- Date/period bugs: month/year arithmetic, boundary conditions (e.g. \"first day of the following month\")\n" +
  "- XSS/injection: user- or data-derived strings written into innerHTML without escHtml()/escAttr()\n" +
  "- Auth/token handling: the Graph token (sent via the x-omnia-graph-token header, never Authorization) leaking into logs, error messages, or the wrong request\n" +
  "- Null/undefined dereferences that crash rendering\n" +
  "- State bugs: STATE mutated in a way that produces a stale or inconsistent re-render\n\n" +
  "Before finalizing findings, apply these checks — the bug classes below are the ones most often missed " +
  "by a holistic read-through because nothing crashes and nothing looks obviously wrong at a glance:\n" +
  "- Do not lower scrutiny for functions commented as test-only, dev-only, or debug — review them with the same rigor as everything else.\n" +
  "- For any function whose comment states an intended behavior (a boundary condition, a date rule, an escaping " +
  "requirement), explicitly restate what the comment promises, then check the code line-by-line against that " +
  "restatement — do not just judge whether the code \"looks reasonable\" on its own.\n\n" +
  "Only report a finding if you can back it with an exact quote from the code below — never describe code from " +
  "memory or assumption.";

const FINDINGS_JSON_INSTRUCTIONS =
  "Return ONLY a JSON array (no other text, no markdown fences) with up to 15 real findings. " +
  "Only report a finding if you can back it with an exact quote from the code below:\n" +
  "[\n  {\n    \"severity\": \"critical|high|medium|low\",\n    \"location\": \"function name\",\n" +
  "    \"title\": \"short title (max 60 chars)\",\n    \"description\": \"what is wrong and why it matters\",\n" +
  "    \"quote\": \"the exact code (verbatim, character for character, 1-6 lines) this finding is about\",\n" +
  "    \"suggestion\": \"how to fix it in plain English — no code snippets or regex patterns\"\n  }\n]\n\n";

// ---------- Function-block indexing (pure string ops, no GAS dependency) ----------

// Scans fullSource for top-level `function name(` declarations and returns
// [{ name, start, end }] line ranges (end = line before next function, or EOF).
function findFunctionBlocks(fullSource) {
  const lines = fullSource.split("\n");
  const blocks = [];
  const re = /^\s*function\s+([A-Za-z0-9_$]+)\s*\(/;
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(re);
    if (m) blocks.push({ name: m[1], start: i, end: lines.length - 1 });
  }
  for (let j = 0; j < blocks.length - 1; j++) blocks[j].end = blocks[j + 1].start - 1;
  return { lines, blocks };
}

// Contiguous // comment line(s) directly above a function declaration, used
// as a one-line summary of what it does when building the function index.
function leadingCommentFor(lines, startIdx) {
  const comments = [];
  let i = startIdx - 1;
  while (i >= 0) {
    const t = lines[i].trim();
    if (t.indexOf("//") !== 0) break;
    comments.unshift(t.replace(/^\/\/\s?/, ""));
    i--;
  }
  return comments.length ? comments.join(" ") : lines[startIdx].trim();
}

// Lightweight index of every top-level function — name + one-line summary,
// not the full body — cheap enough to hand the model in full so it can pick
// relevant functions by meaning rather than keyword/substring matching.
function buildFunctionIndex(source) {
  const found = findFunctionBlocks(source);
  const index = found.blocks.map((b) => ({ name: b.name, summary: leadingCommentFor(found.lines, b.start).slice(0, 180) }));
  return { lines: found.lines, blocks: found.blocks, index };
}

// Pulls the JS out of <script> tags (skipping external/src= includes and
// non-JS types) so it can be syntax-checked without the surrounding HTML
// confusing the parser.
function extractInlineScripts(html) {
  const out = [];
  const re = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    const attrs = m[1] || "";
    if (/\bsrc\s*=/i.test(attrs)) continue;
    if (/type\s*=\s*["'](?!text\/javascript)[^"']*["']/i.test(attrs)) continue;
    out.push(m[2]);
  }
  return out.join("\n;\n");
}

// True if jsCode parses as valid JavaScript — doesn't execute it, just
// checks it compiles. This is what actually disproves a "this causes a
// SyntaxError" claim, rather than trusting the model's description.
function checkSyntaxValid(jsCode) {
  try {
    new Function(jsCode);
    return true;
  } catch (e) {
    return false;
  }
}

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Loose existence check for a finding's claimed "location" (function name).
function locationExistsInSource(source, location) {
  const name = String(location || "").trim();
  if (!name) return true;
  const esc = escapeRegExp(name);
  const patterns = [
    new RegExp("function\\s+" + esc + "\\s*\\("),
    new RegExp("(const|let|var)\\s+" + esc + "\\s*="),
    new RegExp(esc + "\\s*[:=]\\s*function"),
    new RegExp(esc + "\\s*=\\s*\\("),
  ];
  return patterns.some((re) => re.test(source));
}

function normalizeForMatch(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

// A finding's "quote" must be a real, verbatim (whitespace-insensitive)
// substring of the reviewed file — otherwise it's describing code that
// isn't actually there.
function quoteExistsInSource(quote, source) {
  const nq = normalizeForMatch(quote);
  if (!nq) return false;
  return normalizeForMatch(source).indexOf(nq) !== -1;
}

const SYNTAX_CLAIM_RE = /syntax\s*error|redeclar|declared\s+twice|duplicate\s+(let|const|var|declaration)|won'?t\s+(compile|parse|load)|throws?\s+(a\s+)?syntaxerror|fail(s)?\s+to\s+(compile|parse|load)/i;

function claimsSyntaxBreak(finding) {
  const text = (finding.title || "") + " " + (finding.description || "");
  return SYNTAX_CLAIM_RE.test(text);
}

// Layer 1 verification: plain string/regex checks, no API call.
// - every finding must quote real, verbatim code from the file
// - the function name it points to must actually exist
// - any finding claiming a syntax-breaking bug is auto-rejected if the
//   inline scripts still compile cleanly (new Function() over the real source)
function staticVerifyFindings(findings, source) {
  const scripts = extractInlineScripts(source);
  const syntaxOk = checkSyntaxValid(scripts);

  const kept = [];
  const dropped = [];
  findings.forEach((f) => {
    const reasons = [];
    if (!f.quote) {
      reasons.push("no verbatim quote supplied");
    } else if (!quoteExistsInSource(f.quote, source)) {
      reasons.push("quote does not appear verbatim in the file");
    }
    if (f.location && !locationExistsInSource(source, f.location)) {
      reasons.push('location "' + f.location + '" was not found in the file');
    }
    if (claimsSyntaxBreak(f) && syntaxOk) {
      reasons.push("claims a syntax-breaking bug but the file compiles cleanly as-is");
    }
    if (reasons.length) dropped.push({ finding: f, reasons });
    else kept.push(f);
  });
  return { kept, dropped };
}

// Grabs just the function body a finding's "location" points to (plus a
// couple lines of leading context) so the batched verify call doesn't need
// to resend the whole file per finding.
function snippetForLocation(source, location, fallback) {
  const found = findFunctionBlocks(source);
  const block = found.blocks.find((b) => b.name === location);
  if (!block) return fallback || "";
  const start = Math.max(0, block.start - 2);
  const end = Math.min(found.lines.length - 1, block.end);
  return found.lines.slice(start, end + 1).join("\n");
}

// Pulls up to maxLines other lines in the file that mention the same
// identifier(s) as the finding's location/quote, outside the primary
// snippet — lets the verifier see a design-intent comment or a second call
// site living elsewhere in the file, cheaply (plain regex, no extra API cost).
function relatedMentions(source, location, quote, primarySnippet, maxLines) {
  maxLines = maxLines || 12;
  const tokens = new Set();
  (String(location || "").match(/[A-Za-z_$][A-Za-z0-9_$]{2,}/g) || []).forEach((t) => tokens.add(t));
  (String(quote || "").match(/[A-Za-z_$][A-Za-z0-9_$]{2,}/g) || []).forEach((t) => tokens.add(t));
  const names = Array.from(tokens).filter((t) => !/^(function|const|let|var|return|null|undefined|true|false|this)$/.test(t));
  if (!names.length) return "";

  const lines = source.split("\n");
  const res = [];
  for (let i = 0; i < lines.length && res.length < maxLines; i++) {
    const line = lines[i];
    if (primarySnippet.indexOf(line) !== -1) continue;
    if (names.some((n) => line.indexOf(n) !== -1)) res.push(i + 1 + ": " + line.trim());
  }
  return res.join("\n");
}

// Layer 2 verification: one batched call covering every layer-1 survivor,
// each judged against its own snippet plus nearby related lines — added
// after a review call confirmed findings a design-intent comment or a
// second call site would have disproved, because it never saw those lines.
async function verifyFindingsWithLLM(findings, source) {
  if (!findings.length) return [];

  const items = findings
    .map((f, i) => {
      const snippet = snippetForLocation(source, f.location, f.quote);
      const related = relatedMentions(source, f.location, f.quote, snippet);
      return (
        "--- Finding #" + i + " ---\n" +
        "Location: " + f.location + "\nTitle: " + f.title + "\n" +
        "Description: " + f.description + "\n" +
        "Code:\n```\n" + snippet + "\n```\n" +
        (related ? "Other lines elsewhere in the file mentioning the same identifiers (may confirm, explain, or contradict the finding):\n```\n" + related + "\n```\n" : "")
      );
    })
    .join("\n");

  const prompt =
    "You are a skeptical senior engineer double-checking a colleague's code review findings before they " +
    "reach a human, specifically to catch cases where the reviewer misremembered or fabricated code structure, " +
    "or invented a failure scenario that sounds plausible but cannot actually happen.\n\n" +
    'For each finding, judge it against the code snippet AND the "other lines elsewhere" block for that finding, ' +
    "if one is given. Rules:\n" +
    "1. Do not accept the finding's own description of what the code does — re-derive it yourself from the " +
    "snippet. If the finding depends on arithmetic, an index/loop bound, or a formula, recompute it by hand " +
    "line by line with concrete example values before deciding; if your own recomputation contradicts the " +
    "finding's claimed result, REFUTE and state what you actually got.\n" +
    '2. If the finding claims a failure scenario ("if X happens, then Y breaks"), REFUTE unless the snippet ' +
    "shows X is actually reachable given the real logic — an unverified hypothetical is not a bug.\n" +
    "3. If the finding claims two names/fields/locations are duplicated, orphaned, or inconsistent, check the " +
    '"other lines elsewhere" block first for a comment or second usage that explains them as intentionally ' +
    "distinct — if one exists, REFUTE.\n" +
    "4. If the finding claims something about ORDER (one line/call happening before or after another), quote " +
    "the exact two lines in your reason and confirm their relative order as they literally appear — if you " +
    "cannot see both locations to compare, REFUTE as unverifiable rather than trusting the description.\n" +
    "5. Default to REFUTED whenever you are not certain — an unverified finding reaching the user is worse " +
    "than a real one being dropped, since real bugs get caught again on the next review pass.\n\n" +
    items + "\n\n" +
    'Return ONLY a JSON array (no other text, no markdown fences), exactly one entry per finding, same order:\n' +
    '[\n  {\n    "index": 0,\n    "verdict": "CONFIRMED|REFUTED",\n    "reason": "one sentence"\n  }\n]';

  let verdicts;
  try {
    verdicts = await askForJson(prompt, VerdictsSchema, "array", "high", 2048);
  } catch (e) {
    // If the verify call itself fails, don't silently drop everything —
    // pass findings through unverified rather than blocking the whole review.
    return findings.map((f) => ({ ...f, verify: { verdict: "UNVERIFIED", reason: "verification call failed: " + (e && e.message) } }));
  }

  return findings.map((f, i) => {
    const v = verdicts.find((x) => x.index === i);
    return {
      ...f,
      verify: v ? { verdict: v.verdict, reason: v.reason } : { verdict: "UNVERIFIED", reason: "verification call returned no verdict for this finding" },
    };
  });
}

// Runs verification and returns only the survivors, plus a report of what
// was rejected and why. The layer-1 static checks (free, no API call)
// always run; layer 2 (the batched LLM re-verify call) only runs for scoped
// reviews (see runTargetedCodeReview) — small enough to comfortably fit
// within Azure Static Web Apps' managed-API proxy timeout.
async function verifyAndFilterFindings(findings, source) {
  const staticResult = staticVerifyFindings(findings, source);
  const verified = await verifyFindingsWithLLM(staticResult.kept, source);
  const refuted = verified.filter((f) => f.verify && f.verify.verdict === "REFUTED");
  const final = verified.filter((f) => !f.verify || f.verify.verdict !== "REFUTED");
  return {
    findings: final,
    report: {
      generated: findings.length,
      staticRejected: staticResult.dropped.map((d) => ({ title: d.finding.title, reasons: d.reasons })),
      llmRefuted: refuted.map((f) => ({ title: f.title, reason: f.verify.reason })),
    },
  };
}

// Forces the model to explicitly consider every function instead of a
// single holistic pass randomly "noticing" only some of them.
function checklistFor(functionNames) {
  return (
    "\nBefore producing the findings array, go through this exact list of functions one by one and explicitly " +
    "check each of them against every focus area above — do not skip any:\n" +
    functionNames.map((n) => "- " + n).join("\n") +
    "\n(It is fine — expected, even — for most of these to have no real issue. Only include an entry for functions " +
    "where you found an actual, quotable problem.)\n\n"
  );
}

// Runs the review prompt and stamps every finding with the fixed target
// file — the model is never asked for a "file" field at all, since there's
// only ever one possible value here and asking for it is one more thing
// that could come back malformed.
async function runReviewPrompt(promptBody, effort) {
  const findings = await askForJson(promptBody, FindingsSchema, "array", effort, 8000);
  return findings.map((f) => ({ ...f, file: REVIEW_TARGET_FILE }));
}

// Asks Claude to pick which function(s) — out of the whole-file index — are
// actually relevant to a free-text problem description, instead of naive
// keyword matching.
async function selectRelevantFunctionsViaLLM(indexEntries, description) {
  if (!indexEntries.length) return { matched: [], note: "" };

  const listing = indexEntries.map((e) => "- " + e.name + " — " + e.summary).join("\n");
  const prompt =
    "Here is an index of every function in a JavaScript web app (name and one-line summary only, not the code):\n\n" +
    listing +
    '\n\nA user described this problem/area to review:\n"' + description + '"\n\n' +
    "Pick ONLY the function(s) from the list above that are actually relevant to that description — do not " +
    'invent function names that are not in the list. If nothing in the list is genuinely relevant, return an ' +
    'empty "matched" array and use "note" to briefly explain what\'s missing or suggest a more specific way to ' +
    "describe the area, since no match means nothing gets reviewed this time.\n\n" +
    "Return ONLY a JSON object (no other text, no markdown fences):\n" +
    '{\n  "matched": ["exact function name from the list", ...],\n' +
    '  "note": "why these were picked, or what theme to focus on if none matched"\n}';

  let parsed;
  try {
    parsed = await askForJson(prompt, SelectionSchema, "object", "medium", 700);
  } catch (e) {
    return { matched: [], note: "" };
  }

  // Defensive: only trust matches that actually exist in the index sent —
  // never let a hallucinated function name silently scope the review.
  const validNames = new Set(indexEntries.map((e) => e.name));
  const matched = parsed.matched.filter((n) => validNames.has(n));
  return { matched, note: parsed.note || "" };
}

// Scoped to the function(s) judged relevant to a user-supplied description.
// Deliberately does NOT fall back to reviewing the whole file when nothing
// matches — a full-file review is one large Claude call over ~5,000+ lines,
// which was confirmed live to exceed Azure Static Web Apps' managed-API
// proxy timeout (it returned the proxy's own "Backend call failure" rather
// than an application error). Rather than occasionally hitting that same
// failure whenever a description doesn't match a real function, a no-match
// here just returns no findings with a note explaining why.
async function runTargetedCodeReview(description) {
  description = String(description || "").trim();
  if (!description) throw new HttpError(400, "description is required");

  const { content: source } = await fetchFileContent(REVIEW_TARGET_FILE);
  const { lines, blocks, index } = buildFunctionIndex(source);

  let selection = { matched: [], note: "" };
  try {
    selection = await selectRelevantFunctionsViaLLM(index, description);
  } catch (e) {
    // Selection step failing is treated the same as a genuine no-match below.
  }
  const matched = selection.matched.length > 0;

  if (!matched) {
    return {
      findings: [],
      reviewedAt: new Date().toISOString(),
      filesReviewed: [],
      scope: {
        description,
        matched: false,
        note: selection.note || "No function in semantic-app/index.html matched that description — try naming a specific function, UI element, or behavior.",
      },
      verification: { generated: 0, staticRejected: [], llmRefuted: [] },
    };
  }

  const sections = selection.matched
    .map((name) => {
      const b = blocks.find((x) => x.name === name);
      if (!b) return null;
      const start = Math.max(0, b.start - 3);
      const end = Math.min(lines.length - 1, b.end);
      return "// lines " + (start + 1) + "-" + (end + 1) + "\n" + lines.slice(start, end + 1).join("\n");
    })
    .filter(Boolean);
  const content = "=== " + REVIEW_TARGET_FILE + " (relevant sections only) ===\n" + sections.join("\n\n// ...\n\n");

  const prompt =
    "You are a senior JavaScript engineer reviewing the Omnia Reporting web app (a single-page dashboard, " +
    "backed by a Power BI semantic model and a small Azure Functions API).\n\n" +
    'The user has described a specific problem area to focus on:\n"' + description + '"\n\n' +
    "Below are the section(s) of the codebase judged most relevant to that description. Only review this scoped " +
    "code — do not invent issues about code you cannot see.\n\n" +
    checklistFor(selection.matched) +
    REVIEW_FOCUS +
    "\n\n" +
    FINDINGS_JSON_INSTRUCTIONS +
    content;

  // Scoped to a handful of functions (never the whole file, see above), so
  // this can afford both the higher-effort review and the full two-pass
  // LLM verification within Azure's proxy timeout.
  const findings = await runReviewPrompt(prompt, "high");
  const verifyResult = await verifyAndFilterFindings(findings, source);

  return {
    findings: verifyResult.findings,
    reviewedAt: new Date().toISOString(),
    filesReviewed: selection.matched.map((n) => REVIEW_TARGET_FILE + ":" + n),
    scope: { description, matched: true, note: selection.note || "" },
    verification: verifyResult.report,
  };
}

// Given a specific finding, asks Claude for an exact old_string → new_string
// patch. Sends only the relevant function + surrounding context, not the
// full file, to stay well under input-size limits.
async function generateCodeFix(finding) {
  const { content: source } = await fetchFileContent(REVIEW_TARGET_FILE);
  const lines = source.split("\n");
  const funcName = finding.location || "";
  let funcIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].indexOf("function " + funcName) !== -1 || lines[i].indexOf(funcName + ":") !== -1 || lines[i].indexOf(funcName + " =") !== -1) {
      funcIdx = i;
      break;
    }
  }
  let fileContent;
  if (funcIdx !== -1) {
    const s = Math.max(0, funcIdx - 5);
    const e = Math.min(lines.length - 1, funcIdx + 200);
    fileContent = (s > 0 ? "// [file truncated above]\n" : "") + lines.slice(s, e + 1).join("\n") + (e < lines.length - 1 ? "\n// [file truncated below]" : "");
  } else {
    fileContent = source; // function not found — fall back to full file
  }

  const prompt =
    "You are fixing a specific bug in a JavaScript web app.\n\n" +
    "Finding:\n" +
    "- Title: " + finding.title + "\n" +
    "- Location: " + finding.location + "\n" +
    "- Description: " + finding.description + "\n" +
    "- Fix: " + finding.suggestion + "\n\n" +
    "File (" + REVIEW_TARGET_FILE + "):\n```\n" + fileContent + "\n```\n\n" +
    "Return ONLY a JSON object (no other text, no markdown fences):\n" +
    "{\n" +
    '  "old_string": "exact text to replace — include 3-8 surrounding lines so it is unique in the file",\n' +
    '  "new_string": "the replacement text",\n' +
    '  "explanation": "one sentence explaining the change"\n' +
    "}\n\n" +
    "The old_string must be an exact substring of the file above — character for character.";

  const patch = await askForJson(prompt, PatchSchema, "object", "high", 2048);
  return { patch, file: REVIEW_TARGET_FILE };
}

module.exports = { runTargetedCodeReview, generateCodeFix };
