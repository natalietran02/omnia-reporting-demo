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
// Runs the free layer-1 static checks (no API call) and returns the
// survivors, plus a report of what was rejected and why. There used to
// also be a layer-2 LLM re-verification pass (a second Claude call judging
// each finding against its source), but even a small/scoped review chaining
// two-to-three sequential Claude calls in one HTTP request was confirmed
// live to exceed Azure Static Web Apps' managed-API proxy timeout (it
// returned the proxy's own generic "Backend call failure", not an
// application error) — reliably fitting inside that timeout mattered more
// here than the extra confidence a second AI pass gave, so this app keeps
// only the free check.
function verifyAndFilterFindings(findings, source) {
  const staticResult = staticVerifyFindings(findings, source);
  return {
    findings: staticResult.kept,
    report: {
      generated: findings.length,
      staticRejected: staticResult.dropped.map((d) => ({ title: d.finding.title, reasons: d.reasons })),
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

// Picks function(s) relevant to a free-text description using local keyword
// matching against each function's name and leading-comment summary — no
// API call. This used to be a Claude call (semantic matching, better at
// e.g. matching "how discounts are calculated" to a function with no
// literal word overlap), but running it before the review call itself
// meant every targeted review chained two-to-three sequential Claude calls
// in one HTTP request, which was confirmed live to exceed Azure Static Web
// Apps' managed-API proxy timeout even though each individual call was
// small and fast. A plain keyword match is less clever, but the review
// call that follows still gets the surrounding lines of any function it
// picks, so a slightly imprecise match is rarely a wasted one.
const SELECTION_STOPWORDS = new Set([
  "the", "a", "an", "and", "or", "of", "in", "on", "to", "for", "is", "are", "was", "were",
  "this", "that", "it", "its", "with", "how", "does", "do", "when", "why", "what",
  "flow", "page", "code", "logic", "area", "review", "function", "bug", "issue",
]);

// Splits an identifier or sentence into lowercase word tokens, breaking on
// camelCase boundaries as well as spaces/punctuation — "openCodeFixModal"
// becomes ["open","code","fix","modal"]. This is what makes token matching
// (below) safe: comparing whole tokens instead of raw substrings means the
// term "fix" matches the real word "Fix" inside a camelCase name, but does
// NOT also match "fix" hiding inside "toFixed", "prefix", or "suffix" —
// none of those split into a standalone "fix" token, since they're a single
// lowercase run with no internal case change to split on. An earlier
// version of this function used plain substring search and matched "fix"
// (from a "fix PR" description) against `n.toFixed(1)` in unrelated
// formatting helpers — confirmed live, not just a theoretical risk.
function tokenize(text) {
  return String(text || "")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .toLowerCase()
    .match(/[a-z0-9]+/g) || [];
}

// True for an exact token match, or a simple singular/plural or verb-form
// variant ("picker"~"pickers", "review"~"reviewing") — never for one short
// token that just happens to prefix a longer, unrelated one. Both a minimum
// length and a maximum length gap are required: without them, a token like
// the trailing "p" that camelCase-splitting pulls out of "pctBadgeP" would
// trivially prefix-match almost anything ("picker".startsWith("p")) —
// confirmed live, this is what made "fix" match "toFixed" comments before
// tokenizing, and what made a single-letter split token match unrelated
// words after tokenizing but before this length guard.
function tokensRelated(a, b) {
  if (a === b) return true;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;
  if (shorter.length < 4 || longer.length - shorter.length > 4) return false;
  return longer.startsWith(shorter);
}

function selectRelevantFunctionsLocally(indexEntries, description) {
  const terms = Array.from(new Set(tokenize(description).filter((w) => w.length > 2 && !SELECTION_STOPWORDS.has(w))));
  if (!terms.length) {
    return { matched: [], note: "Describe the area using specific words — a function name, a UI label, or a behavior — rather than general terms." };
  }

  const scored = indexEntries
    .map((e) => {
      const haystackTokens = tokenize(e.name).concat(tokenize(e.summary));
      const score = terms.reduce((s, t) => {
        const hit = haystackTokens.some((tok) => tokensRelated(tok, t));
        return s + (hit ? 1 : 0);
      }, 0);
      return { name: e.name, score };
    })
    .filter((e) => e.score > 0)
    .sort((a, b) => b.score - a.score);

  if (!scored.length) {
    return { matched: [], note: 'Nothing in semantic-app/index.html matched a word in "' + description + '" — try naming a specific function, UI label, or behavior.' };
  }
  return { matched: scored.slice(0, 8).map((e) => e.name), note: "" };
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

  const selection = selectRelevantFunctionsLocally(index, description);
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

  // One Claude call total for this request (see selectRelevantFunctionsLocally
  // and verifyAndFilterFindings above for why) — scoped to a handful of
  // functions, never the whole file, so it can afford high effort.
  const findings = await runReviewPrompt(prompt, "high");
  const verifyResult = verifyAndFilterFindings(findings, source);

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
