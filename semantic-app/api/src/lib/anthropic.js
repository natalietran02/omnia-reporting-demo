const Anthropic = require("@anthropic-ai/sdk");
const { HttpError } = require("./httpError");

// Model choice is fixed, not env-configurable — keeps the pricing table
// below (and every cost calculation that depends on it) trustworthy.
const CLAUDE_MODEL = "claude-opus-5";

// USD per 1M tokens (https://www.anthropic.com/pricing). Add an entry here
// if CLAUDE_MODEL above ever changes.
const PRICING_PER_MTOK = {
  "claude-opus-5": { input: 5, output: 25 },
};

function callCostUsd(usage, model) {
  const pricing = PRICING_PER_MTOK[model];
  if (!pricing || !usage) return 0;
  const perMtok = (n) => (Number(n) || 0) / 1e6;
  return (
    perMtok(usage.input_tokens) * pricing.input +
    perMtok(usage.output_tokens) * pricing.output +
    perMtok(usage.cache_creation_input_tokens) * pricing.input * 1.25 +
    perMtok(usage.cache_read_input_tokens) * pricing.input * 0.1
  );
}

// ---------- AI spend budget guard ----------
// Same purpose as Code.js's PropertiesService-backed guard (assertAiBudgetOk_/
// recordAiSpend_/aiBudgetStatus_): block the next call once a configured
// daily/monthly USD ceiling is reached, computed from each response's real
// `usage` block rather than a per-request count.
//
// NOT durable like the Apps Script version: this is a plain module-level
// object, so it resets on cold start and is NOT shared across concurrent
// Function instances under load. That's an acceptable soft ceiling for a
// demo app with a handful of admin users hitting these routes by hand; if
// this app scales up or the budget genuinely needs to hold, move this to
// Azure Table Storage (the Function App's own storage account, already
// available via the AzureWebJobsStorage connection string) instead.
const spend = {
  day: { key: "", usd: 0 },
  month: { key: "", usd: 0 },
};

function todayKey() {
  return new Date().toISOString().slice(0, 10); // YYYY-MM-DD (UTC)
}
function monthKeyNow() {
  return new Date().toISOString().slice(0, 7); // YYYY-MM (UTC)
}

function assertBudgetOk() {
  const todayK = todayKey();
  if (spend.day.key !== todayK) {
    spend.day.key = todayK;
    spend.day.usd = 0;
  }
  const monthK = monthKeyNow();
  if (spend.month.key !== monthK) {
    spend.month.key = monthK;
    spend.month.usd = 0;
  }

  const dailyLimit = Number(process.env.AI_DAILY_BUDGET_USD);
  if (dailyLimit > 0 && spend.day.usd >= dailyLimit) {
    throw new HttpError(
      429,
      "Daily AI budget reached ($" + spend.day.usd.toFixed(2) + " of $" + dailyLimit.toFixed(2) + " used). Try again tomorrow, or raise AI_DAILY_BUDGET_USD."
    );
  }
  const monthlyLimit = Number(process.env.AI_MONTHLY_BUDGET_USD);
  if (monthlyLimit > 0 && spend.month.usd >= monthlyLimit) {
    throw new HttpError(
      429,
      "Monthly AI budget reached ($" + spend.month.usd.toFixed(2) + " of $" + monthlyLimit.toFixed(2) + " used). Raise AI_MONTHLY_BUDGET_USD."
    );
  }
}

function recordSpend(costUsd) {
  if (!(costUsd > 0)) return;
  assertBudgetOk(); // also rolls the day/month buckets over if the clock ticked past between calls
  spend.day.usd += costUsd;
  spend.month.usd += costUsd;
}

function budgetStatus() {
  const todayK = todayKey();
  const monthK = monthKeyNow();
  return {
    dailySpentUsd: spend.day.key === todayK ? spend.day.usd : 0,
    dailyLimitUsd: Number(process.env.AI_DAILY_BUDGET_USD) || null,
    monthlySpentUsd: spend.month.key === monthK ? spend.month.usd : 0,
    monthlyLimitUsd: Number(process.env.AI_MONTHLY_BUDGET_USD) || null,
    durable: false, // see the module-level comment above
  };
}

// ---------- Client ----------
let client = null;
function getClient() {
  if (!process.env.ANTHROPIC_API_KEY) {
    throw new HttpError(500, "ANTHROPIC_API_KEY is not configured on this Function app");
  }
  if (!client) client = new Anthropic();
  return client;
}

// Maps the SDK's typed exceptions to an HttpError so every calling route
// can do the same `err instanceof HttpError ? err.status : 500` as the
// existing fix-PR routes, without each guessing at what went wrong.
function mapAnthropicError(err) {
  if (err instanceof Anthropic.AuthenticationError) return new HttpError(500, "Anthropic API rejected the configured API key");
  if (err instanceof Anthropic.RateLimitError) return new HttpError(429, "Anthropic API is rate-limiting this app — try again shortly");
  if (err instanceof Anthropic.APIError) return new HttpError(502, "Anthropic API error: " + err.message);
  return err;
}

// Plain-text generation (no structured output) — used for AI commentary.
async function anthropicCreate(params) {
  assertBudgetOk();
  const c = getClient();
  let response;
  try {
    response = await c.messages.create({ model: CLAUDE_MODEL, ...params });
  } catch (err) {
    throw mapAnthropicError(err);
  }
  recordSpend(callCostUsd(response.usage, CLAUDE_MODEL));
  return response;
}

module.exports = { CLAUDE_MODEL, anthropicCreate, budgetStatus };
