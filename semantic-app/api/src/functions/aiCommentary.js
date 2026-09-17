const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { anthropicCreate } = require("../lib/anthropic");
const { HttpError } = require("../lib/httpError");

// Ported from Code.js's getClaudeCommentary(). Unlike Apps Script — which
// read data_key_metrics_monthly/data_pl/data_powerbi_unified_fact directly
// off the spreadsheet — this app's reporting data already lives client-side
// (fetched from the Power BI semantic model via DAX queries), so the client
// builds the same per-OpCo/per-month summary text Code.js used to compute
// server-side and sends it as `dataSummary`, the same pattern the app
// already uses for the chat-style AI context elsewhere.
const SYSTEM_PROMPT =
  "You are a data quality analyst for a group of Australian professional-services operating companies. " +
  "Review the months of business metrics below for each operating company and flag anything anomalous or " +
  "potentially incorrect, either within one company or across companies. Be concise and specific, and group " +
  "findings by company. If everything looks normal, say so plainly.";

// POST /api/ai-commentary — body: { dataSummary } (plain text block built
// client-side, one section per OpCo).
app.http("aiCommentary", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "ai-commentary",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const body = await request.json().catch(() => ({}));
      const dataSummary = String(body.dataSummary || "").trim();
      if (!dataSummary) throw new HttpError(400, "dataSummary is required");

      const response = await anthropicCreate({
        max_tokens: 4096,
        output_config: { effort: "medium" },
        system: SYSTEM_PROMPT,
        messages: [{ role: "user", content: dataSummary }],
      });
      const textBlock = response.content.find((b) => b.type === "text");
      const commentary = textBlock ? textBlock.text : "(No commentary text returned.)";
      return { jsonBody: { commentary } };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
