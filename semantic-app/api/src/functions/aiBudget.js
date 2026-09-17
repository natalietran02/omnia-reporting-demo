const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { budgetStatus } = require("../lib/anthropic");
const { HttpError } = require("../lib/httpError");

// GET /api/ai-budget — current AI spend vs configured daily/monthly caps,
// for the Health page's KPI cards.
app.http("aiBudget", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "ai-budget",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      return { jsonBody: { hasApiKey: !!process.env.ANTHROPIC_API_KEY, ...budgetStatus() } };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
