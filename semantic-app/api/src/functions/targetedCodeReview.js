const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { runTargetedCodeReview } = require("../lib/codeReview");
const { HttpError } = require("../lib/httpError");

// POST /api/code-review/targeted — reviews just the function(s) an LLM
// judges relevant to a free-text description, body: { description }.
app.http("targetedCodeReview", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "code-review/targeted",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const body = await request.json().catch(() => ({}));
      const result = await runTargetedCodeReview(body.description);
      return { jsonBody: result };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
