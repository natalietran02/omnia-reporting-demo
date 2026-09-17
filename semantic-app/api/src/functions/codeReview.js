const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { runCodeReview } = require("../lib/codeReview");
const { HttpError } = require("../lib/httpError");

// POST /api/code-review — runs a full AI review of semantic-app/index.html.
app.http("codeReview", {
  methods: ["POST"],
  authLevel: "anonymous", // access is gated by requireAdmin below, not the Functions key
  route: "code-review",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const result = await runCodeReview();
      return { jsonBody: result };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
