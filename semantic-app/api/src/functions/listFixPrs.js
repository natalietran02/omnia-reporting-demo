const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { listFixPRs } = require("../lib/github");
const { HttpError } = require("../lib/httpError");

// GET /api/fix-prs — lists open pull requests on fix/* branches.
app.http("listFixPrs", {
  methods: ["GET"],
  authLevel: "anonymous", // access is gated by requireAdmin below, not the Functions key
  route: "fix-prs",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const prs = await listFixPRs();
      return { jsonBody: prs };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
