const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { closeFixPR } = require("../lib/github");
const { HttpError } = require("../lib/httpError");

// POST /api/fix-prs/{prNumber}/reject — closes the PR without merging.
app.http("rejectFixPr", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "fix-prs/{prNumber}/reject",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const prNumber = Number(request.params.prNumber);
      if (!prNumber) throw new HttpError(400, "Invalid PR number");
      await closeFixPR(prNumber);
      return { jsonBody: { closed: true } };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
