const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { mergeFixPR } = require("../lib/github");
const { HttpError } = require("../lib/httpError");

// POST /api/fix-prs/{prNumber}/approve — squash-merges the PR.
app.http("approveFixPr", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "fix-prs/{prNumber}/approve",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const prNumber = Number(request.params.prNumber);
      if (!prNumber) throw new HttpError(400, "Invalid PR number");
      const result = await mergeFixPR(prNumber);
      return { jsonBody: { merged: true, sha: result.sha } };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
