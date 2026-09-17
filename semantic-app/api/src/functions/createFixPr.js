const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { createFixPR } = require("../lib/github");
const { HttpError } = require("../lib/httpError");

// POST /api/fix-prs — body: { finding, patch } — creates a branch, commits
// the patched file, and opens a PR (the create half of the existing
// list/approve/reject trio on this same route).
app.http("createFixPr", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "fix-prs",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const body = await request.json().catch(() => ({}));
      if (!body.finding) throw new HttpError(400, "finding object required");
      if (!body.patch) throw new HttpError(400, "patch object required");
      const result = await createFixPR(body.finding, body.patch);
      return { jsonBody: result };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
