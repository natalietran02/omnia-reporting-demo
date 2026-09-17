const { app } = require("@azure/functions");
const { requireAdmin } = require("../lib/auth");
const { generateCodeFix } = require("../lib/codeReview");
const { HttpError } = require("../lib/httpError");

// POST /api/code-fix — body: { finding } (one finding object from a review
// response) — returns { patch: { old_string, new_string, explanation }, file }.
app.http("codeFix", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "code-fix",
  handler: async (request, context) => {
    try {
      await requireAdmin(request);
      const body = await request.json().catch(() => ({}));
      if (!body.finding) throw new HttpError(400, "finding object required");
      const result = await generateCodeFix(body.finding);
      return { jsonBody: result };
    } catch (err) {
      const status = err instanceof HttpError ? err.status : 500;
      context.error(err);
      return { status, jsonBody: { error: err.message } };
    }
  },
});
