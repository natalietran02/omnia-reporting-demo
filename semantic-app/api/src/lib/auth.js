const { HttpError } = require("./httpError");

// Kept identical to semantic-app/index.html's own constants — this is a
// server-side mirror of checkAccess(), not a separate source of truth.
// If the AccessUsers workbook ever moves, update both places.
const BOOTSTRAP_ADMIN_EMAILS = ["natalie@edgered.com.au", "wil@edgered.com.au", "nathan@edgered.com.au"];
const FEEDBACK_DRIVE_ID = "b!nxzo85sDP0SbRvi-9_y_KVyL_Ahv4PJPifSrs1JXiuHYUPJ6ZuClQJmQ-gq4ilWE";
const FEEDBACK_ITEM_ID = "01MFUWM2QFPLMV4GOQBNCZP5ZXH6WOPKVV";
const ACCESS_TABLE_NAME = "AccessUsersTable";
const WORKBOOK_BASE = "https://graph.microsoft.com/v1.0/drives/" + FEEDBACK_DRIVE_ID + "/items/" + FEEDBACK_ITEM_ID;

// The caller's Graph token travels in a custom header, NOT Authorization —
// confirmed via a live diagnostic that Azure Static Web Apps' proxy
// overwrites the standard Authorization header with its own internal
// SWA-to-Function service token before this function ever sees the
// request, so a client-supplied bearer token in that header never survives
// the trip. Deliberately NOT prefixed "x-ms-" either — that prefix is
// Azure's own reserved namespace (x-ms-client-principal and friends), so a
// header spelled that way risks the exact same silent-overwrite problem.
const GRAPH_TOKEN_HEADER = "x-omnia-graph-token";

// Verifies the caller is a signed-in admin before any fix-PR function does
// anything. The token is the same Microsoft Graph token the frontend
// already holds for the feedback/access-list workbook (Files.ReadWrite +
// User.Read scopes) — passed straight through rather than re-issued here.
//
// We never validate the JWT signature ourselves: Microsoft Graph does that
// the instant we use the token to call /me, so a forged, expired, or
// wrong-audience token simply fails that call with a 401 instead of needing
// our own signature/JWKS check.
async function requireAdmin(request) {
  const token = (request.headers.get(GRAPH_TOKEN_HEADER) || "").trim();
  if (!token) throw new HttpError(401, "Missing " + GRAPH_TOKEN_HEADER + " header");

  const meResp = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName", {
    headers: { Authorization: "Bearer " + token },
  });
  if (!meResp.ok) throw new HttpError(401, "Invalid or expired sign-in token");
  const me = await meResp.json();
  const email = String(me.mail || me.userPrincipalName || "").toLowerCase();
  if (!email) throw new HttpError(401, "Could not determine caller identity");

  if (BOOTSTRAP_ADMIN_EMAILS.indexOf(email) !== -1) return email;

  // Same AccessUsersTable lookup as checkAccess() in index.html: row must
  // exist, be active, and be role "admin" — anything else is refused.
  const rowsResp = await fetch(WORKBOOK_BASE + "/workbook/tables('" + ACCESS_TABLE_NAME + "')/rows", {
    headers: { Authorization: "Bearer " + token },
  });
  if (!rowsResp.ok) throw new HttpError(403, "Admin access required");
  const data = await rowsResp.json();
  const row = (data.value || [])
    .map((r) => r.values[0])
    .find((v) => String(v[0] || "").toLowerCase() === email);
  const role = row ? row[1] : null;
  const status = row ? row[2] : null;
  if (!row || status !== "active" || role !== "admin") throw new HttpError(403, "Admin access required");

  return email;
}

module.exports = { requireAdmin };
