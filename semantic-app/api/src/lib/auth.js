const { HttpError } = require("./httpError");

// Kept identical to semantic-app/index.html's own constants — this is a
// server-side mirror of checkAccess(), not a separate source of truth.
// If the AccessUsers workbook ever moves, update both places.
const BOOTSTRAP_ADMIN_EMAILS = ["natalie@edgered.com.au", "wil@edgered.com.au", "nathan@edgered.com.au"];
const FEEDBACK_DRIVE_ID = "b!nxzo85sDP0SbRvi-9_y_KVyL_Ahv4PJPifSrs1JXiuHYUPJ6ZuClQJmQ-gq4ilWE";
const FEEDBACK_ITEM_ID = "01MFUWM2QFPLMV4GOQBNCZP5ZXH6WOPKVV";
const ACCESS_TABLE_NAME = "AccessUsersTable";
const WORKBOOK_BASE = "https://graph.microsoft.com/v1.0/drives/" + FEEDBACK_DRIVE_ID + "/items/" + FEEDBACK_ITEM_ID;

// Verifies the caller is a signed-in admin before any fix-PR function does
// anything. The bearer token is the same Microsoft Graph token the frontend
// already holds for the feedback/access-list workbook (Files.ReadWrite
// scope) — passed straight through in the Authorization header rather than
// re-issued here.
//
// We never validate the JWT signature ourselves: Microsoft Graph does that
// the instant we use the token to call /me, so a forged, expired, or
// wrong-audience token simply fails that call with a 401 instead of needing
// our own signature/JWKS check.
async function requireAdmin(request) {
  const authHeader = request.headers.get("authorization") || "";
  const token = authHeader.replace(/^Bearer\s+/i, "").trim();
  if (!token) throw new HttpError(401, "Missing Authorization header");

  const meResp = await fetch("https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName", {
    headers: { Authorization: "Bearer " + token },
  });
  if (!meResp.ok) {
    // Temporary: surface Graph's actual error instead of a generic message
    // while we track down why /me is being rejected.
    const bodyText = await meResp.text().catch(() => "");
    throw new HttpError(401, "Graph /me failed (" + meResp.status + "): " + bodyText.slice(0, 300));
  }
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
