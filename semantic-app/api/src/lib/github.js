const { HttpError } = require("./httpError");

// Fixed to this app's own repo — "Pending fix PRs" only ever targets the
// semantic-app codebase, per the decision to keep fixes scoped to the repo
// they apply to rather than juggling two.
const GITHUB_OWNER = "natalietran02";
const GITHUB_REPO = "omnia-reporting-demo-semantic";
const GITHUB_API = "https://api.github.com";

function getGitHubToken() {
  const t = process.env.GITHUB_TOKEN;
  if (!t) throw new HttpError(500, "GITHUB_TOKEN is not configured on this Function app");
  return t;
}

async function githubRequest(method, path, body) {
  const resp = await fetch(GITHUB_API + path, {
    method,
    headers: {
      Authorization: "Bearer " + getGitHubToken(),
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await resp.text();
  const data = text ? JSON.parse(text) : {};
  if (!resp.ok) {
    throw new HttpError(resp.status, "GitHub " + resp.status + ": " + (data.message || text));
  }
  return data;
}

// Lists open PRs on branches starting with fix/ — mirrors listFixPRs() in Code.js.
async function listFixPRs() {
  const prs = await githubRequest("GET", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/pulls?state=open&per_page=50");
  return prs
    .filter((pr) => pr.head && pr.head.ref && pr.head.ref.indexOf("fix/") === 0)
    .map((pr) => ({
      pr_number: pr.number,
      pr_url: pr.html_url,
      title: pr.title,
      branch: pr.head.ref,
      created: pr.created_at,
    }));
}

// Squash-merges a PR (Approve button) — mirrors mergeFixPR() in Code.js.
async function mergeFixPR(prNumber, commitTitle) {
  return githubRequest("PUT", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/pulls/" + prNumber + "/merge", {
    commit_title: commitTitle || "fix: merge PR #" + prNumber,
    merge_method: "squash",
  });
}

// Closes a PR without merging (Reject button) — mirrors closeFixPR() in Code.js.
async function closeFixPR(prNumber) {
  return githubRequest("PATCH", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/pulls/" + prNumber, {
    state: "closed",
  });
}

module.exports = { listFixPRs, mergeFixPR, closeFixPR };
