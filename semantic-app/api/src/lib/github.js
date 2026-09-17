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

// The single file AI code review targets — this app is one SPA file plus a
// handful of Function files, unlike Code.js's original Code.js+Index.html
// pair, so review scope is kept simple rather than juggling several files.
const REVIEW_TARGET_FILE = "semantic-app/index.html";

// Fetches one file's current content + blob sha from the main branch —
// mirrors the read half of fetchCodeFromGitHub() in Code.js, generalised to
// an arbitrary repo-relative path instead of two hardcoded filenames.
async function fetchFileContent(path) {
  const data = await githubRequest("GET", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/contents/" + path);
  return {
    content: Buffer.from(data.content, "base64").toString("utf-8"),
    sha: data.sha,
  };
}

async function getMainBranchSha() {
  const ref = await githubRequest("GET", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/git/ref/heads/main");
  return ref.object.sha;
}

async function createBranch(branchName, fromSha) {
  return githubRequest("POST", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/git/refs", {
    ref: "refs/heads/" + branchName,
    sha: fromSha,
  });
}

async function commitFile(branch, filePath, content, currentSha, commitMessage) {
  return githubRequest("PUT", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/contents/" + filePath, {
    message: commitMessage,
    content: Buffer.from(content, "utf-8").toString("base64"),
    sha: currentSha,
    branch,
  });
}

async function openPullRequest(title, body, headBranch) {
  return githubRequest("POST", "/repos/" + GITHUB_OWNER + "/" + GITHUB_REPO + "/pulls", {
    title,
    body,
    head: headBranch,
    base: "main",
  });
}

// Creates a branch, commits the patched file, and opens a PR — mirrors
// createFixPR() in Code.js. Re-fetches the file fresh (rather than trusting
// a sha the client may be holding stale) so a concurrent edit to main is
// caught as a real conflict instead of silently overwritten.
async function createFixPR(finding, patch) {
  const filePath = finding.file || REVIEW_TARGET_FILE;
  const current = await fetchFileContent(filePath);

  if (current.content.indexOf(patch.old_string) === -1) {
    throw new HttpError(409, "patch.old_string not found in " + filePath + " — the file has changed since this fix was generated, re-run code_fix");
  }
  const patched = current.content.replace(patch.old_string, patch.new_string);

  const mainSha = await getMainBranchSha();
  const slug = String(finding.title || "fix")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 38);
  const ts = Date.now().toString().slice(-4);
  const branchName = "fix/" + slug + "-" + ts;

  await createBranch(branchName, mainSha);
  await commitFile(branchName, filePath, patched, current.sha, "fix(" + (finding.location || filePath) + "): " + finding.title);

  const prBody =
    "## Finding\n\n" +
    "**Severity:** " + finding.severity + "  \n" +
    "**File:** " + finding.file + "  \n" +
    "**Location:** `" + finding.location + "`\n\n" +
    "**Issue:** " + finding.description + "\n\n" +
    "**Suggested fix:** " + finding.suggestion + "\n\n" +
    "---\n_Auto-generated by Omnia AI code review · " + new Date().toISOString() + "_";

  const pr = await openPullRequest("[Omnia Fix] " + finding.title, prBody, branchName);
  return { pr_url: pr.html_url, pr_number: pr.number, branch: branchName };
}

module.exports = {
  listFixPRs,
  mergeFixPR,
  closeFixPR,
  REVIEW_TARGET_FILE,
  fetchFileContent,
  createFixPR,
};
