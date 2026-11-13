#!/usr/bin/env bash
# SessionStart hook: bring this checkout up to date with GitHub, or say why not.
# Output goes to Claude's context, so it can tell the user what happened.
# Never changes local work: only a fast-forward pull is attempted.

cd "${CLAUDE_PROJECT_DIR:-.}" 2>/dev/null || exit 0

if ! command -v git >/dev/null 2>&1; then
  echo "[git-sync] git is not installed on this machine. Nothing was synced."
  exit 0
fi

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "[git-sync] This folder is not a git repository. Ask the user whether to set up git and a private GitHub repository."
  exit 0
fi

if ! git remote get-url origin >/dev/null 2>&1; then
  echo "[git-sync] No GitHub remote (origin) is configured. Ask the user whether to create a private GitHub repository and push."
  exit 0
fi

branch=$(git symbolic-ref --quiet --short HEAD 2>/dev/null)
if [ -z "$branch" ]; then
  echo "[git-sync] HEAD is detached. Nothing was synced."
  exit 0
fi

if ! git rev-parse --abbrev-ref '@{u}' >/dev/null 2>&1; then
  echo "[git-sync] Branch '$branch' has no upstream on GitHub yet. Nothing was pulled; it needs 'git push -u origin $branch'."
  exit 0
fi

before=$(git rev-parse HEAD)
if out=$(git pull --ff-only --quiet 2>&1); then
  after=$(git rev-parse HEAD)
  if [ "$before" = "$after" ]; then
    echo "[git-sync] '$branch' is already up to date with GitHub."
  else
    count=$(git rev-list --count "$before..$after")
    echo "[git-sync] Pulled $count new commit(s) from GitHub into '$branch':"
    git log --oneline "$before..$after" | head -n 10
  fi
else
  echo "[git-sync] Could NOT fast-forward '$branch'. Local files were left untouched. Tell the user and help resolve it."
  echo "$out" | head -n 5
fi

ahead=$(git rev-list --count '@{u}..HEAD' 2>/dev/null || echo 0)
behind=$(git rev-list --count 'HEAD..@{u}' 2>/dev/null || echo 0)
dirty=$(git status --porcelain 2>/dev/null | wc -l | tr -d ' ')
if [ "$ahead" != "0" ]; then
  echo "[git-sync] $ahead local commit(s) are not pushed to GitHub yet."
fi
if [ "$behind" != "0" ]; then
  echo "[git-sync] Still $behind commit(s) behind GitHub."
fi
if [ "$dirty" != "0" ]; then
  echo "[git-sync] $dirty file(s) have uncommitted changes."
fi
exit 0
