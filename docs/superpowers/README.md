# docs/superpowers — local scratch space

This directory is **gitignored** (see `.gitignore` line "`docs/superpowers/`"). Nothing under it is tracked by default.

## What lives here

- `specs/` — design specs written during brainstorming sessions
- `plans/` — implementation plans generated from those specs
- `poc/` — one-off proof-of-concept notes (as needed)

## Why it's ignored

The three sibling MCPs (`ms365-intent-mcp`, `slack-mcp`, `confluence-jira-mcp`) share the same convention: spec/plan artifacts are working documents from the brainstorming and planning phases, produced in service of shipping code. Once the code lands, the code itself is the durable truth. Keeping specs/plans out of git avoids:

- Divergence between what the spec said and what the code ended up doing
- Merge conflicts on artifacts nobody's reading after implementation
- Noise in `git log` from spec revisions that don't affect the shipped surface

## Force-adding when you really want it versioned

If a spec or plan is exceptionally long-lived (e.g. the initial v1 architecture doc), force-add it explicitly:

```
git add -f docs/superpowers/specs/<file>.md
```

The reference precedent is `confluence-jira-mcp`, which tracks only its initial 2026-03-31 v1 spec+plan and lets everything else stay local. Prefer NOT tracking. The bar for force-add is: *"someone six months from now, without access to my local machine or Claude session transcript, will need to read this to understand the codebase."* Most working specs don't clear that bar.

## Failure mode to avoid

`git status` silently omits files under a gitignored directory. When you write a new file into `docs/superpowers/`, the write returns success but the file will NOT appear in `git status`. If you meant to commit it, use `git add -f` explicitly.

Same trap fired during the 2026-07-08 intent-rewrite planning session — the first commit attempt said "the following paths are ignored by one of your .gitignore files" and required `-f`. That's the guardrail; don't work around it without a reason.
