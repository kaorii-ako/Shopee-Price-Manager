---
name: git-commit
description: Write clear, conventional commit messages from staged changes
license: MIT
compatibility: opencode
metadata:
  audience: developers
  workflow: github
---

## What I do
- Inspect staged changes (`git diff --cached`) to understand what changed and why
- Draft a commit message following the Conventional Commits spec (`feat:`, `fix:`, `chore:`, etc.)
- Include a concise subject line (≤72 chars) and an optional body for non-trivial changes
- Flag if the staged diff spans unrelated concerns and suggest splitting into multiple commits

## When to use me
Use this when you have staged changes and want a well-formed commit message ready to go.
Ask clarifying questions if the intent behind the changes is ambiguous or the scope is unclear.