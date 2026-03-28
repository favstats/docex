---
name: improve-loop
description: Continuously improve and test the project for a specified duration without stopping or asking questions
---

# Improve Loop

Run continuous improvement cycles for a specified duration. Do NOT stop to ask questions. Do NOT wait for confirmation. Just keep working.

## Usage
```
/improve-loop 5h    # run for 5 hours
/improve-loop 2h    # run for 2 hours
/improve-loop 30m   # run for 30 minutes
```

## Behavior

1. **Never stop to ask the user anything.** Make decisions autonomously.
2. **Work in cycles:** audit → fix → test → commit → next issue.
3. **Prioritize by impact:** critical bugs > visual polish > new features > docs.
4. **Run tests after every change.** If tests break, fix them before moving on.
5. **Commit and push frequently.** Small, focused commits.
6. **Use parallel agents** for independent tasks.
7. **Do visual regression checks** with browser screenshots after UI changes.
8. **Log what you did** at the end of each cycle so the user can review.

## What to improve (in priority order)

1. **Bugs and broken things** — anything that doesn't work correctly
2. **Visual polish** — alignment, spacing, consistency, responsiveness
3. **Missing features** — gaps in functionality
4. **Performance** — speed, bundle size, loading time
5. **Accessibility** — ARIA, keyboard nav, contrast
6. **Documentation** — README, API docs, examples
7. **Tests** — coverage gaps, edge cases
8. **Code quality** — dead code, inconsistencies, simplification

## Rules

- If unsure between two approaches, pick the simpler one.
- If a change is risky, use a git worktree or branch.
- If something is working, don't touch it.
- Prefer editing existing files over creating new ones.
- Keep commits small and descriptive.
