# docex Release Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Release docex as a complete, documented, tested library with a Claude Code skill, ready for public use.

**Architecture:** docex is a zero-dependency Node.js library for programmatic .docx editing. The release bundles the library, CLI, OOXML reference, Claude Code skill, README, and integration with the CampaIgn Tracker paper workflow.

**Tech Stack:** Node.js (CommonJS), node:test, git, Claude Code skills

---

## File Structure

```
/mnt/storage/docex/
  README.md                          # CREATE: User-facing documentation
  LICENSE                            # CREATE: MIT license
  package.json                       # MODIFY: add repository, description, keywords
  src/
    docex.js                         # EXISTS
    workspace.js                     # MODIFY: add safe-modify.sh integration
    xml.js                           # EXISTS
    paragraphs.js                    # EXISTS
    comments.js                      # EXISTS
    figures.js                       # EXISTS
    tables.js                        # EXISTS
    textmap.js                       # EXISTS
    citations.js                     # CREATE: port from inject_citations.js
    latex.js                         # CREATE: port from docx-to-latex.js
  cli/
    docex-cli.js                     # MODIFY: add citations + latex commands, safe-modify flag
  test/
    docex.test.js                    # EXISTS (49 tests)
    integration.test.js              # EXISTS (11 tests)
    citations.test.js                # CREATE: tests for citation injection
  reference/
    ooxml-reference.md               # EXISTS (2,967 lines)
  skill/
    SKILL.md                         # CREATE: the docex Claude Code skill (replaces docx-surgery)

ALSO MODIFY:
  /home/favstats/.claude/skills/docx-surgery/SKILL.md    # Rename to docex, rewrite
  /home/favstats/.claude/skills/revise-and-resubmit/SKILL.md  # Update references
  /home/favstats/.claude/projects/-mnt-storage/memory/    # Update memory files

CLEANUP:
  /mnt/storage/docx-engine/          # DELETE: orphaned scaffolding
```

---

### Task 1: Git Init + LICENSE + README

**Files:**
- Create: `/mnt/storage/docex/README.md`
- Create: `/mnt/storage/docex/LICENSE`
- Create: `/mnt/storage/docex/.gitignore`
- Modify: `/mnt/storage/docex/package.json`

- [ ] **Step 1: Initialize git repo**

```bash
cd /mnt/storage/docex
git init
```

- [ ] **Step 2: Create .gitignore**

```
node_modules/
test/output/
*.bak
.DS_Store
```

- [ ] **Step 3: Create MIT LICENSE**

Standard MIT license, copyright 2026 Fabio Votta / CampaIgn Tracker.

- [ ] **Step 4: Write README.md**

Cover: what docex is (LaTeX for .docx), install, quick start (API + CLI), full API reference, test instructions, OOXML reference pointer.

Show the key API examples:
```js
const docex = require('docex');
const doc = docex("manuscript.docx");
doc.author("Fabio Votta");
doc.replace("old", "new");
doc.after("Methods").insert("New paragraph.");
doc.at("text").comment("Note", { by: "Reviewer 2" });
await doc.save();
```

And CLI:
```bash
docex replace manuscript.docx "old" "new" --author "Fabio"
docex list manuscript.docx headings
```

- [ ] **Step 5: Update package.json**

Add: repository URL, author, description, engines (node >= 18).

- [ ] **Step 6: Commit**

```bash
git add -A && git commit -m "feat: initialize docex repository with README, LICENSE, gitignore"
```

---

### Task 2: Safe-modify.sh Integration

**Files:**
- Modify: `/mnt/storage/docex/src/workspace.js`
- Modify: `/mnt/storage/docex/cli/docex-cli.js`

- [ ] **Step 1: Add safe-modify option to workspace.save()**

Add optional `safeModify` parameter to `save()`. When provided with a path to safe-modify.sh and a description, workspace.save() calls safe-modify.sh instead of direct overwrite.

The workspace save should:
1. Save to a temp file first
2. Call safe-modify.sh with the temp file as the command that copies it over the original

- [ ] **Step 2: Add --safe flag to CLI**

```bash
docex replace manuscript.docx "old" "new" --safe /path/to/safe-modify.sh
```

When --safe is provided, the CLI wraps the save through safe-modify.sh.

- [ ] **Step 3: Test the integration**

Run against a real manuscript with safe-modify.sh and verify git commit + backup are created.

- [ ] **Step 4: Commit**

```bash
git add -A && git commit -m "feat: integrate safe-modify.sh for manuscript protection"
```

---

### Task 3: Port Citations Module

**Files:**
- Create: `/mnt/storage/docex/src/citations.js`
- Modify: `/mnt/storage/docex/src/docex.js` (add citations API)
- Modify: `/mnt/storage/docex/cli/docex-cli.js` (add cite command)
- Create: `/mnt/storage/docex/test/citations.test.js`
- Source: `/mnt/storage/nl_local_2026/paper/build/inject_citations.js` (567 lines)

- [ ] **Step 1: Read and understand inject_citations.js**

Read the existing 567-line script. Understand how it:
- Finds (Author, Year) patterns in text
- Queries Zotero API to match citations
- Builds OOXML field codes for Zotero integration

- [ ] **Step 2: Write failing test**

Test that `Citations.inject(ws)` finds citation patterns and replaces them with ZOTERO_CITATION fields.

- [ ] **Step 3: Port to Citations class**

Static methods: `Citations.list(ws)`, `Citations.inject(ws, options)`.

- [ ] **Step 4: Add to docex API**

```js
doc.citations.inject();  // find and replace plain-text citations
doc.citations.list();    // list current citations
```

- [ ] **Step 5: Add CLI command**

```bash
docex cite manuscript.docx                    # inject citations
docex cite manuscript.docx --list             # list citations
```

- [ ] **Step 6: Run tests, commit**

---

### Task 4: Port LaTeX Conversion Module

**Files:**
- Create: `/mnt/storage/docex/src/latex.js`
- Modify: `/mnt/storage/docex/cli/docex-cli.js` (add latex command)
- Source: `/mnt/storage/nl_local_2026/paper/build/docx-to-latex.js` (1,735 lines)

- [ ] **Step 1: Read and understand docx-to-latex.js**

Read the existing 1,735-line script. Understand how it converts OOXML to LaTeX.

- [ ] **Step 2: Port to Latex class**

Static methods: `Latex.convert(ws, options)` returns LaTeX string.

- [ ] **Step 3: Add CLI command**

```bash
docex latex manuscript.docx                   # output .tex file
docex latex manuscript.docx --output paper.tex
```

- [ ] **Step 4: Test, commit**

---

### Task 5: Rewrite docx-surgery Skill as docex Skill

**Files:**
- Create: `/mnt/storage/docex/skill/SKILL.md` (canonical copy)
- Modify: `/home/favstats/.claude/skills/docx-surgery/SKILL.md` (symlink or replace)

- [ ] **Step 1: Write the docex skill**

The skill should:
- Be named "docex" (not "docx-surgery")
- Document when to use it (same triggers as docx-surgery)
- Show the full API with examples
- Include the CLI reference
- Reference the OOXML manual
- Keep all safety rules (safe-modify.sh, never regenerate, never loop)
- Include the R&R workflow (from revise-and-resubmit)
- Include comment reply patterns
- Be comprehensive enough to replace BOTH docx-surgery and revise-and-resubmit

- [ ] **Step 2: Replace the old docx-surgery skill**

Either symlink or copy: `/home/favstats/.claude/skills/docx-surgery/SKILL.md` -> docex skill

- [ ] **Step 3: Update revise-and-resubmit skill**

Replace references to suggest-edit.js and docx-patch.js with docex API calls. Reference the docex skill for OOXML patterns.

- [ ] **Step 4: Commit**

---

### Task 6: Update Memory + Cleanup

**Files:**
- Modify: `/home/favstats/.claude/projects/-mnt-storage/memory/MEMORY.md`
- Create: `/home/favstats/.claude/projects/-mnt-storage/memory/docex_library.md`
- Modify: `/home/favstats/.claude/projects/-mnt-storage/memory/feedback_use_docx_skills.md`
- Modify: `/home/favstats/.claude/projects/-mnt-storage/memory/academic_writing_stack.md`
- Delete: `/mnt/storage/docx-engine/` (orphaned)

- [ ] **Step 1: Create docex memory file**

Record: what docex is, where it lives, its API, test status, CLI location.

- [ ] **Step 2: Update feedback memory**

Update feedback_use_docx_skills.md to reference docex instead of docx-surgery.

- [ ] **Step 3: Update academic writing stack memory**

Add docex to the writing infrastructure reference.

- [ ] **Step 4: Update MEMORY.md index**

Add docex_library.md entry.

- [ ] **Step 5: Delete orphaned docx-engine/**

```bash
rm -rf /mnt/storage/docx-engine
```

- [ ] **Step 6: Final test run**

```bash
cd /mnt/storage/docex
node --test test/docex.test.js
node --test test/integration.test.js
```

Verify all 60 tests still pass.

- [ ] **Step 7: Final git commit**

```bash
git add -A && git commit -m "chore: docex v0.1.0 release"
```

---

## Execution Order

Tasks 1-2 are prerequisites for everything else.
Tasks 3-4 are independent (can run in parallel).
Task 5 depends on 3-4 (skill should document citations + latex).
Task 6 is the final cleanup.

```
Task 1 (git + readme) ──┐
Task 2 (safe-modify)  ──┼── Task 3 (citations) ──┐
                        │   Task 4 (latex)     ──┼── Task 5 (skill) ── Task 6 (cleanup)
                        └────────────────────────┘
```
