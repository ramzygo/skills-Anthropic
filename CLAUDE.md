# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repo Is

A collection of **Agent Skills** for Claude — self-contained instruction folders that Claude loads dynamically to improve performance on specialized tasks. The repo is also registered as a Claude Code Plugin marketplace.

## Installing Skills in Claude Code

```bash
# Register the marketplace
/plugin marketplace add anthropics/skills

# Install a plugin bundle
/plugin install document-skills@anthropic-agent-skills
/plugin install example-skills@anthropic-agent-skills
```

## Skill Structure

Every skill is a folder under `skills/` with this layout:

```
skill-name/
├── SKILL.md          # Required: YAML frontmatter + markdown instructions
├── scripts/          # Optional: Python/JS tools Claude can invoke
├── references/       # Optional: docs loaded on demand
├── assets/           # Optional: templates, fonts, icons
└── templates/        # Optional: boilerplate code
```

`SKILL.md` frontmatter requires only two fields:

```yaml
---
name: skill-name          # unique, lowercase, hyphens
description: ...          # primary trigger mechanism — what it does AND when to use it
---
```

Claude loads skill content in three tiers:
1. **Metadata** (name + description) — always loaded, used for trigger matching
2. **SKILL.md body** — loaded when the skill is triggered (keep under ~500 lines)
3. **Bundled resources** — loaded on demand as needed

## Plugin Bundles

Defined in `.claude-plugin/marketplace.json`:
- **`document-skills`** — `xlsx`, `docx`, `pptx`, `pdf` (source-available, not Apache 2.0)
- **`example-skills`** — 12 creative/technical/enterprise skills (Apache 2.0)
- **`claude-api`** — Claude API/SDK documentation skill (Apache 2.0)

## Skill Evaluation Tooling (skill-creator)

The `skills/skill-creator/` skill ships Python scripts for testing and optimizing skills:

```bash
cd skills/skill-creator/scripts

# Run evals against a skill
python run_eval.py

# Run the iterative improvement loop
python run_loop.py

# Optimize the skill's description for better trigger accuracy
python improve_description.py

# Package a skill into a .skill file
python package_skill.py

# Quick validation check
python quick_validate.py

# Aggregate benchmark results
python aggregate_benchmark.py

# Generate an HTML report from eval results
python generate_report.py
```

The `eval-viewer/generate_review.py` script produces an HTML page for reviewing eval outputs side-by-side.

## Writing Effective Skill Descriptions

The `description` field is the sole trigger mechanism — Claude uses it to decide when to activate a skill. Write descriptions that are slightly "pushy": list specific phrases, contexts, and synonyms that should trigger the skill, not just what it does. All "when to use" guidance belongs in the description, not the body.

## Licensing

Most example skills are Apache 2.0. The four document skills (`docx`, `pdf`, `pptx`, `xlsx`) are source-available under their own license — usable as a reference but not open source.

## Key Reference Files

- `spec/agent-skills-spec.md` — points to the full spec at agentskills.io/specification
- `template/SKILL.md` — minimal starting template for new skills
- `skills/skill-creator/SKILL.md` — detailed guide for the full skill creation and eval workflow
