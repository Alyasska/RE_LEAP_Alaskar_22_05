# workflow.md — How we work together

## Roles

| | Aliaskar (you) | Claude (me) |
|---|---|---|
| Runs on | Windows + VS Code + LEAP GUI | Claude.ai chat |
| Owns | Repo, .leap files, ground-truth data | Diagnoses, scripts, docs |
| Workflow | git pull → run → test → push | Read repo → diagnose → write files for you |

## The cycle (15-30 min target)

```
┌────────────── ALIASKAR ──────────────┐    ┌──────────── CLAUDE ────────────┐
│                                       │    │                                  │
│  1. git pull                          │    │                                  │
│     cat HANDOFF.md (resume context)   │    │                                  │
│                                       │    │                                  │
│  2. Run today's script:               │    │                                  │
│     .\scripts\NN_x\Sxx_thing.ps1      │    │                                  │
│                                       │    │                                  │
│  3. Open LEAP, test in UI             │    │                                  │
│                                       │    │                                  │
│  4. Write logs\cycle_NNN.md:          │    │                                  │
│     - what you ran                    │    │                                  │
│     - what happened                   │    │                                  │
│     - any errors (copy-paste exact)   │    │                                  │
│                                       │    │                                  │
│  5. git add . && git commit -m        │    │                                  │
│     "cycle NNN: <one line summary>"   │    │                                  │
│     && git push                       │    │                                  │
│                                       │    │                                  │
│  6. In chat: "cycle NNN done, see     │ ─▶ │                                  │
│     log + (uploaded .leap if needed)" │    │  Reads cycle_NNN.md from repo    │
│                                       │    │  Reads any uploaded files        │
│                                       │    │  Diagnoses                       │
│                                       │    │  Writes:                         │
│                                       │    │   - updated HANDOFF.md           │
│                                       │    │   - next script(s)               │
│                                       │    │   - issue entries if new bug     │
│                                       │ ◀─ │  Returns files + instructions    │
│                                       │    │                                  │
│  7. Download files from chat,         │    │                                  │
│     drop into repo, commit            │    │                                  │
│                                       │    │                                  │
│  → loop                               │    │                                  │
└───────────────────────────────────────┘    └──────────────────────────────────┘
```

## What I (Claude) need from you each cycle

**Always:**
- The latest `cycle_NNN.md` log file (in repo or pasted in chat)
- A short narrative: "I tried X, got Y, expected Z"

**When something broke:**
- Exact error text — screenshot AND text. LEAP errors are often truncated visually; the underlying message has more detail.
- The `.leap` file if the model state matters for diagnosis (upload it directly in chat)
- The LEAP `Diagnostics` view output if available (Tools → Diagnostics)

**When asking for next step:**
- One of:
  - "what's next per the plan" → I'll point to PROTOTYPE_PLAN.md
  - "the script you gave failed" → I'll read the log and revise
  - "I want to try X instead" → I'll evaluate and either agree or push back with reasoning
  - "I'm blocked because Y" → I'll diagnose and produce a fix

## What you'll get back from me

| Artifact | When | Where |
|---|---|---|
| Updated `HANDOFF.md` | Every cycle | I'll output, you commit |
| New `scripts/Nx/Sxx_*.ps1` or `.vbs` | When we need new code | Same |
| Updated `docs/known_issues.md` | When we diagnose a new LEAP bug | Same |
| New `docs/ui_procedures/UPxx_*.md` | When something must be done in UI | Same |
| Diagnoses in chat | Always | In chat directly |

## Naming conventions

**Scripts:** `<tier-letter><sequence>_<action>.<ext>`
- `S01_extract_leap.py` — first scout script
- `R03_disable_transmission.ps1` — third reduce script
- `F01_remove_broken_loadings.vbs` — first fix script
- `X01_calculate.vbs` — first run/execute script

**Cycle logs:** `logs/cycle_001.md`, `logs/cycle_002.md`, ...

**Snapshots:** `data/snapshots/cycle_NNN_<description>.leap`
- e.g. `cycle_003_post_transmission_disable.leap`

**UI procedures:** `docs/ui_procedures/UP01_delete_loading_row.md`

## Commit message format

```
cycle NNN: <action verb> <object> [<status>]

Examples:
cycle 001: scout colleague's .leap, audit complete
cycle 003: disable transmission simulation, nodal error gone
cycle 004: remove 5 broken loading rows, calc reaches transformation
cycle 006: first successful Calculate, results display 🎉
```

## What I cannot do

- Run scripts on your Windows machine — only you can do that
- Read your local files (anything not in the GitHub repo or uploaded to chat)
- Click in your LEAP UI — manual procedures must be done by you
- Guarantee my scripts work first time — LEAP COM is finicky and the model is messy. Expect 1-3 revisions per script.

## What I can do

- Decrypt and inspect any `.leap` file you upload
- Write PowerShell, VBScript, Python scripts using your proven patterns
- Diagnose LEAP errors when given exact error text + context
- Maintain the project's living memory (HANDOFF.md, known_issues.md)
- Search the LEAP help and SEI forums when stuck
- Refactor docs and code when patterns emerge
