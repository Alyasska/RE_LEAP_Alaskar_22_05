# UP01 -- Inspect LEAP API via the built-in Script Editor

**Purpose:** Discover the actual method signatures available on the LEAP COM API for this install. SEI's online docs are incomplete; the Script Editor is the authoritative source.

**When to do this:** when S03 probe is inconclusive, or when we need the signature for a method that S03 didn't enumerate (Save, Refresh, Export, anything else).

**Time:** ~3-5 minutes.

## Steps

1. Open LEAP
2. Open the KAZ_2024 area (index 6)
3. Menu: **Advanced** -> **Edit Scripts** (or **Advanced** -> **Script Editor**, name varies by version)
4. The Script Editor window opens. It has a code-edit pane on the left and an API browser pane on the right.
5. In the API browser pane (right side):
   - Expand `LEAP` (or `LEAPApplication`) at the root
   - Scroll down to find the `Calculate` method (or any method of interest)
   - **Click on the method name once.** Below the tree, the help pane shows:
     - Whether it's a method or property
     - Return type
     - List and types of required parameters
     - A one-line description
6. **Take a screenshot of this help info.** Save to `docs/screenshots/UP01_<method>_api.png` (committed).
7. Repeat for each method of interest. For cycle 006, capture at least:
   - `Calculate`
   - `Save`
   - `Refresh`
   - `Areas` (collection, see how it's indexed)
   - `BranchVariable` (we use this everywhere; confirm signature)

## Alternative: autocomplete probe in the script editor

1. In the code-edit pane, type `LEAP.` (with trailing period)
2. After a moment, an autocomplete dropdown lists all methods and properties
3. Hover over each item to see its tooltip (parameter list)
4. To capture: hover, screenshot, save

## What to report back

Just the screenshot(s), or the verbatim text from the help pane for each method we asked about. We do not need a perfect catalog -- one signature for Calculate unblocks cycle 006.

## Why this is "OK" by project rules

Project rule 4 says audit/scout work is read-only. Reading the Script Editor's API documentation does not modify the model. It is the equivalent of opening a manual.
