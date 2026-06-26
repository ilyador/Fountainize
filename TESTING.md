# Testing Fountainize

`Tests.js` is a test suite that runs **inside Google Apps Script against a real
Google Doc** — it creates a scratch document, runs the real `convert()` on
Fountain-format input, then reads back the *rendered* formatting (indent, bold,
font, heading, alignment, text) and asserts it. Nothing is mocked.

## How to run it

1. Open the script project: from any Google Doc choose **Extensions → Apps
   Script**, or open the project directly at <https://script.google.com>.
2. The tests create a scratch Doc, so the runner needs the broader **documents**
   scope. In `appsscript.json` add `"https://www.googleapis.com/auth/documents"`
   to `oauthScopes`:
   ```json
   "oauthScopes": [
     "https://www.googleapis.com/auth/documents.currentonly",
     "https://www.googleapis.com/auth/documents",
     "https://www.googleapis.com/auth/script.container.ui",
     "https://www.googleapis.com/auth/script.external_request"
   ]
   ```
   (The add-on itself only needs `documents.currentonly`; this extra scope is for
   the test runner, not the published add-on.)
3. In the editor toolbar, pick **`runFountainTests`** from the function dropdown
   and click **Run**.
4. Approve the authorization prompt the first time (Advanced → Go to project → Allow).
5. Read the **PASS / FAIL report** in the **Execution log**.
6. Open the **"Fountainize Test Sample"** doc it leaves in your Drive to visually
   confirm the rendered screenplay reads correctly.

That's it — no command line and no external tooling required.

## Fountain features covered

### Implemented (asserted against the real rendered output)

| Feature | What's checked |
| --- | --- |
| Scene heading `INT.` / `EXT.` | styled as a scene, uppercased, **bold**, Courier Prime, tagged **Heading 3** (shows in the outline) |
| Combined scene heading `INT./EXT.`, `EXT/INT.`, `I/E.`, `EST.` | recognized as a scene, full prefix preserved (not truncated) |
| Scene numbers (option) | heading prefixed with the scene number |
| Action | default paragraph styling |
| Character cue | uppercase line styled as a character |
| Character extension `(V.O.)` / `(CONT'D)` | cue with a bracketed extension still a character |
| Character cue with number/symbol `GUARD #1` | still a character |
| Character auto-detection | detected name added to the shortcut list |
| Dialogue | line after a character/parenthetical styled as dialogue |
| Parenthetical `(beat)` | whole-line `(...)` styled as a parenthetical |
| Inline parenthetical `(beat) line` | styled as dialogue; the following action line is **not** swallowed |
| Transition `… TO:` / `IN:` / `OUT:` | styled as a transition, right-aligned |
| Centered text `>…<` | brackets stripped, centered |
| Spacing | paragraph **space-before / space-after margins** (no blank lines): scene/action/transition/centered 12pt both sides; character 12 before / 6 after; dialogue 0 / 6; parenthetical 0 / 0; zero blank paragraphs |
| Idempotency | re-running Format Script produces the identical result (no extra lines/spaces/headers) |
| Format selection / block | `formatSelection()` styles only the selected paragraphs (or the one under the cursor) and leaves the rest untouched |

### Not implemented yet (flagged by the suite, with current behavior shown)

These are Fountain spec features Fountainize does **not** support; the suite lists
each one and what Fountainize currently does with it, so the gap is visible:

- Forced Scene Heading — `.SNOWGLOBE`
- Forced Action — `!CAPS LINE`
- Forced Character — `@mcclane`
- Forced Transition — `> SMASH CUT`
- Dual Dialogue — `CHARACTER ^`
- Lyrics — `~la la la`
- Emphasis — `*italic*`, `**bold**`, `_underline_`
- Page Break — `===`
- Notes — `[[ note ]]`
- Boneyard (comments) — `/* cut */`
- Sections — `# Act One`
- Synopses — `= synopsis`
- Title Page markup — `Title:` / `Author:` (Fountainize uses the sidebar form instead)
