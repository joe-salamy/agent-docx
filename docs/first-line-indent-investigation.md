# Handoff — First-Line Indent Not Visible in Word (LAW 312 answer.docx)

**Date:** 2026-08-28  
**Reporter:** LAW 312 student (open-law-notes)  
**Artifact:** `C:\Users\joesa\Documents\Law School\2L Fall\LAW 312 Professional Responsibility\open-law-notes\docs\answer.docx` (also `~/code/open-law-notes/docs/answer.docx`, `file:../agent-docx` link)  
**Profile:** `us-district-conventional` (12pt Times, double `480`, `firstLine 720 = 0.5"`)

## Problem

`032da8d feat(layout): default first-line indent 0.5in (720 twips)` set `src/profiles.ts: body: style(240,double,{firstLineIndentTwips:720})`. Deterministic pagination respects it (`lastLineAvailableTwips -720`, `equivalentPages 1.98`), and `word/styles.xml` + `word/document.xml` emit `w:ind firstLine 720`. In Word the doc opens **visibly unformatted** — no indent, and after an open-save cycle most formatting is gone. Reporter indented p2/p3 (2nd/3rd body paras) in Word to show expected (`firstLine 720`).

Current fresh (2026-08-28 17:39, `b536cf…` → `479f06…` after de-dupe) and the last saved Windows copy are **byte-identical** (`11K`), both with `Body style firstLine 0 / direct firstLine 720` — so the save did *not* strip this time, yet reporter still sees no indent. Visual vs XML diverge.

## What Was Tried

### 1. Baseline (032da8d) — style + direct both 720
- `nativeStyles()`: `AgentDocxBody` `ind 720` + each `w:p` direct `paragraphOptions()` also `ind 720` (redundant).
- Fresh: `AgentDocxBody: ind 720, spacing 480, rFonts Times` + `p1-7 direct ind 720`. No `Normal` style, `docDefaults` empty.
- **After Word open-save (2026-08-28 17:05, 17K):** Word injected `Normal w:default` + `latentStyles` + `docDefaults Times`. `Body` collapsed to `spacing 480` only, `p0-2,5-7` lost direct `ind/spacing/sz`, only `p3/p4` (touched) kept `firstLine 720`. 5/7 body paras lost indent → “unformatted”.

### 2. Fix 4c332fb — define Normal
- Added `Normal` (`firstLine 0`, Times 12pt double, `next Normal qFormat`) and `Body basedOn Normal next Normal qFormat`, `ListParagraph` override for `inspect` (`left 720`).
- Fresh now had `Normal` + `Body firstLine 720`. Tests fixed (`ListParagraph` now 720, `1 fail → 0`).
- **After Word save (17:31, 17K):** Worse — `Body` became **empty** (`<w:basedOn Normal/><w:next Normal/><w:qFormat/>` only), `Normal` kept `480+sz24` only, `p0-2,5-7` no direct, `p3/p4` kept `firstLine 720`. All non-touched paras lost indent *and* `Body` lost all `pPr/rPr`.

### 3. Fix 73a992e — de-dupe style vs direct
- `Body` style `firstLine` → `0` (`bodyForStyle`), direct stays `720` via `paragraphOptions()`. Idea: `style 0 vs direct 720` not redundant, Word keeps direct.
- Fresh (17:39, 11K) now: `Body ind 0` style, `p1-7 direct 720`.
- **After Word save (17:39, still 11K, hash 479f06):** This time **no strip** — `Body ind 0` preserved, `p1-7 direct 720` preserved in both WSL and WIN copies (byte-identical). Yet reporter still reports “visibly unformatted” on open. No manual `p3/p4` diff visible (all already 720).

## Current State (2026-08-28 17:39)

```
word/styles.xml:
  Normal: firstLine 0, spacing 480, Times 24
  Body:   basedOn Normal, firstLine 0, spacing 480, Times 24   ← intentionally 0
word/document.xml:
  p0 Heading1  firstLine 0
  p1-7 Body     direct firstLine 720  spacing 480 sz24
```

Both mounts identical, `docs/answer.docx` 2 pages `23/23` lines, `docx/` deleted per request.

## What Failed / Open Questions

1. **Redundancy trigger:** `style 720 + direct 720` → Word strips direct from 5/7 and then strips style. `style 720` alone → Word strips style entirely. `style 0 + direct 720` → Word now keeps both, but still not rendered per user.
2. **Normal injection:** Without `Normal`, Word injects `Calibri` Normal and optimizes; with `Normal` defined, Word still normalizes `rFonts` to `docDefaults` and may hide `qFormat` styles from gallery, but should render.
3. **Visual vs XML:** `w:ind firstLine` is formatting, not `w:t` leading spaces/tabs. If user checks via text extraction or `Normal.dotm` with “Automatically update document styles” enabled, local Normal may override. Group Policy / `Settings → Advanced → Show document content → Show text wrapped within document window` also affects first-line rendering in Draft view.
4. **Direct vs style precedence:** Word’s `pPr` direct should win, but if `w:adjustRightInd`/`w:autoSpaceDE` or compatibility `w:compatSetting w:val="15"` is set, firstLine may be suppressed. Our `settings.xml` is minimal (`updateFields`), no `compat`.
5. **Theme/body vs direct Tabs:** First-line via `w:ind` vs `w:tabs w:val="clear"` — Word may expect a `w:tab` char for first line, not `w:ind`. Unlikely, but legal templates often use `w:ind`.

## Reproduction

```bash
cd ~/code/open-law-notes
node ./node_modules/agent-docx/dist/cli.js measure docs/answer.md --output /tmp/fresh.docx
python3 -c "import zipfile,re; d=re.search(...)"
# compare to Windows copy after Word save:
# diff <(unzip -p /tmp/fresh.docx word/styles.xml) <(unzip -p /mnt/c/.../docs/answer.docx word/styles.xml)
```

## Proposed Next Steps (not yet implemented)

- Try **style-only** (no direct `firstLine`): remove `paragraphOptions` indent for body, keep `720` only in `Body` style (`firstLine 720`). See if Word preserves style when no direct exists.
- Try **w:tabs** instead of `w:ind` for first line, or add `w:ind` via `w:pPrDefault` in `docDefaults`.
- Add `w:semiHidden / w:unhideWhenUsed` false and `w:locked` false, test with `w:customStyle` explicitly (docx lib doesn’t expose, may need raw `ImportedXmlComponent`).
- Test with Word’s “Open and Repair” vs `libreoffice --headless --convert-to pdf` to see if LibreOffice renders indent (isolates Word-specific stripping).
- Check `word/settings.xml` for `w:defaultTabStop 720` — Word’s default tab stop is 0.5", matches `720`; ensure `w:ind` not colliding with `w:tabs`.
- Consider shipping a `.dotx` template instead of generated styles, or using `project init` template path.
