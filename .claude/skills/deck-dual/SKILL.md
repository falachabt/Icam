---
name: deck-dual
description: Build a presentation in TWO identical versions from one spec — an online Slides artifact (web) and a native .pptx with real Morph transitions — in a clean neo-brutalist academic style. Use when the user wants a soutenance/pitch/report deck, says "PowerPoint avec morph", "version web et PPTX", or wants animated transitions that survive in PowerPoint (the Slides export drops animations).
---

# deck-dual — one spec, two decks (web + native PPTX with Morph)

Why this exists: exporting the web Slides deck to .pptx keeps the design but **loses the animations**. This skill keeps the two versions in sync by placing every element once, in 1920×1080 px coordinates, and emitting it to both formats. The PPTX gets real `<p159:morph>` transitions.

## Workflow

1. **Read the source** (report, brief). Write a `spec.json` (see `examples/riptis.v1.spec.json`). One idea per slide, real numbers only — never invent statistics; label estimates as estimates. Ask the user only what the source cannot answer.
2. **Design**: defaults are in `references/design.md` (palette, fonts, rules). Override under `spec.theme`. Keep it sober: one accent colour, no decorative icons unless the user asks (they read as "AI-made"; this user explicitly removed them).
3. **Assets**: put screenshots/logos in an assets folder. Crop from a PDF with `pdftoppm -r 200` + Pillow; make logos transparent by thresholding near-white. Stack/brand logos: `node scripts/fetch-brand-logos.js assets slug:HEX ...` (Simple Icons). **Redact or mask sensitive data before using a screenshot; never use an unmasked one.**
4. **Build**: `node scripts/build.js --spec spec.json --assets assets --out out [--urls urls.json]`
   - `out/project/deck.json` + `out/project/slides/*.html` → the web deck
   - `out/<filename>.pptx` → the native deck with Morph on slides 2..N
   - needs `pptxgenjs` (brings `jszip`); PPTX icons additionally need `react react-dom react-icons sharp`.
5. **Web**: create the artifact from the Slides type (Artifact tool, `type_url`, follow the instructions it returns), upload the images with `asset:true`, write their urls to `urls.json` (`{"file.png":"/_blob/<id>"}`), rebuild, then publish `deck.json` first + slides (see the type's instructions).
6. **PPTX**: validate with the pptx skill's `scripts/office/validate.py`, hand over the file (SendUserFile). Variants (v1/v2…) = separate spec files sharing assets.
7. Tell the user honestly what was **not** visually checked (LibreOffice often cannot render in the sandbox): they confirm Morph by running the slideshow in PowerPoint.

## How Morph is made (the part that matters)

- PPTX = zip of XML. pptxgenjs has no transitions, so `build.js` post-processes each slide ≥2 and inserts, right after `</p:clrMapOvr>`, an `mc:AlternateContent` with `<p159:morph option="byObject"/>` (Fallback = fade for old PowerPoint). The transition belongs to the slide you **arrive at**.
- Morph pairs objects that carry the **same name**; a `!!` prefix forces the pairing. `build.js` names the recurring chrome `!!mk-sq`, `!!mk-bar`, `!!mk-chip`, `!!mk-logos`, `!!mk-logo-N` and moves/resizes/rotates them per slide (big on cover/closing, small elsewhere; progress bar grows). Do not rename them in PowerPoint.
- Web version: same idea with `data-transition="magic"` on the section and the same `id` on matching pinned elements (the Slides type calls it magic move).

## Layouts (in `scripts/build.js`, `LAYOUT`)

`cover` · `cards` (2–4 cards) · `stats` (2–3 big numbers) · `statement` (dark quote) · `matrix` (weighted multicriteria matrix or plain table; `totalRow`, `boldCols`) · `twocols` (bullets | kv | paras; kv values may be `[{n,logo}]` = stack logos) · `image_list` · `images2` · `image_facts` · `callout` (2 cards + bar) · `kpi` (2×2) · `rows` (numbered; item `{t,badge}` = highlighted "validated" row) · `closing`.
Add a layout = one function that calls the primitives (`card`, `text`, `img`, `icon`, `pic`, `title`, `label`); both formats follow automatically.

## Pitfalls learned (keep)

- Text fits are estimates (~0.55×font-size per char). Titles at 72 px hold ~38 chars per line: keep them one line, or set `titleLines:2` where the layout supports it. Big-number values ≤ ~8 chars at 96 px per card.
- pptxgenjs: hex colours without `#`; no shared option objects; `charSpacing` not `letterSpacing`; shadow offset ≥ 0; `isTextBox:true`; px→pt is ÷2 at 1920 px = 13.333 in.
- PPTX fonts are not embedded: the user must install the Google Fonts used (Space Grotesk, IBM Plex Sans, JetBrains Mono) or pick Arial/Calibri.
- Web HTML subset: inline styles only, no margin/class/em; ≥24 px text; everything here is `position:absolute` on the section (no per-section limit documented; ≤24 pinned children per **div**).
- Table cells cannot be filled in the web subset (only `tr` background) — the PPTX mirrors that (total row tinted, key column bold).
- Don't invent data for matrices: say in the notes/footnote that scores and weights are the author's judgement, and let the user edit them.
- Never use `localStorage`-style tricks, background processes, or claim something was verified if it was only schema-validated.
