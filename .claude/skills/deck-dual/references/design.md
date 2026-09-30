# Design defaults (neo-brutalist, sober, academic)

Palette (hex, no `#`): ink `16181D` · paper `F3F0E8` (slide bg) · card `FBF9F3` · accent `0B4F7A` (single colour) · tint `DCE6EE` · muted `4B5159`. Dark statement slide = accent background, paper text.
Fonts: headings Space Grotesk 700 · body IBM Plex Sans 400/600 · labels JetBrains Mono 700 uppercase with 2 px tracking.
Scale (px @1920): 240 cover · 96 big numbers · 72 titles · 56 quote · 40 card headings · 32 body · 28/26 secondary · 24 labels/footnotes (never smaller).
Signature: 4 px ink borders, hard offset shadows (8–10 px, no blur), no rounded corners, 128 px margins, footer progress bar at y=992.
Recurring chrome (animated by Morph): square (big on cover/closing, 72 px elsewhere, rotates 0/15/−15/30°), progress bar, section chip (`NN / TOTAL — SECTION`), logo plate (top right, big on cover/closing).
Images: screenshots in a bordered card with shadow, `object-fit:contain`; logos transparent PNG.
Content rules: one idea per slide; statements/tables/big numbers instead of bullet walls; titles introduce the topic (same grammar throughout); every status also in words (colour never alone); figures only from the source; estimates labelled.
Speaker notes: 2–4 sentences per slide, in the speaker's voice, put in `notes`.
