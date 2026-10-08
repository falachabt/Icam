#!/usr/bin/env node
/*
 deck-dual build: one spec.json  ->  (1) web Slides files (project/deck.json + project/slides/*.html)
                                     (2) native .pptx with real Morph transitions.
 Every element is placed ONCE in 1920x1080 px coordinates and emitted to both formats, so the
 two versions stay identical.

 usage: node build.js --spec spec.json --assets ./img --out ./out [--urls urls.json]
   --assets  folder holding the local image files named in the spec (used by the PPTX, and by icons)
   --urls    JSON {"file.png": "/_blob/<id>"} = the uploaded asset urls (web version). Optional until you publish.
 deps: pptxgenjs (+ jszip, shipped with it). Icons in the PPTX also need react, react-dom, react-icons, sharp.
*/
const fs = require('fs'), path = require('path');
const A = (() => { const a = process.argv.slice(2), o = {}; for (let i = 0; i < a.length; i += 2) o[a[i].replace(/^--/, '')] = a[i + 1]; return o; })();
if (!A.spec || !A.out) { console.error('usage: node build.js --spec spec.json --assets dir --out dir [--urls urls.json]'); process.exit(1); }
const spec = JSON.parse(fs.readFileSync(A.spec, 'utf8'));
const ASSETS = A.assets || path.dirname(A.spec);
const URLS = A.urls && fs.existsSync(A.urls) ? JSON.parse(fs.readFileSync(A.urls, 'utf8')) : {};
const OUT = A.out; fs.mkdirSync(path.join(OUT, 'project/slides'), { recursive: true });
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

// ---------- theme ----------
const th = Object.assign({ ink: '16181D', paper: 'F3F0E8', card: 'FBF9F3', accent: '0B4F7A', tint: 'DCE6EE', muted: '4B5159' }, spec.theme || {});
const FN = Object.assign({ H: 'Space Grotesk', B: 'IBM Plex Sans', M: 'JetBrains Mono' }, (spec.theme || {}).fonts || {});
const FCSS = { H: `'${FN.H}', Arial, sans-serif`, B: `'${FN.B}', Arial, sans-serif`, M: `'${FN.M}', 'Courier New', monospace` };
const gf = (n, w) => `https://fonts.googleapis.com/css2?family=${n.replace(/ /g, '+')}:wght@${w}&display=swap`;
const faces = {}; [['H', '500..700'], ['B', '400;600'], ['M', '500;700']].forEach(([k, w]) => { faces[FN[k].toLowerCase().replace(/ /g, '-')] = { family: FN[k], href: gf(FN[k], w) }; });
// 'showcase' = ONE screen that stays while screenshots + feature list change step by step (Morph carousel). Expanded into N slides.
function expand(list) {
  const out = [];
  list.forEach(s => {
    if (s.layout !== 'showcase') return out.push(s);
    const d = s.data, groups = d.groups, nShots = d.shots.length, rot = [0, 15, -15, 0, 30, -15, 15];
    const withShot = g => g.items.map((it, i) => ({ ...it, g: groups.indexOf(g), i }));
    const all = groups.flatMap(withShot);
    for (let k = 0; k <= nShots; k++) { // steps 0..nShots-1 = one screenshot each; last step = items WITHOUT a screenshot, all at once
      const final = k === nShots, active = final ? 0 : k;
      const vis = all.filter(it => final || (it.shot != null && it.shot <= k));
      const fresh = new Set(all.filter(it => final ? it.shot == null : it.shot === k).map(it => it.g + ':' + it.i));
      out.push({ id: s.id + (k + 1), label: s.label, rot: rot[k % rot.length], layout: 'showcase_step',
        notes: final ? d.finalNotes : d.shots[k].notes, data: { title: d.title, shots: d.shots, active, groups: groups.map((g, gi) => ({ label: g.label, items: vis.filter(it => it.g === gi).map(it => ({ t: it.t, fresh: fresh.has(it.g + ':' + it.i), key: gi + '_' + it.i })) })) } });
    }
  });
  return out;
}
const slides = expand(spec.slides), TOTAL = slides.length;
const ICON = { Activity: 'LuActivity', Book: 'LuBook', Chart: 'LuChartColumn', Chat: 'LuMessageSquare', Check: 'LuCheck', CheckCircle: 'LuCircleCheck', Clock: 'LuClock', Cloud: 'LuCloud', Code: 'LuCode', Database: 'LuDatabase', Globe: 'LuGlobe', GraduationCap: 'LuGraduationCap', Home: 'LuHouse', Key: 'LuKey', Lightbulb: 'LuLightbulb', Lightning: 'LuZap', Link: 'LuLink', Lock: 'LuLock', PaperPlane: 'LuSend', Play: 'LuPlay', Search: 'LuSearch', Settings: 'LuSettings', Star: 'LuStar', ThumbsUp: 'LuThumbsUp', Tool: 'LuWrench', Trust: 'LuShieldCheck', Users: 'LuUsers', Verified: 'LuBadgeCheck', Warning: 'LuTriangleAlert', Wrench: 'LuWrench' };

// ---------- primitives (px, 1920x1080) ----------
function maker(P, fg, dark) {
  const o = {
    rect: (x, y, w, h, q = {}) => P.push(Object.assign({ t: 'rect', x, y, w, h, fill: th.card, line: th.ink, bw: 4, so: 0, rot: 0 }, q)),
    card: (x, y, w, h, q = {}) => o.rect(x, y, w, h, Object.assign({ so: 10 }, q)),
    text: (x, y, w, h, text, q = {}) => P.push(Object.assign({ t: 'text', x, y, w, h, text, size: 32, font: 'B', color: fg, lh: 1.25, align: 'left', valign: 'top' }, q)),
    img: (x, y, w, h, file, alt, q = {}) => P.push(Object.assign({ t: 'img', x, y, w, h, file, alt }, q)),
    icon: (x, y, s, name, color) => P.push({ t: 'icon', x, y, s, name, color: color || th.accent }),
    pic: (x, y, w, h, file, alt) => { o.card(x, y, w, h); o.img(x + 4, y + 4, w - 8, h - 8, file, alt); },
    title: (t, lines = 1) => o.text(128, 128, 1664, lines === 2 ? 160 : 90, t, { font: 'H', size: 72, bold: true, lh: 1.05, tag: 'h2', color: dark ? th.paper : th.ink }),
    label: (x, y, w, t, color) => o.text(x, y, w, 34, t, { font: 'M', size: 24, bold: true, color: color || th.accent, spacing: 2, upper: true }),
  };
  return o;
}
function cardHead(o, x, y, w, c, fg) { // optional mono label + bold heading + icon at top-right
  if (c.icon) o.icon(x + w - 36 - 48, y + 34, 48, c.icon, fg === th.paper ? th.paper : th.accent);
}

// ---------- layouts: each one calls the primitives; edit geometry here once, both formats follow ----------
const LAYOUT = {
  cover(o, d) {
    o.text(128, 128, 1040, 40, d.kicker || '', { font: 'M', size: 28, bold: true, color: th.accent, spacing: 2, upper: true });
    o.text(128, 330, 1040, 250, d.title, { font: 'H', size: 240, bold: true, lh: 1, tag: 'h1' });
    o.text(128, 600, 1000, 130, d.subtitle || '', { size: 48, bold: true, lh: 1.15 });
    o.text(128, 790, 1040, 48, d.author || '', { size: 36, bold: true });
    (d.lines || []).forEach((l, i) => o.text(128, 840 + i * 37, 1040, 36, l, { size: 28, color: th.muted }));
  },
  cards(o, d) { // 2-4 cards in a row: label/number, heading, text, icon
    o.title(d.title);
    const n = d.cards.length, gap = 32, w = Math.floor((1664 - gap * (n - 1)) / n);
    d.cards.forEach((c, i) => { const x = 128 + i * (w + gap); o.card(x, 300, w, 340);
      o.text(x + 36, 336, w - 140, 40, c.label || '', { font: 'M', size: 28, bold: true, color: th.accent });
      o.text(x + 36, 392, w - 72, 56, c.title, { font: 'H', size: 40, bold: true, lh: 1.15, tag: 'h3' });
      o.text(x + 36, 470, w - 72, 150, c.text || '', { size: 28, color: th.muted, lh: 1.3 });
      cardHead(o, x, 300, w, c); if (c.to) o.P.push({ t: 'nav', x, y: 300, w, h: 340, to: c.to }); }); // c.to = slide id the card jumps to (PPTX click)
  },
  stats(o, d) { // 2-3 big-number cards + optional note
    const two = d.titleLines === 2; o.title(d.title, two ? 2 : 1);
    const n = d.stats.length, gap = 40, w = Math.floor((1664 - gap * (n - 1)) / n), y = two ? 340 : 250;
    d.stats.forEach((c, i) => { const x = 128 + i * (w + gap); o.card(x, y, w, 350);
      o.text(x + 36, y + 26, w - 150, 120, c.value, { font: 'H', size: 96, bold: true, lh: 1.1, color: th.accent });
      o.text(x + 36, y + 160, w - 72, 170, c.text, { size: 32, lh: 1.3 });
      if (c.icon) o.icon(x + w - 36 - 56, y + 36, 56, c.icon); });
    if (d.note) o.text(128, y + 400, 1664, 100, d.note, { size: 32, color: th.muted, lh: 1.3 });
  },
  statement(o, d) { // dark slide: big quote + tags
    o.label(128, 128, 900, d.label, th.paper);
    o.text(128, 200, 1560, 470, d.quote, { font: 'H', size: 56, bold: true, lh: 1.2 });
    const n = d.tags.length, gap = 40, w = Math.floor((1664 - gap * (n - 1)) / n);
    d.tags.forEach((t, i) => { const x = 128 + i * (w + gap); o.card(x, 720, w, 130); o.text(x + 36, 720, w - 72, 130, t, { font: 'H', size: 40, bold: true, color: th.ink, valign: 'middle' }); });
  },
  matrix(o, d) { // weighted multicriteria matrix OR plain table. cols:[{t,w,align}], rows:[[...]], totalRow:true, boldCols:[i]
    o.title(d.title);
    const size = d.size || 28, headH = d.headH || 84, rowH = d.rowH || 64, sum = d.cols.reduce((a, c) => a + c.w, 0);
    const colW = d.cols.map(c => Math.round(1664 * c.w / sum));
    P_table(o, { x: 128, y: 250, w: 1664, colW, head: d.cols.map(c => c.t), align: d.cols.map(c => c.align || 'left'), rows: d.rows, size, headH, rowH, totalRow: !!d.totalRow, boldCols: d.boldCols || [] });
    let y = 250 + headH + rowH * d.rows.length + 40;
    if (d.conclusion) { o.text(128, y, 1664, 50, d.conclusion, { font: 'H', size: 32, bold: true }); y += 60; }
    if (d.note) o.text(128, y, 1664, 36, d.note, { size: 24, color: th.muted });
  },
  twocols(o, d) { // two cards side by side; each {title, icon, dark, bullets[] | kv[[k,v]] | paras[{t,bold}]}
    o.title(d.title);
    [d.left, d.right].forEach((c, i) => { const x = 128 + i * 852, dk = !!c.dark, fg = dk ? th.paper : th.ink;
      if (c.stack) { let yy = 250; c.stack.forEach(sc => { o.card(x, yy, 812, sc.h, sc.tint ? { fill: th.tint } : {}); // column of stacked cards, e.g. human skills first, technical below
        o.text(x + 36, yy + 36, 740, 56, sc.title, { font: 'H', size: 40, bold: true, tag: 'h3' }); o.text(x + 36, yy + 102, 740, sc.h - 120, '', { bullets: sc.bullets, size: 28, lh: 1.25 }); yy += sc.h + 30; }); return; }
      o.card(x, 250, 812, c.h || 560, dk ? { fill: th.ink } : {});
      o.text(x + 36, 286, 660, 56, c.title, { font: 'H', size: 40, bold: true, color: fg, tag: 'h3' });
      if (c.icon) o.icon(x + 812 - 36 - 48, 282, 48, c.icon, dk ? th.paper : th.accent);
      if (c.bullets) o.text(x + 36, 370, 740, 420, '', { bullets: c.bullets, size: 32, lh: 1.3, color: fg });
      if (c.kv) c.kv.forEach((kv, j) => { const y = 390 + j * 88; o.text(x + 36, y + 10, 210, 36, kv[0], { font: 'M', size: 24, bold: true, color: th.accent, upper: true });
        if (typeof kv[1] === 'string') return o.text(x + 260, y, 520, 50, kv[1], { size: 32 });
        let cx = x + 232; kv[1].forEach(tech => { o.img(cx, y, 48, 48, tech.logo, 'Logo ' + tech.n); const w = Math.ceil(tech.n.length * 15) + 6; o.text(cx + 58, y + 8, w, 36, tech.n, { size: 26 }); cx += 58 + w + 22; }); }); // kv value = "text" or [{n,logo}] -> logo row
      if (c.paras) { let y = 370; c.paras.forEach(p => { o.text(x + 36, y, 740, 170, p.t, { size: 32, lh: 1.3, bold: !!p.bold, color: fg }); y += 190; }); }
    });
  },
  image_list(o, d) { // screenshot left; right = heading + bullets, or heading + labelled groups [{label,bullets}]
    o.title(d.title); o.pic(128, 250, d.w || 960, d.h || 637, d.image, d.alt);
    o.text(1136, 250, 656, 56, d.head, { font: 'H', size: 40, bold: true, color: th.accent, tag: 'h3' });
    if (d.groups) { let y = 322; d.groups.forEach(g => { o.label(1136, y, 656, g.label); o.text(1136, y + 40, 656, g.bullets.length * 46 + 10, '', { bullets: g.bullets, size: 28, lh: 1.3 }); y += 40 + g.bullets.length * 46 + 36; }); }
    else o.text(1136, 340, 656, 420, '', { bullets: d.bullets, size: 32, lh: 1.3 });
  },
  images2(o, d) { // two screenshots + captions
    o.title(d.title);
    d.items.forEach((it, i) => { const x = 128 + i * 856; o.pic(x, 250, 808, 531, it.image, it.alt); o.text(x, 810, 808, 90, it.caption || '', { size: 28, lh: 1.3 }); });
  },
  image_facts(o, d) { // wide figure left + stacked fact cards right
    o.title(d.title); o.pic(128, 250, d.w || 1100, d.h || 571, d.image, d.alt);
    if (d.caption) o.text(128, 250 + (d.h || 571) + 24, d.w || 1100, 36, d.caption, { size: 24, italic: true, color: th.muted });
    d.facts.forEach((f, i) => { const y = 250 + i * 244; o.card(1276, y, 516, 220); o.label(1312, y + 40, 440, f.label); o.text(1312, y + 100, 440, 64, f.value, { font: 'H', size: 40, bold: true }); });
  },
  callout(o, d) { // 2-3 cards + full-width ink bar
    o.title(d.title);
    const n = d.cards.length, gap = n === 3 ? 32 : 40, w = Math.floor((1664 - gap * (n - 1)) / n), big = n < 3;
    d.cards.forEach((c, i) => { const x = 128 + i * (w + gap); o.card(x, 250, w, big ? 400 : 430);
      o.label(x + 36, 286, w - 72, c.label); o.text(x + 36, 336, w - 72, 56, c.title, { font: 'H', size: big ? 40 : 36, bold: true, tag: 'h3' });
      o.text(x + 36, big ? 420 : 410, w - 72, 250, c.text, { size: big ? 32 : 28, lh: 1.3 }); if (c.icon) o.icon(x + w - 36 - 48, 280, 48, c.icon); });
    const by = big ? 720 : 740, bh = big ? 160 : 150;
    o.rect(128, by, 1664, bh, { fill: th.ink, line: null });
    o.text(168, by, 1584, bh, d.bar, { font: 'H', size: 40, bold: true, color: th.paper, valign: 'middle', lh: 1.25 });
  },
  kpi(o, d) { // 2x2 big-number cards + footnote
    o.title(d.title);
    d.stats.forEach((c, i) => { const x = 128 + (i % 2) * 848, y = 250 + Math.floor(i / 2) * 284; o.card(x, y, 816, 256);
      o.text(x + 36, y + 20, 744, 110, c.value, { font: 'H', size: 96, bold: true, lh: 1.1, color: th.accent }); o.text(x + 36, y + 140, 744, 96, c.text, { size: 32, lh: 1.3 }); });
    if (d.note) o.text(128, 840, 1664, 70, d.note, { size: 24, color: th.muted, lh: 1.4 });
  },
  rows(o, d) { // up to 5 numbered rows; item = "text" or {t, badge:"VALIDÉ" (put validated items last)} (tinted row + accent pill)
    o.title(d.title);
    d.items.forEach((it, i) => { const y = 250 + i * 116, t = it.t || it, b = it.badge;
      o.rect(128, y, 96, 96, { fill: th.ink }); // same 4px ink border as the text card, so both have identical outer height
      o.text(128, y, 96, 96, String(i + 1).padStart(2, '0'), { font: 'H', size: 40, bold: true, color: th.paper, align: 'center', valign: 'middle' });
      o.rect(224, y, 1568, 96, { fill: b ? th.tint : th.card }); o.text(256, y, b ? 1260 : 1510, 96, t, { size: 32, valign: 'middle', lh: 1.25 });
      if (b) { o.rect(1604, y + 24, 160, 48, { fill: th.accent, line: null }); o.text(1604, y + 24, 160, 48, b, { font: 'M', size: 24, bold: true, color: th.paper, align: 'center', valign: 'middle', spacing: 2 }); } });
  },
  showcase_step(o, d) { // frame + carousel of screenshots (active one inside the frame, others parked off-slide) + accumulating list
    o.title(d.title); o.card(128, 250, 960, 637, { name: 'shot-frame' });
    d.shots.forEach((sh, j) => o.P.push({ t: 'img', x: 132, y: 254, w: 952, h: 629, file: sh.image, alt: sh.alt, name: 'shot-' + (j + 1),
      park: j === d.active ? null : (j < d.active ? -1100 : 2000), hidden: j !== d.active }));
    d.shots.forEach((sh, j) => { if (sh.caption) o.text(128, 910, 960, 36, sh.caption, { size: 24, italic: true, color: th.muted, lh: 1.3, name: 'cap-' + (j + 1), park: j === d.active ? null : (j < d.active ? -1100 : 2000), hidden: j !== d.active }); });
    let y = 250;
    d.groups.forEach((g, gi) => { if (!g.items.length) return;
      o.text(1136, y, 656, 34, g.label, { font: 'M', size: 24, bold: true, color: th.accent, spacing: 2, upper: true, name: 'it-L' + (gi + 1) }); y += 46;
      g.items.forEach(it => { o.text(1136, y, 656, 36, it.t, { size: 26, bold: it.fresh, color: it.fresh ? th.accent : th.ink, lh: 1.3, name: 'it-' + it.key }); y += 40; }); y += 18; });
  },
  closing(o, d) { // conclusion list + big thanks; chrome is "big"
    o.title(d.title); o.label(128, 330, 900, d.label || '');
    d.items.forEach((t, i) => { const y = 400 + i * 110; o.text(128, y + 4, 60, 40, String(i + 1).padStart(2, '0'), { font: 'M', size: 28, bold: true, color: th.accent }); o.text(200, y, 968, 100, t, { size: 32, lh: 1.3 }); });
    if (d.question) { o.text(128, 770, 1040, 140, d.question, { font: 'H', size: 56, bold: true, lh: 1.1, tag: 'h1' }); o.text(128, 915, 1040, 36, d.thanks || 'Merci de votre attention.', { size: 28, color: th.muted }); }
    else o.text(128, 800, 1040, 130, d.thanks || 'Merci.', { font: 'H', size: 96, bold: true, lh: 1.1, tag: 'h1' });
  },
};
function P_table(o, t) { o.P.push(Object.assign({ t: 'table' }, t)); }
const BIG = new Set(['cover', 'closing']);
const IDX = Object.fromEntries(slides.map((sl, i) => [sl.id, i + 1])); // slide id -> 1-based index (for clickable navigation)

// ---------- chrome shared by all slides (these ids/names are what Morph animates) ----------
const ROT = [0, 15, -15, 0, 30, -15, 15];
function chrome(o, i, s, big, dark, fg) {
  const n = i + 1, bw = Math.round(1664 * n / TOTAL);
  o.rect(128, 992, 1664, 20, { fill: dark ? th.accent : th.paper, line: fg, bw: 3, name: 'mk-track' });
  o.rect(128, 992, bw, 20, { fill: dark ? th.paper : th.accent, line: fg, bw: 3, name: 'mk-bar' });
  const g = big ? { x: 1232, y: 200, w: 480, h: 480 } : { x: 1720, y: 40, w: 72, h: 72 };
  o.rect(g.x, g.y, g.w, g.h, { fill: dark ? th.paper : th.accent, so: big ? 16 : 8, rot: s.rot != null ? s.rot : (big && i === 0 ? 0 : ROT[i % ROT.length]), name: 'mk-sq' });
  const lab = String(s.label || '').trim(), lw = Math.round(lab.length * 16.6) + 48; // page number card + title card, side by side, 8px apart, NO dash
  o.P.push({ t: 'chip', kind: 'num', x: 128, y: 44, w: 170, h: 56, text: `${String(n).padStart(2, '0')} / ${TOTAL}`, name: 'mk-num' });
  o.P.push({ t: 'chip', kind: 'title', x: 306, y: 44, w: lw, h: 56, text: lab, name: 'mk-title' });
  const logos = spec.logos || []; if (!logos.length) return;
  const k = big ? 1.436 : 1, pl = big ? { x: 1232, y: 740, w: 560, h: 80 } : { x: 1290, y: 44, w: 390, h: 56 };
  o.rect(pl.x, pl.y, pl.w, pl.h, { fill: th.card, bw: 3, so: big ? 8 : 6, name: 'mk-logos' });
  const gap = 24 * k, avail = pl.w - 32 * k, sumR = logos.reduce((a, l) => a + l.ratio, 0), h0 = Math.min(40 * k, (avail - gap * (logos.length - 1)) / sumR);
  let x = pl.x + (pl.w - (sumR * h0 + gap * (logos.length - 1))) / 2;
  logos.forEach((l, j) => { const w = l.ratio * h0; o.P.push({ t: 'img', x: Math.round(x), y: Math.round(pl.y + (pl.h - h0) / 2), w: Math.round(w), h: Math.round(h0), file: l.file, alt: l.alt || 'Logo', name: 'mk-logo-' + (j + 1) }); x += w + gap; });
}

// ---------- build all prims ----------
const built = slides.map((s, i) => {
  const dark = s.layout === 'statement', fg = dark ? th.paper : th.ink, P = [], o = maker(P, fg, dark); o.P = P;
  chrome(o, i, s, BIG.has(s.layout), dark, fg);
  if (!LAYOUT[s.layout]) throw new Error('unknown layout ' + s.layout);
  LAYOUT[s.layout](o, s.data || {});
  if (s.id !== (spec.planId || 'plan')) P.push({ t: 'nav', x: P.find(p => p.name === 'mk-sq').x, y: P.find(p => p.name === 'mk-sq').y, w: P.find(p => p.name === 'mk-sq').w, h: P.find(p => p.name === 'mk-sq').h, to: spec.planId || 'plan' }); // blue square = back to the plan (PPTX)
  return { s, P, dark, fg };
});

// ---------- HTML emitter ----------
const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const px = v => Math.round(v * 100) / 100;
function html(b, i) {
  const { s, P, dark, fg } = b; const last = i === TOTAL - 1;
  const out = P.map(p => {
    const pos = `position:absolute;left:${px(p.x)}px;top:${px(p.y)}px;`, id = p.name ? ` id="${p.name}"` : '';
    if (p.t === 'rect') return `<div${id} style="${pos}width:${px(p.w)}px;height:${px(p.h)}px;background:#${p.fill};${p.line ? `border:${p.bw}px solid #${p.line};` : ''}${p.so ? `box-shadow:${p.so}px ${p.so}px 0 #${th.ink};` : ''}${p.rot ? `transform:rotate(${p.rot}deg)` : ''}"></div>`;
    if (p.t === 'chip') { const num = p.kind === 'num'; return `<div${id} style="${pos}width:${p.w}px;height:${p.h}px;display:flex;flex-direction:column;justify-content:center;background:#${num ? th.ink : th.card};border:3px solid #${th.ink}"><p style="font-family:${FCSS.M};font-size:24px;font-weight:700;letter-spacing:2px;text-align:center;${num ? '' : 'text-transform:uppercase;'}color:#${num ? th.paper : th.ink}">${esc(p.text)}</p></div>`; }
    if (p.t === 'nav') return ''; // internal slide links exist only in the PPTX
    if (p.t === 'img') return `<img${id} src="${URLS[p.file] || ('/_blob/MISSING-' + p.file)}" alt="${esc(p.alt || '')}" style="${pos}width:${px(p.w)}px;height:${px(p.h)}px;object-fit:contain${p.hidden ? ';opacity:0' : ''}">`;
    if (p.t === 'icon') return `<x-icon name="${p.name}" style="${pos}width:${p.s}px;height:${p.s}px;color:#${p.color}"></x-icon>`;
    if (p.t === 'table') {
      const hd = p.head.map((h, j) => `<th style="width:${(p.colW[j] / p.w * 100).toFixed(1)}%;color:#${th.paper};text-align:${p.align[j]}">${esc(h)}</th>`).join('');
      const rows = p.rows.map((r, k) => { const tot = p.totalRow && k === p.rows.length - 1; return `<tr style="background:#${tot ? th.tint : th.card}">` + r.map((c, j) => `<td style="text-align:${p.align[j]}">${tot || p.boldCols.includes(j) ? `<b>${esc(c)}</b>` : esc(c)}</td>`).join('') + '</tr>'; }).join('');
      return `<div style="${pos}width:${p.w}px"><table style="width:100%;border:4px solid #${th.ink};font-family:${FCSS.B};font-size:${p.size}px;line-height:1.3"><tr style="background:#${th.ink}">${hd}</tr>${rows}</table></div>`;
    }
    // text
    const tag = p.tag || 'p', wt = p.bold ? (p.font === 'B' ? 600 : 700) : 400;
    const st = `font-family:${FCSS[p.font]};font-size:${p.size}px;font-weight:${wt};color:#${p.color};line-height:${p.lh};text-align:${p.align};${p.spacing ? `letter-spacing:${p.spacing}px;` : ''}${p.upper ? 'text-transform:uppercase;' : ''}${p.italic ? 'font-style:italic;' : ''}${p.hidden ? 'opacity:0;' : ''}`;
    const inner = p.bullets ? `<ul style="${pos}width:${p.w}px;${st}">${p.bullets.map(b => `<li>${esc(b)}</li>`).join('')}</ul>` : null;
    if (inner) return inner;
    if (p.valign === 'middle') return `<div style="${pos}width:${p.w}px;height:${p.h}px;display:flex;flex-direction:column;justify-content:center"><${tag} style="${st}">${esc(p.text)}</${tag}></div>`;
    return `<${tag}${id} style="${pos}width:${p.w}px;${st}">${esc(p.text)}</${tag}>`;
  }).join('\n');
  return `<section id="${s.id}"${last ? '' : ' data-transition="magic"'} style="background:#${dark ? th.accent : th.paper};color:#${fg};font-family:${FCSS.B};display:flex;flex-direction:column;padding:128px 128px 160px">\n${out}\n<aside>${esc(s.notes || '')}</aside>\n</section>`;
}
built.forEach((b, i) => fs.writeFileSync(path.join(OUT, 'project/slides', b.s.id + '.html'), html(b, i)));
const secs = {}; (spec.sections || []).forEach((x, k) => secs['s' + (k + 1)] = { description: x.description, start: x.start });
fs.writeFileSync(path.join(OUT, 'project/deck.json'), JSON.stringify({ v: 4, createdOnFiles: { v: 1, at: new Date().toISOString().replace(/\.\d+Z$/, 'Z') }, lists: 'css', title: spec.title, order: slides.map(s => s.id), sections: secs, faces, designSystems: [] }, null, 1));

// ---------- PPTX emitter ----------
const pt = v => v / 2, inch = v => v / 144;
async function iconPngs() { // render used icons (Lucide) to PNG, in the colours used
  const need = new Map(); built.forEach(b => b.P.filter(p => p.t === 'icon').forEach(p => need.set(p.name + '_' + p.color, p)));
  const dir = path.join(OUT, '.icons'); fs.mkdirSync(dir, { recursive: true }); const res = {};
  if (!need.size) return res;
  let React, S, sharp, lu; try { React = require('react'); S = require('react-dom/server'); sharp = require('sharp'); lu = require('react-icons/lu'); } catch (e) { console.warn('! icons skipped in PPTX (install react react-dom react-icons sharp):', e.message); return res; }
  for (const [k, p] of need) { const C = lu[ICON[p.name]]; if (!C) { console.warn('! no icon for', p.name); continue; }
    const f = path.join(dir, k + '.png'); await sharp(Buffer.from(S.renderToStaticMarkup(React.createElement(C, { color: '#' + p.color, size: 256, strokeWidth: 2.2 }))), { density: 300 }).resize(256, 256).png().toFile(f); res[k] = f; }
  return res;
}
(async () => {
  const icons = await iconPngs();
  const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE'; pres.title = spec.title;
  built.forEach((b, i) => {
    const sl = pres.addSlide(); sl.background = { color: b.dark ? th.accent : th.paper };
    b.P.forEach(p => {
      const nm = p.name ? { objectName: '!!' + p.name } : {};   // "!!" forces Morph to pair same-named objects
      if (p.t === 'rect') sl.addShape(pres.shapes.RECTANGLE, Object.assign({ x: inch(p.x), y: inch(p.y), w: inch(p.w), h: inch(p.h), fill: { color: p.fill }, line: p.line ? { color: p.line, width: p.bw / 2 } : { type: 'none' } }, p.so ? { shadow: { type: 'outer', color: th.ink, blur: 0, offset: p.so / 2 * 1.414, angle: 45, opacity: 1 } } : {}, p.rot ? { rotate: p.rot } : {}, nm));
      else if (p.t === 'chip') { const num = p.kind === 'num'; sl.addText(p.text.toUpperCase(), Object.assign({ x: inch(p.x), y: inch(p.y), w: inch(p.w), h: inch(p.h), fontFace: FN.M, fontSize: 12, bold: true, color: num ? th.paper : th.ink, charSpacing: 2, align: 'center', valign: 'middle', fill: { color: num ? th.ink : th.card }, line: { color: th.ink, width: 1.5 }, margin: 0, isTextBox: true }, nm)); }
      else if (p.t === 'nav') { const t = IDX[p.to] || IDX[p.to + '1']; if (t) sl.addShape(pres.shapes.RECTANGLE, { x: inch(p.x), y: inch(p.y), w: inch(p.w), h: inch(p.h), fill: { color: 'FFFFFF', transparency: 100 }, line: { type: 'none' }, objectName: 'nav:' + t }); } // transparent hit area; link injected after write
      else if (p.t === 'img') sl.addImage(Object.assign({ path: path.join(ASSETS, p.file), x: inch(p.park != null ? p.park : p.x), y: inch(p.y), w: inch(p.w), h: inch(p.h), altText: p.alt || '' }, nm));
      else if (p.t === 'icon') { const f = icons[p.name + '_' + p.color]; if (f) sl.addImage({ path: f, x: inch(p.x), y: inch(p.y), w: inch(p.s), h: inch(p.s), altText: '' }); }
      else if (p.t === 'table') {
        const bd = { type: 'solid', pt: 1.5, color: th.ink }, B = [bd, bd, bd, bd];
        const rows = [p.head.map((h, j) => ({ text: h, options: { bold: true, color: th.paper, fill: { color: th.ink }, fontFace: FN.H, fontSize: pt(p.size), align: p.align[j], valign: 'middle', border: B } }))];
        p.rows.forEach((r, k) => { const tot = p.totalRow && k === p.rows.length - 1; rows.push(r.map((c, j) => ({ text: c, options: { fontFace: FN.B, fontSize: pt(p.size), color: th.ink, bold: tot || p.boldCols.includes(j), fill: { color: tot ? th.tint : th.card }, align: p.align[j], valign: 'middle', border: B } }))); });
        sl.addTable(rows, { x: inch(p.x), y: inch(p.y), w: inch(p.w), colW: p.colW.map(inch), rowH: [inch(p.headH)].concat(p.rows.map(() => inch(p.rowH))) });
      } else { // text
        const o = { x: inch(p.park != null ? p.park : p.x), y: inch(p.y), w: inch(p.w), h: inch(p.h), fontFace: FN[p.font], fontSize: pt(p.size), bold: !!p.bold, italic: !!p.italic, color: p.color, align: p.align, valign: p.valign, margin: 0, isTextBox: true, fit: 'none', lineSpacingMultiple: p.lh / 1.2 };
        if (p.spacing) o.charSpacing = p.spacing;
        const txt = p.bullets ? p.bullets.map((t, j) => ({ text: t, options: { bullet: { indent: 22 }, breakLine: j < p.bullets.length - 1, paraSpaceAfter: 8 } })) : (p.upper ? p.text.toUpperCase() : p.text);
        sl.addText(txt, Object.assign(o, nm));
      }
    });
    sl.addNotes(b.s.notes || '');
  });
  const raw = path.join(OUT, '_raw.pptx'); await pres.writeFile({ fileName: raw });
  // inject native Morph into slides 2..N  (transition lives on the slide that is ARRIVED at)
  const MORPH = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1200"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
  const zip = await JSZip.loadAsync(fs.readFileSync(raw));
  for (const f of Object.keys(zip.files)) { const m = f.match(/^ppt\/slides\/slide(\d+)\.xml$/); if (m && +m[1] >= 2) { const x = await zip.file(f).async('string'); if (!x.includes('</p:clrMapOvr>')) throw new Error('no clrMapOvr in ' + f); zip.file(f, x.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + MORPH)); } }
  for (const f of Object.keys(zip.files)) { const m = f.match(/^ppt\/slides\/slide(\d+)\.xml$/); if (!m) continue; let x = await zip.file(f).async('string'); const rn = `ppt/slides/_rels/slide${m[1]}.xml.rels`; let rels = await zip.file(rn).async('string');
    for (const t of new Set([...x.matchAll(/name="nav:(\d+)"/g)].map(q => q[1]))) { const rid = 'rIdNav' + t;
      rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${t}.xml"/></Relationships>`);
      x = x.replace(new RegExp(`<p:cNvPr ([^>]*?)name="nav:${t}"([^>/]*?)(?:/>|></p:cNvPr>)`, 'g'), (_, a, b) => `<p:cNvPr ${a}name="nav:${t}"${b}><a:hlinkClick r:id="${rid}" action="ppaction://hlinksldjump"/></p:cNvPr>`); }
    zip.file(f, x); zip.file(rn, rels); }
  const dst = path.join(OUT, (spec.filename || 'Presentation') + '.pptx'); fs.writeFileSync(dst, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })); fs.unlinkSync(raw);
  console.log('web  :', path.join(OUT, 'project/deck.json'), '+', TOTAL, 'slides\npptx :', dst);
  const miss = Object.keys(URLS).length ? [] : [...new Set(built.flatMap(b => b.P.filter(p => p.t === 'img').map(p => p.file)))];
  if (miss.length) console.log('note : no --urls given; web <img> point to /_blob/MISSING-*. Upload these assets then rebuild:', miss.join(', '));
})();
