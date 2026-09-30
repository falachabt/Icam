#!/usr/bin/env node
/* Download brand logos (Simple Icons, CC0) as coloured 256px PNGs.
   usage: node fetch-brand-logos.js outdir nextdotjs:16181D react:1EA7C9 docker:2496ED ...
   slug = simple-icons slug (https://simpleicons.org), then ':' and the HEX colour to fill with.
   needs: curl, sharp. Result: outdir/<slug>.png  (reference it in the spec as {"n":"Docker","logo":"docker.png"}) */
const { execFileSync } = require('child_process'), sharp = require('sharp'), fs = require('fs'), path = require('path');
const [out, ...items] = process.argv.slice(2); fs.mkdirSync(out, { recursive: true });
(async () => { for (const it of items) { const [slug, hex] = it.split(':');
  const svg = execFileSync('curl', ['-fsSL', `https://cdn.jsdelivr.net/npm/simple-icons@13/icons/${slug}.svg`]).toString().replace('<svg ', `<svg fill="#${hex}" `);
  await sharp(Buffer.from(svg), { density: 600 }).resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(path.join(out, slug + '.png')); console.log('ok', slug); } })();
