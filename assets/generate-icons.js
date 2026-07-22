/* Generates opaque ribbon icons: green square + white glyph (back / forward / clock). */
const zlib = require("zlib");
const fs = require("fs");
const path = require("path");

function crc32(buf) {
  let c, table = [];
  for (let n = 0; n < 256; n++) {
    c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function writePng(size, pixels, file) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    pixels.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  const png = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0))
  ]);
  fs.writeFileSync(file, png);
  console.log("Wrote", file, png.length, "bytes");
}

function distSeg(px, py, ax, ay, bx, by) {
  const abx = bx - ax, aby = by - ay;
  const apx = px - ax, apy = py - ay;
  const ab2 = abx * abx + aby * aby;
  let t = ab2 === 0 ? 0 : (apx * abx + apy * aby) / ab2;
  t = Math.max(0, Math.min(1, t));
  return Math.hypot(px - (ax + t * abx), py - (ay + t * aby));
}

/** Back chevron: tip on the left. */
function inBackChevron(u, v) {
  const tipX = 0.30, tipY = 0.5, reach = 0.40, thick = 0.12;
  const d1 = distSeg(u, v, tipX, tipY, tipX + reach, tipY - reach);
  const d2 = distSeg(u, v, tipX, tipY, tipX + reach, tipY + reach);
  return Math.min(d1, d2) <= thick / 2;
}

/** Forward chevron: tip on the right (mirror of back). */
function inForwardChevron(u, v) {
  return inBackChevron(1 - u, v);
}

/** Clock face: ring + hour/minute hands. */
function inClock(u, v) {
  const cx = 0.5, cy = 0.5;
  const dx = u - cx, dy = v - cy;
  const r = Math.hypot(dx, dy);
  const outer = 0.34, ringW = 0.07;
  // Outer ring
  if (r >= outer - ringW && r <= outer) return true;
  // Center hub
  if (r <= 0.045) return true;
  // Minute hand (up)
  const minuteTipY = cy - 0.22;
  if (distSeg(u, v, cx, cy, cx, minuteTipY) <= 0.035) return true;
  // Hour hand (about 4 o'clock)
  const hourTipX = cx + 0.14;
  const hourTipY = cy + 0.08;
  if (distSeg(u, v, cx, cy, hourTipX, hourTipY) <= 0.04) return true;
  return false;
}

function render(size, glyphFn) {
  const px = Buffer.alloc(size * size * 4);
  const SS = 4;
  const green = [0x21, 0x73, 0x46];
  const white = [255, 255, 255];
  const radius = 0.12;

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let cover = 0, glyph = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const u = (x + (sx + 0.5) / SS) / size;
          const v = (y + (sy + 0.5) / SS) / size;
          const dx = Math.max(radius - u, 0, u - (1 - radius));
          const dy = Math.max(radius - v, 0, v - (1 - radius));
          const outside = dx * dx + dy * dy > radius * radius && (u < radius || u > 1 - radius) && (v < radius || v > 1 - radius);
          if (!outside) {
            cover++;
            if (glyphFn(u, v)) glyph++;
          }
        }
      }
      const total = SS * SS;
      const t = glyph / total;
      const i = (y * size + x) * 4;
      // Fully opaque pixels for ribbon compatibility
      px[i] = Math.round(green[0] * (1 - t) + white[0] * t);
      px[i + 1] = Math.round(green[1] * (1 - t) + white[1] * t);
      px[i + 2] = Math.round(green[2] * (1 - t) + white[2] * t);
      px[i + 3] = 255;
      void cover;
    }
  }
  return px;
}

const icons = [
  { name: "back", glyph: inBackChevron },
  { name: "forward", glyph: inForwardChevron },
  { name: "history", glyph: inClock }
];

for (const { name, glyph } of icons) {
  for (const size of [16, 32, 64, 80]) {
    writePng(size, render(size, glyph), path.join(__dirname, `icon-${name}-${size}.png`));
  }
}

// App logo aliases (history clock) for IconUrl / HighResolutionIconUrl
for (const size of [16, 32, 64, 80]) {
  const src = path.join(__dirname, `icon-history-${size}.png`);
  const dest = path.join(__dirname, `icon-${size}.png`);
  fs.copyFileSync(src, dest);
  console.log("Copied", dest);
}
