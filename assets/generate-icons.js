/* Generates icon-16/32/80.png: white back-arrow on a green disc. No dependencies. */
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
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 6;  // RGBA
  // Prepend filter byte 0 to each row.
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    pixels.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  const png = Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0))
  ]);
  fs.writeFileSync(file, png);
  console.log("Wrote", file);
}

/** Signed distance-ish rendering of a disc + left arrow, 4x supersampled. */
function render(size) {
  const px = Buffer.alloc(size * size * 4);
  const SS = 4;
  const green = [0x21, 0x73, 0x46];

  function insideArrow(u, v) {
    // Coordinates in unit space (0..1). Arrow head: triangle pointing left; shaft: rectangle.
    const headTipX = 0.22, headBackX = 0.5, headHalf = 0.22;
    const shaftX0 = 0.46, shaftX1 = 0.78, shaftHalf = 0.09;
    const dy = Math.abs(v - 0.5);
    if (u >= headTipX && u <= headBackX) {
      const t = (u - headTipX) / (headBackX - headTipX);
      if (dy <= t * headHalf) return true;
    }
    if (u >= shaftX0 && u <= shaftX1 && dy <= shaftHalf) return true;
    return false;
  }

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let discCov = 0, arrowCov = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const u = (x + (sx + 0.5) / SS) / size;
          const v = (y + (sy + 0.5) / SS) / size;
          const d = Math.hypot(u - 0.5, v - 0.5);
          if (d <= 0.48) {
            discCov++;
            if (insideArrow(u, v)) arrowCov++;
          }
        }
      }
      const total = SS * SS;
      const alpha = Math.round((discCov / total) * 255);
      const aw = arrowCov / total;
      const i = (y * size + x) * 4;
      px[i] = Math.round(green[0] * (1 - aw) + 255 * aw);
      px[i + 1] = Math.round(green[1] * (1 - aw) + 255 * aw);
      px[i + 2] = Math.round(green[2] * (1 - aw) + 255 * aw);
      px[i + 3] = alpha;
    }
  }
  return px;
}

for (const size of [16, 32, 64, 80]) {
  writePng(size, render(size), path.join(__dirname, `icon-${size}.png`));
}
