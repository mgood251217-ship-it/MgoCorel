import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const iconDir = path.join(root, "src", "icons");
const addonDir = path.join(root, "addon");
const appUiPath = path.join(addonDir, "AppUI.xslt");
const dllPath = path.join(addonDir, "Resources.dll");
const configPath = path.join(addonDir, "config.xml");

const FIRST_ICON_ID = 101;
const RT_ICON = 3;
const RT_GROUP_ICON = 14;

const captionIcons = {
  "detail": "detail.ico",
  "detail aplikasi": "detail.ico",
  "pengaturan": "settings.ico",
  "spanduk": "label.ico",
  "buat label spanduk": "label.ico",
  "export spanduk": "export.ico",
  "layout": "susun.ico",
  "susun objek": "susun.ico",
  "imposisi otomatis": "imposisi.ico",
  "rectangle nesting": "rectangleNesting.ico",
  "nesting": "nesting.ico",
  "produksi": "calculator.ico",
  "hitung harga": "calculator.ico",
  "duplicate quantity": "duplicate.ico",
  "numbering otomatis": "numbering.ico",
  "job": "excel.ico",
  "job builder": "excel.ico"
};

const align = (n, a) => Math.ceil(n / a) * a;

function fail(message) {
  console.error(`[FAIL] ${message}`);
  process.exit(1);
}

function parseIco(buf, name) {
  if (buf.length < 6 || buf.readUInt16LE(0) !== 0 || buf.readUInt16LE(2) !== 1) {
    fail(`${name} bukan file .ico yang valid`);
  }

  const count = buf.readUInt16LE(4);
  const images = [];

  for (let i = 0; i < count; i++) {
    const o = 6 + i * 16;
    const size = buf.readUInt32LE(o + 8);
    const offset = buf.readUInt32LE(o + 12);

    if (offset + size > buf.length) {
      fail(`${name} rusak pada gambar ke-${i + 1}`);
    }

    images.push({
      width: buf[o],
      height: buf[o + 1],
      colors: buf[o + 2],
      planes: buf.readUInt16LE(o + 4),
      bits: buf.readUInt16LE(o + 6),
      data: buf.subarray(offset, offset + size)
    });
  }

  return images;
}

function buildResourceItems(iconFiles) {
  const items = [];
  const groupIds = {};
  let nextImageId = 1;

  iconFiles.forEach((file, index) => {
    const images = parseIco(fs.readFileSync(path.join(iconDir, file)), file);
    const groupId = FIRST_ICON_ID + index;
    groupIds[file] = groupId;

    const group = Buffer.alloc(6 + 14 * images.length);
    group.writeUInt16LE(0, 0);
    group.writeUInt16LE(1, 2);
    group.writeUInt16LE(images.length, 4);

    images.forEach((image, k) => {
      const imageId = nextImageId++;
      items.push({ type: RT_ICON, id: imageId, data: image.data });

      const o = 6 + 14 * k;
      group[o] = image.width;
      group[o + 1] = image.height;
      group[o + 2] = image.colors;
      group[o + 3] = 0;
      group.writeUInt16LE(image.planes, o + 4);
      group.writeUInt16LE(image.bits, o + 6);
      group.writeUInt32LE(image.data.length, o + 8);
      group.writeUInt16LE(imageId, o + 12);
    });

    items.push({ type: RT_GROUP_ICON, id: groupId, data: group });
  });

  return { items, groupIds };
}

function buildResourceSection(items, rva) {
  const byType = new Map();

  for (const item of items) {
    if (!byType.has(item.type)) byType.set(item.type, []);
    byType.get(item.type).push(item);
  }

  const types = [...byType.keys()].sort((a, b) => a - b);
  for (const t of types) byType.get(t).sort((a, b) => a.id - b.id);

  let pos = 16 + 8 * types.length;
  const typeDirOff = {};
  const nameDirOff = {};
  const leafOff = {};
  const dataOff = {};

  for (const t of types) {
    typeDirOff[t] = pos;
    pos += 16 + 8 * byType.get(t).length;
  }

  for (const t of types) {
    for (const item of byType.get(t)) {
      nameDirOff[`${t}:${item.id}`] = pos;
      pos += 24;
    }
  }

  for (const t of types) {
    for (const item of byType.get(t)) {
      leafOff[`${t}:${item.id}`] = pos;
      pos += 16;
    }
  }

  for (const t of types) {
    for (const item of byType.get(t)) {
      pos = align(pos, 4);
      dataOff[`${t}:${item.id}`] = pos;
      pos += item.data.length;
    }
  }

  const buf = Buffer.alloc(align(pos, 4));
  const subdir = (off) => ((off | 0x80000000) >>> 0);

  buf.writeUInt16LE(types.length, 14);

  types.forEach((t, i) => {
    const e = 16 + 8 * i;
    buf.writeUInt32LE(t, e);
    buf.writeUInt32LE(subdir(typeDirOff[t]), e + 4);
  });

  for (const t of types) {
    const list = byType.get(t);
    buf.writeUInt16LE(list.length, typeDirOff[t] + 14);

    list.forEach((item, i) => {
      const key = `${t}:${item.id}`;
      const e = typeDirOff[t] + 16 + 8 * i;

      buf.writeUInt32LE(item.id, e);
      buf.writeUInt32LE(subdir(nameDirOff[key]), e + 4);

      buf.writeUInt16LE(1, nameDirOff[key] + 14);
      buf.writeUInt32LE(0x409, nameDirOff[key] + 16);
      buf.writeUInt32LE(leafOff[key], nameDirOff[key] + 20);

      buf.writeUInt32LE(rva + dataOff[key], leafOff[key]);
      buf.writeUInt32LE(item.data.length, leafOff[key] + 4);

      item.data.copy(buf, dataOff[key]);
    });
  }

  return buf;
}

function buildDll(rsrc) {
  const fileAlign = 0x200;
  const sectionAlign = 0x1000;
  const rva = 0x1000;
  const rawSize = align(rsrc.length, fileAlign);
  const imageSize = align(rva + rsrc.length, sectionAlign);
  const out = Buffer.alloc(fileAlign + rawSize);

  out.write("MZ", 0, "latin1");
  out.writeUInt32LE(0x40, 0x3c);

  let o = 0x40;
  out.write("PE\0\0", o, "latin1");
  o += 4;

  out.writeUInt16LE(0x8664, o);
  out.writeUInt16LE(1, o + 2);
  out.writeUInt16LE(240, o + 16);
  out.writeUInt16LE(0x2022, o + 18);
  o += 20;

  out.writeUInt16LE(0x20b, o);
  out[o + 2] = 14;
  out.writeUInt32LE(rawSize, o + 8);
  out.writeBigUInt64LE(0x180000000n, o + 24);
  out.writeUInt32LE(sectionAlign, o + 32);
  out.writeUInt32LE(fileAlign, o + 36);
  out.writeUInt16LE(6, o + 40);
  out.writeUInt16LE(6, o + 48);
  out.writeUInt32LE(imageSize, o + 56);
  out.writeUInt32LE(fileAlign, o + 60);
  out.writeUInt16LE(2, o + 68);
  out.writeUInt16LE(0x0140, o + 70);
  out.writeBigUInt64LE(0x100000n, o + 72);
  out.writeBigUInt64LE(0x1000n, o + 80);
  out.writeBigUInt64LE(0x100000n, o + 88);
  out.writeBigUInt64LE(0x1000n, o + 96);
  out.writeUInt32LE(16, o + 108);
  out.writeUInt32LE(rva, o + 112 + 2 * 8);
  out.writeUInt32LE(rsrc.length, o + 112 + 2 * 8 + 4);
  o += 240;

  out.write(".rsrc", o, "latin1");
  out.writeUInt32LE(rsrc.length, o + 8);
  out.writeUInt32LE(rva, o + 12);
  out.writeUInt32LE(rawSize, o + 16);
  out.writeUInt32LE(fileAlign, o + 20);
  out.writeUInt32LE(0x40000040, o + 36);

  rsrc.copy(out, fileAlign);

  return out;
}

function patchAppUi(source, iconFiles, groupIds) {
  const eol = source.includes("\r\n") ? "\r\n" : "\n";
  const entries = [];
  const missing = new Set();
  let patched = 0;

  const result = source.replace(/<itemData\b[^>]*?\/>/gs, (element) => {
    const guid = /\sguid="([^"]+)"/.exec(element);
    const caption = /\suserCaption="([^"]*)"/.exec(element);

    if (!guid || !caption) return element;

    const file = captionIcons[caption[1].trim().toLowerCase()];

    if (!file) return element;

    if (!iconFiles.includes(file)) {
      missing.add(file);
      return element;
    }

    entries.push({ guid: guid[1], caption: caption[1], file, id: groupIds[file] });

    if (/\sicon="/.test(element)) return element;

    patched++;

    return element.replace(
      guid[0],
      `${guid[0]}${eol}                icon="guid://${guid[1]}"`
    );
  });

  return { result, entries, missing: [...missing], patched };
}

function renderConfig(entries) {
  const lines = [
    '<?xml version="1.0" encoding="utf-8"?>',
    "<config>",
    "  <resources>",
    '    <resource name="MgoCorel.Resources" path="Resources.dll">',
    "      <resourceMap>"
  ];

  for (const entry of entries) {
    lines.push(`        <resEntry id="${entry.guid}" icon="${entry.id}"/>`);
  }

  lines.push(
    "      </resourceMap>",
    "    </resource>",
    "  </resources>",
    "</config>",
    ""
  );

  return lines.join("\r\n");
}

console.log("");
console.log("======================================");
console.log("       MgoCorel Addon Build");
console.log("======================================");
console.log("");

if (!fs.existsSync(appUiPath)) fail(`AppUI.xslt tidak ditemukan: ${appUiPath}`);
if (!fs.existsSync(iconDir)) fail(`Folder icons tidak ditemukan: ${iconDir}`);

const iconFiles = fs
  .readdirSync(iconDir)
  .filter((f) => f.toLowerCase().endsWith(".ico"))
  .sort();

if (iconFiles.length === 0) fail("Tidak ada file .ico di src\\icons");

const { items, groupIds } = buildResourceItems(iconFiles);
const dll = buildDll(buildResourceSection(items, 0x1000));

const source = fs.readFileSync(appUiPath, "utf8").replace(/^\uFEFF/, "");
const { result, entries, missing, patched } = patchAppUi(source, iconFiles, groupIds);

for (const file of missing) {
  console.log(`[WARN] Ikon tidak ada di src\\icons: ${file}`);
}

if (entries.length === 0) fail("Tidak ada itemData yang cocok dengan daftar caption ikon");

fs.writeFileSync(dllPath, dll);
fs.writeFileSync(configPath, renderConfig(entries), "utf8");
fs.writeFileSync(appUiPath, result, "utf8");

console.log(`[OK] Resources.dll: ${iconFiles.length} ikon, ${items.length} resource, ${dll.length} bytes`);
console.log(`[OK] config.xml: ${entries.length} item`);
console.log(`[OK] AppUI.xslt: ${patched} item diberi atribut icon`);
console.log("");

for (const entry of entries) {
  console.log(`  ${String(entry.id).padStart(3)}  ${entry.file.padEnd(22)} ${entry.caption}`);
}

console.log("");
