// office.js - xử lý file Word có hình ảnh và định dạng/tính toán Excel cho bot tt
// Cần cài thêm: npm install jszip exceljs
const JSZip = require('jszip');
const ExcelJS = require('exceljs');

// ======================= PHẦN CHUNG =======================
const CMD_RE = /^\/(wordvanbanhinhanh|wordhvanbanhinhanh|exceltinhtoan|excel)(?:@\w+)?(?:\s+([\s\S]*))?$/i;
function parseOfficeCommand(text) {
  const m = CMD_RE.exec(String(text || '').trim());
  if (!m) return null;
  let cmd = m[1].toLowerCase();
  if (cmd === 'wordhvanbanhinhanh') cmd = 'wordvanbanhinhanh';
  return { cmd, args: (m[2] || '').trim() };
}

// ======================= WORD CÓ HÌNH ẢNH =======================
// Khung ảnh chuẩn: 10 cm x 7,5 cm (EMU). Ảnh được cắt khung, không kéo giãn nên không méo.
const FW = 3600000;
const FH = 2700000;
const V_BIAS = 0.35; // 0 = lấy phần trên, 0.5 = lấy giữa

function imgSize(buf) {
  try {
    if (buf[0] === 0x89 && buf[1] === 0x50) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }; // PNG
    if (buf[0] === 0xff && buf[1] === 0xd8) { // JPEG
      let i = 2;
      while (i < buf.length) {
        if (buf[i] !== 0xff) { i++; continue; }
        const mk = buf[i + 1];
        if (mk >= 0xc0 && mk <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(mk)) {
          return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
        }
        i += 2 + buf.readUInt16BE(i + 2);
      }
    }
  } catch (e) { /* bỏ qua */ }
  return null;
}

const PARA_RE = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/g;
const PPR_IMG =
  '<w:pPr><w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>' +
  '<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr>';

async function fixWordImages(buf) {
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0]);
    const tg = /\bTarget="([^"]+)"/.exec(m[0]);
    if (id && tg) rels[id[1]] = tg[1];
  }
  // đọc kích thước thật của từng ảnh
  const sizes = {};
  for (const m of xml.matchAll(/r:embed="([^"]+)"/g)) {
    const rid = m[1];
    if (sizes[rid] !== undefined || !rels[rid]) continue;
    const path = 'word/' + rels[rid].replace(/^\.?\//, '');
    const f = zip.file(path);
    sizes[rid] = f ? imgSize(Buffer.from(await f.async('uint8array'))) : null;
  }

  let count = 0;
  const fixDrawing = (d) => {
    const rid = (/r:embed="([^"]+)"/.exec(d) || [])[1];
    const s = sizes[rid];
    let w = FW, h = FH, crop = '';
    if (s && s.w && s.h) {
      const r = s.w / s.h, t = FW / FH;
      if (r < t) { // ảnh đứng: cắt trên/dưới
        const cut = (1 - r / t) * 100000;
        crop = `<a:srcRect t="${Math.round(cut * V_BIAS)}" b="${Math.round(cut * (1 - V_BIAS))}"/>`;
      } else if (r > t) { // ảnh quá ngang: cắt hai bên
        const cut = (1 - t / r) * 100000;
        crop = `<a:srcRect l="${Math.round(cut / 2)}" r="${Math.round(cut / 2)}"/>`;
      }
    } else { // không đọc được kích thước: giữ tỉ lệ hiện có, thu vào khung
      const e = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(d);
      if (e) { const k = Math.min(FW / +e[1], FH / +e[2]); w = Math.round(+e[1] * k); h = Math.round(+e[2] * k); }
    }
    d = d.replace(/<wp:extent[^>]*\/>/, `<wp:extent cx="${w}" cy="${h}"/>`);
    d = d.replace(/<wp:effectExtent[^>]*\/>/, '<wp:effectExtent l="0" t="0" r="0" b="0"/>');
    d = d.replace(/<a:ext cx="\d+" cy="\d+"\/>/g, `<a:ext cx="${w}" cy="${h}"/>`);
    d = d.replace(/<a:srcRect[^>]*\/>/g, '');
    if (crop) d = d.replace(/(<a:blip\b[^>]*?\/>|<a:blip\b[^>]*>[\s\S]*?<\/a:blip>)/, (x) => x + crop);
    count++;
    return d;
  };

  xml = xml.replace(PARA_RE, (p) => {
    if (!p.includes('<wp:inline') || p.includes('<w:txbxContent')) return p;
    p = p.replace(/<wp:inline\b[\s\S]*?<\/wp:inline>/g, fixDrawing);
    const hasText = /<w:t(?:\s[^>]*)?>\s*[^<\s][^<]*<\/w:t>/.test(p);
    if (!hasText && !/<w:sectPr/.test(p)) {
      if (/<w:pPr>[\s\S]*?<\/w:pPr>/.test(p)) p = p.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, PPR_IMG);
      else p = p.replace(/^(<w:p(?:\s[^>]*)?>)/, '$1' + PPR_IMG);
    }
    return p;
  });

  zip.file('word/document.xml', xml);
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { buffer: out, count };
}

// ======================= EXCEL =======================
const ALLOWED_FN = ['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT'];
const thin = { style: 'thin', color: { argb: 'FF000000' } };
const BORDER = { top: thin, left: thin, bottom: thin, right: thin };
const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const colNumber = (L) => String(L).toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0);
const cellText = (c) => { try { return String(c.text ?? '').trim(); } catch (e) { return ''; } };
const isNum = (c) => typeof c.value === 'number' || (c.value && typeof c.value === 'object' && typeof c.value.result === 'number');
const numVal = (c) => (typeof c.value === 'number' ? c.value : c.value && typeof c.value.result === 'number' ? c.value.result : null);
const SKIP_HEAD = /^(stt|tt|số\s*tt|số thứ tự|no\.?)$/i;

// Mô tả bảng để gửi cho AI (tối đa 60 dòng đầu)
async function excelPreview(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  const lines = [`Trang tính: ${ws.name}; tổng ${ws.rowCount} dòng, ${ws.columnCount} cột`];
  for (let r = 1; r <= Math.min(ws.rowCount, 60); r++) {
    const row = ws.getRow(r), parts = [];
    for (let c = 1; c <= ws.columnCount; c++) { const t = cellText(row.getCell(c)); if (t) parts.push(`${colLetter(c)}=${t}`); }
    if (parts.length) lines.push(`Dòng ${r}: ${parts.join(' | ')}`);
  }
  return lines.join('\n');
}

const SPEC_PROMPT = `Bạn nhận bảng Excel và yêu cầu tính toán của NHÂN. CHỈ trả về JSON hợp lệ:
{
  "header_rows": số dòng tiêu đề tính từ dòng 1 (mặc định 1; dòng cuối của phần này là dòng tên cột),
  "data_end_row": dòng dữ liệu cuối cùng (không tính dòng tổng có sẵn nếu có),
  "new_columns": [ {"header": "tên cột mới", "formula": "công thức cho từng dòng, dùng {r} thay số dòng, ví dụ =C{r}*D{r}"} ],
  "summary": [ {"label": "Tổng cộng", "label_col": "A", "cells": {"E": "SUM", "F": "AVERAGE"}} ],
  "orientation": "doc hoặc ngang hoặc rỗng nếu NHÂN không nêu"
}
Quy tắc: chỉ dùng hàm SUM, AVERAGE, MAX, MIN, COUNT trong "summary". Cột mới nằm sau cột cuối hiện có, nên cột tổng cho cột mới dùng chữ cái cột kế tiếp. Chỉ làm đúng điều NHÂN yêu cầu, không tự thêm. Không có gì cần thêm thì để mảng rỗng. Không dùng markdown.`;

async function excelSpecFromAI(callGemini, preview, request) {
  const res = await callGemini(
    [{ role: 'user', parts: [{ text: `${SPEC_PROMPT}\n\nYÊU CẦU CỦA NHÂN:\n${request}\n\nBẢNG:\n${preview}` }] }],
    { system: 'Bạn là chuyên viên Excel. Chỉ trả về JSON.', json: true }
  );
  const raw = (res.text || '').replace(/```json|```/g, '').trim();
  try { return JSON.parse(raw); } catch (e) { const m = raw.match(/\{[\s\S]*\}/); if (m) return JSON.parse(m[0]); throw new Error('AI trả về sai định dạng'); }
}

const SAFE_FORMULA = /^=?[A-Za-z0-9_+\-*/^().,:$ %<>=&"{}]+$/;

// opts: { orient: 'doc'|'ngang'|null, tinhToan: bool, spec: object|null }
// Không có spec (không có yêu cầu tính toán) thì chỉ căn chỉnh bảng, KHÔNG thêm cột/dòng tổng.
async function processExcel(buf, opts = {}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  wb.calcProperties = { fullCalcOnLoad: true };
  const info = [];

  wb.worksheets.forEach((ws, idx) => {
    if (ws.rowCount < 1 || ws.columnCount < 1) return;
    const spec = idx === 0 ? opts.spec : null;
    const hr = Math.max(1, Math.min(10, Number(spec && spec.header_rows) || 1));

    // dòng cuối có dữ liệu
    let lastRow = ws.rowCount;
    while (lastRow > hr && !ws.getRow(lastRow).hasValues) lastRow--;
    let dataEnd = Math.min(lastRow, Number(spec && spec.data_end_row) || lastRow);
    let lastCol = ws.columnCount;

    if (opts.tinhToan) {
      // 1) cột mới theo công thức
      for (const nc of (spec && spec.new_columns) || []) {
        if (!nc || !nc.formula || !SAFE_FORMULA.test(String(nc.formula))) continue;
        lastCol++;
        ws.getCell(hr, lastCol).value = String(nc.header || '');
        for (let r = hr + 1; r <= dataEnd; r++) {
          ws.getCell(r, lastCol).value = { formula: String(nc.formula).replace(/\{r\}/g, r).replace(/^=/, '') };
        }
      }
      // 2) dòng tổng hợp
      const summary = (spec && spec.summary) || []; // chỉ làm đúng yêu cầu của NHÂN, không tự thêm dòng tổng
      let r = dataEnd + 1;
      for (const s of summary || []) {
        if (!s || !s.cells) continue;
        // chèn ngay sau dữ liệu; nếu đã có dòng tổng cũ thì ghi đè vị trí đó
        ws.getCell(r, colNumber(s.label_col || 'A')).value = String(s.label || '');
        for (const [L, fnRaw] of Object.entries(s.cells)) {
          const fn = String(fnRaw).toUpperCase();
          if (!ALLOWED_FN.includes(fn)) continue;
          const c = colNumber(L);
          if (!c) continue;
          const vals = [];
          for (let k = hr + 1; k <= dataEnd; k++) { const v = numVal(ws.getCell(k, c)); if (v !== null) vals.push(v); }
          let result;
          if (vals.length) {
            result = { SUM: vals.reduce((a, b) => a + b, 0), AVERAGE: vals.reduce((a, b) => a + b, 0) / vals.length, MAX: Math.max(...vals), MIN: Math.min(...vals), COUNT: vals.length }[fn];
          }
          const cell = ws.getCell(r, c);
          cell.value = { formula: `${fn}(${L}${hr + 1}:${L}${dataEnd})`, ...(result !== undefined ? { result } : {}) };
          cell.numFmt = fn === 'AVERAGE' ? '#,##0.00' : '#,##0';
        }
        r++;
      }
      lastRow = Math.max(lastRow, r - 1);
    }

    // ---- Kẻ viền, tiêu đề, độ rộng cột ----
    let sumWidth = 0;
    for (let c = 1; c <= lastCol; c++) {
      let maxLen = 6;
      for (let r = 1; r <= lastRow; r++) {
        const cell = ws.getCell(r, c);
        const isMergedTitle = r < hr || cell.isMerged;
        if (r >= 1 && cell.value !== null && cell.value !== undefined && cell.value !== '') {
          cell.border = BORDER;
          cell.alignment = {
            vertical: 'middle',
            wrapText: true,
            horizontal: r <= hr ? 'center' : (cell.alignment && cell.alignment.horizontal) || undefined,
          };
          if (r <= hr) cell.font = { ...(cell.font || {}), bold: true };
          if (r === hr) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
        }
        if (!isMergedTitle) {
          const t = typeof cell.value === 'object' && cell.value && cell.value.formula ? '00,000,000' : cellText(cell);
          const longest = t.split('\n').reduce((a, b) => Math.max(a, b.length), 0);
          maxLen = Math.max(maxLen, longest);
        }
      }
      const w = Math.min(Math.max(maxLen + 2, 8), 45);
      ws.getColumn(c).width = w;
      sumWidth += w;
    }

    // ---- Thiết lập trang in: A4, vừa chiều rộng, không mất cột/mất trang ----
    let orient = opts.orient || (spec && /ngang|landscape/i.test(spec.orientation || '') ? 'ngang' : spec && /doc|đứng/i.test(spec.orientation || '') ? 'doc' : null);
    if (!orient) orient = lastCol >= 7 || sumWidth > 100 ? 'ngang' : 'doc';
    ws.pageSetup = {
      paperSize: 9,
      orientation: orient === 'ngang' ? 'landscape' : 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
      printArea: `A1:${colLetter(lastCol)}${lastRow}`,
      printTitlesRow: `1:${hr}`,
    };
    info.push(`${ws.name}: ${lastRow} dòng, ${lastCol} cột, A4 ${orient === 'ngang' ? 'ngang' : 'đứng'}`);
  });

  const out = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer: out, info: info.join('; ') };
}

module.exports = { parseOfficeCommand, fixWordImages, excelPreview, excelSpecFromAI, processExcel };
