'use strict';
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

// ===== NHẬN DIỆN LỆNH =====
const CMD_ALIAS = {
  wordvanbanhinhanh: 'wordvanbanhinhanh',
  wordhvanbanhinhanh: 'wordvanbanhinhanh',
  excel: 'excel',
  exceltinhtoan: 'exceltinhtoan',
};

function parseOfficeCommand(text) {
  const m = String(text || '').trim().match(/^\/([A-Za-z0-9_]+)(?:@\w+)?(?:\s+([\s\S]*))?$/);
  if (!m) return null;
  const cmd = CMD_ALIAS[m[1].toLowerCase()];
  if (!cmd) return null;
  return { cmd, args: (m[2] || '').trim() };
}

// ===== WORD: CHỈNH ẢNH + KHỔ GIẤY, LỀ THEO NGHỊ ĐỊNH 30 =====
const EMU_W = 3600000; // 10 cm
const EMU_H = 2700000; // 7,5 cm

function insertBefore(inner, xml, tags) {
  const re = new RegExp('<w:(?:' + tags.join('|') + ')[\\s/>]');
  const m = inner.match(re);
  if (!m) return inner + xml;
  return inner.slice(0, m.index) + xml + inner.slice(m.index);
}
const AFTER_IND = ['contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const AFTER_JC = ['textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

function centerParagraph(p) {
  const open = p.match(/^<w:p(?:\s[^>]*)?>/)[0];
  let rest = p.slice(open.length).replace(/^<w:pPr\s*\/>/, '');
  const IND = '<w:ind w:left="0" w:right="0" w:firstLine="0"/>';
  const JC = '<w:jc w:val="center"/>';
  const m = rest.match(/^<w:pPr>([\s\S]*?)<\/w:pPr>/);
  if (!m) return open + '<w:pPr>' + IND + JC + '</w:pPr>' + rest;
  let inner = m[1].replace(/<w:ind\b[^>]*\/>/g, '').replace(/<w:jc\b[^>]*\/>/g, '');
  inner = insertBefore(inner, IND, AFTER_IND);
  inner = insertBefore(inner, JC, AFTER_JC);
  return open + '<w:pPr>' + inner + '</w:pPr>' + rest.slice(m[0].length);
}

async function fixWordImages(buf) {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file('word/document.xml');
  if (!f) throw new Error('File Word không hợp lệ');
  let xml = await f.async('string');
  let count = 0;

  xml = xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (p) => {
    if (!p.includes('<wp:inline')) return p;
    let n = 0;
    let out = p.replace(/<wp:inline\b[\s\S]*?<\/wp:inline>/g, (blk) => {
      const e = blk.match(/<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/);
      if (!e) return blk;
      const cx = +e[1];
      const cy = +e[2];
      if (!cx || !cy) return blk;
      const s = Math.min(EMU_W / cx, EMU_H / cy); // giữ nguyên tỉ lệ, không méo
      const nx = Math.round(cx * s);
      const ny = Math.round(cy * s);
      n++;
      return blk
        .replace(/<wp:extent\s+cx="\d+"\s+cy="\d+"\s*\/>/, `<wp:extent cx="${nx}" cy="${ny}"/>`)
        .replace(/<a:ext\s+cx="\d+"\s+cy="\d+"\s*\/>/g, `<a:ext cx="${nx}" cy="${ny}"/>`);
    });
    if (!n) return p;
    count += n;
    const hasText = /<w:t(?:\s[^>]*)?>[^<]+<\/w:t>/.test(out);
    if (!hasText) out = centerParagraph(out); // đoạn chỉ có ảnh thì căn giữa
    return out;
  });

  // Khổ A4 đứng, lề: trên 2cm, dưới 2cm, trái 3cm, phải 1,5cm (không đổi nếu văn bản đang để giấy ngang)
  xml = xml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, (s) => {
    const sz = s.match(/<w:pgSz\b[^>]*>/);
    if (sz) {
      const w = +((sz[0].match(/w:w="(\d+)"/) || [])[1] || 0);
      const h = +((sz[0].match(/w:h="(\d+)"/) || [])[1] || 0);
      if (w && h && w > h) return s;
    }
    let out = s;
    if (sz) out = out.replace(/<w:pgSz\b[^>]*>/, '<w:pgSz w:w="11906" w:h="16838"/>');
    out = out.replace(
      /<w:pgMar\b[^>]*>/,
      '<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1701" w:header="567" w:footer="567" w:gutter="0"/>'
    );
    return out;
  });

  zip.file('word/document.xml', xml);
  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { buffer, count };
}

// ===== EXCEL: HÀM DÙNG CHUNG =====
const pad = (n) => String(n).padStart(2, '0');

function cellText(c) {
  const v = c.value;
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return `${pad(v.getUTCDate())}/${pad(v.getUTCMonth() + 1)}/${v.getUTCFullYear()}`;
  if (typeof v === 'object' && (v.formula || v.sharedFormula)) {
    const r = v.result === undefined || v.result === null ? '' : ' -> ' + v.result;
    return '=' + (v.formula || '(công thức dùng chung)') + r;
  }
  try {
    return String(c.text === undefined || c.text === null ? '' : c.text);
  } catch (e) {
    return '';
  }
}

function rowHasValue(row) {
  let has = false;
  row.eachCell({ includeEmpty: false }, (c) => {
    if (cellText(c) !== '') has = true;
  });
  return has;
}

function usedRange(ws) {
  let lastRow = 0;
  let lastCol = 0;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (c, col) => {
      if (cellText(c) !== '') {
        lastRow = Math.max(lastRow, r);
        lastCol = Math.max(lastCol, col);
      }
    });
  });
  return { lastRow, lastCol };
}

// ===== EXCEL: CĂN CHỈNH BẢNG THEO KHỔ A4 =====
function findHeaderRow(ws, lastCol) {
  const need = Math.max(2, Math.ceil(lastCol * 0.6));
  const maxR = Math.min(ws.rowCount, 30);
  for (let r = 1; r <= maxR; r++) {
    let cnt = 0;
    ws.getRow(r).eachCell({ includeEmpty: false }, (c) => {
      if (c.isMerged && c.master !== c) return;
      if (cellText(c) !== '') cnt++;
    });
    if (cnt >= need) return r;
  }
  return 1;
}

function formatSheet(ws, orientOpt) {
  const { lastRow, lastCol } = usedRange(ws);
  if (!lastRow || !lastCol) return null;
  const head = findHeaderRow(ws, lastCol);
  let end = head;
  for (let r = head + 1; r <= lastRow; r++) {
    if (!rowHasValue(ws.getRow(r))) break; // dòng trống: hết bảng
    end = r;
  }

  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const sttCol = /^(stt|tt|số\s*tt|số thứ tự)$/i.test(cellText(ws.getCell(head, 1)).trim());

  for (let r = head; r <= end; r++) {
    for (let c = 1; c <= lastCol; c++) {
      const cell = ws.getCell(r, c);
      cell.border = border;
      if (r === head) {
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.font = { ...(cell.font || {}), bold: true };
      } else {
        const v = cell.value;
        const isDate = v instanceof Date;
        const isNum =
          typeof v === 'number' ||
          (v && typeof v === 'object' && v.formula !== undefined && typeof v.result === 'number');
        let h = 'left';
        if (c === 1 && sttCol) h = 'center';
        else if (isDate) h = 'center';
        else if (isNum) h = 'right';
        cell.alignment = { horizontal: h, vertical: 'middle', wrapText: true };
      }
    }
  }

  // Độ rộng cột theo nội dung
  for (let c = 1; c <= lastCol; c++) {
    let w = 0;
    let hw = 0;
    for (let r = head; r <= end; r++) {
      const cell = ws.getCell(r, c);
      if (cell.isMerged) continue;
      const len = Math.max(0, ...cellText(cell).split('\n').map((s) => s.length));
      if (r === head) hw = Math.max(hw, len);
      else w = Math.max(w, len);
    }
    const min = c === 1 && sttCol ? 5 : 6;
    ws.getColumn(c).width = Math.min(40, Math.max(min, Math.max(w, Math.min(hw, 14)) * 1.15 + 2));
  }

  // Chiều cao dòng khi chữ xuống dòng
  for (let r = head; r <= end; r++) {
    const row = ws.getRow(r);
    let need = 0;
    for (let c = 1; c <= lastCol; c++) {
      const cell = ws.getCell(r, c);
      if (cell.isMerged) continue;
      const wcol = ws.getColumn(c).width || 8.43;
      const lines = cellText(cell)
        .split('\n')
        .reduce((a, s) => a + Math.max(1, Math.ceil(s.length / Math.max(1, wcol - 1))), 0);
      const size = (cell.font && cell.font.size) || 11;
      need = Math.max(need, lines > 1 ? lines * size * 1.35 : 0);
    }
    if (need) row.height = Math.max(need, row.height || 0);
  }

  // Khổ giấy A4
  let total = 0;
  for (let c = 1; c <= lastCol; c++) total += ws.getColumn(c).width || 8.43;
  const land = orientOpt === 'ngang' ? true : orientOpt === 'doc' ? false : total > 90;
  ws.pageSetup = {
    paperSize: 9,
    orientation: land ? 'landscape' : 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: { left: 0.79, right: 0.59, top: 0.79, bottom: 0.79, header: 0.3, footer: 0.3 },
  };
  ws.pageSetup.printArea = `A1:${ws.getColumn(lastCol).letter}${lastRow}`;
  if (end - head > 35) ws.pageSetup.printTitlesRow = `${head}:${head}`;
  return { land };
}

// ===== EXCEL: TÍNH TOÁN THEO LỆNH TIẾNG VIỆT =====
function getWs(wb, name, create) {
  const found = name ? wb.getWorksheet(String(name)) : null;
  if (found) return found;
  if (create && name) return wb.addWorksheet(String(name).slice(0, 31));
  return wb.worksheets[0];
}

function cloneStyle(from, to) {
  try {
    to.style = JSON.parse(JSON.stringify(from.style || {}));
  } catch (e) {
    /* bỏ qua */
  }
}

const cleanFormula = (f) => String(f).trim().replace(/^=/, '');

function applySpec(wb, spec) {
  const ops = Array.isArray(spec && spec.ops) ? spec.ops : [];
  let n = 0;
  for (const op of ops) {
    try {
      const ws = getWs(wb, op.sheet, op.create);
      if (!ws) continue;
      if (op.type === 'fill' && op.col && op.formula) {
        const from = Math.max(1, parseInt(op.from, 10) || 2);
        const to = Math.min(from + 5000, parseInt(op.to, 10) || from);
        const colNum = ws.getColumn(String(op.col).toUpperCase()).number;
        if (op.header && from > 1) {
          const hc = ws.getCell(from - 1, colNum);
          hc.value = String(op.header);
          if (colNum > 1) cloneStyle(ws.getCell(from - 1, colNum - 1), hc);
        }
        for (let r = from; r <= to; r++) {
          const cell = ws.getCell(r, colNum);
          cell.value = { formula: cleanFormula(op.formula).replace(/\{r\}/g, String(r)) };
          if (colNum > 1) cloneStyle(ws.getCell(r, colNum - 1), cell);
          if (op.numFmt) cell.numFmt = String(op.numFmt);
        }
        n++;
      } else if (op.type === 'cell' && op.cell) {
        const cell = ws.getCell(String(op.cell).toUpperCase());
        if (op.formula) cell.value = { formula: cleanFormula(op.formula) };
        else if (op.value !== undefined) cell.value = op.value;
        else continue;
        if (cell.col > 1) {
          const left = ws.getCell(cell.row, cell.col - 1);
          if (left.border && left.border.top && !(cell.border && cell.border.top)) cloneStyle(left, cell);
        }
        if (op.bold) cell.font = { ...(cell.font || {}), bold: true };
        if (op.numFmt) cell.numFmt = String(op.numFmt);
        n++;
      }
    } catch (e) {
      console.warn('Bỏ qua 1 thao tác tính toán:', String((e && e.message) || e).slice(0, 100));
    }
  }
  if (!n) throw new Error('Em chưa tạo được phép tính nào từ yêu cầu này, NHÂN nêu rõ hơn giúp em (tên cột, vị trí đặt kết quả)');
  return n;
}

// ===== HÀM XUẤT RA CHO index.js =====
async function processExcel(buf, opts = {}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  let info;
  if (opts.tinhToan) {
    const n = applySpec(wb, opts.spec || {});
    wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
    info = `Đã thực hiện ${n} thao tác tính toán${opts.spec && opts.spec.note ? ': ' + opts.spec.note : ''}. NHÂN mở bằng Excel để xem kết quả`;
  } else {
    let done = 0;
    const lands = new Set();
    wb.eachSheet((ws) => {
      const r = formatSheet(ws, opts.orient);
      if (r) {
        done++;
        lands.add(r.land ? 'ngang' : 'đứng');
      }
    });
    if (!done) throw new Error('File Excel không có dữ liệu để căn chỉnh');
    info = `Đã kẻ bảng, căn chỉnh và đặt khổ A4 ${[...lands].join('/')}, vừa 1 trang theo chiều ngang (${done} trang tính). Lề trên, dưới, trái 2 cm, phải 1,5 cm`;
  }
  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  return { buffer, info };
}

async function excelPreview(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const lines = [];
  wb.eachSheet((ws) => {
    const { lastRow, lastCol } = usedRange(ws);
    lines.push(`### Trang tính "${ws.name}" (dòng cuối có dữ liệu: ${lastRow}, cột cuối: ${lastCol})`);
    const show = (r) => {
      const cells = [];
      ws.getRow(r).eachCell({ includeEmpty: false }, (c) => {
        const t = cellText(c);
        if (t !== '') cells.push(`${c.address}=${t.slice(0, 60)}`);
      });
      if (cells.length) lines.push(`Dòng ${r}: ` + cells.join(' | '));
    };
    const head = Math.min(lastRow, 80);
    for (let r = 1; r <= head; r++) show(r);
    if (lastRow > head) {
      lines.push('... (lược bớt các dòng giữa) ...');
      for (let r = Math.max(head + 1, lastRow - 2); r <= lastRow; r++) show(r);
    }
  });
  return lines.join('\n').slice(0, 30000);
}

const SPEC_PROMPT = `Bạn nhận nội dung một bảng Excel (liệt kê ô theo địa chỉ) và yêu cầu tính toán bằng tiếng Việt của NHÂN. Hãy chuyển yêu cầu thành các thao tác ghi công thức Excel vào ô. CHỈ trả về JSON hợp lệ, đúng cấu trúc:
{"ops":[ ... ],"note":"một câu ngắn nêu đã làm gì"}
Các loại thao tác:
1) Điền công thức xuống một cột: {"type":"fill","sheet":"tên trang tính","col":"F","from":2,"to":15,"formula":"D{r}*E{r}","header":"Thành tiền","numFmt":"#,##0"}
   ({r} là số dòng. "header" là tiêu đề cột mới, đặt ở dòng from-1; bỏ đi nếu cột đã có tiêu đề.)
2) Ghi một ô: {"type":"cell","sheet":"tên trang tính","cell":"F16","formula":"SUM(F2:F15)","bold":true,"numFmt":"#,##0"}
   hoặc ghi chữ: {"type":"cell","sheet":"tên trang tính","cell":"E16","value":"Cộng","bold":true}
   Muốn lập bảng kết quả ở trang tính mới thì thêm "create":true và đặt tên trang tính mới cho các thao tác đó.
Quy tắc:
- Công thức viết bằng tên hàm tiếng Anh (SUM, SUMIF, COUNTIF, COUNTA, AVERAGE, IF, DATEDIF, DATE, YEAR, ROUND, ...), ngăn cách đối số bằng dấu phẩy.
- Chỉ làm đúng và chỉ đúng điều NHÂN yêu cầu. Không sửa, không xóa dữ liệu gốc. Chỉ ghi vào ô trống hoặc cột mới. Không tự thêm cột, dòng tổng hay thống kê ngoài yêu cầu.
- Dùng đúng địa chỉ ô, đúng dòng bắt đầu và dòng cuối dữ liệu theo bảng cung cấp. Không bịa số liệu.
- Mốc tuổi Đảng được trao Huy hiệu: 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95 năm; các đợt trao tặng: 03/02, 30/04, 02/09, 07/11 hằng năm. Nếu yêu cầu liên quan, đối chiếu đúng các mốc này.
- Nếu yêu cầu chưa rõ hoặc không xác định được cột/dòng cần tính thì chỉ trả về {"thieu":"một câu hỏi ngắn gửi NHÂN"}.`;

async function excelSpecFromAI(callGemini, preview, args) {
  const today = new Date().toLocaleDateString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  const prompt = `${SPEC_PROMPT}\n\nHôm nay là ${today}.\n\nNỘI DUNG BẢNG EXCEL:\n${preview}\n\nYÊU CẦU CỦA NHÂN:\n${args}`;
  const res = await callGemini([{ role: 'user', parts: [{ text: prompt }] }], {
    system: 'Bạn là chuyên viên xử lý bảng tính Excel. Chỉ trả về JSON.',
    json: true,
  });
  const raw = (res.text || '').replace(/```json|```/g, '').trim();
  let spec;
  try {
    spec = JSON.parse(raw);
  } catch (e) {
    const m = raw.match(/\{[\s\S]*\}/);
    if (!m) throw new Error('AI trả về sai định dạng, NHÂN gửi lại yêu cầu giúp em');
    spec = JSON.parse(m[0]);
  }
  if (spec.thieu) throw new Error(String(spec.thieu));
  return spec;
}

module.exports = { fixWordImages, parseOfficeCommand, excelPreview, excelSpecFromAI, processExcel };
