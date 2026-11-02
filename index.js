require('dotenv').config();
const http = require('http');
const https = require('https');
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
} = require('docx');
const office = require('./office');

// ===== CẤU HÌNH =====
const TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const GEMINI_KEY = process.env.GEMINI_API_KEY;
// Ưu tiên model Flash-Lite vì gói miễn phí cho nhiều lượt/ngày hơn (khoảng 500 so với 20)
const MODELS = ['gemini-3.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-3.8-flash', 'gemini-3.6-flash'];
const ALLOWED_ID = Number(process.env.ALLOWED_USER_ID || 0);
const PORT = process.env.PORT || 10000;
const DOMAIN = process.env.RENDER_EXTERNAL_URL;
const HOOK_PATH = `/webhook/${TOKEN}`;
const MAX_TEXT = 400000; // giới hạn ký tự đọc từ 1 file

if (!TOKEN || !GEMINI_KEY) {
  console.error('Thiếu TELEGRAM_BOT_TOKEN hoặc GEMINI_API_KEY trong Environment.');
  process.exit(1);
}

const bot = new Telegraf(TOKEN, {
  handlerTimeout: 180000,
  // Không giữ kết nối cũ, tránh lỗi "socket hang up" khi gửi file về Telegram
  telegram: { agent: new https.Agent({ keepAlive: false }) },
});
const ai = new GoogleGenAI({ apiKey: GEMINI_KEY });

function buildSystem() {
  const now = new Date().toLocaleString('vi-VN', {
    timeZone: 'Asia/Ho_Chi_Minh',
    dateStyle: 'full',
    timeStyle: 'short',
  });
  return `Bạn là "tt" - trợ lý ảo cá nhân của NHÂN (Trịnh Hoài Nhân), chuyên về quản trị kinh doanh, hồ sơ, công văn hành chính, báo cáo, lập kế hoạch và lên ý tưởng. Bây giờ là ${now} (giờ Việt Nam).
Quy tắc:
- Luôn gọi người dùng là NHÂN (không dùng "anh" hay "chị"), xưng "em". Trả lời gọn, trực diện, không mở đầu sáo rỗng.
- Mặc định (khi NHÂN chỉ hỏi, nhờ tóm tắt hoặc trao đổi thông thường): trả lời và tóm tắt bình thường, ngắn gọn, không dàn trang thể thức văn bản. Chỉ khi NHÂN yêu cầu rõ "trình bày theo thể thức văn bản" (hoặc dùng lệnh /vanban) thì hệ thống mới tạo file Word đúng thể thức, em không cần tự dàn trang trong tin nhắn.
- Câu hỏi kiến thức chung, tư vấn: trả lời đầy đủ. Bạn KHÔNG tra cứu được internet, nên với thời tiết, tin tức, giá cả, quy định mới nhất, hãy nói rõ em chưa tra cứu được thông tin thời gian thực và khuyên NHÂN kiểm tra nguồn chính thức, không tự bịa số liệu.
- Với hồ sơ, số liệu, tài liệu NHÂN cung cấp: bám sát 100%, không tự thêm bớt số liệu, không suy diễn thiếu căn cứ. Thiếu thông tin thì hỏi lại.
- Soạn văn bản hành chính theo thể thức Nghị định 30/2020/NĐ-CP, bố cục rõ ràng.
- Huy hiệu Đảng: các mốc 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95 năm tuổi Đảng; các đợt trao tặng 03/02, 30/04, 02/09, 07/11 hằng năm. Khi tính toán phải đối chiếu đúng mốc.
- Không dùng ký hiệu markdown (như ** hoặc #). Viết văn bản thuần, liệt kê bằng dấu gạch đầu dòng "-".
- Khi soạn một văn bản hoàn chỉnh (công văn, báo cáo, quyết định, tờ trình, thông báo, kế hoạch, biên bản...): lời dẫn hoặc ghi chú (nếu có) viết ngắn gọn NGOÀI văn bản; toàn bộ nội dung văn bản đặt giữa hai dòng đánh dấu [[VAN_BAN]] và [[HET_VAN_BAN]] (mỗi dấu nằm trên một dòng riêng). Bên trong hai dấu này chỉ có nội dung văn bản, không có lời dẫn, ghi chú hay câu hỏi. Khi NHÂN yêu cầu chỉnh sửa, luôn gửi lại toàn văn đã chỉnh trong hai dấu này. Câu trả lời thông thường thì không dùng dấu này.`;
}

const history = new Map(); // chatId -> [{role, parts}]
const lastReply = new Map(); // chatId -> câu trả lời gần nhất
const lastDoc = new Map(); // chatId -> văn bản soạn gần nhất (chỉ phần nội dung văn bản)

// ===== CHỈ CHO PHÉP NHÂN DÙNG =====
bot.use((ctx, next) => {
  if (ALLOWED_ID && ctx.from?.id !== ALLOWED_ID) {
    return ctx.reply('Bot riêng, không phục vụ người ngoài.');
  }
  return next();
});

// ===== XỬ LÝ LỖI =====
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const errText = (e) => String((e && e.message) || e);
const isModelGone = (m) => /404|NOT_FOUND|no longer available/i.test(m);
const isTransient = (m) =>
  /503|UNAVAILABLE|high demand|429|RESOURCE_EXHAUSTED|overloaded|500|INTERNAL|socket hang up|ECONNRESET|ETIMEDOUT|fetch failed|timeout/i.test(m);

function friendlyError(e) {
  const m = errText(e);
  if (/429|RESOURCE_EXHAUSTED/i.test(m)) return 'Hạn mức AI đang tạm hết, NHÂN thử lại sau ít phút nhé.';
  if (isTransient(m)) return 'Hệ thống AI đang quá tải, em đã thử lại nhiều lần chưa được. NHÂN nhắn lại sau 1 đến 2 phút nhé.';
  if (/API key|API_KEY|PERMISSION_DENIED|403/i.test(m)) return 'Khóa GEMINI_API_KEY có vấn đề, NHÂN kiểm tra lại trên Render nhé.';
  return 'Em gặp lỗi: ' + m.slice(0, 200);
}

// ===== GỌI AI (tự thử lại, tự đổi model, tự bật/tắt Google Search) =====
// opts.system: thay lời chỉ dẫn hệ thống; opts.json: bắt AI trả về JSON
async function callGemini(contents, opts = {}) {
  let useSearch = false; // đổi thành true để bật Google Search (cần tài khoản có thanh toán)
  let lastErr;
  for (const model of MODELS) {
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        return await ai.models.generateContent({
          model,
          contents,
          config: {
            systemInstruction: opts.system || buildSystem(),
            ...(opts.json ? { responseMimeType: 'application/json' } : {}),
            ...(useSearch && !opts.json ? { tools: [{ googleSearch: {} }] } : {}),
          },
        });
      } catch (e) {
        lastErr = e;
        const m = errText(e);
        if (useSearch && /search|grounding|tool/i.test(m) && !/503|UNAVAILABLE|high demand/i.test(m)) {
          useSearch = false; // tắt tìm kiếm nếu không dùng được, thử lại ngay
          attempt--;
          console.warn('Tắt Google Search:', m.slice(0, 120));
          continue;
        }
        if (/429|RESOURCE_EXHAUSTED/i.test(m)) {
          console.warn(`Model ${model} hết hạn mức, đổi model.`);
          break; // không thử lại cùng model để khỏi tốn thêm lượt
        }
        if (isModelGone(m)) {
          console.warn(`Model ${model} không dùng được, đổi model.`);
          break;
        }
        if (isTransient(m)) {
          console.warn(`Model ${model} lỗi tạm thời (lần ${attempt}):`, m.slice(0, 120));
          await sleep(attempt * 1500);
          continue;
        }
        throw e;
      }
    }
  }
  throw lastErr;
}

async function ask(chatId, parts, savedText) {
  const h = history.get(chatId) || [];
  const res = await callGemini([...h, { role: 'user', parts }]);
  const text = res.text || 'Em chưa có phản hồi, NHÂN thử lại giúp em.';
  h.push({ role: 'user', parts: [{ text: savedText }] });
  h.push({ role: 'model', parts: [{ text }] });
  history.set(chatId, h.slice(-20));

  // Tách riêng phần văn bản (giữa hai dấu) để xuất Word chỉ lấy nội dung văn bản
  let shown = text;
  const m = text.match(/\[\[VAN_BAN\]\]([\s\S]*?)(?:\[\[HET_VAN_BAN\]\]|$)/);
  if (m && m[1].trim()) {
    lastDoc.set(chatId, m[1].trim());
    shown = text.replace(/\[\[(?:HET_)?VAN_BAN\]\]/g, '').replace(/\n{3,}/g, '\n\n').trim();
  }
  lastReply.set(chatId, shown);
  return shown;
}

async function sendLong(ctx, text) {
  for (let i = 0; i < text.length; i += 4000) {
    await ctx.reply(text.slice(i, i + 4000));
  }
}

// Giữ dòng "đang nhập..." suốt lúc AI xử lý
async function withTyping(ctx, action, fn) {
  await ctx.sendChatAction(action).catch(() => {});
  const timer = setInterval(() => ctx.sendChatAction(action).catch(() => {}), 4000);
  try {
    return await fn();
  } finally {
    clearInterval(timer);
  }
}

async function downloadFile(ctx, fileId) {
  const link = await ctx.telegram.getFileLink(fileId);
  const r = await fetch(link.href);
  if (!r.ok) throw new Error('Không tải được file từ Telegram (' + r.status + ')');
  return Buffer.from(await r.arrayBuffer());
}

// Gửi file trực tiếp qua API Telegram bằng fetch có sẵn của Node (ổn định hơn khi mạng yếu)
async function sendDocumentDirect(chatId, buffer, filename) {
  const form = new FormData();
  form.append('chat_id', String(chatId));
  form.append('document', new Blob([buffer]), filename);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 60000);
  try {
    const r = await fetch(`https://api.telegram.org/bot${TOKEN}/sendDocument`, {
      method: 'POST',
      body: form,
      signal: controller.signal,
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.ok) throw new Error('Telegram từ chối: ' + (j.description || r.status));
  } finally {
    clearTimeout(timer);
  }
}

const clip = (s) => (s.length > MAX_TEXT ? s.slice(0, MAX_TEXT) + '\n[...đã cắt bớt do quá dài]' : s);

// Đọc .xlsx bằng ExcelJS để dùng chung thư viện với luồng xử lý Excel.
async function excelToText(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook.worksheets.map((ws) => {
    const rows = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
      const values = row.values.slice(1).map((value) => {
        if (value === null || value === undefined) return '';
        if (value instanceof Date) return value.toLocaleDateString('vi-VN');
        if (typeof value === 'object') {
          if (Array.isArray(value.richText)) return value.richText.map((x) => x.text || '').join('');
          if (value.text !== undefined) return String(value.text);
          if (value.result !== undefined) return String(value.result ?? '');
          if (value.formula) return '=' + value.formula;
        }
        return String(value);
      });
      rows.push(values.join('\t'));
    });
    return `--- Trang tính: ${ws.name} ---\n${rows.join('\n')}`;
  }).join('\n\n');
}

// ===== XUẤT WORD THEO THỂ THỨC (Nghị định 30/2020/NĐ-CP) =====
const FONT = 'Times New Roman';
const NOB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NOBS = { top: NOB, bottom: NOB, left: NOB, right: NOB };
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 28, ...o });
const para = (children, o = {}) =>
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, ...o, children });
const cell = (w, children) =>
  new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: NOBS,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children,
  });
const hiddenTable = (widths, rows) =>
  new Table({
    width: { size: widths[0] + widths[1], type: WidthType.DXA },
    columnWidths: widths,
    borders: { ...NOBS, insideHorizontal: NOB, insideVertical: NOB },
    rows: rows.map((r) => new TableRow({ cantSplit: true, children: [cell(widths[0], r[0]), cell(widths[1], r[1])] })),
  });
const underline = (indent) =>
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    indent: { left: indent, right: indent },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
    children: [run('', { size: 4 })],
  });

function buildFormalDoc(d) {
  const loai = String(d.loai || '').trim().toUpperCase();
  const congVan = !loai;
  const kyHieu = String(d.ky_hieu || '').trim();
  const soText = kyHieu ? `Số: ....../${kyHieu}-......` : 'Số: ....../......';
  const trichYeu = String(d.trich_yeu || '').trim();

  // --- Phần đầu: bảng 2 cột ẩn viền ---
  const left1 = [para([run('')]), para([run('')]), para([run('')])]; // góc trái để trống cho NHÂN điền tên đơn vị
  const left2 = [para([run(soText, { size: 26 })], { spacing: { before: 120, after: 0 } })];
  if (congVan && trichYeu) {
    const cv = /^v\/v/i.test(trichYeu) ? trichYeu : 'V/v ' + trichYeu.replace(/^về việc\s*/i, '');
    left2.push(para([run(cv, { size: 24 })], { spacing: { before: 60, after: 0 } }));
  }
  const right1 = [
    para([run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { bold: true, size: 24 })]),
    para([run('Độc lập - Tự do - Hạnh phúc', { bold: true, size: 26 })]),
    underline(850),
  ];
  const ngay = String(d.dia_danh_ngay || '').trim() || '[Địa danh], ngày ... tháng ... năm 20...';
  const right2 = [para([run(ngay, { italics: true, size: 28 })], { spacing: { before: 120, after: 0 } })];

  const children = [hiddenTable([3600, 5187], [[left1, right1], [left2, right2]])];

  // --- Tên loại, trích yếu ---
  if (!congVan) {
    children.push(para([run(loai, { bold: true })], { spacing: { before: 360, after: 0 } }));
    if (trichYeu) {
      children.push(para([run(trichYeu, { bold: true })]));
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 240 },
          indent: { left: 2700, right: 2700 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
          children: [run('', { size: 4 })],
        })
      );
    } else {
      children.push(para([run('')], { spacing: { after: 240 } }));
    }
  } else {
    children.push(para([run('')], { spacing: { after: 240 } }));
  }

  // --- Kính gửi ---
  const kg = [].concat(d.kinh_gui || []).map(String).filter(Boolean);
  if (kg.length === 1) {
    children.push(para([run('Kính gửi: ' + kg[0])], { spacing: { after: 240 } }));
  } else if (kg.length > 1) {
    children.push(para([run('Kính gửi:')]));
    kg.forEach((x, i) =>
      children.push(para([run('- ' + x.replace(/^-\s*/, ''))], { spacing: { after: i === kg.length - 1 ? 240 : 0 } }))
    );
  }

  // --- Nội dung ---
  const body = [].concat(d.noi_dung || []).map(String);
  body.forEach((line) => {
    const t = line.trim();
    if (!t) return;
    const roman = /^[IVX]+\.\s/.test(t);
    const sub = /^\d+\.\s/.test(t) && t.length < 90 && !/[.;,:]$/.test(t);
    children.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: roman ? 120 : 0, after: 60, line: 360 },
        indent: roman ? undefined : { firstLine: 720 },
        children: [run(t, { bold: roman || sub })],
      })
    );
  });

  // --- Nơi nhận, chức vụ, chữ ký ---
  const noiNhan = [].concat(d.noi_nhan || []).map(String).filter(Boolean);
  const chucVu = String(d.chuc_vu_ky || '').trim() || '[CHỨC VỤ NGƯỜI KÝ]';
  const nguoiKy = String(d.nguoi_ky || '').trim() || '[Họ và tên]';
  const nn = [para([run('Nơi nhận:', { bold: true, italics: true, size: 24 })], { alignment: AlignmentType.LEFT })];
  noiNhan.forEach((x) =>
    nn.push(para([run(x.startsWith('-') ? x : '- ' + x, { size: 22 })], { alignment: AlignmentType.LEFT }))
  );
  const ky = [
    para([run(chucVu.toUpperCase(), { bold: true })]),
    para([run('(Ký, ghi rõ họ tên, đóng dấu)', { italics: true, size: 24 })]),
    para([run('')]),
    para([run('')]),
    para([run('')]),
    para([run(nguoiKy, { bold: true })]),
  ];
  children.push(para([run('')], { spacing: { before: 240, after: 0 } }));
  children.push(hiddenTable([4300, 4487], [[nn, ky]]));

  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 28 } } } },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, bottom: 1134, left: 1985, right: 1134 },
          },
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(doc);
}

const FORMAL_PROMPT = `Chuyển yêu cầu/nội dung dưới đây thành văn bản hành chính theo thể thức Nghị định 30/2020/NĐ-CP. CHỈ trả về JSON hợp lệ, đúng cấu trúc:
{
  "loai": "BÁO CÁO | THÔNG BÁO | KẾ HOẠCH | TỜ TRÌNH | QUYẾT ĐỊNH | BIÊN BẢN... (để chuỗi rỗng nếu là CÔNG VĂN)",
  "ky_hieu": "BC | TB | KH | TTr | QĐ | BB... (để rỗng nếu là công văn)",
  "dia_danh_ngay": "chỉ điền nếu NHÂN nêu rõ, ví dụ: Đồng Nai, ngày 05 tháng 11 năm 2026; nếu không thì để rỗng",
  "trich_yeu": "Về việc ... (báo cáo, tờ trình, kế hoạch...) hoặc V/v ... (công văn)",
  "kinh_gui": "tên nơi nhận, để rỗng nếu không cần (có thể là mảng nếu nhiều nơi nhận)",
  "noi_dung": ["từng đoạn văn, mỗi phần tử một đoạn; tiêu đề mục lớn viết dạng 'I. TÊN MỤC'"],
  "noi_nhan": ["- Như trên;", "- Lưu: VT."],
  "chuc_vu_ky": "để rỗng nếu NHÂN không nêu",
  "nguoi_ky": "để rỗng nếu NHÂN không nêu"
}
Quy tắc:
- Bám sát 100% số liệu, dữ liệu NHÂN cung cấp; không tự thêm bớt, không bịa. Chỗ thiếu thông tin ghi "......".
- Không dùng ký hiệu markdown (như ** hoặc #).
- Nếu yêu cầu có nội dung riêng thì dùng nội dung đó; nếu không có thì dùng "VĂN BẢN GẦN NHẤT" bên dưới.
- Luôn trình bày thành văn bản hoàn chỉnh theo thể thức, kể cả khi nội dung NHÂN đưa chưa phải văn bản hành chính (tin nhắn, ý tưởng, tóm tắt, khẩu hiệu, ghi chú...). Khi đó tự chọn loại văn bản phù hợp nhất (nếu NHÂN không nêu loại thì dùng BÁO CÁO), tự viết trích yếu và sắp xếp nội dung thành các mục rõ ràng, giữ nguyên ý và số liệu, không bịa thêm.
- Chỉ trả về {"thieu":"một câu hỏi ngắn gọn gửi NHÂN"} khi hoàn toàn không có bất kỳ nội dung nào, kể cả "VĂN BẢN GẦN NHẤT" cũng trống.`;

async function structureDoc(request, prev) {
  const prompt =
    `${FORMAL_PROMPT}\n\nYÊU CẦU/NỘI DUNG CỦA NHÂN:\n${clip(request || '(không có)')}\n\nVĂN BẢN GẦN NHẤT:\n${clip(prev || '(không có)')}`;
  const res = await callGemini([{ role: 'user', parts: [{ text: prompt }] }], {
    system: 'Bạn là chuyên viên văn thư, soạn văn bản hành chính Việt Nam chuẩn thể thức. Chỉ trả về JSON.',
    json: true,
  });
  const raw = (res.text || '').replace(/```json|```/g, '').trim();
  try {
    return JSON.parse(raw);
  } catch (e) {
    const m = raw.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    throw new Error('AI trả về sai định dạng');
  }
}

async function sendDocWithRetry(ctx, buffer, filename) {
  let lastErr;
  for (let i = 1; i <= 3; i++) {
    try {
      await sendDocumentDirect(ctx.chat.id, buffer, filename);
      return;
    } catch (e) {
      lastErr = e;
      console.warn(`Gửi file Word lỗi (lần ${i}):`, errText(e).slice(0, 120));
      await sleep(i * 1500);
    }
  }
  try {
    await ctx.replyWithDocument({ source: buffer, filename });
    return;
  } catch (e) {
    lastErr = e;
  }
  throw lastErr;
}

// ===== XUẤT WORD THƯỜNG (không Quốc hiệu - Tiêu ngữ, chỉ nội dung) =====
function buildPlainDoc(text) {
  const paragraphs = String(text)
    .split('\n')
    .map(
      (line) =>
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 120, line: 360 },
          children: [new TextRun({ text: line, font: FONT, size: 28 })],
        })
    );
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, bottom: 1134, left: 1985, right: 1134 },
          },
        },
        children: paragraphs,
      },
    ],
  });
  return Packer.toBuffer(doc);
}

async function makePlainDoc(ctx, request) {
  const id = ctx.chat.id;
  const text = String(request).trim() || lastDoc.get(id) || lastReply.get(id) || '';
  if (!text) return ctx.reply('Chưa có nội dung để xuất. NHÂN gửi nội dung hoặc nhờ em soạn trước nhé.');
  try {
    const buffer = await buildPlainDoc(text);
    await ctx.sendChatAction('upload_document').catch(() => {});
    await sendDocWithRetry(ctx, buffer, 'noi_dung_tt.docx');
    await ctx.reply('Em đã gửi file Word, chỉ có nội dung, không có tiêu đề.');
  } catch (e) {
    console.error('Lỗi xuất docx thường:', e);
    await ctx.reply('Chưa gửi được file Word do đường truyền, NHÂN nhắn lại lệnh sau ít giây nhé.');
  }
}

async function makeFormalDoc(ctx, request) {
  const id = ctx.chat.id;
  const prev = lastDoc.get(id) || lastReply.get(id) || '';
  if (!String(request).trim() && !prev) {
    return ctx.reply('Chưa có nội dung để trình bày. NHÂN gửi nội dung hoặc nhờ em soạn trước nhé.');
  }
  let d;
  try {
    d = await withTyping(ctx, 'upload_document', () => structureDoc(request, prev));
  } catch (e) {
    console.error('Lỗi dựng văn bản:', e);
    return ctx.reply(friendlyError(e));
  }
  if (d.thieu) return ctx.reply(String(d.thieu));
  try {
    const buffer = await buildFormalDoc(d);
    await ctx.sendChatAction('upload_document').catch(() => {});
    await sendDocWithRetry(ctx, buffer, 'van_ban_tt.docx');
    const plain = [].concat(d.noi_dung || []).join('\n');
    if (plain) lastDoc.set(id, plain);
    const h = history.get(id) || [];
    h.push({ role: 'user', parts: [{ text: request || '[Yêu cầu trình bày thể thức văn bản]' }] });
    h.push({ role: 'model', parts: [{ text: `[Đã gửi file Word thể thức: ${d.loai || 'CÔNG VĂN'} ${d.trich_yeu || ''}]\n${plain}` }] });
    history.set(id, h.slice(-20));
    await ctx.reply('Em đã gửi file Word theo thể thức. Góc trái trên để trống cho NHÂN điền tên đơn vị.');
  } catch (e) {
    console.error('Lỗi xuất docx:', e);
    await ctx.reply('Chưa gửi được file Word do đường truyền, NHÂN nhắn lại lệnh sau ít giây nhé.');
  }
}

// ===== WORD CÓ HÌNH ẢNH / EXCEL ĐỊNH DẠNG - TÍNH TOÁN =====
const pendingOffice = new Map(); // chatId -> {cmd, args, at}: NHÂN gõ lệnh trước, gửi file sau
const pendingExcelFlow = new Map(); // chatId -> {stage, buffer, filename, at, timer}
const EXCEL_FLOW_TTL = 15 * 60 * 1000;

function saveExcelFlow(chatId, data) {
  const old = pendingExcelFlow.get(chatId);
  if (old?.timer) clearTimeout(old.timer);
  const state = { ...data, at: Date.now() };
  state.timer = setTimeout(() => pendingExcelFlow.delete(chatId), EXCEL_FLOW_TTL);
  state.timer.unref?.();
  pendingExcelFlow.set(chatId, state);
  return state;
}

function readExcelFlow(chatId) {
  const state = pendingExcelFlow.get(chatId);
  if (!state) return null;
  if (Date.now() - state.at >= EXCEL_FLOW_TTL) {
    if (state.timer) clearTimeout(state.timer);
    pendingExcelFlow.delete(chatId);
    return null;
  }
  return state;
}

async function runExcelCalculation(ctx, state, request) {
  try {
    const result = await withTyping(ctx, 'upload_document', async () => {
      const preview = await office.excelPreview(state.buffer);
      const spec = await office.excelSpecFromAI(callGemini, preview, request);
      return office.processExcel(state.buffer, { tinhToan: true, spec });
    });
    saveExcelFlow(ctx.chat.id, {
      stage: 'orientation',
      buffer: result.buffer,
      filename: state.filename,
      calculated: true,
    });
    await ctx.reply(
      `Em đã lập phép tính theo yêu cầu (${result.info}). NHÂN muốn đặt giấy đứng hay ngang? Gửi /excel đứng hoặc /excel ngang để hoàn tất. Excel sẽ tính kết quả khi NHÂN mở file.`
    );
  } catch (e) {
    console.error('Lỗi tính toán Excel theo hội thoại:', e);
    await ctx.reply(isTransient(errText(e)) ? friendlyError(e) : 'Em chưa thực hiện được phép tính: ' + errText(e).slice(0, 180));
  }
}

async function finishExcelOrientation(ctx, state, args) {
  const orient = /^(đứng|dọc|portrait)$/i.test(String(args).trim())
    ? 'doc'
    : /^(ngang|landscape)$/i.test(String(args).trim())
      ? 'ngang'
      : null;
  if (!orient) return ctx.reply('NHÂN chọn hướng giấy bằng /excel đứng hoặc /excel ngang nhé.');
  try {
    const result = await withTyping(ctx, 'upload_document', () =>
      office.processExcel(state.buffer, { orient, tinhToan: false })
    );
    await sendDocWithRetry(
      ctx,
      result.buffer,
      (state.filename || 'bang_tinh.xlsx').replace(/\.xlsx$/i, '') + '_da_xu_ly.xlsx'
    );
    pendingExcelFlow.delete(ctx.chat.id);
    if (state.timer) clearTimeout(state.timer);
    await ctx.reply(`Em đã hoàn tất bảng Excel khổ giấy ${orient === 'ngang' ? 'ngang' : 'đứng'}. ${result.info}`);
  } catch (e) {
    console.error('Lỗi định dạng Excel theo hướng giấy:', e);
    await ctx.reply(isTransient(errText(e)) ? friendlyError(e) : 'Em chưa định dạng được file: ' + errText(e).slice(0, 180));
  }
}

async function handleExcelFlowText(ctx, text) {
  const state = readExcelFlow(ctx.chat.id);
  if (!state) return false;
  const t = String(text || '').trim();
  if (state.stage === 'orientation') {
    await ctx.reply('NHÂN gửi /excel đứng hoặc /excel ngang để nhận file hoàn chỉnh nhé.');
    return true;
  }
  if (state.stage === 'calculation_request') {
    await runExcelCalculation(ctx, state, t);
    return true;
  }

  const yes = /^(có|co|có nhé|đồng ý|muốn|có,?.*)$/i.test(t);
  const no = /^(không|khong|không cần|không tính(?: toán)?|không,?.*)$/i.test(t);
  if (no) {
    saveExcelFlow(ctx.chat.id, { ...state, stage: 'orientation' });
    await ctx.reply('Dạ. NHÂN chọn hướng giấy bằng /excel đứng hoặc /excel ngang để em căn chỉnh và gửi file nhé.');
    return true;
  }
  if (yes) {
    const request = t.replace(/^(có|co|có nhé|đồng ý|muốn)\s*[, :]*/i, '').trim();
    if (request) {
      await runExcelCalculation(ctx, state, request);
    } else {
      saveExcelFlow(ctx.chat.id, { ...state, stage: 'calculation_request' });
      await ctx.reply('NHÂN ghi phép tính cần làm bằng tiếng Việt nhé, ví dụ: tính thành tiền bằng số lượng nhân đơn giá, rồi tính tổng cột thành tiền.');
    }
    return true;
  }
  // Cho phép NHÂN trả lời thẳng phép tính mà không cần gõ riêng chữ “có”.
  await runExcelCalculation(ctx, state, t);
  return true;
}

const OFFICE_HELP = {
  wordvanbanhinhanh: 'Em đã sẵn sàng. NHÂN gửi file Word (.docx) có hình ảnh, em chỉnh ảnh đồng đều, căn giữa, không làm méo hình và không đổi nội dung chữ.',
  excel: 'Em đã sẵn sàng. NHÂN gửi file Excel (.xlsx), em kẻ bảng, căn chỉnh và đặt khổ A4 vừa trang in. Thêm chữ "ngang" hoặc "đứng" sau lệnh nếu muốn chọn hướng giấy.',
  exceltinhtoan: 'Em đã sẵn sàng. NHÂN gửi file Excel (.xlsx) kèm yêu cầu tính toán ở phần chú thích (hoặc gõ yêu cầu ngay sau lệnh), ví dụ: tính thành tiền = SL x Đơn giá, tổng cột Thành tiền, giấy ngang. Em chỉ làm đúng yêu cầu, không tự thêm cột hay dòng tổng.',
};

async function handleOfficeFile(ctx, doc, cmd, args) {
  const name = doc.file_name || 'file';
  const lower = name.toLowerCase();
  const isWord = cmd === 'wordvanbanhinhanh';
  if (isWord && !lower.endsWith('.docx')) return ctx.reply('Lệnh này cần file Word .docx. File .doc NHÂN lưu lại thành .docx rồi gửi giúp em nhé.');
  if (!isWord && !lower.endsWith('.xlsx')) return ctx.reply('Lệnh này cần file Excel .xlsx. File .xls NHÂN lưu lại thành .xlsx rồi gửi giúp em nhé.');
  if (doc.file_size > 20 * 1024 * 1024) return ctx.reply('File vượt 20MB, Telegram không cho bot tải.');

  try {
    const res = await withTyping(ctx, 'upload_document', async () => {
      const src = await downloadFile(ctx, doc.file_id);
      if (isWord) {
        const r = await office.fixWordImages(src);
        return {
          buffer: r.buffer,
          note: r.count
            ? `Em đã chỉnh ${r.count} ảnh: cùng khung 10 x 7,5 cm, căn giữa, không méo hình. Nội dung chữ giữ nguyên.`
            : 'File không có hình ảnh nội dòng để chỉnh, em gửi lại nguyên bản.',
        };
      }
      if (cmd === 'exceltinhtoan' && args.replace(/\b(ngang|đứng|dọc)\b/gi, '').trim().length <= 3) {
        return { ask: true };
      }
      const orient = /ngang|landscape/i.test(args) ? 'ngang' : /đứng|dọc|portrait/i.test(args) ? 'doc' : null;
      let spec = null;
      if (cmd === 'exceltinhtoan') {
        const preview = await office.excelPreview(src);
        spec = await office.excelSpecFromAI(callGemini, preview, args);
      }
      const r = await office.processExcel(src, { orient, tinhToan: cmd === 'exceltinhtoan', spec });
      return {
        buffer: r.buffer,
        note: 'Em đã xử lý xong. ' + r.info + (cmd === 'exceltinhtoan' ? '. Công thức Excel còn nguyên, NHÂN kiểm tra lại kết quả trước khi dùng.' : '. Chỉ căn chỉnh bảng, không thêm cột hay dòng tổng.'),
      };
    });
    if (res.ask) return ctx.reply('NHÂN chưa nêu yêu cầu tính toán. NHÂN gửi lại file kèm yêu cầu ở chú thích, ví dụ: tính thành tiền = SL x Đơn giá, tổng cột Thành tiền. Nếu chỉ cần căn chỉnh bảng thì dùng lệnh /excel.');
    const { buffer, note } = res;
    const ext = isWord ? '.docx' : '.xlsx';
    await sendDocWithRetry(ctx, buffer, name.replace(/\.[^.]+$/, '') + '_da_xu_ly' + ext);
    await ctx.reply(note);
  } catch (e) {
    console.error('Lỗi xử lý file văn phòng:', e);
    await ctx.reply(isTransient(errText(e)) ? friendlyError(e) : 'Em chưa xử lý được file này: ' + errText(e).slice(0, 150));
  }
}

bot.command(['wordvanbanhinhanh', 'wordhvanbanhinhanh', 'excel', 'exceltinhtoan'], async (ctx) => {
  const p = office.parseOfficeCommand(ctx.message.text);
  if (p?.cmd === 'excel') {
    const state = readExcelFlow(ctx.chat.id);
    if (state?.stage === 'orientation') return finishExcelOrientation(ctx, state, p.args);
    if (state?.stage === 'ask_calculation' || state?.stage === 'calculation_request') {
      return ctx.reply('NHÂN trả lời bot có muốn tính không và nêu phép tính trước nhé. Sau đó mình chọn hướng giấy bằng /excel đứng hoặc /excel ngang.');
    }
  }
  pendingOffice.set(ctx.chat.id, { cmd: p.cmd, args: p.args, at: Date.now() });
  return ctx.reply(OFFICE_HELP[p.cmd]);
});

// ===== LỆNH =====
// Câu nói tự nhiên kích hoạt chế độ thể thức, ví dụ: "trình bày theo thể thức văn bản", "xuất file word"
const FORMAL_RE = /(trình bày|soạn|xuất|tạo|làm|lập|chuyển)[^\n]{0,40}(thể thức|file word|file docx)/i;
// Xuất Word chỉ có nội dung, KHÔNG tiêu đề: ví dụ "xuất file word không tiêu đề"
const PLAIN_RE = /(xuất|tạo|làm|gửi|chuyển)[^\n]{0,40}(word|docx)[^\n]{0,40}(không\s+(có\s+)?(tiêu đề|thể thức|quốc hiệu)|chỉ (có )?nội dung|thông thường)/i;

bot.start((ctx) =>
  ctx.reply(
    'Dạ, em là tt, trợ lý của NHÂN.\n\n' +
      '- Nhắn trực tiếp để hỏi, nhờ tóm tắt, soạn thảo, lập kế hoạch.\n' +
      '- Gửi file PDF, DOCX, XLSX, TXT, CSV hoặc ảnh chụp văn bản, ghi yêu cầu ở phần chú thích.\n' +
      '- Gửi tin nhắn thoại, em nghe và xử lý.\n\n' +
      'Muốn có file Word đúng thể thức văn bản: nhắn "trình bày theo thể thức văn bản" (kèm nội dung hoặc yêu cầu) hoặc gõ /vanban.\n' +
      'Muốn file Word chỉ có nội dung, không tiêu đề: nhắn "xuất file word không tiêu đề" hoặc gõ /word.\n' +
      'Word có hình ảnh: gõ /wordvanbanhinhanh rồi gửi file (hoặc ghi lệnh ở chú thích khi gửi file).\n' +
      'Excel: gửi file .xlsx trước; em sẽ hỏi có cần tính toán không. Sau đó chọn /excel đứng hoặc /excel ngang để căn chỉnh khổ giấy.\n' +
      '/reset : xóa ngữ cảnh, bắt đầu việc mới'
  )
);

bot.command('reset', (ctx) => {
  history.delete(ctx.chat.id);
  lastReply.delete(ctx.chat.id);
  lastDoc.delete(ctx.chat.id);
  const excelFlow = pendingExcelFlow.get(ctx.chat.id);
  if (excelFlow?.timer) clearTimeout(excelFlow.timer);
  pendingExcelFlow.delete(ctx.chat.id);
  return ctx.reply('Đã xóa ngữ cảnh hội thoại.');
});

// /vanban [nội dung hoặc yêu cầu] (/docx và /vb là lệnh tắt cùng chức năng)
bot.command(['vanban', 'docx', 'vb'], (ctx) => {
  const request = ctx.message.text.replace(/^\/\w+(@\w+)?\s*/, '');
  return makeFormalDoc(ctx, request);
});

// /word [nội dung] : xuất file Word chỉ có nội dung, không Quốc hiệu - Tiêu ngữ
bot.command('word', (ctx) => {
  const request = ctx.message.text.replace(/^\/\w+(@\w+)?\s*/, '');
  return makePlainDoc(ctx, request);
});

// ===== TIN NHẮN VĂN BẢN =====
bot.on('text', async (ctx) => {
  const t = ctx.message.text;
  if (await handleExcelFlowText(ctx, t)) return;
  if (PLAIN_RE.test(t)) return makePlainDoc(ctx, '');
  if (FORMAL_RE.test(t)) return makeFormalDoc(ctx, t);
  try {
    const out = await withTyping(ctx, 'typing', () => ask(ctx.chat.id, [{ text: t }], t));
    await sendLong(ctx, out);
  } catch (e) {
    console.error('Lỗi Gemini:', e);
    await ctx.reply(friendlyError(e));
  }
});

// ===== NHẬN FILE =====
bot.on('document', async (ctx) => {
  const doc = ctx.message.document;
  // Lệnh nằm ở chú thích file, hoặc lệnh gõ trước đó (còn hiệu lực 10 phút)
  const capCmd = office.parseOfficeCommand(ctx.message.caption || '');
  const pend = pendingOffice.get(ctx.chat.id);
  if (capCmd || (pend && Date.now() - pend.at < 10 * 60 * 1000)) {
    pendingOffice.delete(ctx.chat.id);
    const c = capCmd || pend;
    return handleOfficeFile(ctx, doc, c.cmd, c.args);
  }
  if ((doc.file_name || '').toLowerCase().endsWith('.xlsx')) {
    if (doc.file_size > 20 * 1024 * 1024) return ctx.reply('File vượt 20MB, Telegram không cho bot tải.');
    try {
      const buffer = await withTyping(ctx, 'upload_document', () => downloadFile(ctx, doc.file_id));
      saveExcelFlow(ctx.chat.id, {
        stage: 'ask_calculation',
        buffer,
        filename: doc.file_name || 'bang_tinh.xlsx',
      });
      return ctx.reply('NHÂN có muốn tính toán trong bảng Excel này không? Trả lời “có” rồi ghi phép tính bằng tiếng Việt, hoặc “không”. Sau đó bot sẽ hỏi hướng giấy để NHÂN chọn /excel đứng hoặc /excel ngang.');
    } catch (e) {
      console.error('Lỗi nhận file Excel:', e);
      return ctx.reply(friendlyError(e));
    }
  }
  const req = ctx.message.caption || 'Tóm tắt các điểm chính của tài liệu này.';
  if (doc.file_size > 20 * 1024 * 1024) {
    return ctx.reply('File vượt 20MB, Telegram không cho bot tải.');
  }
  try {
    const out = await withTyping(ctx, 'typing', async () => {
      const buf = await downloadFile(ctx, doc.file_id);
      const name = (doc.file_name || '').toLowerCase();
      let parts;

      if (doc.mime_type === 'application/pdf' || name.endsWith('.pdf')) {
        parts = [
          { inlineData: { mimeType: 'application/pdf', data: buf.toString('base64') } },
          { text: req },
        ];
      } else if (name.endsWith('.docx')) {
        const { value } = await mammoth.extractRawText({ buffer: buf });
        parts = [{ text: `Nội dung file ${doc.file_name}:\n${clip(value)}\n\nYêu cầu: ${req}` }];
      } else if (name.endsWith('.xlsx')) {
        const txt = await excelToText(buf);
        parts = [{ text: `Nội dung file Excel ${doc.file_name}:\n${clip(txt)}\n\nYêu cầu: ${req}` }];
      } else if (name.endsWith('.txt') || name.endsWith('.csv')) {
        parts = [{ text: `Nội dung file ${doc.file_name}:\n${clip(buf.toString('utf8'))}\n\nYêu cầu: ${req}` }];
      } else if (/^image\//.test(doc.mime_type || '')) {
        parts = [{ inlineData: { mimeType: doc.mime_type, data: buf.toString('base64') } }, { text: req }];
      } else {
        return null;
      }
      return ask(ctx.chat.id, parts, `[Đã gửi file ${doc.file_name}] ${req}`);
    });
    if (out === null) return ctx.reply('Hiện hỗ trợ: PDF, DOCX, XLSX, TXT, CSV và ảnh. File XLS, DOC, PPTX NHÂN lưu lại thành XLSX, DOCX hoặc PDF rồi gửi nhé.');
    await sendLong(ctx, out);
  } catch (e) {
    console.error('Lỗi xử lý file:', e);
    await ctx.reply(friendlyError(e));
  }
});

// ===== NHẬN ẢNH CHỤP (1 ảnh hoặc nhiều ảnh gửi cùng lúc, đọc theo đúng thứ tự gửi) =====
const REQ_ONE =
  'Đọc và chép lại đầy đủ nội dung trong ảnh. Nếu là văn bản hành chính, nêu thêm số hiệu, ngày ban hành, trích yếu.';
const REQ_MANY =
  'Đọc và chép lại đầy đủ nội dung tất cả các trang theo đúng thứ tự từ trang 1 đến hết, không bỏ sót, gộp thành một nội dung liền mạch. Nếu là văn bản hành chính, nêu thêm số hiệu, ngày ban hành, trích yếu.';
const albums = new Map(); // media_group_id -> {items, caption, ctx, timer}

async function processPhotos(ctx, fileIds, req) {
  const n = fileIds.length;
  try {
    const out = await withTyping(ctx, 'typing', async () => {
      const parts = [];
      for (let i = 0; i < n; i++) {
        const buf = await downloadFile(ctx, fileIds[i]);
        if (n > 1) parts.push({ text: `Ảnh/trang số ${i + 1}:` });
        parts.push({ inlineData: { mimeType: 'image/jpeg', data: buf.toString('base64') } });
      }
      const order =
        n > 1
          ? `\n(Có ${n} ảnh, được gửi theo đúng thứ tự trang 1 đến trang ${n}. Xử lý theo đúng thứ tự này.)`
          : '';
      parts.push({ text: req + order });
      return ask(ctx.chat.id, parts, `[Đã gửi ${n} ảnh] ${req}`);
    });
    await sendLong(ctx, out);
  } catch (e) {
    console.error('Lỗi xử lý ảnh:', e);
    await ctx.reply(friendlyError(e));
  }
}

bot.on('photo', async (ctx) => {
  const m = ctx.message;
  const fileId = m.photo[m.photo.length - 1].file_id;
  const gid = m.media_group_id;
  if (!gid) return processPhotos(ctx, [fileId], m.caption || REQ_ONE);

  // Nhiều ảnh gửi cùng lúc: gom lại, chờ 2,5 giây không có ảnh mới thì xử lý một lần
  let a = albums.get(gid);
  if (!a) {
    a = { items: [], caption: '', ctx, timer: null };
    albums.set(gid, a);
  }
  a.items.push({ id: m.message_id, fileId });
  a.ctx = ctx;
  if (m.caption) a.caption = m.caption;
  clearTimeout(a.timer);
  a.timer = setTimeout(() => {
    albums.delete(gid);
    a.items.sort((x, y) => x.id - y.id); // đúng thứ tự gửi
    processPhotos(a.ctx, a.items.map((i) => i.fileId), a.caption || REQ_MANY);
  }, 2500);
});

// ===== NHẬN TIN NHẮN THOẠI =====
bot.on(['voice', 'audio'], async (ctx) => {
  const media = ctx.message.voice || ctx.message.audio;
  try {
    const out = await withTyping(ctx, 'typing', async () => {
      const buf = await downloadFile(ctx, media.file_id);
      const parts = [
        { inlineData: { mimeType: media.mime_type || 'audio/ogg', data: buf.toString('base64') } },
        { text: 'Nghe tin nhắn thoại này và thực hiện đúng yêu cầu trong đó.' },
      ];
      return ask(ctx.chat.id, parts, '[Tin nhắn thoại]');
    });
    await sendLong(ctx, out);
  } catch (e) {
    console.error('Lỗi xử lý thoại:', e);
    await ctx.reply(friendlyError(e));
  }
});

bot.catch((err) => console.error('Lỗi bot:', err));

// ===== SERVER WEBHOOK =====
const handler = bot.webhookCallback(HOOK_PATH);
http
  .createServer((req, res) => {
    if (req.method === 'POST' && req.url === HOOK_PATH) return handler(req, res);
    res.writeHead(200);
    res.end('OK');
  })
  .listen(PORT, async () => {
    console.log('Server chạy cổng', PORT);
    try {
      await bot.telegram.setMyCommands([
        { command: 'vanban', description: 'Trình bày văn bản theo thể thức, xuất file Word' },
        { command: 'word', description: 'Xuất file Word chỉ có nội dung, không tiêu đề' },
        { command: 'wordvanbanhinhanh', description: 'Word có hình ảnh: chỉnh ảnh đồng đều, không méo' },
        { command: 'excel', description: 'Excel: định dạng bảng, khổ A4 đứng hoặc ngang' },
        { command: 'exceltinhtoan', description: 'Excel: định dạng và tính toán theo yêu cầu' },
        { command: 'reset', description: 'Xóa ngữ cảnh, bắt đầu việc mới' },
        { command: 'start', description: 'Xem hướng dẫn' },
      ]);
      if (DOMAIN) {
        await bot.telegram.setWebhook(`${DOMAIN}${HOOK_PATH}`);
        console.log('Đã đăng ký webhook:', DOMAIN);
      } else {
        console.warn('Chưa có RENDER_EXTERNAL_URL, chưa đăng ký webhook.');
      }
    } catch (e) {
      console.error('Lỗi khởi tạo webhook:', e);
    }
  });

// Tự gọi chính mình mỗi 10 phút để Render ít "ngủ" hơn
if (DOMAIN) {
  setInterval(() => fetch(DOMAIN).catch(() => {}), 10 * 60 * 1000);
}

process.once('SIGTERM', () => {
  console.log('Nhận tín hiệu tắt, chờ 10 giây để hoàn tất việc đang làm.');
  setTimeout(() => process.exit(0), 10000);
});
