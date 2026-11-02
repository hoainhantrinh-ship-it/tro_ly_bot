require('dotenv').config();
const http = require('http');
const https = require('https');
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
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

// ===== LỆNH =====
// Câu nói tự nhiên kích hoạt chế độ thể thức, ví dụ: "trình bày theo thể thức văn bản", "xuất file word"
const FORMAL_RE = /(trình bày|soạn|xuất|tạo|làm|lập|chuyển)[^\n]{0,40}(thể thức|file word|file docx)/i;

bot.start((ctx) =>
  ctx.reply(
    'Dạ, em là tt, trợ lý của NHÂN.\n\n' +
      '- Nhắn trực tiếp để hỏi, nhờ tóm tắt, soạn thảo, lập kế hoạch.\n' +
      '- Gửi file PDF, DOCX, XLSX, TXT, CSV hoặc ảnh chụp văn bản, ghi yêu cầu ở phần chú thích.\n' +
      '- Gửi tin nhắn thoại, em nghe và xử lý.\n\n' +
      'Muốn có file Word đúng thể thức văn bản: nhắn "trình bày theo thể thức văn bản" (kèm nội dung hoặc yêu cầu) hoặc gõ /vanban.\n' +
      '/reset : xóa ngữ cảnh, bắt đầu việc mới'
  )
);

bot.command('reset', (ctx) => {
  history.delete(ctx.chat.id);
  lastReply.delete(ctx.chat.id);
  lastDoc.delete(ctx.chat.id);
  return ctx.reply('Đã xóa ngữ cảnh hội thoại.');
});

// /vanban [nội dung hoặc yêu cầu] (/docx và /vb là lệnh tắt cùng chức năng)
bot.command(['vanban', 'docx', 'vb'], (ctx) => {
  const request = ctx.message.text.replace(/^\/\w+(@\w+)?\s*/, '');
  return makeFormalDoc(ctx, request);
});

// ===== TIN NHẮN VĂN BẢN =====
bot.on('text', async (ctx) => {
  const t = ctx.message.text;
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
      } else if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        const wb = XLSX.read(buf, { type: 'buffer' });
        const txt = wb.SheetNames.map((n) => `--- Trang tính: ${n} ---\n${XLSX.utils.sheet_to_csv(wb.Sheets[n])}`).join('\n\n');
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
    if (out === null) return ctx.reply('Hiện hỗ trợ: PDF, DOCX, XLSX, TXT, CSV và ảnh. File DOC, PPTX NHÂN lưu lại thành DOCX hoặc PDF rồi gửi nhé.');
    await sendLong(ctx, out);
  } catch (e) {
    console.error('Lỗi xử lý file:', e);
    await ctx.reply(friendlyError(e));
  }
});

// ===== NHẬN ẢNH CHỤP =====
bot.on('photo', async (ctx) => {
  const req =
    ctx.message.caption ||
    'Đọc và chép lại đầy đủ nội dung trong ảnh. Nếu là văn bản hành chính, nêu thêm số hiệu, ngày ban hành, trích yếu.';
  try {
    const photos = ctx.message.photo;
    const out = await withTyping(ctx, 'typing', async () => {
      const buf = await downloadFile(ctx, photos[photos.length - 1].file_id);
      const parts = [{ inlineData: { mimeType: 'image/jpeg', data: buf.toString('base64') } }, { text: req }];
      return ask(ctx.chat.id, parts, `[Đã gửi 1 ảnh] ${req}`);
    });
    await sendLong(ctx, out);
  } catch (e) {
    console.error('Lỗi xử lý ảnh:', e);
    await ctx.reply(friendlyError(e));
  }
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
