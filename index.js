require('dotenv').config();
const http = require('http');
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

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

const bot = new Telegraf(TOKEN, { handlerTimeout: 180000 });
const ai = new GoogleGenAI({ apiKey: GEMINI_KEY });

function buildSystem() {
  const now = new Date().toLocaleString('vi-VN', {
    timeZone: 'Asia/Ho_Chi_Minh',
    dateStyle: 'full',
    timeStyle: 'short',
  });
  return `Bạn là "tt" - trợ lý ảo cá nhân của NHÂN (Trịnh Hoài Nhân), chuyên về quản trị kinh doanh, hồ sơ, công văn hành chính, báo cáo, lập kế hoạch và lên ý tưởng. Bây giờ là ${now} (giờ Việt Nam).
Quy tắc:
- Gọi người dùng là NHÂN, xưng "em". Trả lời gọn, trực diện, không mở đầu sáo rỗng.
- Câu hỏi kiến thức chung, thời sự, thời tiết, giá cả, quy định mới, tư vấn: trả lời đầy đủ, dùng Google Search khi cần thông tin mới hoặc theo thời gian thực. Nếu NHÂN hỏi thời tiết mà không nêu địa điểm thì mặc định là TP. Hồ Chí Minh.
- Với hồ sơ, số liệu, tài liệu NHÂN cung cấp: bám sát 100%, không tự thêm bớt số liệu, không suy diễn thiếu căn cứ. Thiếu thông tin thì hỏi lại.
- Soạn văn bản hành chính theo thể thức Nghị định 30/2020/NĐ-CP, bố cục rõ ràng.
- Huy hiệu Đảng: các mốc 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95 năm tuổi Đảng; các đợt trao tặng 03/02, 30/04, 02/09, 07/11 hằng năm. Khi tính toán phải đối chiếu đúng mốc.
- Không dùng ký hiệu markdown (như ** hoặc #). Viết văn bản thuần, liệt kê bằng dấu gạch đầu dòng "-".`;
}

const history = new Map(); // chatId -> [{role, parts}]
const lastReply = new Map(); // chatId -> câu trả lời gần nhất

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
async function callGemini(contents) {
  let useSearch = true;
  let lastErr;
  for (const model of MODELS) {
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        return await ai.models.generateContent({
          model,
          contents,
          config: {
            systemInstruction: buildSystem(),
            ...(useSearch ? { tools: [{ googleSearch: {} }] } : {}),
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
  lastReply.set(chatId, text);
  return text;
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

const clip = (s) => (s.length > MAX_TEXT ? s.slice(0, MAX_TEXT) + '\n[...đã cắt bớt do quá dài]' : s);

// ===== XUẤT WORD =====
async function toDocx(text) {
  const paragraphs = text.split('\n').map(
    (line) =>
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 120, line: 360 },
        children: [new TextRun({ text: line, font: 'Times New Roman', size: 28 })],
      })
  );
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, bottom: 1134, left: 1701, right: 850 },
          },
        },
        children: paragraphs,
      },
    ],
  });
  return Packer.toBuffer(doc);
}

// ===== LỆNH =====
bot.start((ctx) =>
  ctx.reply(
    'Dạ, em là tt, trợ lý của NHÂN.\n\n' +
      '- Nhắn trực tiếp để hỏi, nhờ soạn thảo, lập kế hoạch, tra cứu thông tin mới.\n' +
      '- Gửi file PDF, DOCX, XLSX, TXT, CSV hoặc ảnh chụp văn bản, ghi yêu cầu ở phần chú thích.\n' +
      '- Gửi tin nhắn thoại, em nghe và xử lý.\n' +
      '/docx : xuất câu trả lời gần nhất thành file Word\n' +
      '/reset : xóa ngữ cảnh, bắt đầu việc mới'
  )
);

bot.command('reset', (ctx) => {
  history.delete(ctx.chat.id);
  lastReply.delete(ctx.chat.id);
  return ctx.reply('Đã xóa ngữ cảnh hội thoại.');
});

bot.command('docx', async (ctx) => {
  const text = lastReply.get(ctx.chat.id);
  if (!text) return ctx.reply('Chưa có nội dung để xuất. NHÂN nhờ em soạn trước nhé.');
  try {
    await ctx.sendChatAction('upload_document');
    const buffer = await toDocx(text);
    await ctx.replyWithDocument({ source: buffer, filename: 'van_ban_tt.docx' });
  } catch (e) {
    console.error('Lỗi xuất docx:', e);
    await ctx.reply('Lỗi xuất file Word: ' + errText(e).slice(0, 200));
  }
});

// ===== TIN NHẮN VĂN BẢN =====
bot.on('text', async (ctx) => {
  try {
    const t = ctx.message.text;
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
        { command: 'docx', description: 'Xuất câu trả lời gần nhất ra file Word' },
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

process.once('SIGTERM', () => process.exit(0));
