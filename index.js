require('dotenv').config();
const http = require('http');
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');
const mammoth = require('mammoth');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

// ===== CẤU HÌNH =====
const TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const GEMINI_KEY = process.env.GEMINI_API_KEY;
// Thử lần lượt từng model, nếu gặp lỗi 404 thì chuyển sang model kế tiếp
const MODELS = ['gemini-3.8-flash', 'gemini-3.6-flash', 'gemini-3.5-flash'];
const ALLOWED_ID = Number(process.env.ALLOWED_USER_ID || 0);
const PORT = process.env.PORT || 10000;
const DOMAIN = process.env.RENDER_EXTERNAL_URL; // Render tự cấp
const HOOK_PATH = `/webhook/${TOKEN}`;

if (!TOKEN || !GEMINI_KEY) {
  console.error('Thiếu TELEGRAM_BOT_TOKEN hoặc GEMINI_API_KEY trong Environment.');
  process.exit(1);
}

const bot = new Telegraf(TOKEN);
const ai = new GoogleGenAI({ apiKey: GEMINI_KEY });

const SYSTEM = `Bạn là "tt" - trợ lý AI cá nhân của NHÂN (Trịnh Hoài Nhân), chuyên về quản trị kinh doanh, hồ sơ, công văn hành chính, báo cáo, lập kế hoạch và lên ý tưởng.
Quy tắc:
- Gọi người dùng là NHÂN, xưng "em". Trả lời gọn, trực diện, không mở đầu sáo rỗng.
- Chỉ dựa trên dữ liệu NHÂN cung cấp, không tự thêm bớt số liệu, không suy diễn thiếu căn cứ. Thiếu thông tin thì hỏi lại.
- Soạn văn bản hành chính theo thể thức Nghị định 30/2020/NĐ-CP, bố cục rõ ràng.
- Huy hiệu Đảng: các mốc 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95 năm tuổi Đảng; các đợt trao tặng 03/02, 30/04, 02/09, 07/11 hằng năm. Khi tính toán phải đối chiếu đúng mốc.
- Không dùng ký hiệu markdown (như ** hoặc #) trong câu trả lời, viết văn bản thuần.`;

const history = new Map(); // chatId -> [{role, parts}]
const lastReply = new Map(); // chatId -> câu trả lời gần nhất

// ===== CHỈ CHO PHÉP NHÂN DÙNG =====
bot.use((ctx, next) => {
  if (ALLOWED_ID && ctx.from?.id !== ALLOWED_ID) {
    return ctx.reply('Bot riêng, không phục vụ người ngoài.');
  }
  return next();
});

// ===== HÀM GỌI AI =====
async function ask(chatId, parts, savedText) {
  const h = history.get(chatId) || [];
  let res;
  let lastErr;
  for (const model of MODELS) {
    try {
      res = await ai.models.generateContent({
        model,
        contents: [...h, { role: 'user', parts }],
        config: { systemInstruction: SYSTEM },
      });
      break;
    } catch (e) {
      lastErr = e;
      const msg = String(e.message || e);
      if (!(msg.includes('404') || msg.includes('NOT_FOUND') || msg.includes('no longer available'))) throw e;
      console.warn(`Model ${model} không dùng được, thử model tiếp theo.`);
    }
  }
  if (!res) throw lastErr;
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
            size: { width: 11906, height: 16838 }, // A4
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
      'Nhắn trực tiếp để hỏi hoặc nhờ soạn thảo.\n' +
      'Gửi file PDF, DOCX, TXT, CSV kèm yêu cầu trong phần chú thích.\n' +
      '/docx : xuất câu trả lời gần nhất thành file Word\n' +
      '/reset : xóa ngữ cảnh hội thoại'
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
    await ctx.reply('Lỗi xuất file Word: ' + String(e.message).slice(0, 200));
  }
});

// ===== TIN NHẮN VĂN BẢN =====
bot.on('text', async (ctx) => {
  await ctx.sendChatAction('typing');
  try {
    const t = ctx.message.text;
    await sendLong(ctx, await ask(ctx.chat.id, [{ text: t }], t));
  } catch (e) {
    console.error('Lỗi Gemini:', e);
    await ctx.reply('Lỗi gọi AI: ' + String(e.message).slice(0, 300));
  }
});

// ===== NHẬN FILE =====
bot.on('document', async (ctx) => {
  const doc = ctx.message.document;
  const req = ctx.message.caption || 'Tóm tắt các điểm chính của tài liệu này.';
  if (doc.file_size > 20 * 1024 * 1024) {
    return ctx.reply('File vượt 20MB, Telegram không cho bot tải.');
  }
  await ctx.sendChatAction('typing');
  try {
    const link = await ctx.telegram.getFileLink(doc.file_id);
    const buf = Buffer.from(await (await fetch(link.href)).arrayBuffer());
    const name = (doc.file_name || '').toLowerCase();
    let parts;

    if (doc.mime_type === 'application/pdf' || name.endsWith('.pdf')) {
      parts = [
        { inlineData: { mimeType: 'application/pdf', data: buf.toString('base64') } },
        { text: req },
      ];
    } else if (name.endsWith('.docx')) {
      const { value } = await mammoth.extractRawText({ buffer: buf });
      parts = [{ text: `Nội dung file ${doc.file_name}:\n${value}\n\nYêu cầu: ${req}` }];
    } else if (name.endsWith('.txt') || name.endsWith('.csv')) {
      parts = [{ text: `Nội dung file ${doc.file_name}:\n${buf.toString('utf8')}\n\nYêu cầu: ${req}` }];
    } else {
      return ctx.reply('Hiện hỗ trợ: PDF, DOCX, TXT, CSV.');
    }

    await sendLong(ctx, await ask(ctx.chat.id, parts, `[Đã gửi file ${doc.file_name}] ${req}`));
  } catch (e) {
    console.error('Lỗi xử lý file:', e);
    await ctx.reply('Lỗi xử lý file: ' + String(e.message).slice(0, 300));
  }
});

bot.catch((err) => console.error('Lỗi bot:', err));

// ===== SERVER WEBHOOK =====
const handler = bot.webhookCallback(HOOK_PATH);
http
  .createServer((req, res) => {
    if (req.method === 'POST' && req.url === HOOK_PATH) return handler(req, res);
    res.writeHead(200);
    res.end('OK'); // để Render / UptimeRobot kiểm tra bot còn sống
  })
  .listen(PORT, async () => {
    console.log('Server chạy cổng', PORT);
    if (DOMAIN) {
      try {
        await bot.telegram.setWebhook(`${DOMAIN}${HOOK_PATH}`);
        console.log('Đã đăng ký webhook:', DOMAIN);
      } catch (e) {
        console.error('Lỗi đăng ký webhook:', e);
      }
    } else {
      console.warn('Chưa có RENDER_EXTERNAL_URL, chưa đăng ký webhook.');
    }
  });

process.once('SIGTERM', () => process.exit(0));
