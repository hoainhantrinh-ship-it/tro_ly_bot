require('dotenv').config();
const http = require('http');
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');
const mammoth = require('mammoth');

const TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const MODEL = process.env.GEMINI_MODEL || 'gemini-2.5-flash';
const ALLOWED_ID = Number(process.env.ALLOWED_USER_ID || 0);
const PORT = process.env.PORT || 10000;
const HOOK_PATH = `/webhook/${TOKEN}`;
const DOMAIN = process.env.RENDER_EXTERNAL_URL; // Render tự cấp

const bot = new Telegraf(TOKEN);
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

const SYSTEM = "Bạn là 'tt' - trợ lý AI cá nhân của Trịnh Hoài Nhân. Am hiểu hành chính, soạn thảo văn bản theo thể thức Nghị định 30/2020/NĐ-CP, lập kế hoạch. Trả lời gọn, chính xác, chỉ dựa trên dữ liệu được cung cấp, không tự thêm số liệu.";

const history = new Map(); // chatId -> [{role, parts:[{text}]}]

// Chỉ cho phép Nhân dùng
bot.use((ctx, next) => {
  if (ALLOWED_ID && ctx.from?.id !== ALLOWED_ID) return ctx.reply('Bot riêng, không phục vụ.');
  return next();
});

async function ask(chatId, parts, savedText) {
  const h = history.get(chatId) || [];
  const res = await ai.models.generateContent({
    model: MODEL,
    contents: [...h, { role: 'user', parts }],
    config: { systemInstruction: SYSTEM },
  });
  const text = res.text || 'Em chưa có phản hồi, Nhân thử lại giúp em.';
  h.push({ role: 'user', parts: [{ text: savedText }] });
  h.push({ role: 'model', parts: [{ text }] });
  history.set(chatId, h.slice(-20));
  return text;
}

async function sendLong(ctx, text) {
  for (let i = 0; i < text.length; i += 4000) await ctx.reply(text.slice(i, i + 4000));
}

bot.command('reset', (ctx) => { history.delete(ctx.chat.id); ctx.reply('Đã xóa ngữ cảnh hội thoại.'); });

bot.on('text', async (ctx) => {
  await ctx.sendChatAction('typing');
  try {
    const t = ctx.message.text;
    await sendLong(ctx, await ask(ctx.chat.id, [{ text: t }], t));
  } catch (e) {
    console.error('Lỗi Gemini:', e);
    await ctx.reply('Lỗi gọi AI: ' + String(e.message).slice(0, 200));
  }
});

bot.on('document', async (ctx) => {
  const doc = ctx.message.document;
  const req = ctx.message.caption || 'Tóm tắt các điểm chính của tài liệu này.';
  if (doc.file_size > 20 * 1024 * 1024) return ctx.reply('File vượt 20MB, Telegram không cho bot tải.');
  await ctx.sendChatAction('typing');
  try {
    const link = await ctx.telegram.getFileLink(doc.file_id);
    const buf = Buffer.from(await (await fetch(link.href)).arrayBuffer());
    const name = (doc.file_name || '').toLowerCase();
    let parts;
    if (doc.mime_type === 'application/pdf') {
      parts = [{ inlineData: { mimeType: 'application/pdf', data: buf.toString('base64') } }, { text: req }];
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
    await ctx.reply('Lỗi xử lý file: ' + String(e.message).slice(0, 200));
  }
});

bot.catch((err) => console.error('Lỗi bot:', err));

const handler = bot.webhookCallback(HOOK_PATH);
http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === HOOK_PATH) return handler(req, res);
  res.writeHead(200); res.end('OK'); // để Render/UptimeRobot kiểm tra sống
}).listen(PORT, async () => {
  console.log('Server chạy cổng', PORT);
  if (DOMAIN) await bot.telegram.setWebhook(`${DOMAIN}${HOOK_PATH}`);
});
