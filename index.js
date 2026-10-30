require('dotenv').config();
const http = require('http');
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');

const bot = new Telegraf(process.env.TELEGRAM_BOT_TOKEN);
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Đổi model bằng biến GEMINI_MODEL trên Render, không cần sửa code
const MODEL = process.env.GEMINI_MODEL || 'gemini-3.8-flash';

const SYSTEM_PROMPT = "Bạn là 'tt' - trợ lý AI cá nhân của Trịnh Hoài Nhân (sinh năm 1984), gọi là Nhân. Bạn am hiểu hành chính, soạn thảo văn bản, lập kế hoạch và luôn nói chuyện gọn gàng, chuẩn xác, trung thành, thân thiện.";

async function runChat(userMessage) {
    try {
        const response = await ai.models.generateContent({
            model: MODEL,
            contents: userMessage,
            config: { systemInstruction: SYSTEM_PROMPT }
        });
        return response.text || 'Em chưa tạo được câu trả lời, Nhân thử lại giúp em.';
    } catch (error) {
        console.error('Lỗi chi tiết Gemini API:', error);
        return 'Hệ thống đang gặp sự cố, Nhân vui lòng thử lại sau ít phút.';
    }
}

bot.on('text', async (ctx) => {
    await ctx.sendChatAction('typing');
    const replyText = await runChat(ctx.message.text);
    await ctx.reply(replyText.slice(0, 4000));
});

bot.launch();
console.log("Trợ lý 'tt' đã khởi động thành công trên Telegram!");

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

const PORT = process.env.PORT || 3000;
http.createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end('Bot is running!');
}).listen(PORT, () => console.log(`Server is listening on port ${PORT}`));
