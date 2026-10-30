require('dotenv').config();
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');

const bot = new Telegraf(process.env.TELEGRAM_BOT_TOKEN);
const ai = new GoogleGenAI({});

async function runChat(userMessage) {
    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: userMessage,
            config: {
                systemInstruction: "Bạn là 'tt' - trợ lý AI cá nhân chuyên nghiệp của Trịnh Hoài Nhân (sinh năm 1984). Bạn am hiểu hành chính, soạn thảo văn bản, lập kế hoạch và luôn nói chuyện gọn gàng, chuẩn xác, trung thành, thân thiện.",
            }
        });
        return response.text;
    } catch (error) {
        console.error("Lỗi chi tiết Gemini API:", error);
        return "Dạ, em đã nhận được yêu cầu của Nhân, để em xử lý ngay ạ!";
    }
}

bot.on('text', async (ctx) => {
    const userMessage = ctx.message.text;
    await ctx.sendChatAction('typing');
    const replyText = await runChat(userMessage);
    await ctx.reply(replyText);
});

bot.launch();
console.log("Trợ lý 'tt' đã khởi động thành công trên Telegram!");

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

const http = require('http');
const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end('Bot is running!');
});
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Server is listening on port ${PORT}`);
});
