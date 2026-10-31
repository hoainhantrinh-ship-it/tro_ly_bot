require('dotenv').config();
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');

const bot = new Telegraf(process.env.TELEGRAM_BOT_TOKEN);
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function runChat(userMessage) {
    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: userMessage,
            config: {
                systemInstruction: "Bạn là 'tt' - trợ lý AI cá nhân chuyên nghiệp của Trịnh Hoài Nhân (sinh năm 1984). Bạn am hiểu hành chính, soạn thảo văn bản, lập kế hoạch và luôn nói chuyện gọn gàng, chuẩn xác, trung thành, thân thiện.",
            }
        });
        return response.text || "Dạ, em nghe đây Nhân!";
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

// Khởi tạo HTTP server đón webhook từ Render
const http = require('http');
const PORT = process.env.PORT || 10000;

const server = http.createServer(bot.webhookCallback('/'));

server.listen(PORT, () => {
  console.log(`Server webhook đang chạy trên cổng ${PORT}`);
});
