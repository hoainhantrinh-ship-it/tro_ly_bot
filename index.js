require('dotenv').config();
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');

// Lấy Token và API Key từ biến môi trường
const bot = new Telegraf(process.env.TELEGRAM_BOT_TOKEN);
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Khởi tạo phiên trò chuyện mang tính cách của "tt"
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
        console.error("Lỗi gọi Gemini API:", error);
        return "Dạ, trợ lý đang gặp chút sự cố kết nối, Nhân vui lòng thử lại sau nhé!";
    }
}

// Lắng nghe tin nhắn từ Telegram
bot.on('text', async (ctx) => {
    const userMessage = ctx.message.text;
    
    // Gửi trạng thái "đang soạn tin..."
    await ctx.sendChatAction('typing');
    
    // Lấy câu trả lời từ tt (Gemini)
    const replyText = await runChat(userMessage);
    
    // Gửi lại cho Nhân
    await ctx.reply(replyText);
});

// Khởi động bot
bot.launch();
console.log("Trợ lý 'tt' đã khởi động thành công trên Telegram!");

// Bật tắt an toàn
process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
// Tạo web server đơn giản để Render nhận diện cổng mạng
const http = require('http');
const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end('Bot is running!');
});
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Server is listening on port ${PORT}`);
});