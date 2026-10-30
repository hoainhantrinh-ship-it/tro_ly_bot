require('dotenv').config();
const { Telegraf } = require('telegraf');
const { GoogleGenAI } = require('@google/genai');

const bot = new Telegraf(process.env.TELEGRAM_BOT_TOKEN);

// Khởi tạo Gemini
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY
});

async function runChat(userMessage) {
  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.0-flash', // hoặc 'gemini-1.5-flash' nếu bạn muốn giữ
      contents: [
        {
          role: 'user',
          parts: [{ text: userMessage }]
        }
      ],
      config: {
        systemInstruction: `Bạn là 'tt' - trợ lý AI cá nhân chuyên nghiệp của Trịnh Hoài Nhân (sinh năm 1984). 
Bạn am hiểu hành chính, soạn thảo văn bản, lập kế hoạch và luôn nói chuyện gọn gàng, chuẩn xác, trung thành, thân thiện.`,
        temperature: 0.7,
      }
    });

    // Lấy text an toàn
    const text = response?.text 
      || response?.candidates?.[0]?.content?.parts?.[0]?.text 
      || "Dạ, em nghe đây Nhân!";

    return text;
  } catch (error) {
    console.error("Lỗi chi tiết Gemini API:", error);
    return "Dạ, em đã nhận được yêu cầu, nhưng đang gặp chút trở ngại kết nối API. Nhân thử lại giúp em nhé!";
  }
}

// Xử lý tin nhắn
bot.on('text', async (ctx) => {
  const userMessage = ctx.message.text;
  await ctx.sendChatAction('typing');
  
  try {
    const replyText = await runChat(userMessage);
    await ctx.reply(replyText);
  } catch (err) {
    console.error("Lỗi khi trả lời:", err);
    await ctx.reply("Dạ, em gặp lỗi rồi ạ. Nhân thử lại sau nhé!");
  }
});

// Bắt lỗi toàn cục của bot
bot.catch((err, ctx) => {
  console.error(`Lỗi bot với ${ctx.updateType}:`, err);
});

bot.launch();
console.log("Trợ lý 'tt' đã khởi động thành công trên Telegram!");

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

// HTTP server keep-alive (cho Render/Railway)
const http = require('http');
const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end('Bot is running!');
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Server is listening on port ${PORT}`);
});
