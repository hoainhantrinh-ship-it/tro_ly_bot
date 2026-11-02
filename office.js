// =====================================================================
// HƯỚNG DẪN GHÉP VÀO index.js (5 chỗ). Cài thêm: npm install jszip exceljs
// =====================================================================

// ---------- (1) Ngay dưới dòng require('docx')... thêm: ----------
const office = require('./office');

// ---------- (2) Dán khối này ngay TRÊN dòng "// ===== LỆNH =====" ----------
// ===== WORD CÓ HÌNH ẢNH / EXCEL ĐỊNH DẠNG - TÍNH TOÁN =====
const pendingOffice = new Map(); // chatId -> {cmd, args, at}: NHÂN gõ lệnh trước, gửi file sau

const OFFICE_HELP = {
  wordvanbanhinhanh: 'Em đã sẵn sàng. NHÂN gửi file Word (.docx) có hình ảnh, em chỉnh ảnh đồng đều, căn giữa, không làm méo hình và không đổi nội dung chữ.',
  excel: 'Em đã sẵn sàng. NHÂN gửi file Excel (.xlsx), em kẻ bảng, căn chỉnh và đặt khổ A4 vừa trang in. Thêm chữ "ngang" hoặc "đứng" sau lệnh nếu muốn chọn hướng giấy.',
  exceltinhtoan: 'Em đã sẵn sàng. NHÂN gửi file Excel (.xlsx) kèm yêu cầu tính toán ở phần chú thích, ví dụ: tính thành tiền = SL x Đơn giá, tổng cột Thành tiền, giấy ngang. Không nêu yêu cầu thì em cộng tổng các cột số.',
};

async function handleOfficeFile(ctx, doc, cmd, args) {
  const name = doc.file_name || 'file';
  const lower = name.toLowerCase();
  const isWord = cmd === 'wordvanbanhinhanh';
  if (isWord && !lower.endsWith('.docx')) return ctx.reply('Lệnh này cần file Word .docx. File .doc NHÂN lưu lại thành .docx rồi gửi giúp em nhé.');
  if (!isWord && !lower.endsWith('.xlsx')) return ctx.reply('Lệnh này cần file Excel .xlsx. File .xls NHÂN lưu lại thành .xlsx rồi gửi giúp em nhé.');
  if (doc.file_size > 20 * 1024 * 1024) return ctx.reply('File vượt 20MB, Telegram không cho bot tải.');

  try {
    const { buffer, note } = await withTyping(ctx, 'upload_document', async () => {
      const src = await downloadFile(ctx, doc.file_id);
      if (isWord) {
        const r = await office.fixWordImages(src);
        return { buffer: r.buffer, note: r.count ? `Em đã chỉnh ${r.count} ảnh: cùng khung 10 x 7,5 cm, căn giữa, không méo hình. Nội dung chữ giữ nguyên.` : 'File không có hình ảnh nội dòng để chỉnh, em gửi lại nguyên bản.' };
      }
      const orient = /ngang|landscape/i.test(args) ? 'ngang' : /đứng|dọc|portrait/i.test(args) ? 'doc' : null;
      let spec = null;
      if (cmd === 'exceltinhtoan' && args.replace(/\b(ngang|đứng|dọc)\b/gi, '').trim().length > 3) {
        const preview = await office.excelPreview(src);
        spec = await office.excelSpecFromAI(callGemini, preview, args);
      }
      const r = await office.processExcel(src, { orient, tinhToan: cmd === 'exceltinhtoan', spec });
      return { buffer: r.buffer, note: 'Em đã xử lý xong. ' + r.info + (cmd === 'exceltinhtoan' ? '. Công thức Excel còn nguyên, NHÂN kiểm tra lại kết quả trước khi dùng.' : '') };
    });
    const ext = isWord ? '.docx' : '.xlsx';
    await sendDocWithRetry(ctx, buffer, name.replace(/\.[^.]+$/, '') + '_da_xu_ly' + ext);
    await ctx.reply(note);
  } catch (e) {
    console.error('Lỗi xử lý file văn phòng:', e);
    await ctx.reply(isTransient(errText(e)) || /Gemini|AI/i.test(errText(e)) ? friendlyError(e) : 'Em chưa xử lý được file này: ' + errText(e).slice(0, 150));
  }
}

bot.command(['wordvanbanhinhanh', 'wordhvanbanhinhanh', 'excel', 'exceltinhtoan'], (ctx) => {
  const p = office.parseOfficeCommand(ctx.message.text);
  pendingOffice.set(ctx.chat.id, { cmd: p.cmd, args: p.args, at: Date.now() });
  return ctx.reply(OFFICE_HELP[p.cmd]);
});

// ---------- (3) Trong bot.on('document', ...) thêm ngay SAU dòng "const doc = ctx.message.document;" ----------
  // Lệnh nằm ở chú thích file, hoặc lệnh gõ trước đó (còn hiệu lực 10 phút)
  const capCmd = office.parseOfficeCommand(ctx.message.caption || '');
  const pend = pendingOffice.get(ctx.chat.id);
  if (capCmd || (pend && Date.now() - pend.at < 10 * 60 * 1000)) {
    pendingOffice.delete(ctx.chat.id);
    const c = capCmd || pend;
    return handleOfficeFile(ctx, doc, c.cmd, c.args);
  }

// ---------- (4) Trong bot.start(...) thêm 2 dòng vào chuỗi hướng dẫn: ----------
      'Word có hình ảnh: gõ /wordvanbanhinhanh rồi gửi file (hoặc ghi lệnh ở chú thích khi gửi file).\n' +
      'Excel: /excel (định dạng thường) hoặc /exceltinhtoan (kèm yêu cầu tính toán), có thể thêm "ngang" hoặc "đứng".\n' +

// ---------- (5) Trong bot.telegram.setMyCommands([...]) thêm: ----------
        { command: 'wordvanbanhinhanh', description: 'Word có hình ảnh: chỉnh ảnh đồng đều, không méo' },
        { command: 'excel', description: 'Excel: định dạng bảng, khổ A4 đứng hoặc ngang' },
        { command: 'exceltinhtoan', description: 'Excel: định dạng và tính toán theo yêu cầu' },
