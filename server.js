const express = require('express');
const cors = require('cors');
const multer = require('multer');
const fetch = require('node-fetch');
const FormData = require('form-data');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = require('docx');
const ExcelJS = require('exceljs');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const TG_TOKEN = process.env.TG_TOKEN || '6575215253:AAEs92rEfReGD8E7bHcXKQEFQO3Bb0avfn8';
const TG_CHAT = process.env.TG_CHAT || '5816903954';
const GEMINI_KEY = process.env.GEMINI_KEY || 'AIzaSyAZjzUwrbspHInJYiUiIluUWLpsWOjgYh8';

app.use(cors());
app.use(express.json({ limit: '50mb' }));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const tmpDir = '/tmp/kompyordam';
if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });

// HELPERS
async function sendTG(text) {
  const res = await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendMessage`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: TG_CHAT, text, parse_mode: 'HTML' })
  });
  return res.json();
}

async function sendTGFile(filePath, fileName, caption) {
  const form = new FormData();
  form.append('chat_id', TG_CHAT);
  form.append('caption', caption || '');
  form.append('document', fs.createReadStream(filePath), { filename: fileName });
  const res = await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendDocument`, { method: 'POST', body: form });
  return res.json();
}

async function gemini(prompt) {
  const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_KEY}`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.8, maxOutputTokens: 4096 } })
  });
  const d = await res.json();
  return d.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

function parseJSON(text) {
  try { return JSON.parse(text.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim()); }
  catch(e) { return null; }
}

// ROUTES
app.get('/', (req, res) => res.json({ status: 'KompYordam Server OK ✅' }));

// Telegram notify
app.post('/api/notify', async (req, res) => {
  try { await sendTG(req.body.message); res.json({ success: true }); }
  catch(e) { res.status(500).json({ success: false, error: e.message }); }
});

// AI Humanize
app.post('/api/humanize', async (req, res) => {
  try {
    const { text, style } = req.body;
    const styles = {
      talaba: "o'zbek talabasi yozgandek — oddiy, jonli, ba'zida kichik grammatik xato bo'lishi mumkin",
      rasmiy: "rasmiy hujjat uslubida, lekin tabiiy inson yozganday",
      oddiy: "oddiy, kundalik suhbat uslubida, sodda va iliq",
      ilmiy: "ilmiy uslubda, lekin tabiiy va inson yozganday"
    };
    const prompt = `Quyidagi AI matnni ${styles[style]||styles.talaba} qilib qayta yoz. AI izlari, robot uslubi, takroriy iboralar (birinchidan, ikkinchidan, xulosa qilib aytganda) larni olib tashla. Tabiiy, inson yozganday qil. Faqat qayta yozilgan matnni ber:\n\n${text}`;
    const result = await gemini(prompt);
    res.json({ success: true, content: result });
  } catch(e) { res.status(500).json({ success: false, error: e.message }); }
});

// Create Word
app.post('/api/create/word', upload.single('image'), async (req, res) => {
  try {
    const { topic, pages, extra, phone } = req.body;
    const prompt = `"${topic}" mavzusida referat yoz. Hajm: ${pages||'10-15 bet'}. ${extra||''}\nKirish, 3 ta asosiy bo'lim, xulosa, adabiyotlar. O'zbek tilida, inson uslubida. # sarlavhalar bilan ajrat.`;
    const content = await gemini(prompt);
    const lines = content.split('\n');
    const children = [
      new Paragraph({ text: topic, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, spacing: { after: 400 } })
    ];
    for (const line of lines) {
      const t = line.trim();
      if (!t) continue;
      if (t.startsWith('# ')) children.push(new Paragraph({ text: t.replace('# ',''), heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 } }));
      else if (t.startsWith('## ')) children.push(new Paragraph({ text: t.replace('## ',''), heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 } }));
      else children.push(new Paragraph({ children: [new TextRun({ text: t, size: 24, font: 'Times New Roman' })], spacing: { after: 120 }, indent: { firstLine: 720 } }));
    }
    const doc = new Document({ sections: [{ properties: { page: { margin: { top: 1440, right: 1080, bottom: 1440, left: 1800 } } }, children }] });
    const fileName = `referat_${Date.now()}.docx`;
    const filePath = path.join(tmpDir, fileName);
    fs.writeFileSync(filePath, await Packer.toBuffer(doc));
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    await sendTG(`📝 <b>WORD BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Mavzu: ${topic}\n📖 Hajm: ${pages}\n📞 Telefon: ${phone}\n⏰ ${now}`);
    await sendTGFile(filePath, fileName, `📝 Tayyor Word hujjat — ${topic}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) { console.error(e); res.status(500).json({ success: false, error: e.message }); }
});

// Create Excel
app.post('/api/create/excel', upload.single('image'), async (req, res) => {
  try {
    const { description, phone } = req.body;
    const prompt = `Excel jadval uchun JSON: {"title":"...","sheets":[{"name":"...","headers":["..."],"rows":[["..."]],"formulas":[{"cell":"D2","formula":"=SUM(A2:C2)"}]}]}\nFaqat JSON. So'rov: ${description}. Kamida 10 qator real ma'lumot. O'zbek tilida.`;
    const aiResp = await gemini(prompt);
    let data = parseJSON(aiResp) || { title: description, sheets: [{ name: "Ma'lumotlar", headers: ["№","Nomi","Miqdori","Narxi","Jami"], rows: Array.from({length:10},(_,i)=>[i+1,`Element ${i+1}`,Math.floor(Math.random()*100),Math.floor(Math.random()*50000),'']), formulas: Array.from({length:10},(_,i)=>({cell:`E${i+2}`,formula:`=C${i+2}*D${i+2}`})) }] };
    
    const wb = new ExcelJS.Workbook();
    for (const sheet of data.sheets || []) {
      const ws = wb.addWorksheet(sheet.name || 'Sheet1');
      if (data.title) {
        ws.mergeCells(1,1,1,(sheet.headers||[]).length||5);
        const tc = ws.getCell('A1');
        tc.value = data.title;
        tc.font = { bold:true, size:14, color:{argb:'FFFFFFFF'} };
        tc.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF1a1e35'} };
        tc.alignment = { horizontal:'center', vertical:'middle' };
        ws.getRow(1).height = 32;
      }
      if (req.file) {
        try {
          const imgId = wb.addImage({ buffer: req.file.buffer, extension: (req.file.mimetype.split('/')[1]||'jpeg').replace('jpg','jpeg') });
          ws.addImage(imgId, { tl:{col:0,row:2}, ext:{width:200,height:150} });
        } catch(ie) { console.log('img err',ie.message); }
      }
      const sr = data.title ? 2 : 1;
      if (sheet.headers) {
        const hr = ws.getRow(sr);
        sheet.headers.forEach((h,i) => {
          const c = hr.getCell(i+1);
          c.value = h; c.font = {bold:true,color:{argb:'FFFFFFFF'},size:11};
          c.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FF2563eb'}};
          c.alignment = {horizontal:'center',vertical:'middle'};
          c.border = {top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};
          ws.getColumn(i+1).width = Math.max(h.length+5,15);
        });
        hr.height = 25;
      }
      if (sheet.rows) {
        sheet.rows.forEach((row,ri) => {
          const dr = ws.getRow(sr+1+ri);
          row.forEach((val,ci) => {
            const c = dr.getCell(ci+1);
            if (typeof val === 'string' && val.startsWith('=')) c.value = {formula:val.substring(1)};
            else c.value = val;
            c.fill = {type:'pattern',pattern:'solid',fgColor:{argb:ri%2===0?'FFF8F9FF':'FFFFFFFF'}};
            c.border = {top:{style:'thin',color:{argb:'FFe0e0e0'}},bottom:{style:'thin',color:{argb:'FFe0e0e0'}},left:{style:'thin',color:{argb:'FFe0e0e0'}},right:{style:'thin',color:{argb:'FFe0e0e0'}}};
          });
        });
      }
      if (sheet.formulas) {
        sheet.formulas.forEach(f => {
          const c = ws.getCell(f.cell);
          c.value = {formula:f.formula.replace('=','')};
          c.font = {bold:true,color:{argb:'FF16a34a'}};
        });
      }
    }
    const fileName = `excel_${Date.now()}.xlsx`;
    const filePath = path.join(tmpDir, fileName);
    await wb.xlsx.writeFile(filePath);
    const now = new Date().toLocaleString('uz-UZ', {timeZone:'Asia/Tashkent'});
    await sendTG(`📊 <b>EXCEL BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Tavsif: ${description?.substring(0,100)}\n📞 Telefon: ${phone}\n⏰ ${now}`);
    await sendTGFile(filePath, fileName, `📊 Tayyor Excel — ${data.title}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) { console.error(e); res.status(500).json({ success: false, error: e.message }); }
});

// Create PowerPoint
app.post('/api/create/pptx', async (req, res) => {
  try {
    const { topic, slides: slideCount, style, phone } = req.body;
    const prompt = `"${topic}" uchun PowerPoint. JSON: {"slides":[{"title":"...","points":["...","...","..."]}]}. ${slideCount||'12 ta'} slayd. O'zbek tilida. Faqat JSON.`;
    const aiResp = await gemini(prompt);
    let slidesData = parseJSON(aiResp) || { slides: [{ title: topic, points: ['Kirish', 'Asosiy qism', 'Xulosa'] }] };
    
    const colors = style === 'Klassik / Rasmiy' ? {bg:'FFFFFF',text:'1a1a2e',accent:'2563eb'} : style === 'Yorqin / Ijodiy' ? {bg:'1a0533',text:'FFFFFF',accent:'ff6b9d'} : {bg:'0d0f1a',text:'f0f2ff',accent:'f0c060'};
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    for (const [idx, slide] of (slidesData.slides||[]).entries()) {
      const s = pptx.addSlide();
      s.background = { color: colors.bg };
      s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:0.1, h:'100%', fill:{color:colors.accent} });
      s.addText(`${idx+1}/${(slidesData.slides||[]).length}`, { x:8.5, y:0.1, w:1, h:0.4, fontSize:11, color:colors.accent, align:'right', bold:true });
      s.addText(slide.title||'', { x:0.4, y:0.3, w:9, h:1.2, fontSize:idx===0?36:26, bold:true, color:colors.text, fontFace:'Calibri' });
      s.addShape(pptx.ShapeType.rect, { x:0.4, y:1.5, w:8.8, h:0.03, fill:{color:colors.accent}, line:{color:colors.accent} });
      if (slide.points?.length) {
        s.addText(slide.points.map(p=>`• ${p}`).join('\n'), { x:0.4, y:1.7, w:9, h:4.5, fontSize:15, color:colors.text, fontFace:'Calibri', lineSpacingMultiple:1.6, valign:'top' });
      }
      s.addText('KompYordam', { x:0, y:6.8, w:'100%', h:0.3, fontSize:9, color:colors.accent, align:'center', italic:true });
    }
    const fileName = `taqdimot_${Date.now()}.pptx`;
    const filePath = path.join(tmpDir, fileName);
    await pptx.writeFile({ fileName: filePath });
    const now = new Date().toLocaleString('uz-UZ', {timeZone:'Asia/Tashkent'});
    await sendTG(`🖥️ <b>POWERPOINT BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Mavzu: ${topic}\n📊 Slaydlar: ${(slidesData.slides||[]).length} ta\n📞 Telefon: ${phone}\n⏰ ${now}`);
    await sendTGFile(filePath, fileName, `🖥️ Tayyor PowerPoint — ${topic}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) { console.error(e); res.status(500).json({ success: false, error: e.message }); }
});

// Order + payment notify
app.post('/api/order', upload.single('screenshot'), async (req, res) => {
  try {
    const { orderType, phone, paidAmount, ...rest } = req.body;
    const now = new Date().toLocaleString('uz-UZ', {timeZone:'Asia/Tashkent'});
    const icons = {excel:'📊',word:'📝',ppt:'🖥️',ai:'🤖',other:'💡'};
    let msg = `💰 <b>YANGI BUYURTMA + TO'LOV</b>\n━━━━━━━━━━━━━━━━\n`;
    msg += `${icons[orderType]||'📋'} <b>Xizmat:</b> ${orderType}\n`;
    if (paidAmount) msg += `💵 <b>To'lov:</b> ${parseInt(paidAmount).toLocaleString()} so'm\n`;
    msg += `━━━━━━━━━━━━━━━━\n`;
    Object.entries(rest).forEach(([k,v]) => { if(v) msg += `📌 <b>${k}:</b> ${v}\n`; });
    msg += `📞 <b>Bog'lanish:</b> ${phone}\n⏰ <b>Vaqt:</b> ${now}`;
    await sendTG(msg);
    if (req.file) {
      const sp = path.join(tmpDir, `ss_${Date.now()}.jpg`);
      fs.writeFileSync(sp, req.file.buffer);
      await sendTGFile(sp, 'screenshot.jpg', `💳 To'lov screenshoti — ${phone}`);
      fs.unlinkSync(sp);
    }
    res.json({ success: true });
  } catch(e) { res.status(500).json({ success: false, error: e.message }); }
});

app.listen(PORT, async () => {
  console.log(`✅ Server port ${PORT} da ishlamoqda`);
  try { await sendTG(`🚀 <b>KompYordam Server ishga tushdi!</b>\n⏰ ${new Date().toLocaleString('uz-UZ',{timeZone:'Asia/Tashkent'})}`); } catch(e){}
});
