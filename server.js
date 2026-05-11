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
const XAI_KEY = process.env.XAI_KEY || 'xai-DsvWTca3tMCc3PSJNVOtvSpadfBbnxAYWqzBcmqh94Lfzr4MWrTw51WTYZoF6fWElaWN44XfTT2a5jbL';

app.use(cors());
app.use(express.json({ limit: '50mb' }));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const tmpDir = '/tmp/kompyordam';
if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });

// ===== HELPERS =====
async function sendTG(text) {
  try {
    await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendMessage`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: TG_CHAT, text, parse_mode: 'HTML' })
    });
  } catch(e) { console.error('TG error:', e.message); }
}

async function sendTGFile(filePath, fileName, caption) {
  try {
    const form = new FormData();
    form.append('chat_id', TG_CHAT);
    form.append('caption', caption || '');
    form.append('document', fs.createReadStream(filePath), { filename: fileName });
    await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendDocument`, { method: 'POST', body: form });
  } catch(e) { console.error('TG file error:', e.message); }
}

// Gemini with 60 second timeout
async function gemini(prompt) {
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 55000);
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_KEY}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.7, maxOutputTokens: 4096 }
        }),
        signal: controller.signal
      }
    );
    clearTimeout(timeout);
    const d = await res.json();
    if (d.error) { console.error('Gemini API error:', JSON.stringify(d.error)); return null; }
    return d.candidates?.[0]?.content?.parts?.[0]?.text || null;
  } catch(e) {
    console.error('Gemini fetch error:', e.message);
    return null;
  }
}

// ===== GROK AI (xAI) =====
async function grok(prompt) {
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 55000);
    const res = await fetch('https://api.x.ai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${XAI_KEY}`
      },
      body: JSON.stringify({
        model: 'grok-beta',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 4096,
        temperature: 0.7
      }),
      signal: controller.signal
    });
    clearTimeout(timeout);
    const d = await res.json();
    if (d.error) { console.error('Grok error:', d.error); return null; }
    return d.choices?.[0]?.message?.content || null;
  } catch(e) {
    console.error('Grok fetch error:', e.message);
    return null;
  }
}

// Smart AI: Gemini first, Grok fallback
async function smartAI(prompt) {
  console.log('Trying Gemini...');
  const geminiResult = await gemini(prompt);
  if (geminiResult) {
    console.log('Gemini SUCCESS');
    return geminiResult;
  }
  console.log('Gemini failed, trying Grok...');
  const grokResult = await grok(prompt);
  if (grokResult) {
    console.log('Grok SUCCESS');
    return grokResult;
  }
  console.log('Both AI failed');
  return null;
}

// ===== ROUTES =====
app.get('/', (req, res) => res.json({ status: 'KompYordam Server v4.0 OK ✅' }));

app.get('/test-ppt', async (req, res) => {
  try {
    const pptx = new PptxGenJS();
    const s = pptx.addSlide();
    s.background = { color: '0d0f1a' };
    s.addText('Test slayd - OK!', { x:1, y:2, w:8, h:2, fontSize:32, bold:true, color:'f0c060', align:'center' });
    const fp = `/tmp/test_${Date.now()}.pptx`;
    await pptx.writeFile({ fileName: fp });
    res.download(fp, 'test.pptx', () => { try { fs.unlinkSync(fp); } catch(e){} });
  } catch(e) { res.json({ error: e.message }); }
});

// ===== ORDER NOTIFY =====
app.post('/api/order', upload.single('screenshot'), async (req, res) => {
  try {
    const { orderType, phone, paidAmount, ...rest } = req.body;
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    const icons = { excel:'📊', word:'📝', ppt:'🖥️', ai:'🤖', other:'💡' };
    const names = { excel:'Excel Hujjat', word:'Word/Referat', ppt:'PowerPoint', ai:'AI Humanizer', other:'Boshqa' };
    let msg = `💰 <b>YANGI BUYURTMA + TO'LOV</b>\n━━━━━━━━━━━━━━━━\n`;
    msg += `${icons[orderType]||'📋'} <b>Xizmat:</b> ${names[orderType]||orderType}\n`;
    if (paidAmount) msg += `💵 <b>To'lov:</b> ${parseInt(paidAmount).toLocaleString()} so'm\n`;
    msg += `━━━━━━━━━━━━━━━━\n`;
    Object.entries(rest).forEach(([k,v]) => { if(v && typeof v==='string') msg += `📌 <b>${k}:</b> ${v.substring(0,200)}\n`; });
    msg += `📞 <b>Bog'lanish:</b> ${phone}\n⏰ <b>Vaqt:</b> ${now}`;
    await sendTG(msg);
    if (req.file) {
      const sp = path.join(tmpDir, `ss_${Date.now()}.jpg`);
      fs.writeFileSync(sp, req.file.buffer);
      await sendTGFile(sp, 'screenshot.jpg', `💳 To'lov screenshoti — ${phone}`);
      try { fs.unlinkSync(sp); } catch(e) {}
    }
    res.json({ success: true });
  } catch(e) { res.status(500).json({ success: false, error: e.message }); }
});

// ===== HUMANIZE =====
app.post('/api/humanize', async (req, res) => {
  try {
    const { text, style, phone, paidAmount } = req.body;
    if (!text) return res.status(400).json({ success: false, error: 'Matn kiritilmagan' });
    
    const styleDesc = {
      talaba: "oddiy o'zbek talabasi yozgandek — ba'zida kichik xato, jonli, tabiiy. Rasmiy iboralar yo'q.",
      rasmiy: "rasmiy hujjat uslubida, lekin inson yozganday tabiiy. Professional, aniq.",
      oddiy: "oddiy suhbat uslubida, sodda, iliq. Qisqa va lo'nda gaplar.",
      ilmiy: "ilmiy uslubda, lekin tabiiy inson yozganday. Terminlar to'g'ri ishlatilsin."
    };
    const styleNames = { talaba:"Talaba uslubi", rasmiy:"Rasmiy hujjat", oddiy:"Oddiy suhbat", ilmiy:"Ilmiy maqola" };
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    
    await sendTG(`🤖 <b>AI HUMANIZER BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📝 Uslub: ${styleNames[style]||style}\n📏 Matn: ${text.length} belgi\n📞 Tel: ${phone||'—'}\n💵 ${paidAmount ? parseInt(paidAmount).toLocaleString()+' so\'m' : '—'}\n⏰ ${now}\n⏳ AI ishlayapti...`);
    
    const prompt = `Sen AI matnni inson uslubiga o'tkazish mutaxassisisisan.

VAZIFA: Quyidagi AI matnni ${styleDesc[style]||styleDesc.talaba} qilib to'liq qayta yoz.

MAJBURIY QOIDALAR:
1. "Birinchidan", "Ikkinchidan", "Xulosa qilib", "Shuni ta'kidlash joizki", "Bundan tashqari" — ISHLATMA
2. Har bir gapni boshqacha tuzil
3. Ba'zi gaplarda shaxsiy kuzatuv bo'lsin
4. Gaplar uzunligi xilma-xil bo'lsin
5. FAQAT qayta yozilgan matnni ber — izoh YOZMA

MATN:
${text}`;

    const result = await smartAI(prompt);
    if (!result) {
      await sendTG(`❌ AI Humanizer: Gemini javob bermadi\n📞 ${phone||'—'}`);
      return res.status(500).json({ success: false, error: "AI javob bermadi, qayta urinib ko'ring" });
    }
    
    // Send result to Telegram (split if too long)
    const header = `✅ <b>AI HUMANIZER NATIJA</b>\n━━━━━━━━━━━━━\n📝 ${styleNames[style]||style} | 📞 ${phone||'—'}\n⏰ ${now}\n━━━━━━━━━━━━━\n`;
    const maxLen = 3500;
    if (result.length <= maxLen) {
      await sendTG(header + result);
    } else {
      await sendTG(header + result.substring(0, maxLen) + '\n...(davomi)');
      await sendTG(`✅ <b>NATIJA davomi:</b>\n━━━━━━━━━━━━━\n${result.substring(maxLen)}`);
    }
    
    res.json({ success: true, content: result });
  } catch(e) {
    await sendTG(`❌ AI Humanizer xatosi: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ===== WORD =====
app.post('/api/create/word', upload.single('image'), async (req, res) => {
  try {
    const { topic, pages, extra, phone } = req.body;
    if (!topic) return res.status(400).json({ success: false, error: 'Mavzu kiritilmagan' });
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    await sendTG(`📝 <b>WORD BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Mavzu: ${topic}\n📖 Hajm: ${pages||'10-15 bet'}\n📞 Tel: ${phone}\n⏰ ${now}\n⏳ Tayyorlanmoqda...`);

    const prompt = `"${topic}" mavzusida referat yoz. Hajm: ${pages||'10-15 bet'}. ${extra||''}
AYNAN quyidagi tuzilmada yoz:
# KIRISH
[kirish matni - 2-3 paragraf]

## 1. [Birinchi bo'lim]
[matn - 3-4 paragraf]

## 2. [Ikkinchi bo'lim]
[matn - 3-4 paragraf]

## 3. [Uchinchi bo'lim]
[matn - 3-4 paragraf]

# XULOSA
[xulosa - 2-3 paragraf]

# FOYDALANILGAN ADABIYOTLAR
1. [manba]
2. [manba]
3. [manba]
4. [manba]
5. [manba]

O'zbek tilida, inson uslubida yoz.`;

    const content = await smartAI(prompt);
    if (!content) {
      await sendTG(`❌ Word xatosi: Gemini javob bermadi\n📌 Mavzu: ${topic}`);
      return res.status(500).json({ success: false, error: 'AI kontent yaratmadi, qayta urinib ko\'ring' });
    }

    const children = [
      new Paragraph({ text: topic, heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, spacing: { after: 400 } }),
      new Paragraph({ children: [new TextRun({ text: `Bajardi: ___________`, size: 24, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT, spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: `Toshkent — ${new Date().getFullYear()}`, size: 24, font: 'Times New Roman' })], alignment: AlignmentType.CENTER, spacing: { after: 600 } })
    ];

    for (const line of content.split('\n')) {
      const t = line.trim();
      if (!t) { children.push(new Paragraph({ text: '', spacing: { after: 60 } })); continue; }
      if (t.startsWith('# ')) {
        children.push(new Paragraph({ text: t.replace('# ','').toUpperCase(), heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { before: 400, after: 200 } }));
      } else if (t.startsWith('## ')) {
        children.push(new Paragraph({ text: t.replace('## ',''), heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 150 } }));
      } else if (t.match(/^\d+\./)) {
        children.push(new Paragraph({ children: [new TextRun({ text: t, size: 24, font: 'Times New Roman' })], spacing: { after: 80 }, indent: { left: 360 } }));
      } else {
        children.push(new Paragraph({ children: [new TextRun({ text: t, size: 24, font: 'Times New Roman' })], spacing: { after: 100 }, indent: { firstLine: 720 }, alignment: AlignmentType.BOTH }));
      }
    }

    const doc = new Document({ sections: [{ properties: { page: { margin: { top:1440, right:1080, bottom:1440, left:1800 } } }, children }] });
    const fileName = `referat_${Date.now()}.docx`;
    const filePath = path.join(tmpDir, fileName);
    fs.writeFileSync(filePath, await Packer.toBuffer(doc));
    await sendTGFile(filePath, fileName, `📝 <b>WORD TAYYOR!</b>\n📌 ${topic}\n📞 ${phone}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) {
    console.error('Word error:', e.message);
    await sendTG(`❌ Word xatosi: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ===== EXCEL =====
app.post('/api/create/excel', upload.single('image'), async (req, res) => {
  try {
    const { description, phone } = req.body;
    if (!description) return res.status(400).json({ success: false, error: 'Tavsif kiritilmagan' });
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    await sendTG(`📊 <b>EXCEL BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 ${description.substring(0,150)}\n📞 ${phone}\n⏰ ${now}\n⏳ Tayyorlanmoqda...`);

    // Parse description to extract requirements
    const desc = description.toLowerCase();
    const isPayroll = desc.includes('oylik') || desc.includes('maosh') || desc.includes('ishchi') || desc.includes('xodim');
    const isBaholar = desc.includes('baho') || desc.includes('talaba') || desc.includes('o\'quvchi');
    const isInventory = desc.includes('mahsulot') || desc.includes('tovar') || desc.includes('ombor');
    const isBudget = desc.includes('xarajat') || desc.includes('daromad') || desc.includes('byudjet');

    // Extract numbers from description
    const nums = description.match(/\d+/g) || [];
    const rowCount = parseInt(nums[0]) || 10;

    let title = description;
    let headers = [];
    let rows = [];
    let formulas = [];

    if (isPayroll) {
      // Generate actual payroll names with Gemini
      const namePrompt = `${rowCount} ta o'zbek ismini va lavozimini quyidagi formatda yoz (har biri yangi qatorda):
Ism Familiya | Lavozim | Oylik maoshi
Masalan: Karimov Akbar | Direktor | 3500000
Faqat ro'yxatni ber, boshqa narsa yozma.`;
      const namesText = await smartAI(namePrompt);
      
      headers = ['№', 'F.I.O.', 'Lavozim', 'Oylik maoshi', 'Ishlagan kun', 'Hisoblangan', 'Ushlab qolish (5%)', 'Qo\'lga tegishi'];
      
      if (namesText) {
        const nameLines = namesText.split('\n').filter(l => l.includes('|')).slice(0, rowCount);
        nameLines.forEach((line, i) => {
          const parts = line.split('|').map(p => p.trim());
          const name = parts[0] || `Xodim ${i+1}`;
          const position = parts[1] || 'Xodim';
          const salary = parseInt((parts[2] || '2000000').replace(/[^0-9]/g, '')) || 2000000;
          rows.push([i+1, name, position, salary, 22, '', '', '']);
          formulas.push({ cell: `F${i+2}`, formula: `D${i+2}/22*E${i+2}` });
          formulas.push({ cell: `G${i+2}`, formula: `F${i+2}*0.05` });
          formulas.push({ cell: `H${i+2}`, formula: `F${i+2}-G${i+2}` });
        });
      } else {
        // Fallback names
        const names = ['Karimov Akbar','Rahimova Malika','Toshmatov Jasur','Nazarova Dilnoza','Yusupov Sanjar','Mirzayeva Gulnora','Abdullayev Timur','Qodirov Sherzod','Ergasheva Nodira','Xolmatov Firdavs','Hasanova Zulfiya','Tursunov Bobur'];
        const positions = ['Direktor','Buxgalter','Menejer','Kotiba','Dasturchi','Muhandis','Xavfsizlik','Haydovchi','Ombor','Texnik','Iqtisodchi','Maslahatchi'];
        const salaries = [3500000,2800000,2200000,1900000,3200000,2600000,1600000,1800000,1700000,2000000,2400000,2100000];
        for (let i = 0; i < Math.min(rowCount, names.length); i++) {
          rows.push([i+1, names[i], positions[i], salaries[i], 22, '', '', '']);
          formulas.push({ cell: `F${i+2}`, formula: `D${i+2}/22*E${i+2}` });
          formulas.push({ cell: `G${i+2}`, formula: `F${i+2}*0.05` });
          formulas.push({ cell: `H${i+2}`, formula: `F${i+2}-G${i+2}` });
        }
      }
    } else if (isBaholar) {
      headers = ['№', 'Talaba F.I.O.', 'Matematika', 'Fizika', 'Kimyo', 'Tarix', 'Adabiyot', "O'rtacha baho", 'Natija'];
      const namePrompt = `${rowCount} ta o'zbek talaba ismini yoz, har biri yangi qatorda. Faqat ismlar.`;
      const namesText = await smartAI(namePrompt);
      const nameLines = (namesText || '').split('\n').filter(l => l.trim()).slice(0, rowCount);
      for (let i = 0; i < rowCount; i++) {
        const name = nameLines[i] || `Talaba ${i+1}`;
        const grades = [Math.floor(Math.random()*30)+70, Math.floor(Math.random()*30)+70, Math.floor(Math.random()*30)+70, Math.floor(Math.random()*30)+70, Math.floor(Math.random()*30)+70];
        rows.push([i+1, name.trim(), ...grades, '', '']);
        formulas.push({ cell: `H${i+2}`, formula: `AVERAGE(C${i+2}:G${i+2})` });
        formulas.push({ cell: `I${i+2}`, formula: `IF(H${i+2}>=90,"A'lo",IF(H${i+2}>=70,"Yaxshi","Qoniqarli"))` });
      }
    } else {
      // Generic - ask Gemini for relevant data
      const dataPrompt = `"${description}" uchun Excel jadval ma'lumotlarini yoz.
Quyidagi AYNAN shu formatda yoz (| bilan ajrat):
SARLAVHA: [jadval nomi]
USTUNLAR: Ustun1 | Ustun2 | Ustun3 | Ustun4 | Ustun5
QATOR1: qiymat1 | qiymat2 | 100 | 50000 | 
QATOR2: qiymat3 | qiymat4 | 200 | 30000 |
(kamida ${rowCount} ta qator, real ma'lumotlar bilan)
O'zbek tilida.`;
      
      const dataText = await smartAI(dataPrompt);
      
      if (dataText) {
        const lines = dataText.split('\n');
        for (const line of lines) {
          const t = line.trim();
          if (t.startsWith('SARLAVHA:')) title = t.replace('SARLAVHA:','').trim();
          else if (t.startsWith('USTUNLAR:')) headers = t.replace('USTUNLAR:','').split('|').map(h => h.trim()).filter(h => h);
          else if (t.startsWith('QATOR')) {
            const vals = t.replace(/QATOR\d+:/,'').split('|').map(v => v.trim());
            const parsedVals = vals.map(v => isNaN(v) || v==='' ? v : Number(v));
            rows.push(parsedVals);
          }
        }
      }
      
      if (headers.length === 0) {
        headers = ['№', 'Nomi', 'Kategoriya', 'Miqdori', 'Narxi (so\'m)', 'Jami (so\'m)'];
        rows = Array.from({length: rowCount}, (_, i) => [i+1, `Element ${i+1}`, 'Asosiy', Math.floor(Math.random()*100)+10, Math.floor(Math.random()*50000)+10000, '']);
        formulas = Array.from({length: rowCount}, (_, i) => ({ cell: `F${i+2}`, formula: `D${i+2}*E${i+2}` }));
      }
    }

    // Build Excel
    const wb = new ExcelJS.Workbook();
    wb.creator = 'KompYordam';
    const ws = wb.addWorksheet("Ma'lumotlar");
    const colCount = headers.length;

    // Title row
    ws.mergeCells(1, 1, 1, colCount);
    const tc = ws.getCell('A1');
    tc.value = title;
    tc.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1a1e35' } };
    tc.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 32;

    // Date
    ws.mergeCells(2, 1, 2, colCount);
    const dc = ws.getCell('A2');
    dc.value = `Sana: ${new Date().toLocaleDateString('uz-UZ')}  |  KompYordam`;
    dc.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
    dc.alignment = { horizontal: 'right' };
    ws.getRow(2).height = 18;

    // Image
    if (req.file) {
      try {
        const ext = (req.file.mimetype.split('/')[1]||'jpeg').replace('jpg','jpeg');
        const imgId = wb.addImage({ buffer: req.file.buffer, extension: ext });
        ws.addImage(imgId, { tl: { col: 0, row: 2 }, ext: { width: 200, height: 150 } });
      } catch(ie) { console.log('Image err:', ie.message); }
    }

    // Headers row 3
    const hr = ws.getRow(3);
    headers.forEach((h, i) => {
      const c = hr.getCell(i+1);
      c.value = h;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563eb' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      c.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
      ws.getColumn(i+1).width = Math.max((h||'').length + 6, 14);
    });
    hr.height = 28;

    // Data rows
    rows.forEach((row, ri) => {
      const dr = ws.getRow(4 + ri);
      (row||[]).forEach((val, ci) => {
        const c = dr.getCell(ci+1);
        if (typeof val === 'string' && val.startsWith('=')) c.value = { formula: val.substring(1) };
        else c.value = val;
        c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: ri%2===0 ? 'FFF0F4FF' : 'FFFFFFFF' } };
        c.border = { top:{style:'thin',color:{argb:'FFd0d0d0'}}, bottom:{style:'thin',color:{argb:'FFd0d0d0'}}, left:{style:'thin',color:{argb:'FFd0d0d0'}}, right:{style:'thin',color:{argb:'FFd0d0d0'}} };
        if (typeof val === 'number') c.alignment = { horizontal: 'center' };
      });
      dr.height = 22;
    });

    // Formulas
    formulas.forEach(f => {
      try {
        const c = ws.getCell(f.cell);
        c.value = { formula: f.formula };
        c.font = { bold: true, color: { argb: 'FF16a34a' } };
        c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: 'FFf0fff8' } };
        c.numFmt = '#,##0';
      } catch(fe) { console.log('Formula err:', fe.message); }
    });

    // Total row
    const totalRow = ws.getRow(4 + rows.length);
    ws.mergeCells(4+rows.length, 1, 4+rows.length, colCount-1);
    const tl = totalRow.getCell(1);
    tl.value = 'JAMI:';
    tl.font = { bold: true, size: 12 };
    tl.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: 'FFfff3cd' } };

    const fileName = `excel_${Date.now()}.xlsx`;
    const filePath = path.join(tmpDir, fileName);
    await wb.xlsx.writeFile(filePath);
    await sendTGFile(filePath, fileName, `📊 <b>EXCEL TAYYOR!</b>\n📌 ${description.substring(0,100)}\n📞 ${phone}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) {
    console.error('Excel error:', e.message);
    await sendTG(`❌ Excel xatosi: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ===== POWERPOINT =====
app.post('/api/create/pptx', async (req, res) => {
  try {
    const { topic, slides: slideCount, style, phone } = req.body;
    if (!topic) return res.status(400).json({ success: false, error: 'Mavzu kiritilmagan' });
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    await sendTG(`🖥️ <b>PPT BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Mavzu: ${topic}\n📊 ${slideCount}\n📞 ${phone}\n⏰ ${now}\n⏳ Tayyorlanmoqda...`);

    const numSlides = Math.min(parseInt((slideCount||'12').replace(/[^0-9]/g,''))||12, 15);

    const prompt = `"${topic}" mavzusida ${numSlides} ta slayd. Har biri AYNAN bu formatda:

SLAYD1
SARLAVHA: [sarlavha]
- [nuqta 1]
- [nuqta 2]
- [nuqta 3]
- [nuqta 4]

SLAYD2
SARLAVHA: [sarlavha]
- [nuqta 1]
- [nuqta 2]
- [nuqta 3]

O'zbek tilida. 1-slayd kirish, oxirgisi xulosa.`;

    const aiText = await smartAI(prompt);
    const slides = [];

    if (aiText) {
      // Split by SLAYD pattern
      const parts = aiText.split(/SLAYD\d+/i);
      for (const part of parts) {
        if (!part.trim()) continue;
        const lines = part.split('\n').filter(l => l.trim());
        let title = '';
        const points = [];
        for (const line of lines) {
          const t = line.trim();
          if (t.toUpperCase().startsWith('SARLAVHA:')) title = t.replace(/SARLAVHA:/i,'').trim();
          else if (t.startsWith('-') || t.startsWith('•')) points.push(t.replace(/^[-•]\s*/,'').trim());
          else if (t && !t.match(/^[A-Z]+:/) && t.length > 3) points.push(t);
        }
        if (title || points.length > 0) slides.push({ title: title||topic, points: points.slice(0,5) });
      }
    }

    // Fallback
    if (slides.length < 3) {
      const fb = [
        { title: topic, points: ["Taqdimot maqsadi va vazifalari","Mavzuning dolzarbligi","Asosiy yo'nalishlar"] },
        { title: "Kirish", points: ["Mavzu haqida umumiy ma'lumot","Tadqiqot predmeti","Asosiy savollar va muammolar","Ishning ahamiyati"] },
        { title: "Asosiy tushunchalar", points: ["Birinchi asosiy tushuncha ta'rifi","Ikkinchi tushuncha va uning xususiyatlari","Tushunchalar o'rtasidagi bog'liqlik"] },
        { title: "Tahlil", points: ["Mavjud holat tahlili","Asosiy ko'rsatkichlar","Tendentsiyalar va o'zgarishlar","Prognoz"] },
        { title: "Muammolar va yechimlar", points: ["Asosiy muammolar","Ularning sabablari","Taklif etilayotgan yechimlar","Amalga oshirish yo'llari"] },
        { title: "Xalqaro tajriba", points: ["Rivojlangan mamlakatlar tajribasi","Muvaffaqiyatli misollar","O'zbekiston uchun darslar"] },
        { title: "O'zbekistondagi holat", points: ["Hozirgi vaziyat","Amalga oshirilgan ishlar","Rejalashtirilgan tadbirlar","Kutilayotgan natijalar"] },
        { title: "Statistik ma'lumotlar", points: ["Asosiy raqamli ko'rsatkichlar","Grafiklar va jadvallar tahlili","Yillar bo'yicha dinamika"] },
        { title: "Istiqbol", points: ["Qisqa muddatli maqsadlar (1-2 yil)","O'rta muddatli rejalar (3-5 yil)","Uzoq muddatli istiqbol"] },
        { title: "Tavsiyalar", points: ["Davlat tashkilotlari uchun tavsiyalar","Xususiy sektor uchun","Fuqarolik jamiyati uchun"] },
        { title: "Xulosa", points: ["Asosiy natijalar va xulosalar","Muhim topilmalar","Keyingi qadamlar"] },
        { title: "Savol va javoblar", points: ["Savollaringiz bormi?","Muhokama uchun mavzular","Bog'lanish: @kompyordamm"] }
      ];
      slides.length = 0;
      slides.push(...fb.slice(0, numSlides));
    }

    const colors = {
      'Zamonaviy': { bg:'0d0f1a', text:'f0f2ff', accent:'f0c060', muted:'8890b8' },
      'Klassik / Rasmiy': { bg:'F8F9FA', text:'1a1a2e', accent:'1d4ed8', muted:'64748b' },
      'Yorqin / Ijodiy': { bg:'1a0533', text:'FFFFFF', accent:'ff6b9d', muted:'c084fc' }
    };
    const c = colors[style] || colors['Zamonaviy'];

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    for (const [idx, slide] of slides.entries()) {
      const s = pptx.addSlide();
      s.background = { color: c.bg };

      s.addText(`${idx+1}/${slides.length}`, { x:8.5, y:0.08, w:1.3, h:0.3, fontSize:9, color:c.muted, align:'right' });

      if (idx === 0) {
        s.addText(slide.title||topic, { x:0.5, y:1.8, w:9, h:2.2, fontSize:36, bold:true, color:c.text, align:'center', fontFace:'Calibri' });
        s.addText('━━━━━━━━━━━━━━━━━━━━━━━━━━━', { x:1.5, y:3.9, w:7, h:0.3, fontSize:12, color:c.accent, align:'center' });
        if ((slide.points||[]).length > 0) {
          s.addText(slide.points[0], { x:0.5, y:4.3, w:9, h:0.7, fontSize:16, color:c.muted, align:'center', italic:true });
        }
        s.addText(`KompYordam · ${new Date().getFullYear()}`, { x:0, y:6.6, w:'100%', h:0.3, fontSize:9, color:c.accent, align:'center' });
      } else {
        s.addText(slide.title||'', { x:0.3, y:0.15, w:9.4, h:0.85, fontSize:24, bold:true, color:c.text, fontFace:'Calibri', valign:'middle' });
        s.addText('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', { x:0.3, y:1.05, w:9.4, h:0.15, fontSize:6, color:c.accent });
        (slide.points||[]).slice(0,5).forEach((point, pi) => {
          const y = 1.3 + pi * 1.0;
          if (y > 6.3) return;
          s.addText(`▶  ${point}`, { x:0.4, y, w:9.2, h:0.85, fontSize:15, color:c.text, fontFace:'Calibri', valign:'middle' });
        });
        s.addText(`KompYordam`, { x:0, y:6.78, w:'100%', h:0.22, fontSize:8, color:c.muted, align:'center' });
      }
    }

    const fileName = `taqdimot_${Date.now()}.pptx`;
    const filePath = path.join(tmpDir, fileName);
    await pptx.writeFile({ fileName: filePath });
    await sendTGFile(filePath, fileName, `🖥️ <b>PPT TAYYOR!</b>\n📌 ${topic}\n📊 ${slides.length} ta slayd\n📞 ${phone}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) {
    console.error('PPT error:', e.message, e.stack?.substring(0,300));
    await sendTG(`❌ PPT xatosi: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.listen(PORT, async () => {
  console.log(`✅ KompYordam Server v4.0 — port ${PORT}`);
  try { await sendTG(`🚀 <b>KompYordam Server v4.1 ishga tushdi!</b>\n✅ Excel, Word, PPT — hammasi tayyor\n🤖 AI: Gemini + Grok (xAI) ikki AI\n⏰ ${new Date().toLocaleString('uz-UZ',{timeZone:'Asia/Tashkent'})}`); } catch(e) {}
});
