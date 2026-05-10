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

// ===== HELPERS =====
async function sendTG(text) {
  try {
    const res = await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: TG_CHAT, text, parse_mode: 'HTML' })
    });
    return await res.json();
  } catch(e) { console.error('TG error:', e.message); }
}

async function sendTGFile(filePath, fileName, caption) {
  try {
    const form = new FormData();
    form.append('chat_id', TG_CHAT);
    form.append('caption', caption || '');
    form.append('document', fs.createReadStream(filePath), { filename: fileName });
    const res = await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendDocument`, { method: 'POST', body: form });
    return await res.json();
  } catch(e) { console.error('TG file error:', e.message); }
}

async function gemini(prompt, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_KEY}`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
          })
        }
      );
      const d = await res.json();
      if (d.error) { console.error('Gemini error:', d.error); continue; }
      return d.candidates?.[0]?.content?.parts?.[0]?.text || '';
    } catch(e) {
      console.error(`Gemini attempt ${i+1} failed:`, e.message);
      if (i < retries - 1) await new Promise(r => setTimeout(r, 2000));
    }
  }
  return '';
}

function parseJSON(text) {
  if (!text) return null;
  try {
    // Remove markdown code blocks
    let cleaned = text
      .replace(/```json\s*/gi, '')
      .replace(/```\s*/gi, '')
      .trim();
    // Find JSON object
    const start = cleaned.indexOf('{');
    const end = cleaned.lastIndexOf('}');
    if (start !== -1 && end !== -1) {
      cleaned = cleaned.substring(start, end + 1);
    }
    return JSON.parse(cleaned);
  } catch(e) {
    console.error('JSON parse error:', e.message, '\nText:', text?.substring(0, 200));
    return null;
  }
}

// ===== ROUTES =====
app.get('/', (req, res) => res.json({ status: 'KompYordam Server OK ✅', version: '3.0' }));

app.get('/test-ppt', async (req, res) => {
  try {
    const PptxGenJS = require('pptxgenjs');
    const pptx = new PptxGenJS();
    const s = pptx.addSlide();
    s.background = { color: '0d0f1a' };
    s.addText('Test slayd', { x:1, y:2, w:8, h:2, fontSize:32, bold:true, color:'f0c060', align:'center' });
    const filePath = '/tmp/test_' + Date.now() + '.pptx';
    await pptx.writeFile({ fileName: filePath });
    const fs = require('fs');
    res.download(filePath, 'test.pptx', () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) {
    res.json({ error: e.message, stack: e.stack?.substring(0,500) });
  }
});

// ===== HUMANIZE =====
app.post('/api/humanize', async (req, res) => {
  try {
    const { text, style } = req.body;
    if (!text) return res.status(400).json({ success: false, error: 'Matn kiritilmagan' });

    const styleDesc = {
      talaba: "oddiy o'zbek talabasi yozgandek — ba'zida kichik xato, jonli, tabiiy",
      rasmiy: "rasmiy hujjat uslubida, lekin inson yozganday tabiiy",
      oddiy: "oddiy suhbat uslubida, sodda, iliq",
      ilmiy: "ilmiy uslubda, lekin tabiiy"
    };

    const prompt = `Sen AI matnni inson uslubiga o'tkazish mutaxassisisisan.

Qoidalar:
1. AI izlarini to'liq olib tashla (birinchidan/ikkinchidan/xulosa qilib/shuni ta'kidlash joizki kabi iboralar)
2. Har bir gapni qayta qur, faqat ma'noni saqla
3. Uslub: ${styleDesc[style] || styleDesc.talaba}
4. Gaplar uzunligi har xil bo'lsin
5. Ba'zi gaplarda noaniqlik yoki shaxsiy fikr qo'sh
6. FAQAT qayta yozilgan matnni ber, hech qanday izoh yozma

Matn:
${text}`;

    const result = await gemini(prompt);
    if (!result) return res.status(500).json({ success: false, error: 'AI javob bermadi' });
    
    res.json({ success: true, content: result });
  } catch(e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ===== WORD =====
app.post('/api/create/word', upload.single('image'), async (req, res) => {
  try {
    const { topic, pages, extra, phone } = req.body;
    if (!topic) return res.status(400).json({ success: false, error: 'Mavzu kiritilmagan' });

    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    await sendTG(`📝 <b>WORD BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Mavzu: ${topic}\n📖 Hajm: ${pages || '10-15 bet'}\n📞 Telefon: ${phone}\n⏰ ${now}\n\n⏳ Tayyorlanmoqda...`);

    const prompt = `"${topic}" mavzusida to'liq referat yoz.
Hajm: ${pages || '10-15 bet'}. ${extra ? 'Qo\'shimcha: ' + extra : ''}

MUHIM: Quyidagi formatda yoz (# va ## sarlavhalar bilan):

# KIRISH
[2-3 paragraf, mavzuning dolzarbligi, maqsad]

# ASOSIY QISM

## 1. [Birinchi bo'lim nomi]
[3-4 paragraf, batafsil ma'lumot]

## 2. [Ikkinchi bo'lim nomi]
[3-4 paragraf, batafsil ma'lumot]

## 3. [Uchinchi bo'lim nomi]
[3-4 paragraf, batafsil ma'lumot]

# XULOSA
[2-3 paragraf, asosiy fikrlar xulosasi]

# FOYDALANILGAN ADABIYOTLAR
1. [Muallif. Kitob nomi. Yil]
2. [Muallif. Kitob nomi. Yil]
3. [Muallif. Kitob nomi. Yil]
4. [Muallif. Kitob nomi. Yil]
5. [Muallif. Kitob nomi. Yil]

O'zbek tilida, inson uslubida, ilmiy-ommabop tarzda yoz.`;

    const content = await gemini(prompt);
    if (!content) throw new Error('AI kontent yaratmadi');

    const lines = content.split('\n');
    const children = [
      new Paragraph({
        text: topic,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 }
      }),
      new Paragraph({
        children: [new TextRun({ text: `Bajardi: Talaba`, size: 24, font: 'Times New Roman' })],
        alignment: AlignmentType.RIGHT,
        spacing: { after: 200 }
      }),
      new Paragraph({
        children: [new TextRun({ text: `Toshkent — ${new Date().getFullYear()}`, size: 24, font: 'Times New Roman' })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 }
      })
    ];

    for (const line of lines) {
      const t = line.trim();
      if (!t) {
        children.push(new Paragraph({ text: '', spacing: { after: 80 } }));
        continue;
      }
      if (t.startsWith('# ')) {
        children.push(new Paragraph({
          text: t.replace('# ', '').toUpperCase(),
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          spacing: { before: 400, after: 200 }
        }));
      } else if (t.startsWith('## ')) {
        children.push(new Paragraph({
          text: t.replace('## ', ''),
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 300, after: 150 }
        }));
      } else if (t.match(/^\d+\./)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, size: 24, font: 'Times New Roman' })],
          spacing: { after: 100 },
          indent: { left: 360 }
        }));
      } else {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, size: 24, font: 'Times New Roman' })],
          spacing: { after: 120 },
          indent: { firstLine: 720 },
          alignment: AlignmentType.BOTH
        }));
      }
    }

    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 1440, right: 1080, bottom: 1440, left: 1800 } } },
        children
      }]
    });

    const fileName = `referat_${Date.now()}.docx`;
    const filePath = path.join(tmpDir, fileName);
    fs.writeFileSync(filePath, await Packer.toBuffer(doc));
    await sendTGFile(filePath, fileName, `📝 <b>WORD TAYYOR!</b>\n📌 ${topic}\n📞 ${phone}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) {
    console.error('Word error:', e);
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
    await sendTG(`📊 <b>EXCEL BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Tavsif: ${description.substring(0,150)}\n📞 Telefon: ${phone}\n⏰ ${now}\n\n⏳ Tayyorlanmoqda...`);

    const prompt = `Sen Excel mutaxassisisisan. Quyidagi so'rov uchun Excel jadval ma'lumotlarini JSON formatda tayyorla.

So'rov: "${description}"

MUHIM QOIDALAR:
1. Faqat sof JSON qaytargin — boshqa hech narsa yozma
2. Kamida 15 ta REAL va ANIQ ma'lumot qatori
3. Formulalar C va D ustunlar asosida hisoblash uchun
4. O'zbek tilida

JSON format (AYNAN shu formatda):
{"title":"Jadval sarlavhasi","sheets":[{"name":"Varaq1","headers":["Ustun1","Ustun2","Ustun3","Ustun4","Ustun5"],"rows":[["qiymat1","qiymat2",100,50000,""],["qiymat3","qiymat4",200,30000,""]],"formulas":[{"cell":"E2","formula":"C2*D2"},{"cell":"E3","formula":"C3*D3"}]}]}`;

    const aiResp = await gemini(prompt);
    console.log('Excel AI response:', aiResp?.substring(0, 500));
    
    let data = parseJSON(aiResp);
    
    // If AI fails, create smart fallback based on description
    if (!data) {
      console.log('Using fallback for excel');
      const isPayroll = description.toLowerCase().includes('oylik') || description.toLowerCase().includes('maosh') || description.toLowerCase().includes('ishchi');
      const isInventory = description.toLowerCase().includes('mahsulot') || description.toLowerCase().includes('tovar') || description.toLowerCase().includes('ombor');
      
      if (isPayroll) {
        data = {
          title: description,
          sheets: [{
            name: "Oylik hisob",
            headers: ["№", "F.I.O.", "Lavozim", "Oylik maoshi", "Ishlagan kunlar", "Hisoblangan", "Ushlab qolish", "Qo'lga tegishi"],
            rows: [
              [1, "Karimov Akbar", "Direktor", 3000000, 22, "", 150000, ""],
              [2, "Rahimova Malika", "Buxgalter", 2500000, 22, "", 125000, ""],
              [3, "Toshmatov Jasur", "Menejer", 2000000, 20, "", 100000, ""],
              [4, "Nazarova Dilnoza", "Kotiba", 1800000, 22, "", 90000, ""],
              [5, "Yusupov Sanjar", "Dasturchi", 3500000, 22, "", 175000, ""],
              [6, "Mirzayeva Gulnora", "Muhandis", 2800000, 21, "", 140000, ""],
              [7, "Abdullayev Timur", "Xavfsizlik", 1500000, 22, "", 75000, ""],
              [8, "Qodirov Sherzod", "Haydovchi", 1700000, 22, "", 85000, ""],
              [9, "Ergasheva Nodira", "Ombor", 1600000, 20, "", 80000, ""],
              [10, "Xolmatov Firdavs", "Texnik", 1900000, 22, "", 95000, ""]
            ],
            formulas: [
              {cell:"F2",formula:"D2/22*E2"},{cell:"H2",formula:"F2-G2"},
              {cell:"F3",formula:"D3/22*E3"},{cell:"H3",formula:"F3-G3"},
              {cell:"F4",formula:"D4/22*E4"},{cell:"H4",formula:"F4-G4"},
              {cell:"F5",formula:"D5/22*E5"},{cell:"H5",formula:"F5-G5"},
              {cell:"F6",formula:"D6/22*E6"},{cell:"H6",formula:"F6-G6"},
              {cell:"F7",formula:"D7/22*E7"},{cell:"H7",formula:"F7-G7"},
              {cell:"F8",formula:"D8/22*E8"},{cell:"H8",formula:"F8-G8"},
              {cell:"F9",formula:"D9/22*E9"},{cell:"H9",formula:"F9-G9"},
              {cell:"F10",formula:"D10/22*E10"},{cell:"H10",formula:"F10-G10"},
              {cell:"F11",formula:"D11/22*E11"},{cell:"H11",formula:"F11-G11"}
            ]
          }]
        };
      } else {
        data = {
          title: description,
          sheets: [{
            name: "Ma'lumotlar",
            headers: ["№", "Nomi", "Kategoriya", "Miqdori", "Narxi (so'm)", "Jami (so'm)"],
            rows: Array.from({length:15}, (_, i) => [i+1, `Mahsulot ${i+1}`, "Asosiy", Math.floor(Math.random()*100)+10, Math.floor(Math.random()*50000)+10000, ""]),
            formulas: Array.from({length:15}, (_, i) => ({cell:`F${i+2}`, formula:`D${i+2}*E${i+2}`}))
          }]
        };
      }
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = 'KompYordam';
    wb.created = new Date();

    for (const sheet of (data.sheets || [])) {
      const ws = wb.addWorksheet(sheet.name || 'Sheet1');
      const colCount = (sheet.headers || []).length || 6;

      // Title
      ws.mergeCells(1, 1, 1, colCount);
      const tc = ws.getCell('A1');
      tc.value = data.title || description;
      tc.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
      tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1a1e35' } };
      tc.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getRow(1).height = 35;

      // Image if uploaded
      if (req.file) {
        try {
          const ext = (req.file.mimetype.split('/')[1] || 'jpeg').replace('jpg', 'jpeg');
          const imgId = wb.addImage({ buffer: req.file.buffer, extension: ext });
          ws.addImage(imgId, { tl: { col: 0, row: 2 }, ext: { width: 250, height: 180 } });
        } catch(ie) { console.log('Image err:', ie.message); }
      }

      // Date row
      ws.mergeCells(2, 1, 2, colCount);
      const dateCell = ws.getCell('A2');
      dateCell.value = `Sana: ${new Date().toLocaleDateString('uz-UZ')}`;
      dateCell.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
      dateCell.alignment = { horizontal: 'right' };

      const startRow = 3;

      // Headers
      if (sheet.headers) {
        const hr = ws.getRow(startRow);
        sheet.headers.forEach((h, i) => {
          const c = hr.getCell(i + 1);
          c.value = h;
          c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563eb' } };
          c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          c.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
          ws.getColumn(i + 1).width = Math.max((h || '').length + 6, 14);
        });
        hr.height = 28;
      }

      // Data rows
      if (sheet.rows) {
        sheet.rows.forEach((row, ri) => {
          const dr = ws.getRow(startRow + 1 + ri);
          (row || []).forEach((val, ci) => {
            const c = dr.getCell(ci + 1);
            if (typeof val === 'string' && val.startsWith('=')) {
              c.value = { formula: val.substring(1) };
            } else if (val === '' && sheet.formulas?.find(f => f.cell === `${String.fromCharCode(65+ci)}${startRow+1+ri}`)) {
              // will be set by formulas
            } else {
              c.value = val;
            }
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ri % 2 === 0 ? 'FFF0F4FF' : 'FFFFFFFF' } };
            c.border = { top: { style: 'thin', color: { argb: 'FFd0d0d0' } }, bottom: { style: 'thin', color: { argb: 'FFd0d0d0' } }, left: { style: 'thin', color: { argb: 'FFd0d0d0' } }, right: { style: 'thin', color: { argb: 'FFd0d0d0' } } };
            if (typeof val === 'number') c.alignment = { horizontal: 'center' };
          });
          dr.height = 22;
        });
      }

      // Formulas
      if (sheet.formulas) {
        sheet.formulas.forEach(f => {
          try {
            const c = ws.getCell(f.cell);
            c.value = { formula: f.formula.replace(/^=/, '') };
            c.font = { bold: true, color: { argb: 'FF16a34a' } };
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf0fff8' } };
          } catch(fe) { console.log('Formula err:', fe.message); }
        });
      }

      // Total row
      const lastDataRow = startRow + (sheet.rows || []).length + 1;
      ws.mergeCells(lastDataRow, 1, lastDataRow, colCount - 1);
      const totalLabel = ws.getCell(`A${lastDataRow}`);
      totalLabel.value = 'JAMI:';
      totalLabel.font = { bold: true, size: 12 };
      totalLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfff3cd' } };
    }

    const fileName = `excel_${Date.now()}.xlsx`;
    const filePath = path.join(tmpDir, fileName);
    await wb.xlsx.writeFile(filePath);
    await sendTGFile(filePath, fileName, `📊 <b>EXCEL TAYYOR!</b>\n📌 ${description.substring(0,100)}\n📞 ${phone}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });
  } catch(e) {
    console.error('Excel error:', e);
    await sendTG(`❌ Excel xatosi: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ===== POWERPOINT - yangi yondashuv =====
app.post('/api/create/pptx', async (req, res) => {
  try {
    const { topic, slides: slideCount, style, phone } = req.body;
    if (!topic) return res.status(400).json({ success: false, error: 'Mavzu kiritilmagan' });

    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    await sendTG(`🖥️ <b>PPT BUYURTMA KELDI</b>\n━━━━━━━━━━━━━\n📌 Mavzu: ${topic}\n📞 Telefon: ${phone}\n⏰ ${now}\n\n⏳ Tayyorlanmoqda...`);

    const numSlides = Math.min(parseInt((slideCount || '12').replace(/[^0-9]/g, '')) || 12, 15);

    // Get content from Gemini as plain text
    const prompt = `"${topic}" mavzusida ${numSlides} ta PowerPoint slayd uchun kontent tayyorla.

Har bir slayd uchun quyidagi formatda yoz:

===SLAYD===
SARLAVHA: [slayd sarlavhasi]
NUQTA1: [birinchi nuqta matni]
NUQTA2: [ikkinchi nuqta matni]
NUQTA3: [uchinchi nuqta matni]
NUQTA4: [to'rtinchi nuqta matni]

O'zbek tilida, professional uslubda. Birinchi slayd kirish, oxirgisi xulosa bo'lsin.`;

    const aiText = await gemini(prompt);
    console.log('AI response length:', aiText ? aiText.length : 0);

    // Parse slides
    const slides = [];
    const blocks = (aiText || '').split('===SLAYD===').filter(b => b.trim());
    
    for (const block of blocks) {
      const lines = block.split('\n').filter(l => l.trim());
      let title = '';
      const points = [];
      for (const line of lines) {
        const t = line.trim();
        if (t.toUpperCase().startsWith('SARLAVHA:')) title = t.replace(/SARLAVHA:/i, '').trim();
        else if (t.toUpperCase().startsWith('NUQTA')) {
          const val = t.replace(/NUQTA\d+:/i, '').trim();
          if (val) points.push(val);
        }
      }
      if (title) slides.push({ title, points });
    }

    // Fallback slides
    if (slides.length < 3) {
      const fb = [
        { title: topic, points: ['Taqdimot maqsadi', 'Asosiy yo\'nalish', 'Kutilayotgan natijalar'] },
        { title: 'Kirish', points: ['Mavzuning dolzarbligi', 'Tadqiqot maqsadi', 'Asosiy savollar', 'Metodologiya'] },
        { title: 'Tarixiy tahlil', points: ['Kelib chiqishi', 'Rivojlanish bosqichlari', 'Hozirgi holat'] },
        { title: 'Asosiy tushunchalar', points: ['Birinchi tushuncha', 'Ikkinchi tushuncha', 'O\'zaro bog\'liqlik'] },
        { title: 'Muammolar', points: ['Mavjud muammolar', 'Sabablari', 'Oqibatlari', 'Echim yo\'llari'] },
        { title: 'Statistika', points: ['Asosiy ko\'rsatkichlar', 'Tendentsiyalar', 'Prognozlar'] },
        { title: 'Xalqaro tajriba', points: ['Yetakchi mamlakatlar', 'Muvaffaqiyatli misollar', 'O\'rganish mumkin bo\'lgan darslar'] },
        { title: 'O\'zbekistonda holat', points: ['Hozirgi vaziyat', 'Amalga oshirilgan ishlar', 'Rejalar'] },
        { title: 'Istiqbol', points: ['Qisqa muddatli maqsadlar', 'Uzoq muddatli rejalar', 'Kutilayotgan natijalar'] },
        { title: 'Tavsiyalar', points: ['Davlat uchun', 'Biznes uchun', 'Fuqarolar uchun'] },
        { title: 'Xulosa', points: ['Asosiy xulosalar', 'Muhim topilmalar', 'Keyingi qadamlar'] },
        { title: 'Savol va javoblar', points: ['Savollaringiz?', 'Bog\'lanish: @kompyordamm'] }
      ];
      slides.length = 0;
      slides.push(...fb.slice(0, numSlides));
    }

    console.log('Slides count:', slides.length);

    // Colors
    const colors = {
      'Zamonaviy': { bg: '0d0f1a', text: 'f0f2ff', accent: 'f0c060', muted: '8890b8' },
      'Klassik / Rasmiy': { bg: 'F8F9FA', text: '1a1a2e', accent: '1d4ed8', muted: '64748b' },
      'Yorqin / Ijodiy': { bg: '1a0533', text: 'FFFFFF', accent: 'ff6b9d', muted: 'c084fc' }
    };
    const c = colors[style] || colors['Zamonaviy'];

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    for (const [idx, slide] of slides.entries()) {
      const s = pptx.addSlide();
      s.background = { color: c.bg };

      if (idx === 0) {
        // Title slide - no shapes, only text
        s.addText(slide.title || topic, {
          x: 0.5, y: 1.5, w: 9, h: 2.5,
          fontSize: 38, bold: true, color: c.text,
          align: 'center', fontFace: 'Calibri',
          breakLine: false
        });
        s.addText('─────────────────────────', {
          x: 1.5, y: 3.8, w: 7, h: 0.4,
          fontSize: 14, color: c.accent, align: 'center'
        });
        if (slide.points && slide.points[0]) {
          s.addText(slide.points[0], {
            x: 0.5, y: 4.3, w: 9, h: 0.8,
            fontSize: 18, color: c.muted, align: 'center', italic: true
          });
        }
        s.addText(`KompYordam · ${new Date().getFullYear()}`, {
          x: 0, y: 6.6, w: '100%', h: 0.35,
          fontSize: 10, color: c.accent, align: 'center'
        });
        s.addText(`${idx + 1} / ${slides.length}`, {
          x: 8.5, y: 0.1, w: 1.3, h: 0.3,
          fontSize: 9, color: c.muted, align: 'right'
        });
      } else {
        // Content slide
        // Header background simulation with text block
        s.addText(slide.title || '', {
          x: 0.2, y: 0.1, w: 9.6, h: 0.85,
          fontSize: 24, bold: true, color: c.text,
          fontFace: 'Calibri', valign: 'middle'
        });
        s.addText('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', {
          x: 0.2, y: 1.0, w: 9.6, h: 0.25,
          fontSize: 7, color: c.accent, align: 'left'
        });

        const pts = (slide.points || []).slice(0, 5);
        pts.forEach((point, pi) => {
          const y = 1.35 + pi * 1.0;
          if (y > 6.4) return;
          s.addText(`▶  ${point}`, {
            x: 0.4, y, w: 9.2, h: 0.85,
            fontSize: 16, color: c.text,
            fontFace: 'Calibri', valign: 'middle'
          });
        });

        s.addText(`${idx + 1} / ${slides.length}  |  KompYordam`, {
          x: 0, y: 6.75, w: '100%', h: 0.25,
          fontSize: 8, color: c.muted, align: 'center'
        });
      }
    }

    const fileName = `taqdimot_${Date.now()}.pptx`;
    const filePath = path.join(tmpDir, fileName);
    
    console.log('Writing PPTX to:', filePath);
    await pptx.writeFile({ fileName: filePath });
    console.log('PPTX written successfully');
    
    await sendTGFile(filePath, fileName, `🖥️ <b>PPT TAYYOR!</b>\n📌 ${topic}\n📊 ${slides.length} ta slayd\n📞 ${phone}`);
    res.download(filePath, fileName, () => { try { fs.unlinkSync(filePath); } catch(e){} });

  } catch(e) {
    console.error('PPT FULL ERROR:', e.message);
    console.error('STACK:', e.stack);
    await sendTG(`❌ PPT xatosi: ${e.message}\n${e.stack ? e.stack.substring(0,300) : ''}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ===== ORDER NOTIFY =====
app.post('/api/order', upload.single('screenshot'), async (req, res) => {
  try {
    const { orderType, phone, paidAmount, ...rest } = req.body;
    const now = new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' });
    const icons = { excel: '📊', word: '📝', ppt: '🖥️', ai: '🤖', other: '💡' };
    const names = { excel: 'Excel Hujjat', word: 'Word/Referat', ppt: 'PowerPoint', ai: 'AI Humanizer', other: 'Boshqa' };

    let msg = `💰 <b>YANGI BUYURTMA + TO'LOV</b>\n━━━━━━━━━━━━━━━━\n`;
    msg += `${icons[orderType] || '📋'} <b>Xizmat:</b> ${names[orderType] || orderType}\n`;
    if (paidAmount) msg += `💵 <b>To'lov:</b> ${parseInt(paidAmount).toLocaleString()} so'm\n`;
    msg += `━━━━━━━━━━━━━━━━\n`;
    Object.entries(rest).forEach(([k, v]) => { if (v && typeof v === 'string') msg += `📌 <b>${k}:</b> ${v.substring(0, 200)}\n`; });
    msg += `📞 <b>Bog'lanish:</b> ${phone}\n⏰ <b>Vaqt:</b> ${now}`;

    await sendTG(msg);

    if (req.file) {
      const sp = path.join(tmpDir, `ss_${Date.now()}.jpg`);
      fs.writeFileSync(sp, req.file.buffer);
      await sendTGFile(sp, 'tolов_screenshot.jpg', `💳 To'lov screenshoti — ${phone}`);
      try { fs.unlinkSync(sp); } catch(e) {}
    }

    res.json({ success: true });
  } catch(e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

app.listen(PORT, async () => {
  console.log(`✅ KompYordam Server v3.0 — port ${PORT}`);
  try {
    await sendTG(`🚀 <b>KompYordam Server v3.0 ishga tushdi!</b>\n✅ Barcha xizmatlar tayyor\n⏰ ${new Date().toLocaleString('uz-UZ', { timeZone: 'Asia/Tashkent' })}`);
  } catch(e) {}
});
