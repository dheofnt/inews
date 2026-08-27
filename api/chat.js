// api/chat.js
// Backend ini meniru "Claude Project": instructions.md = custom instructions,
// folder knowledge/ = project knowledge. Keduanya dibaca dari server (bukan
// dari browser visitor), jadi SETIAP orang yang chat lewat link ini otomatis
// mendapat jawaban yang berbasis data & gaya yang sudah kamu tentukan.
//
// Otak AI-nya pakai Google Gemini API (bukan Claude) supaya bisa dipakai
// gratis dalam batas kuota — arsitektur & fitur (chart, file) tetap sama.
// Ganti model di GEMINI_MODEL kalau perlu; cek daftar model & syarat kuota
// gratis terbaru di https://ai.google.dev/gemini-api/docs/models sebelum
// deploy, karena ini bisa berubah sewaktu-waktu.
//
// API key GEMINI_API_KEY tetap hanya hidup sebagai environment variable
// di server — tidak pernah dikirim ke browser.

import fs from 'fs';
import path from 'path';

const GEMINI_MODEL = 'gemini-3.6-flash';

// Aturan chart ini SELALU ditambahkan ke system prompt, terpisah dari
// instructions.md, supaya fitur grafik tetap jalan walau owner belum
// menulisnya sendiri di instructions.md.
const CHART_INSTRUCTIONS = `
Kalau pengguna meminta grafik/chart/visualisasi angka, dan kamu punya data
yang cukup akurat (dari data referensi di atas atau dari percakapan),
sertakan SATU blok berikut di balasanmu, di luar teks penjelasan biasa:

\`\`\`chart
{
  "type": "bar",
  "title": "Judul singkat grafik",
  "labels": ["Label1", "Label2"],
  "series": [
    { "name": "Nama seri", "values": [angka1, angka2] }
  ]
}
\`\`\`

Aturan blok chart:
- "type" hanya boleh "bar", "line", atau "pie".
- Untuk "pie", gunakan tepat satu seri di "series".
- Selalu sertakan juga teks penjelasan singkat (sebelum/sesudah blok), jangan
  cuma blok chart saja.
- Kalau data tidak cukup untuk grafik yang akurat, jangan mengarang angka —
  jelaskan keterbatasannya sebagai teks biasa saja, tanpa blok chart.
`.trim();

// Aturan file (docx/pptx/xlsx) — sama seperti chart, selalu ditambahkan
// otomatis supaya tidak perlu ditulis manual di instructions.md.
const FILE_INSTRUCTIONS = `
Kalau pengguna minta dibuatkan file Word, PowerPoint, atau Excel (docx/pptx/xlsx),
dan kamu punya cukup informasi untuk mengisinya, balas dengan teks penjelasan
singkat + SATU blok berikut (pilih skema sesuai jenis file):

Word (type: "docx"):
\`\`\`file
{
  "type": "docx",
  "filename": "nama-file.docx",
  "title": "Judul dokumen",
  "sections": [
    { "heading": "Judul bagian", "level": 1, "paragraphs": ["teks paragraf"] }
  ]
}
\`\`\`

PowerPoint (type: "pptx"):
\`\`\`file
{
  "type": "pptx",
  "filename": "nama-file.pptx",
  "slides": [
    { "title": "Judul slide", "bullets": ["poin 1", "poin 2"] }
  ]
}
\`\`\`

Excel (type: "xlsx"):
\`\`\`file
{
  "type": "xlsx",
  "filename": "nama-file.xlsx",
  "sheets": [
    { "name": "Sheet1", "headers": ["Kolom A", "Kolom B"], "rows": [["a1", "b1"]] }
  ]
}
\`\`\`

Aturan blok file:
- Hanya satu blok file per balasan.
- "type" wajib salah satu dari "docx", "pptx", "xlsx".
- Isi kontennya berdasarkan data referensi/percakapan yang benar-benar ada —
  jangan mengarang isi. Kalau informasinya kurang, tanyakan dulu ke pengguna
  sebagai teks biasa, jangan keluarkan blok file.
`.trim();

let cachedContext = null;

function loadContext() {
  if (cachedContext) return cachedContext;

  const root = process.cwd();

  // 1) Custom instructions (setara "Custom instructions" di Claude Project)
  let instructions = '';
  try {
    instructions = fs.readFileSync(path.join(root, 'instructions.md'), 'utf-8').trim();
  } catch (e) {
    instructions = '';
  }

  // 2) Project knowledge (setara file-file yang kamu upload ke Project)
  let knowledge = '';
  const knowledgeDir = path.join(root, 'knowledge');
  try {
    const files = fs.readdirSync(knowledgeDir);
    for (const file of files) {
      const filePath = path.join(knowledgeDir, file);
      if (fs.statSync(filePath).isFile()) {
        const content = fs.readFileSync(filePath, 'utf-8');
        knowledge += `\n\n--- FILE: ${file} ---\n${content}`;
      }
    }
  } catch (e) {
    // folder knowledge kosong/tidak ada — lanjut tanpa knowledge
  }

  cachedContext = { instructions, knowledge };
  return cachedContext;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: 'GEMINI_API_KEY belum diatur di server. Cek README.' });
    return;
  }

  const { messages } = req.body || {};
  if (!Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: 'Field "messages" kosong atau tidak valid.' });
    return;
  }

  const { instructions, knowledge } = loadContext();

  const systemPrompt = [
    instructions || 'Kamu adalah asisten yang membantu menjawab permintaan pengguna dengan jelas dan ringkas.',
    knowledge ? '\n\nGunakan data referensi berikut sebagai konteks setiap kali relevan dengan pertanyaan pengguna:\n' + knowledge : '',
    '\n\n' + CHART_INSTRUCTIONS,
    '\n\n' + FILE_INSTRUCTIONS
  ].join('');

  // Gemini pakai role "user"/"model" (bukan "user"/"assistant" seperti Claude),
  // dan format contents/parts, bukan messages/content langsung.
  const contents = messages.map(m => ({
    role: m.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: typeof m.content === 'string' ? m.content : String(m.content || '') }]
  }));

  try {
    const geminiRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: systemPrompt }] },
          contents
        })
      }
    );

    const data = await geminiRes.json();

    if (!geminiRes.ok) {
      res.status(geminiRes.status).json({ error: data?.error?.message || 'Terjadi kesalahan dari API Gemini.' });
      return;
    }

    const candidate = data.candidates && data.candidates[0];
    const text = (candidate?.content?.parts || []).map(p => p.text || '').join('');

    if (!text) {
      // Kemungkinan diblokir filter keamanan Gemini, atau habis kuota tapi
      // tetap balas status 200 dengan body kosong — kita tangani manual.
      const reason = candidate?.finishReason ? ` (finishReason: ${candidate.finishReason})` : '';
      res.status(200).json({
        content: [{ type: 'text', text: 'Maaf, tidak ada balasan dari model.' + reason }]
      });
      return;
    }

    // Dinormalisasi ke bentuk yang sama seperti respons Claude supaya
    // index.html tidak perlu diubah sama sekali.
    res.status(200).json({ content: [{ type: 'text', text }] });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
