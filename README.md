# Chatbot Data — versi gratis (Google Gemini)

Chatbot ini meniru cara kerja Claude Project di claude.ai — ada **instruksi
kustom** dan **data referensi (knowledge)** yang sudah dibekukan di backend —
tapi "otak" AI-nya pakai **Google Gemini API**, bukan Claude, supaya bisa
dipakai gratis dalam batas kuota. Siapa pun yang membuka link ini otomatis
chat dengan versi yang sudah kamu bekali data & instruksi itu — mereka tidak
perlu (dan tidak bisa) mengubahnya.

⚠️ **Penting soal "gratis":** Gemini API punya jatah gratis, tapi ada batas
kuota (jumlah request per menit/hari) yang bisa berubah sewaktu-waktu dan
tergantung model yang dipakai. Cek syarat & batas terbaru di
https://ai.google.dev/gemini-api/docs/models dan https://aistudio.google.com
sebelum deploy — jangan cuma percaya catatan di README ini karena ini bisa
kedaluwarsa.

⚠️ **Penting soal "Project":** ini bukan sinkronisasi otomatis dari Project
kamu di claude.ai — tidak ada API resmi untuk itu sejauh yang saya tahu (cek
docs.claude.com kalau mau pastikan status terbaru). Kamu perlu **menyalin
manual** instruksi & datanya sekali di awal, dan karena mesinnya sekarang
Gemini (bukan Claude), gaya jawabannya bisa sedikit berbeda dari yang biasa
kamu lihat di claude.ai.

## Struktur

```
konsol-data-claude/
├── index.html          ← tampilan chat (visitor cuma lihat ini)
├── instructions.md      ← salin "Custom instructions" dari Project kamu ke sini
├── knowledge/            ← taruh file data mentah kamu di sini (.txt/.md/.csv/.json)
├── api/chat.js           ← backend: baca instructions.md + knowledge/, panggil Gemini API
├── vercel.json
├── package.json
└── README.md
```

## Langkah 1 — Pindahkan isi Project kamu

1. Buka claude.ai → Project kamu → Settings → **Custom instructions**.
   Salin teksnya, tempel ke `instructions.md` (ganti isi placeholder-nya).
2. Buka tab **Project knowledge** di Project yang sama. Download/salin
   file-file yang sudah kamu upload di sana, taruh di folder `knowledge/`
   (hapus dulu `contoh-data.txt`).
3. Ini mendekati (bukan 100% identik) — Project di claude.ai kadang
   memakai retrieval pintar untuk data besar, sedangkan di sini semua isi
   `knowledge/` langsung dimasukkan penuh ke setiap percakapan. Aman untuk
   data yang tidak terlalu besar (kira-kira sampai puluhan ribu kata). Kalau
   datamu jauh lebih besar dari itu, bilang saja — perlu pendekatan berbeda
   (pencarian/retrieval), bukan sekadar tempel semua.

## Langkah 2 — Buat API key (gratis)

1. Buka **aistudio.google.com**, login pakai akun Google kamu.
2. Cari menu **Get API key** → **Create API key**. Salin key-nya.
3. Cek halaman **kuota/limits** di sana untuk lihat batas gratis yang
   berlaku saat kamu daftar — ini bisa berubah, jadi jangan asumsikan
   angka dari README lama manapun.

## Langkah 3 — Deploy ke Vercel

```
npm install -g vercel
cd konsol-data-claude
vercel
```

Lalu set API key sebagai environment variable:

```
vercel env add GEMINI_API_KEY
```

Tempel key-mu, pilih environment "Production". Deploy ulang ke production:

```
vercel --prod
```

Kamu akan dapat URL, misalnya `https://chatbot-data.vercel.app`. Itu link
yang bisa dibagikan — semua orang yang buka akan chat dengan versi yang
sudah tahu instruksi & data kamu.

## Fitur grafik/chart

Chatbot ini sekarang bisa menampilkan grafik (bar, line, atau pie) langsung
di dalam bubble percakapan, digambar dengan Chart.js di sisi browser.

Cara kerjanya:

1. `api/chat.js` selalu menambahkan aturan format ke system prompt (terpisah
   dari `instructions.md`, jadi tidak perlu ditulis manual): kalau pengguna
   minta grafik dan Claude punya data yang cukup, Claude membalas dengan
   teks penjelasan + satu blok berformat:
   ````
   ```chart
   { "type": "bar", "title": "...", "labels": [...], "series": [{"name":"...","values":[...]}] }
   ```
   ````
2. `index.html` mendeteksi blok itu di balasan, memisahkannya dari teks
   biasa, lalu menggambarnya sebagai grafik memakai Chart.js (dimuat dari
   CDN jsdelivr — otomatis, tidak perlu instalasi tambahan).
3. Kalau Claude tidak punya data yang cukup akurat, dia diinstruksikan untuk
   tidak mengarang angka dan cukup menjawab dengan teks biasa (tanpa grafik).

Ini tetap "nyambung" dengan data & gaya kamu karena aturan ini digabung
dengan `instructions.md` + isi `knowledge/` yang sama — akurasi grafik
bergantung sepenuhnya pada seberapa lengkap data yang kamu taruh di
`knowledge/`.

**Batasan**: `type` yang didukung cuma `bar`, `line`, `pie`. Kalau butuh
jenis chart lain atau layout dashboard yang lebih kompleks, bisa
dikembangkan lebih lanjut — bilang aja.

## Fitur file (Word / PowerPoint / Excel)

Chatbot ini juga bisa membuatkan file `.docx`, `.pptx`, atau `.xlsx` yang bisa
langsung diunduh dari dalam chat — semuanya dirakit di **browser** (bukan di
server), pakai library `docx`, `pptxgenjs`, dan `exceljs` yang dimuat dari CDN.

Cara kerjanya sama seperti fitur chart:

1. `api/chat.js` menambahkan aturan format ke system prompt (otomatis, tidak
   perlu ditulis manual): kalau pengguna minta file dan Claude punya cukup
   data, Claude membalas dengan teks + satu blok:
   ````
   ```file
   { "type": "docx" | "pptx" | "xlsx", "filename": "...", ... }
   ```
   ````
2. `index.html` mendeteksi blok itu, merakit file sungguhan di browser
   pengguna, lalu menampilkan kartu "Unduh [nama file]" di bubble chat.
3. Kalau datanya tidak cukup, Claude diinstruksikan untuk bertanya balik,
   bukan mengarang isi file.

**Batasan yang perlu kamu tahu:**
- Ini generator ringan — cocok untuk dokumen/slide/spreadsheet sederhana
  (judul, paragraf, bullet, tabel data). Bukan pengganti fitur pembuatan
  dokumen kelas profesional (styling kompleks, gambar, chart native di
  pptx/xlsx, dsb) — kalau butuh itu, perlu pengembangan lebih lanjut.
- Library dimuat dari CDN (jsdelivr) memakai major-version pin (`docx@8`,
  `pptxgenjs@3`, `exceljs@4`). Kalau suatu saat linknya berubah/berhenti
  jalan, cek halaman npm masing-masing library untuk link CDN & versi
  terbaru.
- Karena generate-nya di browser, ukuran file yang sangat besar bisa lambat
  di HP dengan spek rendah — untuk data mentahmu yang besar, pertimbangkan
  Claude cuma ambil ringkasannya saja untuk isi file.

## Fitur suara (voice chat)

Chatbot bisa dipakai suara — input & baca balasan — pakai **Web Speech API**
bawaan browser (gratis, tanpa API key tambahan, tanpa perlu server).

- **Tombol mic 🎤** (di sebelah kotak teks): tekan → bicara → otomatis
  ditranskrip jadi teks dan langsung terkirim.
- **Tombol speaker 🔇/🔊** (di header): toggle baca-otomatis. Kalau aktif,
  setiap balasan asisten langsung dibacakan.
- Tiap bubble balasan juga punya tombol 🔊 kecil untuk memutar ulang
  balasan itu kapan saja.

**Batasan yang perlu kamu tahu:**
- Input suara (mic) paling stabil di **Chrome/Edge** (desktop & Android) dan
  **Safari iOS 14.5+**. Firefox dukungannya minim/tidak ada — di browser yang
  tidak didukung, tombol mic otomatis nonaktif (fitur teks tetap normal).
- Butuh koneksi **HTTPS** (otomatis terpenuhi kalau deploy lewat Vercel) dan
  izin akses mikrofon dari browser pengunjung.
- Bahasa diset ke Indonesia (`id-ID`); akurasi transkripsi tergantung mesin
  speech recognition bawaan browser masing-masing pengunjung, bukan sesuatu
  yang bisa kita kontrol dari kode ini.

## Tampilan desktop vs mobile

Satu file `index.html` yang sama, tapi tampilannya otomatis berubah sesuai
lebar layar — sekarang dua-duanya terinspirasi WhatsApp:

- **Desktop** (>700px): ala **WhatsApp Desktop** — ada sidebar kiri (nama
  app + satu "kontak" mewakili chatbot ini), header putih bersih, wallpaper
  krem di area chat, bubble hijau (kamu) / putih (asisten).
- **Mobile** (≤700px): ala **aplikasi WhatsApp** — sidebar disembunyikan,
  fullscreen, header gradasi hijau tua (khas status bar WhatsApp mobile),
  bubble sama seperti desktop.

⚠️ Sidebar di versi desktop itu **murni dekoratif** — cuma menampilkan satu
"kontak" tetap (chatbot ini sendiri) untuk memperkuat kesan aplikasi chat.
Tidak ada fitur multi-percakapan/multi-kontak sungguhan di baliknya; kotak
pencarian juga belum fungsional (belum ada apa pun untuk dicari).

Grafik (chart) otomatis ikut warna tema (hijau `#25D366` sebagai warna
utama) — diambil dari variabel warna saat digambar, bukan warna tetap.

Tidak perlu build/deploy terpisah — satu link yang sama otomatis tampil
sesuai perangkat pengunjung.

⚠️ Palet & tata letak di sini **terinspirasi** WhatsApp (familiar buat
banyak orang) — bukan reproduksi identik/resmi dari produk tersebut. Kalau
mau warna/branding sendiri, tinggal ganti nilai variabel warna (`--bg`,
`--accent-ai`, dst.) di bagian atas `<style>`, atau ganti teks di
`.sidebar-header`/`.chat-list-item`.

## Siapa yang bisa mengubah data?

Hanya kamu. `api/chat.js` cuma **membaca** `instructions.md` dan isi
`knowledge/` (`fs.readFileSync`) — tidak ada endpoint di aplikasi ini yang
bisa dipakai pengunjung untuk menulis/mengubah file tersebut, sekarang atau
nanti. Satu-satunya jalan update adalah lewat kode sumbernya sendiri.

Cara paling praktis untuk update rutin **tanpa perlu install Vercel CLI tiap
kali**: hubungkan folder project ini ke sebuah repo **GitHub**, aktifkan
auto-deploy di Vercel (Project Settings → Git), lalu untuk update data kamu
tinggal edit file `instructions.md` / isi `knowledge/` langsung lewat web
GitHub (bisa dari HP) → commit → Vercel otomatis redeploy dalam
hitungan detik. Akses dikontrol penuh oleh login GitHub kamu — gratis, tanpa
perlu bikin sistem admin/password sendiri.

## Kustomisasi

- **Ganti model**: edit `GEMINI_MODEL` di bagian atas `api/chat.js` (cek
  daftar model terbaru & mana yang termasuk jatah gratis di
  https://ai.google.dev/gemini-api/docs/models).
- **Update data**: edit file di `knowledge/`, lalu `vercel --prod` lagi
  untuk redeploy (perubahan file tidak otomatis live, harus deploy ulang).
- **Batasi siapa yang bisa pakai**: kalau link ini dibagikan luas dan kamu
  khawatir biaya membengkak, tambahkan password sederhana di `api/chat.js`
  atau rate limiting — bilang aja kalau mau saya tambahkan.

## Troubleshooting

- Kalau chatbot menjawab seolah `instructions.md`/`knowledge/` kosong,
  cek log di dashboard Vercel — biasanya karena file tidak ikut ter-bundle.
  `vercel.json` di sini sudah mengatur `includeFiles`, tapi kalau tetap
  gagal, pindahkan `instructions.md` dan folder `knowledge/` ke dalam
  folder `api/` sebagai fallback.
