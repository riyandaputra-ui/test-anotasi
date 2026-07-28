<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Anotator BIO GANDA Teks Pernyataan Bahasa Indonesia</title>
    
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#f0fdf4',
                            100: '#dcfce7',
                            500: '#22c55e',
                            600: '#16a34a',
                            700: '#15803d',
                            800: '#166534',
                            900: '#14532d',
                        },
                        slate: {
                            850: '#111827',
                            950: '#030712'
                        }
                    }
                }
            }
        }
    </script>

    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>

    <!-- SheetJS for Excel Import/Export -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <!-- React 18 & Babel for JSX -->
    <script src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
    <script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
    <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>

    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
        
        body {
            font-family: 'Inter', sans-serif;
        }

        .font-mono {
            font-family: 'JetBrains Mono', monospace;
        }

        /* Custom Scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        html.dark ::-webkit-scrollbar-track {
            background: #0f172a;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        html.dark ::-webkit-scrollbar-thumb {
            background: #334155;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
        html.dark ::-webkit-scrollbar-thumb:hover {
            background: #475569;
        }

        .glass-panel {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(16px);
            border: 1px solid rgba(0, 0, 0, 0.08);
        }
        html.dark .glass-panel {
            background: rgba(15, 23, 42, 0.75);
            border: 1px solid rgba(255, 255, 255, 0.08);
        }

        .glass-card {
            background: rgba(248, 250, 252, 0.7);
            backdrop-filter: blur(8px);
            border: 1px solid rgba(0, 0, 0, 0.05);
        }
        html.dark .glass-card {
            background: rgba(30, 41, 59, 0.5);
            border: 1px solid rgba(255, 255, 255, 0.05);
        }

        .glow-effect {
            box-shadow: 0 0 25px -5px rgba(34, 197, 94, 0.15);
        }
    </style>
</head>
<body class="min-h-screen flex flex-col antialiased selection:bg-brand-500 selection:text-slate-950 bg-slate-50 dark:bg-[#090d16] text-slate-900 dark:text-[#f8fafc] transition-colors duration-300">
    <div id="root"></div>

    <script type="text/babel">
        const { useState, useEffect, useRef, useMemo } = React;

        const SAMPLE_SENTENCES = [
            {
                id: 1,
                title: "Ekonomi Kuartal III",
                text: 'Presiden Jokowi menyatakan bahwa " pertumbuhan ekonomi Kuartal III sangat memuaskan ".'
            },
            {
                id: 2,
                title: "Kebijakan Pajak",
                text: 'Menteri Keuangan menjelaskan kebijakan insentif pajak baru akan merangsang investasi.'
            },
            {
                id: 3,
                title: "Hambatan Birokrasi",
                text: 'Pengamat menilai hambatan birokrasi berpotensi merugikan industri nasional.'
            }
        ];

        const STATEMENT_TAGS = [
            'O',
            'B-WHO', 'I-WHO',
            'B-STAT', 'I-STAT',
            'B-BREL', 'I-BREL',
            'B-FREL', 'I-FREL'
        ];

        const SENTIMENT_TAGS = [
            'O',
            'B-POS', 'I-POS',
            'B-NEG', 'I-NEG',
            'B-NET', 'I-NET'
        ];

        const STATEMENT_COLORS = {
            'WHO': 'bg-purple-100 dark:bg-purple-950/80 text-purple-700 dark:text-purple-300 border-purple-300 dark:border-purple-700/80',
            'STAT': 'bg-blue-100 dark:bg-blue-950/80 text-blue-700 dark:text-blue-300 border-blue-300 dark:border-blue-700/80',
            'BREL': 'bg-amber-100 dark:bg-amber-950/80 text-amber-700 dark:text-amber-300 border-amber-300 dark:border-amber-700/80',
            'FREL': 'bg-cyan-100 dark:bg-cyan-950/80 text-cyan-700 dark:text-cyan-300 border-cyan-300 dark:border-cyan-700/80',
            'O': 'bg-slate-200 dark:bg-slate-900/60 text-slate-600 dark:text-slate-400 border-slate-300 dark:border-slate-800'
        };

        const SENTIMENT_COLORS = {
            'POS': 'bg-emerald-100 dark:bg-emerald-950/80 text-emerald-700 dark:text-emerald-300 border-emerald-300 dark:border-emerald-700/80',
            'NEG': 'bg-rose-100 dark:bg-rose-950/80 text-rose-700 dark:text-rose-300 border-rose-300 dark:border-rose-700/80',
            'NET': 'bg-slate-200 dark:bg-slate-800/80 text-slate-700 dark:text-slate-300 border-slate-300 dark:border-slate-700',
            'O': 'bg-slate-100 dark:bg-slate-900/60 text-slate-500 dark:text-slate-400 border-slate-200 dark:border-slate-800'
        };

        const DEFAULT_GUIDE_TEXT = `PANDUAN SKEMA BIO GANDA:
1. Struktur Pernyataan (statement_tag):
   - B-WHO / I-WHO: Pelaku / Subjek yang memberikan pernyataan (Jika terdapat jabatan/gelar jika terdapat pada teks maka WAJIB dilabel dalam satu span B-WHO+I-WHO termasuk simbol.
   - B-STAT / I-STAT: Isi utama pernyataan / konten kutipan. Seluruh pernyataan WAJIB digabungkan dalam satu span B-STAT+I-STAT termasuk symbol.
   - B-BREL / I-BREL = Menandai verba/kata kerja pelapor yang menghubungkan subjek dengan isi pernyataan secara langsung.
   - B-FREL / I-FREL = Menandai kata atau frasa penunjuk hubungan tidak langsung atau keterangan konteks kutipan.
   - O = Bukan entitas struktur (TERMASUK SEMUA TANDA KUTIP / TANDA PETIK: ", ', “ ” ‘ ’)
   - PENTING (ATURAN KATA PENGANTAR): Kata depan/preposisi pengantar jabatan seperti "Sebagai", "Selaku", "Mewakili", "Atas nama" WAJIB dilabeli O (TIDAK boleh masuk WHO maupun FREL).
   - PENTING (ATURAN KATA HUBUNG): Kata hubung (konjungsi) penghubung klausa WAJIB dilabeli O.
   - Keterangan yang menggambarkan waktu, hari, tanggal, kegiatan, sarana, atau wadah tempat pernyataan disampaikan di luar STAT WAJIB dilabeli O.

2. Sentimen (sentiment_tag):
   - B-POS / I-POS: Frasa berpolaritas positif.
   - B-NEG / I-NEG: Frasa berpolaritas negatif.
   - B-NET / I-NET: Netral / faktual / deskriptif.
   - O: Di luar frasa sentimen.`;

        const DEFAULT_SYSTEM_PROMPT = `Kamu adalah sistem anotasi bahasa Indonesia berbasis skema BIO GANDA. Tugasmu:
memberi DUA label per token — (1) label STRUKTUR PERNYATAAN dan (2) label SENTIMEN.

=== SKEMA LABEL STRUKTUR (statement_tag) ===
B-WHO / I-WHO = Pelaku / Subjek yang memberikan pernyataan (Jika terdapat jabatan/gelar jika terdapat pada teks maka WAJIB dilabel dalam satu span B-WHO+I-WHO termasuk simbol.
B-STAT / I-STAT = Isi utama pernyataan / konten kutipan. Seluruh pernyataan WAJIB digabungkan dalam satu span B-STAT+I-STAT termasuk symbol
B-BREL / I-BREL = Menandai verba/kata kerja pelapor yang menghubungkan subjek dengan isi pernyataan secara langsung.
B-FREL / I-FREL = Menandai kata atau frasa penunjuk hubungan tidak langsung atau keterangan konteks kutipan.
O = Bukan entitas struktur (TERMASUK SEMUA TANDA KUTIP / TANDA PETIK: ", ', “ ” ‘ ’)
PENTING (ATURAN KATA PENGANTAR): Kata depan/preposisi pengantar jabatan seperti "Sebagai", "Selaku", "Mewakili", "Atas nama" WAJIB dilabeli O (TIDAK boleh masuk WHO maupun FREL).
PENTING (ATURAN KATA HUBUNG): Kata hubung (konjungsi) penghubung klausa WAJIB dilabeli O.
Keterangan yang menggambarkan waktu, hari, tanggal, kegiatan, sarana, atau wadah tempat pernyataan disampaikan di luar STAT WAJIB dilabeli O.

=== SKEMA LABEL SENTIMEN (sentiment_tag) ===
- B-POS / I-POS = frasa positif
- B-NEG / I-NEG = frasa negatif
- B-NET / I-NET = frasa netral/deskriptif/faktual
- O = di luar frasa sentimen

Aturan struktur:
- Setiap tanda kutip / petik ( baik pembuka maupun penutup seperti ", ', “ ” ‘ ’ ) dari statement ATAU di dalam teks SELALU memiliki label O
- B-STAT dimulai dari kata pertama isi pernyataan setelah tanda kutip, dilanjutkan I-STAT untuk seluruh konten pernyataan termasuk koma di dalam kalimat STAT
- Setelah koma pemisah proposisi (yang memisahkan STAT dari verba pelapor), tag kembali ke O atau B-BREL
- Verba pelapor (ujar, kata, ungkap, menyatakan, mengatakan) = B-BREL
- Nama orang pembuat pernyataan = B-WHO / I-WHO
- Nama tempat, tanggal, waktu setelah verba pelapor = O
- PENTING (ATURAN KATA PENGANTAR): Kata depan/preposisi pengantar jabatan seperti "Sebagai", "Selaku", "Mewakili", "Atas nama" WAJIB dilabeli O (TIDAK boleh masuk WHO maupun FREL).

Aturan sentimen (ikuti LANGKAH 1-7 di bawah secara berurutan):

LANGKAH 1 - TOKEN YANG SELALU O:
- Nama orang, tempat, organisasi, jabatan, tanggal/hari
- Tanda baca & tanda kutip (" ' “ ” ‘ ’)
- Kata kerja pelapor: mengatakan, menjelaskan, menyampaikan, ujar
- Konjungsi: dan, tapi, namun, serta, dari, sebagai, meski, meskipun, walau, kendati

LANGKAH 2 - FRASA SENTIMEN EKSPLISIT:
- Kata evaluatif (bagus, buruk, hebat, gagal, sukses, kecewa, senang) → POS/NEG sesuai makna
- Negasi (tidak, tak, bukan, kurang, belum, tanpa) MEMBALIK polaritas kata sesudahnya
- Penguat (sangat, banget, sekali) ikut span, ikuti polaritas kata inti
- Frasa negasi implisit (jauh dari, bebas dari, terhindar dari, terbebas dari, lolos dari) + kata konotasi negatif → SELURUH span = POS

PERINGATAN: Kata "Kebebasan", "Hak", "Independensi", "Transparansi", "Keadilan", "Demokrasi", "Kesejahteraan", "Keamanan" — polaritasnya TIDAK ditentukan dari kata itu sendiri, tapi dari KONTEKS (apakah sedang terancam/dipersoalkan → NEG; atau diperoleh/ditingkatkan → POS).

LANGKAH 3 - VERBA/FRASA DAMPAK (klasifikasikan ke Jenis A/B/C/D dulu):
- Jenis A (dampak sudah terjadi): memungkinkan, mendukung, meningkatkan, membantu, melindungi → POS; menghambat, merugikan, menyerang → NEG
- Jenis B (ajakan/harapan): mendorong, mengajak, meminta, mengimbau, diharapkan, perlu, harus → NET
- Jenis C (sorotan + objek peluang/potensi positif): soroti/singgung + objek manfaat/peluang → POS
- Jenis D (sorotan + objek ancaman/risiko): soroti/singgung + objek dikhawatirkan/terancam → NEG

LANGKAH 4 - KATA BERKONOTASI YANG SEBENARNYA DESKRIPSI FAKTA:
- beban, insiden, protes, krisis, gejolak, tantangan → HANYA NET jika kalimat sekadar menyebut jenis/kategori tanpa akibat eksplisit; → NEG jika ada akibat/konsekuensi eksplisit

LANGKAH 5 - SARKASME: label sesuai makna sebenarnya, bukan harfiah

LANGKAH 6 - PEMISAH FRASA:
- Koma tanpa konjungsi = pemisah dua proposisi berbeda
- Klausa konsesif (meski/walau/kendati): anotasi token tetap independen di level token; untuk sentiment_agregat HANYA hitung klausa utama
- Negasi kausal (tidak membuat, tidak menyebabkan, bukan berarti, tidak serta-merta): premis = tidak dihitung untuk agregat; klausa akibat (dengan negasi kausal yang membalik polaritas) = yang menentukan agregat

LANGKAH 7 - HITUNG SENTIMENT_AGREGAT:
| Kondisi | Hasil |
|---|---|
| n_POS > 0 dan n_NEG = 0 | POS |
| n_NEG > 0 dan n_POS = 0 | NEG |
| n_POS > n_NEG | POS |
| n_NEG > n_POS | NEG |
| n_POS = n_NEG | NET |
| hanya NET atau tidak ada entitas | NET |

=== TOKENISASI ===
- Pisahkan semua kata dengan spasi
- Tanda baca dan tanda kutip (termasuk ", ', “ ” ‘ ’) yang menempel pada kata HARUS dipisahkan menjadi token tersendiri dengan label "O"
- Contoh: "keuangan," → ["\"", "keuangan", ",", "\""]
- SEMUA token dari teks HARUS ada di output, tidak boleh ada yang hilang

=== FORMAT OUTPUT ===
Keluarkan HANYA JSON valid, tanpa markdown code fence, tanpa penjelasan:
{
  "sentences": [
    {
      "sentence_id": 1,
      "text": "teks asli",
      "tokens": [
        {"word": "kata", "statement_tag": "B-STAT", "sentiment_tag": "O"}
      ],
      "statement_entities": [
        {"text": "span teks", "label": "STAT"}
      ],
      "sentiment_entities": [
        {"text": "span teks", "label": "POS"}
      ],
      "sentiment_agregat": "POS"
    }
  ]
}`.trim();

        function clientTokenize(text) {
            if (!text) return [];
            const regex = /([\w-]+|[“„”"«»‘’'()\[\]{},.!?])/g;
            const matches = text.match(regex);
            return matches || text.split(/\s+/);
        }

        function extractEntitiesFromTokens(tokens, tagField) {
            const entities = [];
            let currentEntity = null;

            tokens.forEach((tok) => {
                const tag = tok[tagField] || 'O';
                if (tag.startsWith('B-')) {
                    if (currentEntity) entities.push(currentEntity);
                    const label = tag.split('-')[1];
                    currentEntity = { text: tok.word, label: label };
                } else if (tag.startsWith('I-') && currentEntity) {
                    currentEntity.text += ' ' + tok.word;
                } else {
                    if (currentEntity) {
                        entities.push(currentEntity);
                        currentEntity = null;
                    }
                }
            });
            if (currentEntity) entities.push(currentEntity);
            return entities;
        }

        function computeAggregateSentiment(sentimentEntities) {
            if (!sentimentEntities || sentimentEntities.length === 0) return 'NET';
            let posCount = 0;
            let negCount = 0;
            sentimentEntities.forEach(e => {
                if (e.label === 'POS') posCount++;
                if (e.label === 'NEG') negCount++;
            });
            if (posCount > negCount) return 'POS';
            if (negCount > posCount) return 'NEG';
            if (posCount > 0 && posCount === negCount) return 'POS';
            return 'NET';
        }

        function Toast({ toast, onClose }) {
            if (!toast) return null;
            const bgClass = toast.type === 'error' ? 'bg-rose-100 dark:bg-rose-950 border-rose-300 dark:border-rose-600 text-rose-800 dark:text-rose-200' :
                            toast.type === 'info' ? 'bg-blue-100 dark:bg-blue-950 border-blue-300 dark:border-blue-600 text-blue-800 dark:text-blue-200' :
                            'bg-emerald-100 dark:bg-emerald-950 border-emerald-300 dark:border-emerald-600 text-emerald-800 dark:text-emerald-200';
            return (
                <div className={`fixed bottom-6 right-6 z-50 p-4 rounded-2xl border shadow-2xl flex items-center space-x-3 max-w-md ${bgClass} animate-bounce bg-white dark:bg-transparent`}>
                    <span className="text-sm font-medium flex-1">{toast.message}</span>
                    <button onClick={onClose} className="opacity-70 hover:opacity-100 font-bold text-xs">Tutup</button>
                </div>
            );
        }

        function App() {
            const [activeTab, setActiveTab] = useState('single');
            const [systemPrompt, setSystemPrompt] = useState(() => {
                return localStorage.getItem('custom_system_prompt') || DEFAULT_SYSTEM_PROMPT;
            });
            const [apiKey, setApiKey] = useState(() => {
                return localStorage.getItem('gemini_api_key') || '';
            });
            const [showApiModal, setShowApiModal] = useState(false);
            const [toast, setToast] = useState(null);

            // Tema Light/Dark
            const [isDarkMode, setIsDarkMode] = useState(() => {
                if (typeof window !== 'undefined') {
                    const saved = localStorage.getItem('theme');
                    if (saved) return saved === 'dark';
                    return true;
                }
                return true;
            });

            // Single sentence state
            const [singleText, setSingleText] = useState(SAMPLE_SENTENCES[0].text);
            const [loadingSingle, setLoadingSingle] = useState(false);
            const [annotationResult, setAnnotationResult] = useState(null);
            const [viewMode, setViewMode] = useState('visual');
            const [processingMode, setProcessingMode] = useState('all');

            // Guide text state & edit toggle
            const [guideText, setGuideText] = useState(() => {
                return localStorage.getItem('custom_guide_text') || DEFAULT_GUIDE_TEXT;
            });
            const [isEditingGuide, setIsEditingGuide] = useState(false);

            // Batch processing state
            const [batchInputText, setBatchInputText] = useState('');
            const [batchResults, setBatchResults] = useState([]);
            const [isProcessingBatch, setIsProcessingBatch] = useState(false);
            const [batchProgress, setBatchProgress] = useState({ current: 0, total: 0 });
            const [editingBatchIndex, setEditingBatchIndex] = useState(null);

            useEffect(() => {
                if (window.lucide) {
                    window.lucide.createIcons();
                }
            });

            // Effect toggle Theme
            useEffect(() => {
                if (isDarkMode) {
                    document.documentElement.classList.add('dark');
                    localStorage.setItem('theme', 'dark');
                } else {
                    document.documentElement.classList.remove('dark');
                    localStorage.setItem('theme', 'light');
                }
            }, [isDarkMode]);

            const showToast = (msg, type = 'success') => {
                setToast({ message: msg, type });
                setTimeout(() => setToast(null), 4500);
            };

            const callGeminiLLM = async (text, sentenceId = 1, mode = 'all') => {
                const userPrompt = text;
                const key = apiKey.trim();
                
                const activeKey = key || ''; 
                const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent${activeKey ? `?key=${activeKey}` : ''}`;

                let dynamicInstruction = systemPrompt;
                if (mode === 'structure') {
                    dynamicInstruction += "\n\nINSTRUKSI KHUSUS PENGGUNA: KELUARKAN HANYA LABEL STRUKTUR PERNYATAAN (statement_tag). Untuk sentiment_tag SELALU isi dengan 'O', sentiment_entities harus kosong ([]), dan sentiment_agregat harus 'NET'.";
                } else if (mode === 'sentiment') {
                    dynamicInstruction += "\n\nINSTRUKSI KHUSUS PENGGUNA: KELUARKAN HANYA LABEL SENTIMEN (sentiment_tag). Untuk statement_tag SELALU isi dengan 'O', dan statement_entities harus kosong ([]).";
                }

                const payload = {
                    contents: [{ parts: [{ text: userPrompt }] }],
                    systemInstruction: { parts: [{ text: dynamicInstruction }] }
                };

                let response;
                let retries = 2;
                let delay = 1000;

                for (let i = 0; i < retries; i++) {
                    try {
                        const headers = { 'Content-Type': 'application/json' };
                        response = await fetch(endpoint, {
                            method: 'POST',
                            headers: headers,
                            body: JSON.stringify(payload)
                        });
                        if (response.ok) break;
                        
                        if (response.status === 429 || response.status === 403 || response.status === 401) {
                            const errBody = await response.text();
                            throw new Error(`API Error (${response.status}): Periksa API Key atau kuota Gemini Anda.`);
                        }
                    } catch (e) {
                        if (e.message && e.message.startsWith('API Error')) {
                            throw e;
                        }
                        if (i === retries - 1) throw e;
                    }
                    await new Promise(r => setTimeout(r, delay));
                    delay *= 2;
                }

                if (!response || !response.ok) {
                    throw new Error("Koneksi ke Gemini AI gagal. Periksa jaringan atau kredensial API Anda.");
                }

                const data = await response.json();
                const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text;
                if (!rawText) throw new Error("Format respon LLM tidak berisi teks valid.");

                let cleaned = rawText.trim();
                if (cleaned.startsWith('```')) {
                    cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '');
                }

                const parsed = JSON.parse(cleaned);
                let sentenceData = parsed.sentences ? parsed.sentences[0] : parsed;
                sentenceData.sentence_id = sentenceId;
                
                if (mode === 'structure') {
                    if (sentenceData.tokens) sentenceData.tokens.forEach(t => t.sentiment_tag = 'O');
                    sentenceData.sentiment_entities = [];
                    sentenceData.sentiment_agregat = 'NET';
                } else if (mode === 'sentiment') {
                    if (sentenceData.tokens) sentenceData.tokens.forEach(t => t.statement_tag = 'O');
                    sentenceData.statement_entities = [];
                }

                if (!sentenceData.statement_entities) {
                    sentenceData.statement_entities = extractEntitiesFromTokens(sentenceData.tokens, 'statement_tag');
                }
                if (!sentenceData.sentiment_entities) {
                    sentenceData.sentiment_entities = extractEntitiesFromTokens(sentenceData.tokens, 'sentiment_tag');
                }
                if (!sentenceData.sentiment_agregat) {
                    sentenceData.sentiment_agregat = computeAggregateSentiment(sentenceData.sentiment_entities);
                }

                return sentenceData;
            };

            const handleAnnotateSingle = async () => {
                if (!singleText.trim()) {
                    showToast("Masukkan kalimat terlebih dahulu.", "error");
                    return;
                }

                setLoadingSingle(true);

                try {
                    const res = await callGeminiLLM(singleText.trim(), 1, processingMode);
                    setAnnotationResult(res);
                } catch (err) {
                    console.error(err);
                    showToast(err.message || "Gagal melakukan anotasi.", "error");
                } finally {
                    setLoadingSingle(false);
                }
            };

            const handleTokenTagChange = (tokenIndex, field, value) => {
                if (!annotationResult) return;
                const updatedTokens = [...annotationResult.tokens];
                updatedTokens[tokenIndex] = {
                    ...updatedTokens[tokenIndex],
                    [field]: value
                };

                const statEnts = extractEntitiesFromTokens(updatedTokens, 'statement_tag');
                const sentEnts = extractEntitiesFromTokens(updatedTokens, 'sentiment_tag');
                const agregat = computeAggregateSentiment(sentEnts);

                setAnnotationResult({
                    ...annotationResult,
                    tokens: updatedTokens,
                    statement_entities: statEnts,
                    sentiment_entities: sentEnts,
                    sentiment_agregat: agregat
                });
            };

            const handleBatchTokenTagChange = (sentenceIndex, tokenIndex, field, value) => {
                const updatedResults = [...batchResults];
                const targetSentence = { ...updatedResults[sentenceIndex] };
                const updatedTokens = [...targetSentence.tokens];
                
                updatedTokens[tokenIndex] = {
                    ...updatedTokens[tokenIndex],
                    [field]: value
                };

                targetSentence.tokens = updatedTokens;
                targetSentence.statement_entities = extractEntitiesFromTokens(updatedTokens, 'statement_tag');
                targetSentence.sentiment_entities = extractEntitiesFromTokens(updatedTokens, 'sentiment_tag');
                targetSentence.sentiment_agregat = computeAggregateSentiment(targetSentence.sentiment_entities);

                updatedResults[sentenceIndex] = targetSentence;
                setBatchResults(updatedResults);
            };

            const handleBatchProcess = async () => {
                const lines = batchInputText.split('\n').map(l => l.trim()).filter(Boolean);
                if (lines.length === 0) {
                    showToast("Masukkan setidaknya satu kalimat untuk batch processing.", "error");
                    return;
                }

                setIsProcessingBatch(true);
                setBatchProgress({ current: 0, total: lines.length });
                setBatchResults([]);

                for (let i = 0; i < lines.length; i++) {
                    const text = lines[i];
                    let currentResult;
                    try {
                        currentResult = await callGeminiLLM(text, i + 1, processingMode);
                    } catch (e) {
                        currentResult = {
                            sentence_id: i + 1,
                            text: text,
                            tokens: clientTokenize(text).map(w => ({ word: w, statement_tag: 'O', sentiment_tag: 'O' })),
                            statement_entities: [],
                            sentiment_entities: [],
                            sentiment_agregat: 'NET',
                            error: e.message
                        };
                        showToast(`Baris ${i + 1} gagal diproses.`, "error");
                    }
                    
                    setBatchResults(prev => [...prev, currentResult]);
                    setBatchProgress({ current: i + 1, total: lines.length });
                }

                setIsProcessingBatch(false);
                showToast(`Selesai memproses ${lines.length} kalimat.`, "success");
            };

            const handleFileUpload = (e) => {
                const file = e.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = (evt) => {
                    try {
                        const bstr = evt.target.result;
                        const wb = XLSX.read(bstr, { type: 'binary' });
                        const sheetName = wb.SheetNames.includes("FORMAT B - SENTENCE") ? "FORMAT B - SENTENCE" : wb.SheetNames[0];
                        const ws = wb.Sheets[sheetName];
                        const data = XLSX.utils.sheet_to_json(ws);

                        const extractedTexts = data
                            .map(r => r.text || r.Text || r.kalimat || Object.values(r)[1] || Object.values(r)[0])
                            .filter(Boolean);

                        if (extractedTexts.length > 0) {
                            setBatchInputText(extractedTexts.join('\n'));
                            showToast(`Berhasil mengimpor ${extractedTexts.length} kalimat dari Excel.`, "success");
                        } else {
                            showToast("Tidak dapat menemukan kolom teks dalam file Excel.", "error");
                        }
                    } catch (err) {
                        showToast("Gagal membaca file Excel.", "error");
                    }
                };
                reader.readAsBinaryString(file);
            };

            const exportToExcel = (resultsList, filename = "Anotasi_BIO_Ganda_Output.xlsx") => {
                if (!resultsList || resultsList.length === 0) {
                    showToast("Tidak ada data untuk diunduh.", "error");
                    return;
                }

                const tokenRows = [];
                const sentenceRows = [];

                resultsList.forEach(sent => {
                    const sid = sent.sentence_id;
                    if (sent.tokens) {
                        sent.tokens.forEach((tok, tid) => {
                            tokenRows.push({
                                'Sentence_ID': sid,
                                'Token_ID': tid + 1,
                                'Word': tok.word,
                                'Statement_Tag': tok.statement_tag || 'O',
                                'Sentiment_Tag': tok.sentiment_tag || 'O'
                            });
                        });
                    }

                    sentenceRows.push({
                        'Sentence_ID': sid,
                        'Text': sent.text,
                        'Statement_Entities': JSON.stringify(sent.statement_entities || []),
                        'Sentiment_Entities': JSON.stringify(sent.sentiment_entities || []),
                        'Sentiment_Agregat': sent.sentiment_agregat || 'NET'
                    });
                });

                const wb = XLSX.utils.book_new();
                const wsTokens = XLSX.utils.json_to_sheet(tokenRows);
                const wsSentences = XLSX.utils.json_to_sheet(sentenceRows);

                XLSX.utils.book_append_sheet(wb, wsTokens, "FORMAT A - TOKEN");
                XLSX.utils.book_append_sheet(wb, wsSentences, "FORMAT B - SENTENCE");

                XLSX.writeFile(wb, filename);
                showToast(`File Excel ${filename} berhasil dibuat!`, "success");
            };

            return (
                <div className="flex-1 flex flex-col min-h-screen">
                    <Toast toast={toast} onClose={() => setToast(null)} />

                    {/* Header */}
                    <header className="sticky top-0 z-40 glass-panel border-b border-slate-200 dark:border-slate-800">
                        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
                            <div className="flex items-center space-x-3">
                                <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-brand-600 to-emerald-400 flex items-center justify-center text-white dark:text-slate-950 font-bold shadow-lg shadow-brand-500/20">
                                    <i data-lucide="tags" className="w-5 h-5"></i>
                                </div>
                                <div>
                                    <h1 className="font-bold text-lg text-slate-900 dark:text-slate-100 leading-tight">Anotator BIO GANDA</h1>
                                    <p className="text-xs text-slate-500 dark:text-slate-400">Struktur Pernyataan & Sentimen Bahasa Indonesia</p>
                                </div>
                            </div>

                            {/* Tabs Navigation */}
                            <nav className="hidden md:flex items-center space-x-1 bg-slate-100 dark:bg-slate-900/80 p-1 rounded-xl border border-slate-200 dark:border-slate-800">
                                <button
                                    onClick={() => setActiveTab('single')}
                                    className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center space-x-2 ${
                                        activeTab === 'single' ? 'bg-brand-600 text-white dark:text-slate-950 shadow-md font-semibold' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
                                    }`}
                                >
                                    <i data-lucide="file-text" className="w-4 h-4"></i>
                                    <span>Single Mode</span>
                                </button>
                                <button
                                    onClick={() => setActiveTab('batch')}
                                    className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center space-x-2 ${
                                        activeTab === 'batch' ? 'bg-brand-600 text-white dark:text-slate-950 shadow-md font-semibold' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
                                    }`}
                                >
                                    <i data-lucide="layers" className="w-4 h-4"></i>
                                    <span>Batch Mode</span>
                                </button>
                                <button
                                    onClick={() => setActiveTab('guide')}
                                    className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center space-x-2 ${
                                        activeTab === 'guide' ? 'bg-brand-600 text-white dark:text-slate-950 shadow-md font-semibold' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
                                    }`}
                                >
                                    <i data-lucide="book-open" className="w-4 h-4"></i>
                                    <span>Panduan Rules</span>
                                </button>
                                <button
                                    onClick={() => setActiveTab('prompt')}
                                    className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center space-x-2 ${
                                        activeTab === 'prompt' ? 'bg-brand-600 text-white dark:text-slate-950 shadow-md font-semibold' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
                                    }`}
                                >
                                    <i data-lucide="terminal" className="w-4 h-4"></i>
                                    <span>System Prompt</span>
                                </button>
                            </nav>

                            {/* API Key Status / Settings */}
                            <div className="flex items-center space-x-2 sm:space-x-3">
                                <div className="hidden sm:flex items-center space-x-1.5 px-3 py-1.5 rounded-xl bg-brand-50 dark:bg-brand-950/50 border border-brand-200 dark:border-brand-500/30 text-brand-700 dark:text-brand-300 text-xs font-medium">
                                    <div className="w-2 h-2 rounded-full bg-emerald-500 dark:bg-emerald-400 animate-pulse"></div>
                                    <span>Gemini AI Aktif</span>
                                </div>
                                
                                <button
                                    onClick={() => setIsDarkMode(!isDarkMode)}
                                    className="p-2 rounded-xl border text-xs font-medium flex items-center justify-center transition-all bg-white dark:bg-slate-900 border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800"
                                    title={isDarkMode ? "Switch to Light Mode" : "Switch to Dark Mode"}
                                >
                                    {isDarkMode ? <i data-lucide="sun" className="w-4 h-4"></i> : <i data-lucide="moon" className="w-4 h-4"></i>}
                                </button>

                                <button
                                    onClick={() => setShowApiModal(true)}
                                    className="px-3 py-1.5 rounded-xl border text-xs font-medium flex items-center space-x-2 transition-all bg-white dark:bg-slate-900 border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800"
                                >
                                    <i data-lucide="settings" className="w-3.5 h-3.5"></i>
                                    <span className="hidden sm:inline">{apiKey.trim() ? 'API Aktif' : 'Opsi API'}</span>
                                </button>
                            </div>
                        </div>
                    </header>

                    {/* API Key Modal */}
                    {showApiModal && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 dark:bg-slate-950/80 backdrop-blur-sm p-4">
                            <div className="glass-panel bg-white dark:bg-transparent w-full max-w-md p-6 rounded-3xl space-y-4 shadow-2xl">
                                <div className="flex justify-between items-center">
                                    <h3 className="font-bold text-slate-900 dark:text-slate-100 text-lg flex items-center space-x-2">
                                        <i data-lucide="key" className="w-5 h-5 text-brand-500"></i>
                                        <span>Konfigurasi Gemini API Key</span>
                                    </h3>
                                    <button onClick={() => setShowApiModal(false)} className="text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200">
                                        <i data-lucide="x" className="w-5 h-5"></i>
                                    </button>
                                </div>
                                <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
                                    Jika dikosongkan, aplikasi otomatis memproses menggunakan koneksi standar Gemini Flash AI.
                                </p>
                                <input
                                    type="password"
                                    value={apiKey}
                                    onChange={(e) => setApiKey(e.target.value)}
                                    placeholder="AIzaSy..."
                                    className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl px-4 py-2.5 text-slate-900 dark:text-slate-100 text-sm focus:outline-none focus:border-brand-500 font-mono"
                                />
                                <div className="flex justify-end space-x-3 pt-2">
                                    <button
                                        onClick={() => {
                                            localStorage.setItem('gemini_api_key', apiKey);
                                            setShowApiModal(false);
                                            showToast("API Key disimpan.", "success");
                                        }}
                                        className="px-4 py-2 rounded-xl bg-brand-600 hover:bg-brand-500 text-white dark:text-slate-950 font-bold text-xs"
                                    >
                                        Simpan & Tutup
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* Main Body */}
                    <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 space-y-6">

                        {/* SINGLE MODE */}
                        {activeTab === 'single' && (
                            <div className="space-y-6">
                                <div className="glass-card p-4 rounded-2xl flex flex-wrap items-center justify-between gap-3">
                                    <div className="flex items-center space-x-2 text-xs text-slate-600 dark:text-slate-400">
                                        <i data-lucide="sparkles" className="w-4 h-4 text-brand-500"></i>
                                        <span>Contoh Kalimat Cepat:</span>
                                    </div>
                                    <div className="flex flex-wrap gap-2">
                                        {SAMPLE_SENTENCES.map((s) => (
                                            <button
                                                key={s.id}
                                                onClick={() => {
                                                    setSingleText(s.text);
                                                    setAnnotationResult(null);
                                                }}
                                                className="px-3 py-1 rounded-lg bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 transition-all"
                                            >
                                                {s.title}
                                            </button>
                                        ))}
                                    </div>
                                </div>

                                <div className="glass-panel p-6 rounded-3xl space-y-4 glow-effect">
                                    <div className="flex justify-between items-center">
                                        <label className="text-sm font-semibold text-slate-800 dark:text-slate-200 flex items-center space-x-2">
                                            <i data-lucide="edit-3" className="w-4 h-4 text-brand-500"></i>
                                            <span>Kalimat Pernyataan Bahasa Indonesia:</span>
                                        </label>
                                        <span className="text-xs text-slate-500 dark:text-slate-400 font-mono">{singleText.length} karakter</span>
                                    </div>

                                    <textarea
                                        rows={3}
                                        value={singleText}
                                        onChange={(e) => setSingleText(e.target.value)}
                                        placeholder="Masukkan kalimat pernyataan Bahasa Indonesia di sini..."
                                        className="w-full bg-white/80 dark:bg-slate-950/80 border border-slate-200 dark:border-slate-800 rounded-2xl p-4 text-slate-900 dark:text-slate-100 placeholder-slate-400 dark:placeholder-slate-500 focus:outline-none focus:border-brand-500 focus:ring-1 focus:ring-brand-500 text-sm leading-relaxed"
                                    ></textarea>

                                    <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                                        <div className="flex flex-col space-y-1.5">
                                            <span className="text-[10px] text-slate-500 dark:text-slate-400 font-bold uppercase tracking-wider">Opsi Pemrosesan Data:</span>
                                            <div className="flex flex-wrap items-center gap-1 bg-slate-100 dark:bg-slate-900 p-1 rounded-lg border border-slate-200 dark:border-slate-700/80 w-fit">
                                                <button onClick={() => setProcessingMode('all')} className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-all ${processingMode === 'all' ? 'bg-brand-600 text-white dark:text-slate-950 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>All (Semua)</button>
                                                <button onClick={() => setProcessingMode('structure')} className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-all ${processingMode === 'structure' ? 'bg-purple-600 text-white shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>Struktur Pernyataan</button>
                                                <button onClick={() => setProcessingMode('sentiment')} className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-all ${processingMode === 'sentiment' ? 'bg-emerald-600 text-white dark:text-slate-950 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>Sentimen Token</button>
                                            </div>
                                        </div>
                                        <button
                                            onClick={handleAnnotateSingle}
                                            disabled={loadingSingle}
                                            className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-brand-600 to-emerald-500 hover:from-brand-500 hover:to-emerald-400 text-white dark:text-slate-950 font-bold text-sm flex items-center space-x-2 transition-all shadow-lg shadow-brand-500/20 disabled:opacity-50 w-full md:w-auto justify-center"
                                        >
                                            {loadingSingle ? (
                                                <>
                                                    <i data-lucide="loader-2" className="w-4 h-4 animate-spin"></i>
                                                    <span>Menganalisis...</span>
                                                </>
                                            ) : (
                                                <>
                                                    <i data-lucide="zap" className="w-4 h-4"></i>
                                                    <span>Proses Anotasi</span>
                                                </>
                                            )}
                                        </button>
                                    </div>
                                </div>

                                {annotationResult && (
                                    <div className="glass-panel p-6 rounded-3xl space-y-6">
                                        <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 dark:border-slate-800 pb-4">
                                            <div className="flex items-center space-x-3">
                                                <div className="w-3 h-3 rounded-full bg-brand-500 animate-ping"></div>
                                                <h2 className="font-bold text-lg text-slate-900 dark:text-slate-100">Hasil Anotasi BIO GANDA</h2>
                                                <div className="flex items-center space-x-2">
                                                    <span className="text-xs text-slate-500 dark:text-slate-400 font-medium">Agregat:</span>
                                                    <select
                                                        value={annotationResult.sentiment_agregat || 'NET'}
                                                        onChange={(e) => {
                                                            setAnnotationResult({
                                                                ...annotationResult,
                                                                sentiment_agregat: e.target.value
                                                            });
                                                        }}
                                                        className={`px-3 py-1 rounded-xl font-bold text-xs border focus:outline-none cursor-pointer ${
                                                            annotationResult.sentiment_agregat === 'POS' ? 'bg-emerald-100 dark:bg-emerald-950 text-emerald-800 dark:text-emerald-300 border-emerald-300 dark:border-emerald-500' :
                                                            annotationResult.sentiment_agregat === 'NEG' ? 'bg-rose-100 dark:bg-rose-950 text-rose-800 dark:text-rose-300 border-rose-300 dark:border-rose-500' : 'bg-slate-100 dark:bg-slate-800 text-slate-800 dark:text-slate-300 border-slate-300 dark:border-slate-600'
                                                        }`}
                                                    >
                                                        <option value="POS" className="bg-white dark:bg-slate-900 text-emerald-800 dark:text-emerald-300">POS</option>
                                                        <option value="NEG" className="bg-white dark:bg-slate-900 text-rose-800 dark:text-rose-300">NEG</option>
                                                        <option value="NET" className="bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-300">NET</option>
                                                    </select>
                                                </div>
                                            </div>

                                            <div className="flex items-center space-x-2">
                                                <div className="bg-slate-100 dark:bg-slate-900 p-1 rounded-xl border border-slate-200 dark:border-slate-800 flex space-x-1">
                                                    <button
                                                        onClick={() => setViewMode('visual')}
                                                        className={`px-3 py-1 rounded-lg text-xs font-medium ${viewMode === 'visual' ? 'bg-brand-600 text-white dark:text-slate-950 font-bold' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-800'}`}
                                                    >
                                                        Visual Badge
                                                    </button>
                                                    <button
                                                        onClick={() => setViewMode('edit')}
                                                        className={`px-3 py-1 rounded-lg text-xs font-medium ${viewMode === 'edit' ? 'bg-brand-600 text-white dark:text-slate-950 font-bold' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-800'}`}
                                                    >
                                                        Interactive Edit
                                                    </button>
                                                    <button
                                                        onClick={() => setViewMode('json')}
                                                        className={`px-3 py-1 rounded-lg text-xs font-medium ${viewMode === 'json' ? 'bg-brand-600 text-white dark:text-slate-950 font-bold' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-800'}`}
                                                    >
                                                        JSON
                                                    </button>
                                                </div>

                                                <button
                                                    onClick={() => exportToExcel([annotationResult], "Single_Sentence_Anotasi.xlsx")}
                                                    className="px-3 py-1.5 rounded-xl bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-semibold text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 flex items-center space-x-1.5"
                                                >
                                                    <i data-lucide="download" className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400"></i>
                                                    <span>Export Excel</span>
                                                </button>
                                            </div>
                                        </div>

                                        {viewMode === 'visual' && (
                                            <div className="space-y-6">
                                                <div className="glass-card p-5 rounded-2xl space-y-3">
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-xs font-bold tracking-wider text-purple-600 dark:text-purple-400 uppercase flex items-center space-x-1.5">
                                                            <i data-lucide="git-commit" className="w-4 h-4"></i>
                                                            <span>Layer 1: Struktur Pernyataan (statement_tag)</span>
                                                        </span>
                                                    </div>

                                                    <div className="flex flex-wrap gap-2 pt-2">
                                                        {annotationResult.tokens.map((tok, idx) => {
                                                            const tag = tok.statement_tag || 'O';
                                                            const baseTag = tag.includes('-') ? tag.split('-')[1] : tag;
                                                            const colorStyle = STATEMENT_COLORS[baseTag] || STATEMENT_COLORS['O'];

                                                            return (
                                                                <div key={idx} className={`px-2.5 py-1.5 rounded-xl border flex flex-col items-center space-y-0.5 ${colorStyle}`}>
                                                                    <span className="text-sm font-medium">{tok.word}</span>
                                                                    <span className="text-[10px] font-mono opacity-80">{tag}</span>
                                                                </div>
                                                            );
                                                        })}
                                                    </div>
                                                </div>

                                                <div className="glass-card p-5 rounded-2xl space-y-3">
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-xs font-bold tracking-wider text-emerald-600 dark:text-emerald-400 uppercase flex items-center space-x-1.5">
                                                            <i data-lucide="heart" className="w-4 h-4"></i>
                                                            <span>Layer 2: Sentimen Token (sentiment_tag)</span>
                                                        </span>
                                                    </div>

                                                    <div className="flex flex-wrap gap-2 pt-2">
                                                        {annotationResult.tokens.map((tok, idx) => {
                                                            const tag = tok.sentiment_tag || 'O';
                                                            const baseTag = tag.includes('-') ? tag.split('-')[1] : tag;
                                                            const colorStyle = SENTIMENT_COLORS[baseTag] || SENTIMENT_COLORS['O'];

                                                            return (
                                                                <div key={idx} className={`px-2.5 py-1.5 rounded-xl border flex flex-col items-center space-y-0.5 ${colorStyle}`}>
                                                                    <span className="text-sm font-medium">{tok.word}</span>
                                                                    <span className="text-[10px] font-mono opacity-80">{tag}</span>
                                                                </div>
                                                            );
                                                        })}
                                                    </div>
                                                </div>

                                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                                    <div className="bg-slate-50/80 dark:bg-slate-900/80 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 space-y-2">
                                                        <h3 className="text-xs font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wider flex items-center space-x-2">
                                                            <i data-lucide="box" className="w-4 h-4 text-purple-600 dark:text-purple-400"></i>
                                                            <span>Entitas Struktur Terkstrak</span>
                                                        </h3>
                                                        <div className="space-y-1.5 max-h-48 overflow-y-auto pr-2">
                                                            {annotationResult.statement_entities && annotationResult.statement_entities.length > 0 ? (
                                                                annotationResult.statement_entities.map((ent, idx) => (
                                                                    <div key={idx} className="p-2 rounded-xl bg-white/60 dark:bg-slate-950/60 border border-slate-200 dark:border-slate-800 flex justify-between items-center text-xs">
                                                                        <span className="text-slate-800 dark:text-slate-200 font-medium">"{ent.text}"</span>
                                                                        <span className="px-2 py-0.5 rounded-md bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 font-mono font-bold text-[10px] border border-purple-200 dark:border-purple-800">
                                                                            {ent.label}
                                                                        </span>
                                                                    </div>
                                                                ))
                                                            ) : (
                                                                <p className="text-xs text-slate-500 italic">Tidak ada entitas struktur terkstrak.</p>
                                                            )}
                                                        </div>
                                                    </div>

                                                    <div className="bg-slate-50/80 dark:bg-slate-900/80 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 space-y-2">
                                                        <h3 className="text-xs font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wider flex items-center space-x-2">
                                                            <i data-lucide="smile" className="w-4 h-4 text-emerald-600 dark:text-emerald-400"></i>
                                                            <span>Entitas Sentimen Terkstrak</span>
                                                        </h3>
                                                        <div className="space-y-1.5 max-h-48 overflow-y-auto pr-2">
                                                            {annotationResult.sentiment_entities && annotationResult.sentiment_entities.length > 0 ? (
                                                                annotationResult.sentiment_entities.map((ent, idx) => (
                                                                    <div key={idx} className="p-2 rounded-xl bg-white/60 dark:bg-slate-950/60 border border-slate-200 dark:border-slate-800 flex justify-between items-center text-xs">
                                                                        <span className="text-slate-800 dark:text-slate-200 font-medium">"{ent.text}"</span>
                                                                        <span className={`px-2 py-0.5 rounded-md font-mono font-bold text-[10px] border ${
                                                                            ent.label === 'POS' ? 'bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 border-emerald-200 dark:border-emerald-800' :
                                                                            ent.label === 'NEG' ? 'bg-rose-100 dark:bg-rose-950 text-rose-700 dark:text-rose-300 border-rose-200 dark:border-rose-800' : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-300 dark:border-slate-700'
                                                                        }`}>
                                                                            {ent.label}
                                                                        </span>
                                                                    </div>
                                                                ))
                                                            ) : (
                                                                <p className="text-xs text-slate-500 italic">Tidak ada entitas sentimen terkstrak.</p>
                                                            )}
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        )}

                                        {viewMode === 'edit' && (
                                            <div className="space-y-4">
                                                <p className="text-xs text-slate-600 dark:text-slate-400">
                                                    Anda dapat mengubah tag BIO secara langsung pada tabel di bawah. Perubahan akan langsung memperbarui daftar entitas dan sentimen agregat.
                                                </p>
                                                <div className="overflow-x-auto rounded-2xl border border-slate-200 dark:border-slate-800">
                                                    <table className="w-full text-left text-xs text-slate-700 dark:text-slate-300">
                                                        <thead className="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400 font-mono uppercase text-[10px]">
                                                            <tr>
                                                                <th className="px-4 py-3 border-b border-slate-200 dark:border-slate-800 w-12 text-center">#</th>
                                                                <th className="px-4 py-3 border-b border-slate-200 dark:border-slate-800">Token / Word</th>
                                                                <th className="px-4 py-3 border-b border-slate-200 dark:border-slate-800">Statement Tag (Layer 1)</th>
                                                                <th className="px-4 py-3 border-b border-slate-200 dark:border-slate-800">Sentiment Tag (Layer 2)</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 bg-white/40 dark:bg-slate-950/40 font-mono">
                                                            {annotationResult.tokens.map((tok, idx) => (
                                                                <tr key={idx} className="hover:bg-slate-50 dark:hover:bg-slate-900/50 transition-all">
                                                                    <td className="px-4 py-2.5 text-center text-slate-500">{idx + 1}</td>
                                                                    <td className="px-4 py-2.5 text-slate-900 dark:text-slate-100 font-semibold font-sans text-sm">{tok.word}</td>
                                                                    <td className="px-4 py-2.5">
                                                                        <select
                                                                            value={tok.statement_tag || 'O'}
                                                                            onChange={(e) => handleTokenTagChange(idx, 'statement_tag', e.target.value)}
                                                                            className="bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-lg px-2.5 py-1 text-xs text-purple-700 dark:text-purple-300 focus:outline-none focus:border-purple-500"
                                                                        >
                                                                            {STATEMENT_TAGS.map(t => <option key={t} value={t}>{t}</option>)}
                                                                        </select>
                                                                    </td>
                                                                    <td className="px-4 py-2.5">
                                                                        <select
                                                                            value={tok.sentiment_tag || 'O'}
                                                                            onChange={(e) => handleTokenTagChange(idx, 'sentiment_tag', e.target.value)}
                                                                            className="bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-lg px-2.5 py-1 text-xs text-emerald-700 dark:text-emerald-300 focus:outline-none focus:border-emerald-500"
                                                                        >
                                                                            {SENTIMENT_TAGS.map(t => <option key={t} value={t}>{t}</option>)}
                                                                        </select>
                                                                    </td>
                                                                </tr>
                                                            ))}
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </div>
                                        )}

                                        {viewMode === 'json' && (
                                            <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 overflow-x-auto">
                                                <pre className="font-mono text-xs text-emerald-700 dark:text-emerald-400 leading-relaxed">
                                                    {JSON.stringify(annotationResult, null, 2)}
                                                </pre>
                                            </div>
                                        )}
                                    </div>
                                )}
                            </div>
                        )}

                        {/* BATCH MODE */}
                        {activeTab === 'batch' && (
                            <div className="space-y-6">
                                <div className="glass-panel p-6 rounded-3xl space-y-6">
                                    <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 dark:border-slate-800 pb-4">
                                        <div>
                                            <h2 className="font-bold text-lg text-slate-900 dark:text-slate-100 flex items-center space-x-2">
                                                <i data-lucide="layers" className="w-5 h-5 text-brand-500"></i>
                                                <span>Batch Processing Kalimat Banyak</span>
                                            </h2>
                                            <p className="text-xs text-slate-500 dark:text-slate-400">Masukkan beberapa kalimat (1 kalimat per baris) atau unggah file Excel/CSV.</p>
                                        </div>

                                        <div className="flex items-center space-x-3">
                                            <label className="px-4 py-2 rounded-xl bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 border border-slate-200 dark:border-slate-700 text-xs font-medium text-slate-700 dark:text-slate-200 cursor-pointer flex items-center space-x-2 transition-all">
                                                <i data-lucide="upload" className="w-4 h-4 text-brand-500"></i>
                                                <span>Import File Excel</span>
                                                <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" />
                                            </label>
                                        </div>
                                    </div>

                                    <textarea
                                        rows={6}
                                        value={batchInputText}
                                        onChange={(e) => setBatchInputText(e.target.value)}
                                        placeholder={`Presiden Jokowi menyatakan bahwa pertumbuhan ekonomi Kuartal III sangat memuaskan.\nMenteri Keuangan menjelaskan kebijakan insentif pajak baru akan merangsang investasi.\nPengamat menilai hambatan birokrasi berpotensi merugikan industri.`}
                                        className="w-full bg-white/80 dark:bg-slate-950/80 border border-slate-200 dark:border-slate-800 rounded-2xl p-4 text-slate-900 dark:text-slate-100 placeholder-slate-400 dark:placeholder-slate-600 focus:outline-none focus:border-brand-500 text-sm leading-relaxed font-mono"
                                    ></textarea>

                                    <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                                        <div className="flex flex-col space-y-1.5">
                                            <span className="text-[10px] text-slate-500 dark:text-slate-400 font-bold uppercase tracking-wider">Opsi Pemrosesan Data:</span>
                                            <div className="flex flex-wrap items-center gap-1 bg-slate-100 dark:bg-slate-900 p-1 rounded-lg border border-slate-200 dark:border-slate-700/80 w-fit">
                                                <button onClick={() => setProcessingMode('all')} className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-all ${processingMode === 'all' ? 'bg-brand-600 text-white dark:text-slate-950 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>All (Semua)</button>
                                                <button onClick={() => setProcessingMode('structure')} className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-all ${processingMode === 'structure' ? 'bg-purple-600 text-white shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>Struktur Pernyataan</button>
                                                <button onClick={() => setProcessingMode('sentiment')} className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-all ${processingMode === 'sentiment' ? 'bg-emerald-600 text-white dark:text-slate-950 shadow-sm' : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>Sentimen Token</button>
                                            </div>
                                        </div>

                                        <div className="flex items-center space-x-3 w-full md:w-auto justify-end">
                                            <span className="text-xs text-slate-500 dark:text-slate-400 font-mono hidden md:block">
                                                {batchInputText.split('\n').filter(Boolean).length} kalimat terdeteksi
                                            </span>

                                            <button
                                                onClick={handleBatchProcess}
                                                disabled={isProcessingBatch}
                                                className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-brand-600 to-emerald-500 text-white dark:text-slate-950 font-bold text-sm flex items-center space-x-2 shadow-lg shadow-brand-500/20 disabled:opacity-50 w-full md:w-auto justify-center"
                                            >
                                                {isProcessingBatch ? (
                                                    <>
                                                        <i data-lucide="loader-2" className="w-4 h-4 animate-spin"></i>
                                                        <span>Memproses ({batchProgress.current}/{batchProgress.total})...</span>
                                                    </>
                                                ) : (
                                                    <>
                                                        <i data-lucide="play" className="w-4 h-4 fill-current"></i>
                                                        <span>Jalankan Batch Anotasi</span>
                                                    </>
                                                )}
                                            </button>
                                        </div>
                                    </div>

                                    {isProcessingBatch && (
                                        <div className="space-y-2">
                                            <div className="w-full bg-slate-200 dark:bg-slate-900 rounded-full h-2 overflow-hidden border border-slate-300 dark:border-slate-800">
                                                <div
                                                    className="bg-brand-500 h-2 rounded-full transition-all duration-300"
                                                    style={{ width: `${(batchProgress.current / batchProgress.total) * 100}%` }}
                                                ></div>
                                            </div>
                                        </div>
                                    )}
                                </div>

                                {batchResults.length > 0 && (
                                    <div className="glass-panel p-6 rounded-3xl space-y-4">
                                        <div className="flex justify-between items-center border-b border-slate-200 dark:border-slate-800 pb-4">
                                            <div>
                                                <h3 className="font-bold text-slate-800 dark:text-slate-200 text-base">Hasil Batch Anotasi ({batchResults.length} Kalimat)</h3>
                                                <p className="text-xs text-slate-500 dark:text-slate-400">Hasil kalimat yang telah selesai dipproses ditampilkan secara real-time.</p>
                                            </div>
                                            <button
                                                onClick={() => exportToExcel(batchResults, "Batch_Anotasi_BIO_Ganda.xlsx")}
                                                className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white dark:text-slate-950 font-bold text-xs flex items-center space-x-2 shadow-lg shadow-emerald-500/20"
                                            >
                                                <i data-lucide="download" className="w-4 h-4"></i>
                                                <span>Download Output Excel</span>
                                            </button>
                                        </div>

                                        <div className="space-y-4">
                                            {batchResults.map((r, sIdx) => {
                                                const isEditing = editingBatchIndex === sIdx;
                                                return (
                                                    <div key={r.sentence_id} className="glass-card p-4 rounded-2xl space-y-3 border border-slate-200 dark:border-slate-800">
                                                        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 dark:border-slate-800/60 pb-3">
                                                            <div className="flex items-center space-x-3">
                                                                <span className="w-6 h-6 rounded-full bg-slate-100 dark:bg-slate-900 border border-slate-300 dark:border-slate-700 flex items-center justify-center text-xs font-mono text-slate-600 dark:text-slate-300 font-bold">
                                                                    {r.sentence_id}
                                                                </span>
                                                                <span className="font-semibold text-slate-900 dark:text-slate-100 text-sm">"{r.text}"</span>
                                                            </div>

                                                            <div className="flex items-center space-x-3">
                                                                <div className="flex items-center space-x-1.5">
                                                                    <span className="text-xs text-slate-500 dark:text-slate-400 font-medium">Agregat:</span>
                                                                    <select
                                                                        value={r.sentiment_agregat || 'NET'}
                                                                        onChange={(e) => {
                                                                            const updatedResults = [...batchResults];
                                                                            updatedResults[sIdx] = {
                                                                                ...updatedResults[sIdx],
                                                                                sentiment_agregat: e.target.value
                                                                            };
                                                                            setBatchResults(updatedResults);
                                                                        }}
                                                                        className={`px-2.5 py-1 rounded-xl font-mono font-bold text-xs border focus:outline-none cursor-pointer ${
                                                                            r.sentiment_agregat === 'POS' ? 'bg-emerald-100 dark:bg-emerald-950 text-emerald-800 dark:text-emerald-300 border-emerald-300 dark:border-emerald-500' :
                                                                            r.sentiment_agregat === 'NEG' ? 'bg-rose-100 dark:bg-rose-950 text-rose-800 dark:text-rose-300 border-rose-300 dark:border-rose-500' : 'bg-slate-100 dark:bg-slate-800 text-slate-800 dark:text-slate-300 border-slate-300 dark:border-slate-600'
                                                                        }`}
                                                                    >
                                                                        <option value="POS" className="bg-white dark:bg-slate-900 text-emerald-800 dark:text-emerald-300">POS</option>
                                                                        <option value="NEG" className="bg-white dark:bg-slate-900 text-rose-800 dark:text-rose-300">NEG</option>
                                                                        <option value="NET" className="bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-300">NET</option>
                                                                    </select>
                                                                </div>

                                                                <button
                                                                    onClick={() => setEditingBatchIndex(isEditing ? null : sIdx)}
                                                                    className={`px-3 py-1.5 rounded-xl text-xs font-semibold border flex items-center space-x-1.5 transition-all ${
                                                                        isEditing 
                                                                            ? 'bg-brand-600 text-white dark:text-slate-950 border-brand-500 font-bold' 
                                                                            : 'bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 border-slate-200 dark:border-slate-700'
                                                                    }`}
                                                                >
                                                                    <i data-lucide={isEditing ? "check" : "edit-3"} className="w-3.5 h-3.5"></i>
                                                                    <span>{isEditing ? 'Selesai Edit' : 'Edit Interaktif'}</span>
                                                                </button>
                                                            </div>
                                                        </div>

                                                        {!isEditing ? (
                                                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1 text-xs">
                                                                <div className="bg-slate-50/60 dark:bg-slate-900/60 p-3 rounded-xl border border-slate-200 dark:border-slate-800/80 space-y-1.5">
                                                                    <span className="text-[10px] font-bold text-purple-600 dark:text-purple-400 uppercase tracking-wider">Entitas Struktur</span>
                                                                    <div className="flex flex-wrap gap-1 pt-1">
                                                                        {r.statement_entities?.map((e, i) => (
                                                                            <span key={i} className="px-2 py-0.5 rounded bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 text-[10px] border border-purple-200 dark:border-purple-800">
                                                                                {e.label}: {e.text}
                                                                            </span>
                                                                        ))}
                                                                    </div>
                                                                </div>

                                                                <div className="bg-slate-50/60 dark:bg-slate-900/60 p-3 rounded-xl border border-slate-200 dark:border-slate-800/80 space-y-1.5">
                                                                    <span className="text-[10px] font-bold text-emerald-600 dark:text-emerald-400 uppercase tracking-wider">Entitas Sentimen</span>
                                                                    <div className="flex flex-wrap gap-1 pt-1">
                                                                        {r.sentiment_entities?.map((e, i) => (
                                                                            <span key={i} className={`px-2 py-0.5 rounded text-[10px] border ${
                                                                                e.label === 'POS' ? 'bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 border-emerald-200 dark:border-emerald-800' :
                                                                                e.label === 'NEG' ? 'bg-rose-100 dark:bg-rose-950 text-rose-700 dark:text-rose-300 border-rose-200 dark:border-rose-800' : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-300 dark:border-slate-700'
                                                                            }`}>
                                                                                {e.label}: {e.text}
                                                                            </span>
                                                                        ))}
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        ) : (
                                                            <div className="space-y-3 pt-2">
                                                                <p className="text-xs text-brand-600 dark:text-brand-400 italic">Mode edit aktif untuk kalimat #{r.sentence_id}. Ubah label token di bawah ini:</p>
                                                                <div className="overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-800 max-h-60">
                                                                    <table className="w-full text-left text-xs text-slate-700 dark:text-slate-300">
                                                                        <thead className="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400 font-mono uppercase text-[10px] sticky top-0">
                                                                            <tr>
                                                                                <th className="px-3 py-2 border-b border-slate-200 dark:border-slate-800 w-10 text-center">#</th>
                                                                                <th className="px-3 py-2 border-b border-slate-200 dark:border-slate-800">Token</th>
                                                                                <th className="px-3 py-2 border-b border-slate-200 dark:border-slate-800">Statement Tag</th>
                                                                                <th className="px-3 py-2 border-b border-slate-200 dark:border-slate-800">Sentiment Tag</th>
                                                                            </tr>
                                                                        </thead>
                                                                        <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 bg-white/60 dark:bg-slate-950/60 font-mono">
                                                                            {r.tokens.map((tok, tIdx) => (
                                                                                <tr key={tIdx} className="hover:bg-slate-50 dark:hover:bg-slate-900/50">
                                                                                    <td className="px-3 py-2 text-center text-slate-500">{tIdx + 1}</td>
                                                                                    <td className="px-3 py-2 text-slate-900 dark:text-slate-100 font-semibold font-sans text-sm">{tok.word}</td>
                                                                                    <td className="px-3 py-2">
                                                                                        <select
                                                                                            value={tok.statement_tag || 'O'}
                                                                                            onChange={(e) => handleBatchTokenTagChange(sIdx, tIdx, 'statement_tag', e.target.value)}
                                                                                            className="bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-lg px-2 py-1 text-xs text-purple-700 dark:text-purple-300 focus:outline-none focus:border-purple-500"
                                                                                        >
                                                                                            {STATEMENT_TAGS.map(t => <option key={t} value={t}>{t}</option>)}
                                                                                        </select>
                                                                                    </td>
                                                                                    <td className="px-3 py-2">
                                                                                        <select
                                                                                            value={tok.sentiment_tag || 'O'}
                                                                                            onChange={(e) => handleBatchTokenTagChange(sIdx, tIdx, 'sentiment_tag', e.target.value)}
                                                                                            className="bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded-lg px-2 py-1 text-xs text-emerald-700 dark:text-emerald-300 focus:outline-none focus:border-emerald-500"
                                                                                        >
                                                                                            {SENTIMENT_TAGS.map(t => <option key={t} value={t}>{t}</option>)}
                                                                                        </select>
                                                                                    </td>
                                                                                </tr>
                                                                            ))}
                                                                        </tbody>
                                                                    </table>
                                                                </div>
                                                            </div>
                                                        )}
                                                    </div>
                                                );
                                            })}
                                        </div>
                                    </div>
                                )}
                            </div>
                        )}

                        {/* GUIDE MODE */}
                        {activeTab === 'guide' && (
                            <div className="glass-panel p-6 sm:p-8 rounded-3xl space-y-6">
                                <div className="border-b border-slate-200 dark:border-slate-800 pb-4 flex flex-wrap justify-between items-center gap-3">
                                    <div>
                                        <h2 className="font-bold text-xl text-slate-900 dark:text-slate-100 flex items-center space-x-2">
                                            <i data-lucide="book-open" className="w-6 h-6 text-brand-500"></i>
                                            <span>Panduan Aturan Anotasi BIO GANDA</span>
                                        </h2>
                                        <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">Pedoman komprehensif penabelan entitas struktur pernyataan dan sentimen Bahasa Indonesia.</p>
                                    </div>
                                    <div className="flex items-center space-x-2">
                                        {!isEditingGuide ? (
                                            <button
                                                onClick={() => setIsEditingGuide(true)}
                                                className="px-4 py-2 rounded-xl bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-semibold text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 flex items-center space-x-1.5 transition-all"
                                            >
                                                <i data-lucide="edit-3" className="w-4 h-4 text-brand-600 dark:text-brand-400"></i>
                                                <span>Edit Panduan</span>
                                            </button>
                                        ) : (
                                            <>
                                                <button
                                                    onClick={() => {
                                                        localStorage.setItem('custom_guide_text', guideText);
                                                        setIsEditingGuide(false);
                                                        showToast("Panduan aturan berhasil disimpan!", "success");
                                                    }}
                                                    className="px-4 py-2 rounded-xl bg-brand-600 hover:bg-brand-500 text-white dark:text-slate-950 font-bold text-xs flex items-center space-x-1.5 shadow-lg shadow-brand-500/20"
                                                >
                                                    <i data-lucide="save" className="w-4 h-4"></i>
                                                    <span>Save Panduan</span>
                                                </button>
                                                <button
                                                    onClick={() => {
                                                        setGuideText(DEFAULT_GUIDE_TEXT);
                                                        localStorage.removeItem('custom_guide_text');
                                                        setIsEditingGuide(false);
                                                        showToast("Panduan dikembalikan ke default.", "info");
                                                    }}
                                                    className="px-3 py-2 rounded-xl bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-semibold text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 flex items-center space-x-1.5"
                                                >
                                                    <i data-lucide="rotate-ccw" className="w-3.5 h-3.5"></i>
                                                    <span>Reset</span>
                                                </button>
                                                <button
                                                    onClick={() => setIsEditingGuide(false)}
                                                    className="px-3 py-2 rounded-xl bg-slate-100 dark:bg-slate-900 hover:bg-slate-200 dark:hover:bg-slate-800 text-xs font-semibold text-slate-600 dark:text-slate-400 border border-slate-200 dark:border-slate-800"
                                                >
                                                    Batal
                                                </button>
                                            </>
                                        )}
                                    </div>
                                </div>

                                {!isEditingGuide ? (
                                    <div className="glass-card p-6 rounded-2xl space-y-4">
                                        <pre className="text-xs text-slate-800 dark:text-slate-200 whitespace-pre-wrap font-mono leading-relaxed bg-slate-50/80 dark:bg-slate-950/80 p-5 rounded-xl border border-slate-200 dark:border-slate-800">
                                            {guideText}
                                        </pre>
                                    </div>
                                ) : (
                                    <div className="space-y-3">
                                        <textarea
                                            rows={14}
                                            value={guideText}
                                            onChange={(e) => setGuideText(e.target.value)}
                                            className="w-full bg-white dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-2xl p-5 text-slate-800 dark:text-slate-200 font-mono text-xs leading-relaxed focus:outline-none focus:border-brand-500 shadow-inner"
                                        ></textarea>
                                    </div>
                                )}
                            </div>
                        )}

                        {/* PROMPT EDITOR MODE */}
                        {activeTab === 'prompt' && (
                            <div className="glass-panel p-6 sm:p-8 rounded-3xl space-y-6">
                                <div className="border-b border-slate-200 dark:border-slate-800 pb-4 flex flex-wrap justify-between items-center gap-3">
                                    <div>
                                        <h2 className="font-bold text-xl text-slate-900 dark:text-slate-100 flex items-center space-x-2">
                                            <i data-lucide="terminal" className="w-6 h-6 text-brand-500"></i>
                                            <span>System Prompt LLM Editor</span>
                                        </h2>
                                        <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">Sesuaikan instruksi sistem yang dikirimkan ke model Gemini API.</p>
                                    </div>
                                    <div className="flex items-center space-x-2">
                                        <button
                                            onClick={() => {
                                                localStorage.setItem('custom_system_prompt', systemPrompt);
                                                showToast("System prompt berhasil disimpan!", "success");
                                            }}
                                            className="px-4 py-2 rounded-xl bg-brand-600 hover:bg-brand-500 text-white dark:text-slate-950 font-bold text-xs flex items-center space-x-1.5 shadow-lg shadow-brand-500/20"
                                        >
                                            <i data-lucide="save" className="w-4 h-4"></i>
                                            <span>Save Prompt</span>
                                        </button>
                                        <button
                                            onClick={() => {
                                                setSystemPrompt(DEFAULT_SYSTEM_PROMPT);
                                                localStorage.removeItem('custom_system_prompt');
                                                showToast("System prompt dikembalikan ke default.", "success");
                                            }}
                                            className="px-3 py-2 rounded-xl bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-semibold text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 flex items-center space-x-1.5"
                                        >
                                            <i data-lucide="rotate-ccw" className="w-3.5 h-3.5"></i>
                                            <span>Reset Default</span>
                                        </button>
                                    </div>
                                </div>

                                <div className="relative">
                                    <div className="absolute top-3 right-3 flex items-center space-x-2 bg-slate-100/80 dark:bg-slate-900/80 px-3 py-1 rounded-lg border border-slate-200/80 dark:border-slate-700/80 text-[10px] font-mono text-slate-600 dark:text-slate-400">
                                        <i data-lucide="shield-alert" className="w-3 h-3 text-brand-600 dark:text-brand-400"></i>
                                        <span>Editable LLM Instructions</span>
                                    </div>
                                    <textarea
                                        rows={16}
                                        value={systemPrompt}
                                        onChange={(e) => setSystemPrompt(e.target.value)}
                                        className="w-full bg-white dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-2xl p-4 text-emerald-700 dark:text-emerald-400 font-mono text-xs leading-relaxed focus:outline-none focus:border-brand-500 shadow-inner"
                                    ></textarea>
                                </div>
                            </div>
                        )}
                    </main>
                </div>
            );
        }

        ReactDOM.render(<App />, document.getElementById('root'));
    </script>
</body>
</html>
