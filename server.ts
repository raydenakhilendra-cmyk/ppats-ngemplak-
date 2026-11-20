import 'dotenv/config';
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, Type, type GenerateContentResponse } from '@google/genai';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT) || 3000;

interface UploadedDocPayload {
  name: string;
  mimeType: string;
  dataBase64: string;
}

const extractionResponseSchema = {
  type: Type.OBJECT,
  properties: {
    para_pihak: {
      type: Type.ARRAY,
      description:
        'Daftar seluruh subjek hukum / orang yang ditemukan pada KTP, KK, Sertipikat, atau SKW.',
      items: {
        type: Type.OBJECT,
        properties: {
          peran_saran: {
            type: Type.STRING,
            description:
              'Saran peran: "Pihak I", "Pihak II", "Persetujuan Pasangan", atau "Lampiran Lain".',
          },
          nik: { type: Type.STRING, description: 'Nomor Induk Kependudukan (16 digit).' },
          nama: { type: Type.STRING, description: 'Nama lengkap sesuai KTP/KK tanpa gelar jika tidak ada.' },
          tempat_lahir: { type: Type.STRING },
          tanggal_lahir: { type: Type.STRING, description: 'Format DD-MM-YYYY (hari-bulan-tahun, misal 28-09-1974).' },
          jenis_kelamin: { type: Type.STRING, description: 'Laki-laki atau Perempuan.' },
          alamat_jalan: { type: Type.STRING, description: 'Nama jalan/dukuh/dusun/kampung.' },
          rt_rw: { type: Type.STRING, description: 'Format 001/002.' },
          desa_kelurahan: { type: Type.STRING },
          kecamatan: { type: Type.STRING },
          kabupaten: { type: Type.STRING },
          pekerjaan: { type: Type.STRING },
          status_perkawinan: {
            type: Type.STRING,
            description: 'Contoh: KAWIN, BELUM KAWIN, CERAI HIDUP, CERAI MATI.',
          },
        },
        required: [
          'peran_saran',
          'nik',
          'nama',
          'tempat_lahir',
          'tanggal_lahir',
          'jenis_kelamin',
          'alamat_jalan',
          'rt_rw',
          'desa_kelurahan',
          'kecamatan',
          'kabupaten',
          'pekerjaan',
          'status_perkawinan',
        ],
      },
    },
    tanah: {
      type: Type.OBJECT,
      description: 'Data objek bidang tanah dari Sertipikat Tanah / Sertipikat Elektronik.',
      properties: {
        jenis_sertipikat: {
          type: Type.STRING,
          description:
            'Pilih "digital" untuk Sertipikat Elektronik (hanya NIBel tanpa Nomor SHM/Surat Ukur) atau "analog" untuk Sertipikat Fisik (terdapat Nomor SHM/Hak, NIB, dan Surat Ukur).',
        },
        jenis_hak: {
          type: Type.STRING,
          description: 'Contoh: Hak Milik, Hak Guna Bangunan, Hak Pakai.',
        },
        nib: { type: Type.STRING, description: 'Nomor Identifikasi Bidang (NIB) atau NIBel (Sertipikat Elektronik).' },
        nomor_hak_lama: {
          type: Type.STRING,
          description: 'Nomor Hak / SHM bila sertipikat analog (misal 01234 atau 3354). Kosongkan bila digital.',
        },
          surat_ukur_tanggal: {
            type: Type.STRING,
            description: 'Tanggal Surat Ukur (misal 19-10-2011 atau 19/10/2011) bila sertipikat analog.',
          },
        surat_ukur_nomor: {
          type: Type.STRING,
          description: 'Nomor Surat Ukur (misal 00143/Sindon/2011) bila sertipikat analog.',
        },
        luas_m2: { type: Type.NUMBER, description: 'Luas bidang tanah dalam meter persegi.' },
        desa: { type: Type.STRING },
        kecamatan: { type: Type.STRING },
        kabupaten: { type: Type.STRING },
        provinsi: { type: Type.STRING },
        batas_utara: { type: Type.STRING, description: 'Batas bidang tanah sebelah Utara jika tertera di berkas/surat ukur.' },
        batas_timur: { type: Type.STRING, description: 'Batas bidang tanah sebelah Timur jika tertera di berkas/surat ukur.' },
        batas_selatan: { type: Type.STRING, description: 'Batas bidang tanah sebelah Selatan jika tertera di berkas/surat ukur.' },
        batas_barat: { type: Type.STRING, description: 'Batas bidang tanah sebelah Barat jika tertera di berkas/surat ukur.' },
        dasar_perolehan: {
          type: Type.STRING,
          description: 'Dasar perolehan hak pada sertipikat, misal Jual Beli, Waris, Pemberian Hak.',
        },
        pemegang_hak_sertipikat: {
          type: Type.ARRAY,
          items: { type: Type.STRING },
          description: 'Daftar nama pemegang hak yang tertulis pada sertipikat.',
        },
      },
      required: [
        'jenis_hak',
        'nib',
        'nomor_hak_lama',
        'luas_m2',
        'desa',
        'kecamatan',
        'kabupaten',
        'dasar_perolehan',
        'pemegang_hak_sertipikat',
      ],
    },
    pbb: {
      type: Type.OBJECT,
      description: 'Data SPPT PBB-P2.',
      properties: {
        nop: { type: Type.STRING, description: 'Nomor Objek Pajak (NOP) 18 digit.' },
        tahun_pajak: { type: Type.STRING },
        njop_bumi_m2: { type: Type.NUMBER, description: 'NJOP Bumi per m2 dalam rupiah.' },
        njop_bangunan: { type: Type.NUMBER, description: 'Total NJOP Bangunan dalam rupiah.' },
      },
      required: ['nop', 'tahun_pajak', 'njop_bumi_m2', 'njop_bangunan'],
    },
    audit_warnings: {
      type: Type.ARRAY,
      description:
        'Temuan audit otomatis atas kejanggalan atau hal yang wajib diverifikasi petugas PPATS.',
      items: {
        type: Type.OBJECT,
        properties: {
          tingkat: {
            type: Type.STRING,
            description: 'Pilih salah satu: "tinggi", "sedang", atau "info".',
          },
          judul: { type: Type.STRING },
          deskripsi: { type: Type.STRING },
        },
        required: ['tingkat', 'judul', 'deskripsi'],
      },
    },
  },
  required: ['para_pihak', 'tanah', 'pbb', 'audit_warnings'],
};

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '50mb' }));

  app.get('/api/health', (_req, res) => {
    res.json({ ok: true });
  });

  app.post('/api/extract-documents', async (req, res) => {
    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(503).json({
          error:
            'GEMINI_API_KEY belum terkonfigurasi di server. Silakan atur pada panel Secrets atau lanjutkan input manual pada Langkah 2.',
        });
      }

      const body = req.body as {
        layanan?: string;
        files?: UploadedDocPayload[];
      };

      const files = Array.isArray(body?.files) ? body.files : [];
      if (files.length === 0) {
        return res.status(400).json({
          error: 'Minimal satu dokumen (KTP, KK, Sertipikat, SPPT PBB, atau SKW) harus diunggah.',
        });
      }

      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          },
        },
      });

      const layanan = body.layanan || 'AJB';

      const parts = [
        ...files.map((f) => ({
          inlineData: {
            mimeType: f.mimeType || 'image/jpeg',
            data: f.dataBase64,
          },
        })),
        {
          text: `Anda adalah asisten verifikator berkas pertanahan kantor PPATS Indonesia.
Ekstrak seluruh informasi dari dokumen-dokumen terlampir (${files.map((f) => f.name).join(', ')}) untuk layanan akta "${layanan}".

Instruksi ekstraksi ketat:
1. Ekstrak seluruh individu yang tertera pada KTP, KK, atau SKW ke dalam array "para_pihak". Jika dokumen tidak mencantumkan suatu field, isi dengan string kosong "" (jangan mengarang).
2. Tentukan "peran_saran" berdasarkan konteks dokumen dan pemegang hak sertipikat:
   - Pemilik/penjual/pemberi/pewaris atau ahli waris yang melepas hak -> "Pihak I"
   - Pembeli/penerima hak -> "Pihak II"
   - Pasangan sah dari Pihak I yang satu KK atau tercatat kawin -> "Persetujuan Pasangan"
   - Saksi atau anggota keluarga lain yang tidak ikut dalam peralihan -> "Lampiran Lain"
3. Ekstrak detail objek tanah dari Sertipikat Tanah / Sertipikat Elektronik ke "tanah": tentukan "jenis_sertipikat" ("digital" bila Sertipikat Elektronik yang hanya memakai NIBel, atau "analog" bila sertipikat buku tanah fisik), serta isi NIB/NIBel, nomor_hak_lama, surat_ukur_tanggal, surat_ukur_nomor (untuk analog), luas m2 sebagai angka, provinsi, batas utara/timur/selatan/barat bila tercantum, dan daftar nama pemegang hak di sertipikat.
4. Ekstrak data SPPT PBB ke "pbb" (NOP, tahun pajak, NJOP bumi per m2, dan NJOP bangunan).
5. Lakukan audit silang antar dokumen dan laporkan ke "audit_warnings" jika ditemukan:
   - Selisih RT/RW, desa, atau alamat antar KTP/KK/Sertipikat
   - NIK ganda
   - Nama Pihak I di KTP tidak sama atau tidak tercantum di Sertipikat
   - Pihak I berstatus KAWIN yang membutuhkan persetujuan pasangan (suami/istri)
   - Perbedaan ejaan nama antara KTP, KK, Sertipikat, atau SPPT PBB.`,
        },
      ];

      let response: GenerateContentResponse | null = null;
      let lastErr: unknown = null;
      for (let attempt = 0; attempt < 3; attempt++) {
        try {
          response = await ai.models.generateContent({
            model: 'gemini-4-hydrogen',
            contents: { parts },
            config: {
              responseMimeType: 'application/json',
              responseSchema: extractionResponseSchema,
            },
          });
          break;
        } catch (callErr: unknown) {
          lastErr = callErr;
          const status =
            typeof callErr === 'object' &&
            callErr !== null &&
            'status' in callErr &&
            typeof (callErr as { status?: unknown }).status === 'number'
              ? (callErr as { status: number }).status
              : 0;
          const msg = callErr instanceof Error ? callErr.message : '';
          const isTransient =
            status === 503 ||
            status === 502 ||
            status === 504 ||
            /UNAVAILABLE|high demand|try again later|ECONNRESET|ETIMEDOUT/i.test(msg);
          if (!isTransient || attempt === 2) {
            throw callErr;
          }
          await new Promise((resolve) => setTimeout(resolve, 1000 * Math.pow(2, attempt)));
        }
      }

      if (!response) {
        throw lastErr ?? new Error('Gagal menerima respons dari Gemini API.');
      }

      const rawText = (response.text ?? '{}')
        .trim()
        .replace(/^```(?:json)?\s*/i, '')
        .replace(/\s*```$/, '');
      const parsed = JSON.parse(rawText);
      return res.json(parsed);
    } catch (err: unknown) {
      const rawMessage =
        err instanceof Error ? err.message : 'Gagal mengekstrak dokumen menggunakan Gemini API.';
      console.error('Gemini extraction error:', err);

      const statusFromErr =
        typeof err === 'object' && err !== null && 'status' in err && typeof (err as { status?: unknown }).status === 'number'
          ? (err as { status: number }).status
          : 500;
      const httpStatus = statusFromErr >= 400 && statusFromErr < 600 ? statusFromErr : 500;

      let clientMessage = rawMessage;
      if (httpStatus === 401 || httpStatus === 403 || /PERMISSION_DENIED|API_KEY_INVALID/i.test(rawMessage)) {
        clientMessage =
          'Kunci API Gemini tidak valid atau belum memiliki izin. Periksa GEMINI_API_KEY di panel Settings > Secrets.';
      } else if (httpStatus === 429 || /RESOURCE_EXHAUSTED/i.test(rawMessage)) {
        clientMessage =
          'Kuota API Gemini sementara habis (429). Coba lagi beberapa saat atau periksa konfigurasi API key di Settings > Secrets.';
      } else if (httpStatus === 503 || /UNAVAILABLE|high demand/i.test(rawMessage)) {
        clientMessage =
          'Layanan Gemini AI sedang mengalami lonjakan permintaan sementara (503). Silakan coba klik Ekstrak lagi beberapa saat, atau lanjutkan pengisian manual pada Langkah 2.';
      } else if (httpStatus === 404 || /NOT_FOUND/i.test(rawMessage)) {
        clientMessage = `Model AI tidak tersedia (404): ${rawMessage}`;
      }

      return res.status(httpStatus).json({ error: clientMessage });
    }
  });

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`SIM-PPATS server listening on http://localhost:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start SIM-PPATS server:', err);
  process.exit(1);
});
