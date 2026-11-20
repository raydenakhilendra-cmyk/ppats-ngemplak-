import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import JSZip from 'jszip';
import FileSaver from 'file-saver';
import {
  BerkasRecord,
  CustomTemplateEntry,
  DataTanah,
  OfficeSettings,
  ParaPihak,
  PersonBasicInfo,
  ServiceType,
  TemplateDocKey,
} from '../types/ppats';
import {
  extractMonthYearFromDate,
  formatNumberId,
  formatRupiah,
  formatTanggalDDMMYYYY,
  formatTanggalLengkapHuruf,
  parseFlexibleDate,
  terbilangLuas,
  terbilangRupiah,
} from './indonesianFormatter';
import { generateDefaultDocxBuffer } from './docxTemplates';
import { buildNonAktaLampiranBuffer } from './nonAktaTemplates';
import { TEMPLATE_DOCS_META } from './seedData';
import { isAktaService } from './taxAndAudit';

function saveBlob(blob: Blob, fileName: string) {
  const mod = FileSaver as unknown as {
    saveAs?: (b: Blob, n: string) => void;
    default?: { saveAs?: (b: Blob, n: string) => void } | ((b: Blob, n: string) => void);
  };
  const fn =
    mod.saveAs ??
    (typeof mod === 'function' ? (mod as unknown as (b: Blob, n: string) => void) : null) ??
    (typeof mod.default === 'function' ? mod.default : mod.default?.saveAs);
  if (fn) {
    fn(blob, fileName);
  } else {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

function getSebutan(jk?: string): string {
  return (jk || '').toLowerCase().includes('perempuan') ? 'Nyonya' : 'Tuan';
}

export function resolveJenisSertipikat(
  tanah: Pick<DataTanah, 'jenis_sertipikat' | 'surat_ukur_nomor' | 'nomor_hak_lama'>
): 'digital' | 'analog' {
  if (tanah.jenis_sertipikat === 'digital' || tanah.jenis_sertipikat === 'analog') {
    return tanah.jenis_sertipikat;
  }
  if (
    tanah.surat_ukur_nomor?.trim() ||
    (tanah.nomor_hak_lama &&
      tanah.nomor_hak_lama.trim() !== '' &&
      tanah.nomor_hak_lama.trim() !== '-')
  ) {
    return 'analog';
  }
  return 'digital';
}

export function formatAnalogNomorHak(raw?: string): string {
  const trimmed = (raw || '').trim();
  if (!trimmed || trimmed === '-') return '-';
  const stripped = trimmed
    .replace(/^(?:SHM|SHGB|SHP|HP|Sertipikat)\s*(?:No\.?|Nomor)\s*[:.]?\s*/i, '')
    .trim();
  return stripped || trimmed;
}

function splitRtRw(rtRw?: string): { rt: string; rw: string } {
  const parts = (rtRw || '').split('/').map((s) => s.trim());
  return {
    rt: parts[0] || '-',
    rw: parts[1] || '-',
  };
}

function getJudulAkta(layanan: ServiceType): string {
  switch (layanan) {
    case 'AJB':
      return 'AKTA JUAL BELI';
    case 'APHB':
      return 'AKTA PEMBAGIAN HAK BERSAMA';
    case 'Akta Hibah':
      return 'AKTA HIBAH';
    case 'Turun Waris':
      return 'BERKAS TURUN WARIS';
    case 'Pemecahan':
      return 'BERKAS PEMECAHAN BIDANG';
    case 'Ukur Ulang':
      return 'BERKAS PENGUKURAN ULANG';
    case 'Ralat Data':
      return 'BERKAS RALAT DATA';
    case 'Pemecahan/Ukur Ulang':
      return 'BERKAS PEMECAHAN / UKUR ULANG';
  }
}

export function getNonAktaBundleFileName(layanan: ServiceType): string {
  switch (layanan) {
    case 'Turun Waris':
      return '02_Lampiran_Bundel_Turun_Waris.docx';
    case 'Pemecahan':
      return '02_Lampiran_Bundel_Pemecahan.docx';
    case 'Ukur Ulang':
      return '02_Lampiran_Bundel_Ukur_Ulang.docx';
    case 'Ralat Data':
      return '02_Lampiran_Bundel_Ralat_Data.docx';
    default:
      return '02_Lampiran_Bundel_Non_Akta.docx';
  }
}

const ORDINAL_WORDS_1_TO_20: Record<number, string> = {
  1: 'Pertama',
  2: 'Kedua',
  3: 'Ketiga',
  4: 'Keempat',
  5: 'Kelima',
  6: 'Keenam',
  7: 'Ketujuh',
  8: 'Kedelapan',
  9: 'Kesembilan',
  10: 'Kesepuluh',
  11: 'Kesebelas',
  12: 'Kedua Belas',
  13: 'Ketiga Belas',
  14: 'Keempat Belas',
  15: 'Kelima Belas',
  16: 'Keenam Belas',
  17: 'Ketujuh Belas',
  18: 'Kedelapan Belas',
  19: 'Kesembilan Belas',
  20: 'Kedua Puluh',
};

export function numberToRoman(num: number): string {
  if (!Number.isFinite(num) || num <= 0) return String(num);
  const table: Array<[number, string]> = [
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let rem = Math.floor(num);
  let out = '';
  for (const [val, sym] of table) {
    while (rem >= val) {
      out += sym;
      rem -= val;
    }
  }
  return out;
}

export function romanToNumber(roman: string): number | null {
  const s = roman.trim().toUpperCase();
  if (!s || !/^[IVXLC]+$/.test(s)) return null;
  const map: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100 };
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const cur = map[s[i]];
    const next = map[s[i + 1]] ?? 0;
    if (cur < next) {
      total -= cur;
    } else {
      total += cur;
    }
  }
  return total > 0 ? total : null;
}

export function parseAdditionalPihakMeta(peran: string): {
  order: number;
  roman_numeral: string;
  peran_label_upper: string;
  peran_title: string;
  peran_lower: string;
} {
  const m = /^Pihak\s+([A-Za-z0-9]+)$/i.exec((peran || '').trim());
  const token = (m?.[1] || '').trim();
  const numeric = /^\d+$/.test(token) ? Number(token) : romanToNumber(token);

  if (numeric !== null && numeric >= 1) {
    const roman = numberToRoman(numeric);
    const word = ORDINAL_WORDS_1_TO_20[numeric] ?? `Ke-${numeric}`;
    return {
      order: numeric,
      roman_numeral: roman,
      peran_label_upper: `PIHAK ${word.toUpperCase()}`,
      peran_title: `Pihak ${word}`,
      peran_lower: `Pihak ${word.toLowerCase()}`,
    };
  }

  return {
    order: 999,
    roman_numeral: token.toUpperCase() || '-',
    peran_label_upper: (peran || 'PIHAK LAINNYA').toUpperCase(),
    peran_title: peran || 'Pihak Lainnya',
    peran_lower: peran || 'Pihak lainnya',
  };
}

function mapPartyForTemplate(p: ParaPihak) {
  const { rt, rw } = splitRtRw(p.rt_rw);
  const parts = [
    p.alamat_jalan,
    p.rt_rw ? `Rukun Tetangga (RT) ${rt} Rukun Warga (RW) ${rw}` : '',
    p.desa_kelurahan ? `Desa ${p.desa_kelurahan}` : '',
    p.kecamatan ? `Kecamatan ${p.kecamatan}` : '',
    p.kabupaten,
  ].filter(Boolean);

  const formattedBirthDate = p.tanggal_lahir ? formatTanggalDDMMYYYY(p.tanggal_lahir) : '-';

  return {
    peran: p.peran,
    sebutan: getSebutan(p.jenis_kelamin),
    nik: p.nik || '-',
    nama: p.nama || '-',
    tempat_lahir: p.tempat_lahir || '-',
    tanggal_lahir: formattedBirthDate,
    ttl: [p.tempat_lahir, formattedBirthDate !== '-' ? formattedBirthDate : ''].filter(Boolean).join(', ') || '-',
    jenis_kelamin: p.jenis_kelamin || '-',
    alamat_jalan: p.alamat_jalan || '-',
    rt_rw: p.rt_rw || '-',
    rt,
    rw,
    desa_kelurahan: p.desa_kelurahan || '-',
    kecamatan: p.kecamatan || '-',
    kabupaten: p.kabupaten || '-',
    pekerjaan: p.pekerjaan || '-',
    status_perkawinan: p.status_perkawinan || '-',
    alamat_lengkap: parts.join(', ') || '-',
  };
}

export function formatPersonFullAddress(p?: Partial<PersonBasicInfo>): string {
  if (!p) return '-';
  const hasStructured = Boolean(
    p.alamat_jalan || p.rt || p.rw || p.desa_kelurahan || p.kecamatan || p.kabupaten
  );
  if (!hasStructured) {
    return p.alamat?.trim() || '-';
  }
  const rtRwPart =
    p.rt || p.rw
      ? `Rukun Tetangga (RT) ${p.rt || '-'} Rukun Warga (RW) ${p.rw || '-'}`
      : '';
  const desaPart = p.desa_kelurahan ? `Desa ${p.desa_kelurahan}` : '';
  const kecPart = p.kecamatan ? `Kecamatan ${p.kecamatan}` : '';
  const kabPart = p.kabupaten || '';
  return (
    [p.alamat_jalan, rtRwPart, desaPart, kecPart, kabPart]
      .map((s) => (s || '').trim())
      .filter(Boolean)
      .join(', ') ||
    p.alamat?.trim() ||
    '-'
  );
}

export function formatPersonShortAddress(p?: Partial<PersonBasicInfo>): string {
  if (!p) return '-';
  const hasStructured = Boolean(
    p.alamat_jalan || p.rt || p.rw || p.desa_kelurahan || p.kecamatan || p.kabupaten
  );
  if (!hasStructured) {
    return p.alamat?.trim() || '-';
  }
  const rtRwPart = p.rt || p.rw ? `Rt.${p.rt || '-'}/${p.rw || '-'}` : '';
  const dukuhRt = [p.alamat_jalan, rtRwPart].filter(Boolean).join(' ');
  const desaPart = p.desa_kelurahan ? `Ds. ${p.desa_kelurahan}` : '';
  const kecPart = p.kecamatan ? p.kecamatan.replace(/^Kecamatan\s+/i, '') : '';
  const kabPart = p.kabupaten ? p.kabupaten.replace(/^Kabupaten\s+/i, '') : '';
  return (
    [dukuhRt, desaPart, kecPart, kabPart]
      .map((s) => (s || '').trim())
      .filter(Boolean)
      .join(', ') ||
    p.alamat?.trim() ||
    '-'
  );
}

export function resolvePersonDetail(
  name: string,
  list?: PersonBasicInfo[],
  override?: PersonBasicInfo
): {
  nama: string;
  tempat_lahir: string;
  tanggal_lahir: string;
  ttl: string;
  pekerjaan: string;
  nik: string;
  alamat: string;
  alamat_singkat: string;
} {
  const fromList = (list ?? []).find(
    (item) => item.nama.trim().toLowerCase() === name.trim().toLowerCase()
  );
  const merged: PersonBasicInfo | undefined =
    fromList || override
      ? {
          ...(fromList ?? { nama: name || '-' }),
          ...(override ?? {}),
          nama: override?.nama?.trim() || name || fromList?.nama || '-',
        }
      : undefined;

  if (!merged) {
    return {
      nama: name || '-',
      tempat_lahir: '-',
      tanggal_lahir: '-',
      ttl: '-',
      pekerjaan: 'Staf Kantor / Perangkat',
      nik: '-',
      alamat: '-',
      alamat_singkat: '-',
    };
  }

  const tempatLahir = merged.tempat_lahir?.trim() || '-';
  const tglLahirFormatted = formatTanggalDDMMYYYY(merged.tanggal_lahir);
  const ttlParts = [
    tempatLahir !== '-' ? tempatLahir : '',
    tglLahirFormatted !== '-' ? tglLahirFormatted : '',
  ].filter(Boolean);

  return {
    nama: merged.nama || name || '-',
    tempat_lahir: tempatLahir,
    tanggal_lahir: tglLahirFormatted,
    ttl: ttlParts.join(', ') || '-',
    pekerjaan: merged.pekerjaan?.trim() || '-',
    nik: merged.nik?.trim() || '-',
    alamat: formatPersonFullAddress(merged),
    alamat_singkat: formatPersonShortAddress(merged),
  };
}

export function buildTemplateContext(
  record: BerkasRecord,
  settings: OfficeSettings
): Record<string, unknown> {
  const rawPihakPertama = record.para_pihak
    .filter((p) => p.peran === 'Pihak I')
    .map(mapPartyForTemplate);

  const rawPihakKedua = record.para_pihak
    .filter((p) => p.peran === 'Pihak II')
    .map(mapPartyForTemplate);

  const persetujuanPasangan = record.para_pihak
    .filter((p) => p.peran === 'Persetujuan Pasangan')
    .map(mapPartyForTemplate);

  const rawPihakLanjutan = record.para_pihak
    .filter(
      (p) =>
        p.peran.startsWith('Pihak ') &&
        p.peran !== 'Pihak I' &&
        p.peran !== 'Pihak II'
    )
    .map((p) => ({
      ...mapPartyForTemplate(p),
      ...parseAdditionalPihakMeta(p.peran),
    }))
    .sort((a, b) => a.order - b.order);

  // Penomoran penghadap berurutan 1, 2, 3, ... (untuk AJB/Hibah: Pihak I lalu Pihak II lalu lanjutan)
  let seqCounter = 0;
  const pihakPertama = rawPihakPertama.map((p) => {
    seqCounter += 1;
    return { ...p, no_urut_pihak: seqCounter };
  });
  const pihakKedua = rawPihakKedua.map((p) => {
    seqCounter += 1;
    return { ...p, no_urut_pihak: seqCounter };
  });
  const pihakLanjutan = rawPihakLanjutan.map((p) => {
    seqCounter += 1;
    return { ...p, no_urut_pihak: seqCounter };
  });

  // Untuk APHB: semua pemilik menjadi pihak terpisah (Pihak I, Pihak II, ..., Pihak XX+)
  const pihakTerpisah = record.para_pihak
    .filter((p) => p.peran.startsWith('Pihak '))
    .map((p) => ({
      ...mapPartyForTemplate(p),
      ...parseAdditionalPihakMeta(p.peran),
    }))
    .sort((a, b) => a.order - b.order)
    .map((p, idx) => ({ ...p, no_urut_pihak: idx + 1 }));

  const pihakKetiga = pihakLanjutan.filter((p) => p.roman_numeral === 'III');

  const emptyPerson = {
    nama: '-',
    ttl: '-',
    pekerjaan: '-',
    nik: '-',
    alamat_lengkap: '-',
  };
  const p1Primary = pihakPertama[0] ?? emptyPerson;
  const p2Primary = pihakKedua[0] ?? emptyPerson;

  // Persetujuan pasangan vs keterangan tanpa persetujuan
  const needConsentFlag =
    record.akta.butuh_persetujuan_pasangan !== undefined
      ? record.akta.butuh_persetujuan_pasangan
      : persetujuanPasangan.length > 0;
  const adaPersetujuanPasangan = needConsentFlag && persetujuanPasangan.length > 0;
  const tanpaPersetujuan = record.akta.butuh_persetujuan_pasangan === false;
  const keteranganTanpaPersetujuan =
    record.akta.keterangan_tanpa_persetujuan?.trim() ||
    'Menurut keterangan para pihak, dalam melakukan perbuatan hukum ini tanpa memerlukan persetujuan dari siapapun, dikarenakan tanah tersebut berasal dari Warisan.';

  // Jenis akta & pasal tanah pertanian
  const isAjb = record.layanan === 'AJB';
  const isAphb = record.layanan === 'APHB';
  const isHibah = record.layanan === 'Akta Hibah';
  const isLainnya = !isAjb && !isAphb && !isHibah;
  const isNotAphb = !isAphb;
  const isPertanian = record.tanah.status_tanah === 'Pertanian';
  const pasalShift = isPertanian ? 1 : 0;

  const pihak1Selaku = isAjb
    ? 'Penjual'
    : isHibah
    ? 'Pemberi Hibah'
    : 'Pelepas Hak';
  const pihak2Selaku = isAjb
    ? 'Pembeli'
    : isHibah
    ? 'Penerima Hibah'
    : 'Penerima Hak';

  const kalimatPerbuatanPihak = isAjb
    ? 'Pihak Pertama menerangkan dengan ini menjual kepada Pihak Kedua, dan Pihak Kedua menerangkan dengan ini membeli dari Pihak Pertama:'
    : isHibah
    ? 'Pihak Pertama menerangkan dengan ini menghibahkan kepada Pihak Kedua, dan Pihak Kedua menerangkan dengan ini menerima hibah dari Pihak Pertama:'
    : isAphb
    ? 'Para pihak menerangkan bahwa mereka bersama-sama adalah pemegang hak bersama atas bidang tanah sebagaimana diuraikan di bawah ini:'
    : `Pihak Pertama menerangkan dengan ini melakukan perbuatan hukum ${record.layanan} kepada Pihak Kedua, dan Pihak Kedua menerangkan dengan ini menerima dari Pihak Pertama:`;

  const daftarSebutanParaPihak = (() => {
    if (!isAphb) return 'Pihak Pertama dan Pihak Kedua';
    if (pihakTerpisah.length <= 1) return 'Pihak Pertama';
    if (pihakTerpisah.length <= 6) {
      const titles = pihakTerpisah.map((p) => p.peran_title);
      return `${titles.slice(0, -1).join(', ')} dan ${titles[titles.length - 1]}`;
    }
    return `Para Pihak (${pihakTerpisah[0].peran_title} sampai dengan ${pihakTerpisah[pihakTerpisah.length - 1].peran_title})`;
  })();

  // Pihak penerima hak tunggal untuk APHB
  const preferredAphbRole = record.akta.pihak_penerima_aphb || 'Pihak II';
  const aphbWinnerMapped =
    pihakTerpisah.find((p) => p.peran === preferredAphbRole) ??
    pihakTerpisah.find((p) => p.peran === 'Pihak II') ??
    pihakTerpisah[pihakTerpisah.length - 1];
  const aphbWinnerMeta = parseAdditionalPihakMeta(aphbWinnerMapped?.peran || 'Pihak II');
  const aphbPemberiList = aphbWinnerMapped
    ? pihakTerpisah.filter((p) => p.no_urut_pihak !== aphbWinnerMapped.no_urut_pihak)
    : pihakTerpisah;

  // Penerima / Pembeli untuk lampiran (Permohonan, Surat Kuasa, Ket BPHTB, Penguasaan Fisik, Keabsahan Dokumen)
  const penerimaObj = isAphb ? aphbWinnerMapped ?? p2Primary : p2Primary;
  const penerimaRingkas = isAphb
    ? aphbWinnerMapped?.nama || '-'
    : pihakKedua.map((p) => p.nama).join(', ') || '-';

  // Pemilik awal untuk Surat Pernyataan SPPT/PBB
  const pemilikAwalObj = isAphb
    ? aphbPemberiList[0] ?? pihakTerpisah[0] ?? p1Primary
    : p1Primary;
  const pemilikAwalRingkas = isAphb
    ? (aphbPemberiList.length > 0 ? aphbPemberiList : pihakTerpisah).map((p) => p.nama).join(', ') || '-'
    : pihakPertama.map((p) => p.nama).join(', ') || '-';

  // Label & daftar nama kolom Keterangan di Surat Pengantar (Hal. 1 Lampiran)
  const ketPihak1Label = isAjb
    ? 'Penjual'
    : isHibah
    ? 'Pemberi Hibah'
    : 'Pemberi Hak';
  const ketPihak2Label = isAjb
    ? 'Pembeli'
    : isHibah
    ? 'Penerima Hibah'
    : 'Penerima Hak';
  const ketPihak1Names = isAphb
    ? (aphbPemberiList.length > 0 ? aphbPemberiList : pihakTerpisah).map((p) => p.nama).join('\n') || '-'
    : pihakPertama.map((p) => p.nama).join('\n') || '-';
  const ketPihak2Names = isAphb
    ? aphbWinnerMapped?.nama || '-'
    : pihakKedua.map((p) => p.nama).join('\n') || '-';

  // Daftar penandatangan Pemrosesan Data Pribadi (Hal. 9): seluruh pihak termasuk pembeli/penerima, TIDAK termasuk Persetujuan Pasangan
  const dataPribadiSigners = (
    isAphb ? pihakTerpisah : [...pihakPertama, ...pihakKedua, ...pihakLanjutan]
  ).map((p, idx) => ({
    no: idx + 1,
    nama: p.nama || '-',
  }));

  const sertipikatAn =
    record.tanah.pemegang_hak_sertipikat?.map((s) => s.trim()).filter(Boolean).join(', ') ||
    (isAphb ? pihakTerpisah.map((p) => p.nama).join(', ') : pihakPertama.map((p) => p.nama).join(', ')) ||
    '-';

  const perihalPengantar = isAjb
    ? 'Jual beli'
    : isAphb
    ? 'Pembagian Hak Bersama'
    : isHibah
    ? 'Hibah'
    : record.layanan;
  const perihalKalimat = isAjb
    ? 'Jual Beli'
    : isAphb
    ? 'Pembagian Hak Bersama'
    : isHibah
    ? 'Hibah'
    : record.layanan;
  const lampiranAktaLabel = isAjb
    ? 'Akta Jual Beli'
    : isAphb
    ? 'Akta Pembagian Hak Bersama'
    : isHibah
    ? 'Akta Hibah'
    : getJudulAkta(record.layanan);
  const lampiranAktaLabelLower = isAjb
    ? 'Akta jual beli'
    : isAphb
    ? 'Akta pembagian hak bersama'
    : isHibah
    ? 'Akta hibah'
    : `Akta ${record.layanan}`;
  const pernyataanBphtbPerbuatan = isAjb
    ? 'Membeli'
    : isHibah
    ? 'Menerima Hibah'
    : isAphb
    ? 'Menerima Pembagian Hak Bersama atas'
    : 'Menerima peralihan hak atas';

  // Daftar pihak penandatangan Akta
  const signingParties = record.para_pihak
    .filter((p) => p.peran.startsWith('Pihak '))
    .map((p) => ({
      ...parseAdditionalPihakMeta(p.peran),
      nama: p.nama || '-',
    }))
    .sort((a, b) => a.order - b.order);

  // Mode TTD AJB / Hibah (2 kelompok pihak: Pihak Pertama & Pihak Kedua, mendukung multi-pemilik di Pihak I / II)
  const sigModeGrouped2 = isNotAphb && pihakLanjutan.length === 0;
  const pihakPertamaTtd =
    pihakPertama.length > 0
      ? pihakPertama.map((p, idx) => ({
          no: idx + 1,
          show_no: pihakPertama.length > 1,
          nama: p.nama,
        }))
      : [{ no: 1, show_no: false, nama: '-' }];
  const pihakKeduaTtd =
    pihakKedua.length > 0
      ? pihakKedua.map((p, idx) => ({
          no: idx + 1,
          show_no: pihakKedua.length > 1,
          nama: p.nama,
        }))
      : [{ no: 1, show_no: false, nama: '-' }];

  // Mode TTD APHB (atau akta dengan pihak lanjutan terpisah s/d 20+ pihak)
  const useSeparateSigLayout = !sigModeGrouped2;
  const sigMode2 = useSeparateSigLayout && signingParties.length <= 2;
  const sigModeGrid = useSeparateSigLayout && signingParties.length >= 3;
  const sigRows: Array<{
    c1_label: string;
    c1_nama: string;
    c2_label: string;
    c2_nama: string;
    c3_label: string;
    c3_nama: string;
  }> = [];
  if (sigModeGrid) {
    for (let i = 0; i < signingParties.length; i += 3) {
      const p1 = signingParties[i];
      const p2 = signingParties[i + 1];
      const p3 = signingParties[i + 2];
      sigRows.push({
        c1_label: p1?.peran_title ?? '',
        c1_nama: p1?.nama ?? '',
        c2_label: p2?.peran_title ?? '',
        c2_nama: p2?.nama ?? '',
        c3_label: p3?.peran_title ?? '',
        c3_nama: p3?.nama ?? '',
      });
    }
  }

  const sigMode3 = sigModeGrid && signingParties.length === 3;
  const sigMode4Plus = sigModeGrid && signingParties.length >= 4;
  const sigCol1 = signingParties[0] ?? { peran_title: 'Pihak Pertama', nama: '-' };
  const sigCol2 = signingParties[1] ?? { peran_title: 'Pihak Kedua', nama: '-' };
  const sigCol3 = signingParties[2] ?? { peran_title: 'Pihak Ketiga', nama: '-' };
  const sigExtraParties = signingParties.slice(3).map((s) => ({
    peran_title: s.peran_title,
    nama: s.nama,
  }));

  const perbuatanPertanian = isAphb
    ? 'pembagian hak bersama'
    : isHibah
    ? 'hibah'
    : isAjb
    ? 'jual beli'
    : 'peralihan hak';

  const parsedAktaDate = parseFlexibleDate(record.akta.tanggal_akta);
  const formattedDate = formatTanggalLengkapHuruf(record.akta.tanggal_akta);
  const tanggalMasukFormatted = formatTanggalDDMMYYYY(record.tanggal_masuk);
  const activeYear =
    parsedAktaDate?.year ??
    extractMonthYearFromDate(record.tanggal_masuk)?.year ??
    new Date().getFullYear();

  const rawPertanianDate =
    (record.tanah.tanggal_pernyataan_pertanian ?? '').trim() ||
    (parsedAktaDate ? record.akta.tanggal_akta : '');
  const parsedPertanianDate = parseFlexibleDate(rawPertanianDate);
  const tanggalPertanianDisplay = parsedPertanianDate
    ? formatTanggalLengkapHuruf(rawPertanianDate).tanggal_biasa
    : rawPertanianDate || '......................';

  const saksi1 = resolvePersonDetail(
    record.akta.saksi_1,
    settings.profil_saksi,
    record.akta.saksi_1_detail
  );
  const saksi2 = resolvePersonDetail(
    record.akta.saksi_2,
    settings.profil_saksi,
    record.akta.saksi_2_detail
  );
  const kuasaBpn = resolvePersonDetail(
    record.akta.nama_kuasa_bpn,
    settings.profil_kuasa,
    record.akta.kuasa_bpn_detail
  );

  const namaPpats = record.akta.nama_ppats || settings.nama_ppats;
  const jabatanPpats = record.akta.jabatan_ppats || settings.jabatan_ppats;
  const kecShort = (settings.nama_kecamatan || '').replace(/^Kecamatan\s+/i, '').trim();
  const jabatanCamat = /camat/i.test(jabatanPpats) && !/ppats/i.test(jabatanPpats)
    ? jabatanPpats
    : `Camat ${kecShort || settings.nama_kecamatan}`.trim();
  const kotaKabupaten = (record.tanah.kabupaten || settings.nama_kabupaten || '')
    .replace(/^(Kabupaten|Kota)\s+/i, '')
    .trim() || '-';
  const tglAktaDash = parsedAktaDate ? parsedAktaDate.ddMmYyyy : '..........';
  const jenisSertipikat = resolveJenisSertipikat(record.tanah);
  const isSertipikatDigital = jenisSertipikat === 'digital';
  const isSertipikatAnalog = !isSertipikatDigital;
  const nomorHakAnalog = isSertipikatAnalog
    ? formatAnalogNomorHak(record.tanah.nomor_hak_lama)
    : '-';
  const rawSuratUkurTgl = record.tanah.surat_ukur_tanggal?.trim() || '';
  const suratUkurTanggal = /^\d{4}-\d{2}-\d{2}$/.test(rawSuratUkurTgl)
    ? formatTanggalDDMMYYYY(rawSuratUkurTgl)
    : rawSuratUkurTgl || '-';
  const suratUkurNomor = record.tanah.surat_ukur_nomor?.trim() || '-';
  const nibClean = record.tanah.nib?.trim() || '-';
  const nibNomorDisplay = isSertipikatDigital
    ? nibClean
    : `${nomorHakAnalog} (NIB: ${nibClean})`;
  const labelNomorSertipikat = isSertipikatDigital ? 'Nibel' : 'Nomor Hak / NIB';
  const sertipikatRingkasBaris = isSertipikatDigital
    ? `Nibel : ${nibClean}`
    : `Nomor : ${nomorHakAnalog} (NIB : ${nibClean})`;

  return {
    kode_berkas: record.kode_berkas,
    no_urut: record.no_urut,
    tanggal_masuk: tanggalMasukFormatted,
    layanan: record.layanan,
    layanan_upper: record.layanan.toUpperCase(),
    judul_akta: getJudulAkta(record.layanan),
    perihal_pengantar: perihalPengantar,
    perihal_kalimat: perihalKalimat,
    lampiran_akta_label: lampiranAktaLabel,
    lampiran_akta_label_lower: lampiranAktaLabelLower,
    pernyataan_bphtb_perbuatan: pernyataanBphtbPerbuatan,
    is_ajb: isAjb,
    is_aphb: isAphb,
    is_not_aphb: isNotAphb,
    is_hibah: isHibah,
    is_lainnya: isLainnya,
    pihak1_selaku: pihak1Selaku,
    pihak2_selaku: pihak2Selaku,
    kalimat_perbuatan_pihak: kalimatPerbuatanPihak,
    daftar_sebutan_para_pihak: daftarSebutanParaPihak,
    is_pertanian: isPertanian,
    perbuatan_pertanian: perbuatanPertanian,
    tanggal_pernyataan_pertanian: tanggalPertanianDisplay,
    no_pasal_pengukuran: 3 + pasalShift,
    no_pasal_keabsahan: 4 + pasalShift,
    no_pasal_penerimaan_hibah: 4 + pasalShift,
    no_pasal_penjaminan_hibah: 5 + pasalShift,
    no_pasal_domisili: (isHibah ? 6 : 5) + pasalShift,
    no_pasal_biaya: (isHibah ? 7 : 6) + pasalShift,
    status: record.status,
    nomor_akta: record.akta.nomor_akta || '-',
    nomor_pengantar:
      record.akta.nomor_pengantar || `${String(record.no_urut).padStart(2, '0')}/${activeYear}`,
    tanggal_akta: parsedAktaDate ? formattedDate.tanggal_biasa : '......................',
    tanggal_akta_dash: tglAktaDash,
    hari_akta: parsedAktaDate ? formattedDate.hari : '..........',
    tanggal_angka: parsedAktaDate ? formattedDate.tanggal_angka : '....',
    tanggal_akta_huruf: parsedAktaDate ? formattedDate.tanggal_huruf : '....................',
    bulan_akta: parsedAktaDate ? formattedDate.bulan : '............',
    tahun_angka: parsedAktaDate ? formattedDate.tahun_angka : '......',
    tahun_akta_huruf: parsedAktaDate ? formattedDate.tahun_huruf : '........................',
    kalimat_akta: parsedAktaDate
      ? formattedDate.kalimat_akta
      : 'Hari ini, .........., Tanggal .... (....................) ............ Tahun ...... (........................)',
    pihak_pertama: pihakPertama,
    pihak_kedua: pihakKedua,
    pihak_ketiga: pihakKetiga,
    pihak_lanjutan: pihakLanjutan,
    pihak_terpisah: pihakTerpisah,
    persetujuan_pasangan: persetujuanPasangan,
    ada_persetujuan_pasangan: adaPersetujuanPasangan,
    tanpa_persetujuan: tanpaPersetujuan,
    keterangan_tanpa_persetujuan: keteranganTanpaPersetujuan,
    penerima_aphb_label: aphbWinnerMeta.peran_lower,
    penerima_aphb_nama: (aphbWinnerMapped?.nama || '-').toUpperCase(),
    penerima_nama: penerimaObj.nama,
    penerima_ttl: penerimaObj.ttl,
    penerima_pekerjaan: penerimaObj.pekerjaan,
    penerima_nik: penerimaObj.nik,
    penerima_alamat: penerimaObj.alamat_lengkap,
    penerima_ringkas: penerimaRingkas,
    pemilik_awal_nama: pemilikAwalObj.nama,
    pemilik_awal_ttl: pemilikAwalObj.ttl,
    pemilik_awal_pekerjaan: pemilikAwalObj.pekerjaan,
    pemilik_awal_nik: pemilikAwalObj.nik,
    pemilik_awal_alamat: pemilikAwalObj.alamat_lengkap,
    pemilik_awal_ringkas: pemilikAwalRingkas,
    sertipikat_an: sertipikatAn,
    ket_pihak1_label: ketPihak1Label,
    ket_pihak1_names: ketPihak1Names,
    ket_pihak2_label: ketPihak2Label,
    ket_pihak2_names: ketPihak2Names,
    pihak_data_pribadi_ttd: dataPribadiSigners,
    sig_mode_grouped_2: sigModeGrouped2,
    pihak_pertama_ttd: pihakPertamaTtd,
    pihak_kedua_ttd: pihakKeduaTtd,
    sig_mode_2: sigMode2,
    sig_mode_grid: sigModeGrid,
    sig_rows: sigRows,
    sig_mode_3: sigMode3,
    sig_mode_4plus: sigMode4Plus,
    sig_c1_label: sigCol1.peran_title,
    sig_c1_nama: sigCol1.nama,
    sig_c2_label: sigCol2.peran_title,
    sig_c2_nama: sigCol2.nama,
    sig_c3_label: sigCol3.peran_title,
    sig_c3_nama: sigCol3.nama,
    sig_extra_parties: sigExtraParties,
    pihak_pertama_ringkas: pihakPertama.map((p) => p.nama).join(', ') || '-',
    pihak_kedua_ringkas: pihakKedua.map((p) => p.nama).join(', ') || '-',
    pihak_lanjutan_ringkas: pihakLanjutan.map((p) => p.nama).join(', ') || '-',
    p1_nama: p1Primary.nama,
    p1_ttl: p1Primary.ttl,
    p1_pekerjaan: p1Primary.pekerjaan,
    p1_nik: p1Primary.nik,
    p1_alamat: p1Primary.alamat_lengkap,
    p2_nama: p2Primary.nama,
    p2_ttl: p2Primary.ttl,
    p2_pekerjaan: p2Primary.pekerjaan,
    p2_nik: p2Primary.nik,
    p2_alamat: p2Primary.alamat_lengkap,
    pasangan_nama: persetujuanPasangan.map((p) => p.nama).join(', ') || '-',
    is_sertipikat_digital: isSertipikatDigital,
    is_sertipikat_analog: isSertipikatAnalog,
    jenis_sertipikat: jenisSertipikat,
    jenis_hak: record.tanah.jenis_hak || '-',
    nib: nibClean,
    nomor_hak_lama: record.tanah.nomor_hak_lama || '-',
    nomor_hak_analog: nomorHakAnalog,
    surat_ukur_tanggal: suratUkurTanggal,
    surat_ukur_nomor: suratUkurNomor,
    label_nomor_sertipikat: labelNomorSertipikat,
    sertipikat_ringkas_baris: sertipikatRingkasBaris,
    nib_atau_nomor_hak: nibNomorDisplay,
    luas_m2: formatNumberId(record.tanah.luas_m2),
    luas_terbilang: terbilangLuas(record.tanah.luas_m2),
    batas_utara: record.tanah.batas_utara || '-',
    batas_timur: record.tanah.batas_timur || '-',
    batas_selatan: record.tanah.batas_selatan || '-',
    batas_barat: record.tanah.batas_barat || '-',
    penggunaan_tanah: record.tanah.penggunaan_tanah || 'Pekarangan',
    menguasai_sejak_tahun:
      record.tanah.menguasai_sejak_tahun || String(activeYear),
    desa_tanah: record.tanah.desa || '-',
    kecamatan_tanah: record.tanah.kecamatan || '-',
    kabupaten_tanah: record.tanah.kabupaten || '-',
    kota_kabupaten: kotaKabupaten,
    provinsi_tanah:
      record.tanah.provinsi || settings.nama_provinsi || 'Jawa Tengah',
    dasar_perolehan: record.tanah.dasar_perolehan || '-',
    nop: record.pbb.nop || '-',
    tahun_pajak: record.pbb.tahun_pajak || '-',
    njop_bumi_m2_rp: formatRupiah(record.pbb.njop_bumi_m2),
    njop_bangunan_rp: formatRupiah(record.pbb.njop_bangunan),
    njop_total_rp: formatRupiah(record.pajak.njop_total),
    nilai_transaksi_rp: formatRupiah(record.pajak.nilai_kesepakatan),
    nilai_transaksi_terbilang: terbilangRupiah(record.pajak.nilai_kesepakatan),
    npop_rp: formatRupiah(record.pajak.npop),
    npoptkp_rp: formatRupiah(record.pajak.npoptkp),
    bphtb_rp: formatRupiah(record.pajak.bphtb),
    bphtb_terbilang: terbilangRupiah(record.pajak.bphtb),
    pph_rp: formatRupiah(record.pajak.pph),
    pph_terbilang: terbilangRupiah(record.pajak.pph),
    nama_ppats: namaPpats,
    jabatan_ppats: jabatanPpats,
    jabatan_ppats_upper: jabatanPpats.toUpperCase(),
    jabatan_camat: jabatanCamat,
    jabatan_camat_upper: jabatanCamat.toUpperCase(),
    nip_ppats: record.akta.nip_ppats || settings.nip_ppats,
    sk_ppats_nomor: settings.sk_ppats_nomor || '-',
    sk_ppats_tanggal: settings.sk_ppats_tanggal || '-',
    sk_ppats_penerbit:
      settings.sk_ppats_penerbit ||
      'Kakanwil Pertanahan Nasional Propinsi Jawa Tengah',
    nama_kecamatan: settings.nama_kecamatan,
    nama_kecamatan_upper: settings.nama_kecamatan.toUpperCase(),
    nama_kabupaten: settings.nama_kabupaten,
    nama_kabupaten_upper: settings.nama_kabupaten.toUpperCase(),
    nama_provinsi: settings.nama_provinsi || 'Jawa Tengah',
    alamat_kantor: settings.alamat_kantor,
    telepon_kantor: settings.telepon_kantor,
    nama_kuasa_bpn: kuasaBpn.nama,
    kuasa_bpn_nama: kuasaBpn.nama,
    kuasa_bpn_tempat_lahir: kuasaBpn.tempat_lahir,
    kuasa_bpn_tanggal_lahir: kuasaBpn.tanggal_lahir,
    kuasa_bpn_ttl: kuasaBpn.ttl,
    kuasa_bpn_pekerjaan: kuasaBpn.pekerjaan,
    kuasa_bpn_nik: kuasaBpn.nik,
    kuasa_bpn_alamat: kuasaBpn.alamat_singkat || kuasaBpn.alamat,
    kuasa_bpn_alamat_lengkap: kuasaBpn.alamat,
    saksi_1: saksi1.nama,
    saksi_1_nama: saksi1.nama,
    saksi_1_tempat_lahir: saksi1.tempat_lahir,
    saksi_1_tanggal_lahir: saksi1.tanggal_lahir,
    saksi_1_ttl: saksi1.ttl,
    saksi_1_pekerjaan: saksi1.pekerjaan,
    saksi_1_nik: saksi1.nik,
    saksi_1_alamat: saksi1.alamat,
    saksi_2: saksi2.nama,
    saksi_2_nama: saksi2.nama,
    saksi_2_tempat_lahir: saksi2.tempat_lahir,
    saksi_2_tanggal_lahir: saksi2.tanggal_lahir,
    saksi_2_ttl: saksi2.ttl,
    saksi_2_pekerjaan: saksi2.pekerjaan,
    saksi_2_nik: saksi2.nik,
    saksi_2_alamat: saksi2.alamat,
    catatan_loket: record.akta.catatan_loket || '-',
  };
}

export function base64ToUint8Array(base64: string): Uint8Array {
  if (typeof Buffer !== 'undefined') {
    return new Uint8Array(Buffer.from(base64, 'base64'));
  }
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

export function uint8ArrayToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

export function renderDocxWithTemplate(
  templateBytes: Uint8Array,
  context: Record<string, unknown>
): Uint8Array {
  const zip = new PizZip(templateBytes);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' },
    nullGetter: () => '-',
  });
  doc.render(context);
  return doc.getZip().generate({
    type: 'uint8array',
    mimeType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

export async function generateBundleFiles(
  record: BerkasRecord,
  settings: OfficeSettings,
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>
): Promise<Array<{ fileName: string; bytes: Uint8Array; usedCustomTemplate: boolean }>> {
  const context = buildTemplateContext(record, settings);
  const outputs: Array<{ fileName: string; bytes: Uint8Array; usedCustomTemplate: boolean }> = [];

  if (!isAktaService(record.layanan)) {
    const custom = customTemplates['02_lampiran_bundel'];
    if (custom?.dataBase64) {
      const rendered = renderDocxWithTemplate(base64ToUint8Array(custom.dataBase64), context);
      outputs.push({
        fileName: getNonAktaBundleFileName(record.layanan),
        bytes: rendered,
        usedCustomTemplate: true,
      });
    } else {
      const bytes = await buildNonAktaLampiranBuffer(record, settings);
      outputs.push({
        fileName: getNonAktaBundleFileName(record.layanan),
        bytes,
        usedCustomTemplate: false,
      });
    }
    return outputs;
  }

  for (const meta of TEMPLATE_DOCS_META) {
    const custom = customTemplates[meta.key];
    let templateBytes: Uint8Array;
    let usedCustomTemplate = false;

    if (custom?.dataBase64) {
      templateBytes = base64ToUint8Array(custom.dataBase64);
      usedCustomTemplate = true;
    } else {
      templateBytes = await generateDefaultDocxBuffer(meta.key);
    }

    const rendered = renderDocxWithTemplate(templateBytes, context);
    outputs.push({
      fileName: meta.fileName,
      bytes: rendered,
      usedCustomTemplate,
    });
  }

  return outputs;
}

export async function generateAndDownloadBundleZip(
  record: BerkasRecord,
  settings: OfficeSettings,
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>
): Promise<{ zipFileName: string; fileCount: number }> {
  const files = await generateBundleFiles(record, settings, customTemplates);
  const zip = new JSZip();
  for (const f of files) {
    zip.file(f.fileName, f.bytes);
  }
  const blob = await zip.generateAsync({ type: 'blob' });
  const safeLayanan = record.layanan.replace(/[^a-zA-Z0-9]+/g, '_');
  const safeKode = record.kode_berkas.replace(/[^a-zA-Z0-9-]+/g, '_');
  const zipFileName = `Bundel_${safeKode}_${safeLayanan}.zip`;
  saveBlob(blob, zipFileName);
  return { zipFileName, fileCount: files.length };
}

export async function renderSingleDocxBytes(
  record: BerkasRecord,
  settings: OfficeSettings,
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>,
  docKey: TemplateDocKey
): Promise<{ bytes: Uint8Array; fileName: string; usedCustomTemplate: boolean; context: Record<string, unknown> }> {
  const context = buildTemplateContext(record, settings);
  if (!isAktaService(record.layanan)) {
    const custom = customTemplates['02_lampiran_bundel'];
    const usedCustomTemplate = Boolean(custom?.dataBase64);
    const bytes = custom?.dataBase64
      ? renderDocxWithTemplate(base64ToUint8Array(custom.dataBase64), context)
      : await buildNonAktaLampiranBuffer(record, settings);
    return {
      bytes,
      fileName: getNonAktaBundleFileName(record.layanan),
      usedCustomTemplate,
      context,
    };
  }

  const meta = TEMPLATE_DOCS_META.find((m) => m.key === docKey);
  const custom = customTemplates[docKey];
  const usedCustomTemplate = Boolean(custom?.dataBase64);
  const templateBytes = custom?.dataBase64
    ? base64ToUint8Array(custom.dataBase64)
    : await generateDefaultDocxBuffer(docKey);
  const bytes = renderDocxWithTemplate(templateBytes, context);
  return {
    bytes,
    fileName: meta?.fileName ?? `${docKey}.docx`,
    usedCustomTemplate,
    context,
  };
}

export async function generateAndDownloadSingleDocx(
  record: BerkasRecord,
  settings: OfficeSettings,
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>,
  docKey: TemplateDocKey
): Promise<{ fileName: string }> {
  const { bytes, fileName } = await renderSingleDocxBytes(
    record,
    settings,
    customTemplates,
    docKey
  );
  const safeLayanan = record.layanan.replace(/[^a-zA-Z0-9]+/g, '_');
  const safeKode = record.kode_berkas.replace(/[^a-zA-Z0-9-]+/g, '_');
  const outName = `${safeKode}_${safeLayanan}_${fileName}`;
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  saveBlob(blob, outName);
  return { fileName: outName };
}

export async function downloadDefaultTemplateDocx(key: TemplateDocKey, fileName: string): Promise<void> {
  const bytes = await generateDefaultDocxBuffer(key);
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  saveBlob(blob, `Template_Bawaan_${fileName}`);
}
