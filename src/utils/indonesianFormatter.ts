const SATUAN = [
  '',
  'satu',
  'dua',
  'tiga',
  'empat',
  'lima',
  'enam',
  'tujuh',
  'delapan',
  'sembilan',
  'sepuluh',
  'sebelas',
];

function terbilangInteger(n: number): string {
  const num = Math.floor(Math.abs(n));
  if (num === 0) return '';
  if (num < 12) return SATUAN[num];
  if (num < 20) return `${terbilangInteger(num - 10)} belas`;
  if (num < 100) {
    const tens = Math.floor(num / 10);
    const rest = num % 10;
    return `${SATUAN[tens]} puluh${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  if (num < 200) {
    const rest = num - 100;
    return `seratus${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  if (num < 1000) {
    const hundreds = Math.floor(num / 100);
    const rest = num % 100;
    return `${SATUAN[hundreds]} ratus${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  if (num < 2000) {
    const rest = num - 1000;
    return `seribu${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  if (num < 1_000_000) {
    const thousands = Math.floor(num / 1000);
    const rest = num % 1000;
    return `${terbilangInteger(thousands)} ribu${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  if (num < 1_000_000_000) {
    const millions = Math.floor(num / 1_000_000);
    const rest = num % 1_000_000;
    return `${terbilangInteger(millions)} juta${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  if (num < 1_000_000_000_000) {
    const billions = Math.floor(num / 1_000_000_000);
    const rest = num % 1_000_000_000;
    return `${terbilangInteger(billions)} miliar${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
  }
  const trillions = Math.floor(num / 1_000_000_000_000);
  const rest = num % 1_000_000_000_000;
  return `${terbilangInteger(trillions)} triliun${rest > 0 ? ' ' + terbilangInteger(rest) : ''}`;
}

/**
 * Mengonversi angka (termasuk desimal dengan koma) ke teks terbilang Bahasa Indonesia.
 */
export function terbilang(nilai: number): string {
  if (!Number.isFinite(nilai)) return 'nol';
  if (nilai === 0) return 'nol';

  const prefix = nilai < 0 ? 'minus ' : '';
  const absVal = Math.abs(nilai);
  const intPart = Math.floor(absVal);
  const intText = intPart === 0 ? 'nol' : terbilangInteger(intPart);

  const str = absVal.toString();
  const dotIdx = str.indexOf('.');
  if (dotIdx === -1) {
    return `${prefix}${intText}`.replace(/\s+/g, ' ').trim();
  }

  const decimalDigits = str.slice(dotIdx + 1).replace(/0+$/, '');
  if (!decimalDigits) {
    return `${prefix}${intText}`.replace(/\s+/g, ' ').trim();
  }

  const decText = decimalDigits
    .split('')
    .map((digit) => (digit === '0' ? 'nol' : SATUAN[Number(digit)]))
    .join(' ');

  return `${prefix}${intText} koma ${decText}`.replace(/\s+/g, ' ').trim();
}

export function terbilangRupiah(nilai: number): string {
  const rounded = Math.round(nilai);
  return `${terbilang(rounded)} rupiah`;
}

export function terbilangLuas(luasM2: number): string {
  return `${terbilang(luasM2)} meter persegi`;
}

export function formatRupiah(nilai: number): string {
  const safe = Number.isFinite(nilai) ? Math.round(nilai) : 0;
  return new Intl.NumberFormat('id-ID', {
    style: 'currency',
    currency: 'IDR',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(safe);
}

export function formatNumberId(nilai: number): string {
  const safe = Number.isFinite(nilai) ? nilai : 0;
  return new Intl.NumberFormat('id-ID').format(safe);
}

const HARI_ID = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

export const BULAN_ID = [
  'Januari',
  'Februari',
  'Maret',
  'April',
  'Mei',
  'Juni',
  'Juli',
  'Agustus',
  'September',
  'Oktober',
  'November',
  'Desember',
];

export interface ParsedCalendarDate {
  year: number;
  month: number; // 1-12
  day: number; // 1-31
  iso: string; // YYYY-MM-DD
  ddMmYyyy: string; // DD-MM-YYYY
}

/**
 * Mengurai tanggal dari format DD-MM-YYYY, DD/MM/YYYY, atau YYYY-MM-DD.
 * Mengembalikan null jika kosong atau tidak valid secara kalender.
 */
export function parseFlexibleDate(dateStr?: string): ParsedCalendarDate | null {
  const raw = (dateStr || '').trim();
  if (!raw) return null;

  let year = 0;
  let month = 0;
  let day = 0;

  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(raw);
  if (isoMatch) {
    year = Number(isoMatch[1]);
    month = Number(isoMatch[2]);
    day = Number(isoMatch[3]);
  } else {
    const dmyMatch = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(raw);
    if (!dmyMatch) return null;
    day = Number(dmyMatch[1]);
    month = Number(dmyMatch[2]);
    year = Number(dmyMatch[3]);
  }

  if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }

  const utcDate = new Date(Date.UTC(year, month - 1, day));
  if (
    utcDate.getUTCFullYear() !== year ||
    utcDate.getUTCMonth() !== month - 1 ||
    utcDate.getUTCDate() !== day
  ) {
    return null;
  }

  const dd = String(day).padStart(2, '0');
  const mm = String(month).padStart(2, '0');
  const yyyy = String(year);

  return {
    year,
    month,
    day,
    iso: `${yyyy}-${mm}-${dd}`,
    ddMmYyyy: `${dd}-${mm}-${yyyy}`,
  };
}

export function extractMonthYearFromDate(
  dateStr?: string
): { year: number; month: number } | null {
  const parsed = parseFlexibleDate(dateStr);
  if (parsed) {
    return { year: parsed.year, month: parsed.month };
  }
  const raw = (dateStr || '').trim();
  const ym = /^(\d{4})-(\d{2})/.exec(raw);
  if (ym) {
    return { year: Number(ym[1]), month: Number(ym[2]) };
  }
  const my = /^(\d{2})[-/](\d{4})$/.exec(raw);
  if (my) {
    return { year: Number(my[2]), month: Number(my[1]) };
  }
  return null;
}

export function normalizeDateToIso(dateStr?: string): string {
  return parseFlexibleDate(dateStr)?.iso ?? '';
}

export function todayDdMmYyyy(): string {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = String(d.getFullYear());
  return `${dd}-${mm}-${yyyy}`;
}

export interface FormattedIndonesianDate {
  hari: string;
  tanggal_angka: string;
  tanggal_huruf: string;
  bulan: string;
  tahun_angka: string;
  tahun_huruf: string;
  tanggal_biasa: string;
  tanggal_dash: string;
  kalimat_akta: string;
}

/**
 * Mengonversi tanggal DD-MM-YYYY atau YYYY-MM-DD ke format tanggal akta Bahasa Indonesia lengkap.
 */
export function formatTanggalLengkapHuruf(dateStr?: string): FormattedIndonesianDate {
  const parsed = parseFlexibleDate(dateStr);
  if (!parsed) {
    return {
      hari: '-',
      tanggal_angka: '-',
      tanggal_huruf: '-',
      bulan: '-',
      tahun_angka: '-',
      tahun_huruf: '-',
      tanggal_biasa: '-',
      tanggal_dash: '-',
      kalimat_akta: '-',
    };
  }

  const { year, month, day, ddMmYyyy } = parsed;
  const monthIdx = month - 1;
  const utcDate = new Date(Date.UTC(year, monthIdx, day));
  const hari = HARI_ID[utcDate.getUTCDay()] ?? '-';
  const bulan = BULAN_ID[monthIdx] ?? '-';
  const tanggalHuruf = terbilang(day);
  const tahunHuruf = terbilang(year);
  const dd = String(day).padStart(2, '0');

  return {
    hari,
    tanggal_angka: dd,
    tanggal_huruf: tanggalHuruf,
    bulan,
    tahun_angka: String(year),
    tahun_huruf: tahunHuruf,
    tanggal_biasa: `${day} ${bulan} ${year}`,
    tanggal_dash: ddMmYyyy,
    kalimat_akta: `Hari ini, ${hari} tanggal ${tanggalHuruf} bulan ${bulan} tahun ${tahunHuruf} (${ddMmYyyy})`,
  };
}

export function toRomanMonth(monthOneBased: number): string {
  const romans = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
  return romans[monthOneBased - 1] ?? 'I';
}

export function formatTanggalDDMMYYYY(dateStr?: string): string {
  const raw = (dateStr || '').trim();
  if (!raw) return '-';
  const parsed = parseFlexibleDate(raw);
  if (parsed) {
    return parsed.ddMmYyyy;
  }
  return raw;
}

