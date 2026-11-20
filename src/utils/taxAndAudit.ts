import {
  AuditWarning,
  BerkasRecord,
  DataTanah,
  OfficeSettings,
  ParaPihak,
  ServiceType,
  TaxCalculationResult,
} from '../types/ppats';
import { extractMonthYearFromDate, toRomanMonth } from './indonesianFormatter';

export const AKTA_SERVICE_TYPES: ServiceType[] = ['AJB', 'APHB', 'Akta Hibah'];
export const NON_AKTA_SERVICE_TYPES: ServiceType[] = [
  'Turun Waris',
  'Pemecahan',
  'Ukur Ulang',
  'Ralat Data',
];

export function isAktaService(layanan: ServiceType): boolean {
  return layanan === 'AJB' || layanan === 'APHB' || layanan === 'Akta Hibah';
}

export function getDefaultNpoptkp(
  layanan: ServiceType,
  settings?: Pick<OfficeSettings, 'npoptkp_umum' | 'npoptkp_waris'>
): number {
  if (layanan === 'Turun Waris') {
    return settings?.npoptkp_waris ?? 300_000_000;
  }
  return settings?.npoptkp_umum ?? 80_000_000;
}

export interface CalculateTaxInput {
  luas_m2: number;
  harga_per_m2: number;
  njop_bumi_m2: number;
  njop_bangunan: number;
  npoptkp: number;
  layanan: ServiceType;
  nilai_transaksi_override?: number | null;
}

export function calculateLandTaxes(input: CalculateTaxInput): TaxCalculationResult {
  const luas = Math.max(0, Number(input.luas_m2) || 0);
  const hargaM2 = Math.max(0, Number(input.harga_per_m2) || 0);
  const njopBumiM2 = Math.max(0, Number(input.njop_bumi_m2) || 0);
  const njopBangunan = Math.max(0, Number(input.njop_bangunan) || 0);
  const npoptkp = Math.max(0, Number(input.npoptkp) || 0);

  const nilai_kesepakatan =
    input.nilai_transaksi_override !== undefined &&
    input.nilai_transaksi_override !== null &&
    input.nilai_transaksi_override >= 0
      ? Math.round(input.nilai_transaksi_override)
      : Math.round(luas * hargaM2);

  const njop_bumi_total = Math.round(luas * njopBumiM2);
  const njop_total = njop_bumi_total + Math.round(njopBangunan);

  const npop = Math.max(nilai_kesepakatan, njop_total);
  const npopkp = Math.max(0, npop - npoptkp);
  const bphtb = Math.round(npopkp * 0.05);
  const pph = Math.round(npop * 0.025);

  return {
    nilai_kesepakatan,
    njop_bumi_total,
    njop_total,
    npop,
    npoptkp,
    npopkp,
    bphtb,
    pph,
  };
}

export interface TaxBreakdownSummary {
  total_pajak: number;
  dasar_npop_sumber: 'Kesepakatan' | 'NJOP' | 'Setara';
  selisih_kesepakatan_njop: number;
}

export function summarizeTaxBreakdown(
  result: TaxCalculationResult
): TaxBreakdownSummary {
  const total_pajak = Math.max(0, result.bphtb) + Math.max(0, result.pph);
  const selisih_kesepakatan_njop =
    result.nilai_kesepakatan - result.njop_total;
  const dasar_npop_sumber: TaxBreakdownSummary['dasar_npop_sumber'] =
    result.nilai_kesepakatan > result.njop_total
      ? 'Kesepakatan'
      : result.njop_total > result.nilai_kesepakatan
      ? 'NJOP'
      : 'Setara';

  return {
    total_pajak,
    dasar_npop_sumber,
    selisih_kesepakatan_njop,
  };
}

export interface ApplyTaxDraftInput {
  layanan?: ServiceType;
  luas_m2: number;
  njop_bumi_m2: number;
  njop_bangunan: number;
  harga_per_m2: number;
  nilai_transaksi_override: number | null;
  npoptkp: number;
}

export function applyTaxCalculationToRecord(
  record: BerkasRecord,
  input: ApplyTaxDraftInput,
  updatedAt: string = new Date().toISOString()
): BerkasRecord {
  const layanan = input.layanan ?? record.layanan;
  const luas_m2 = Math.max(0, Number(input.luas_m2) || 0);
  const njop_bumi_m2 = Math.max(0, Number(input.njop_bumi_m2) || 0);
  const njop_bangunan = Math.max(0, Number(input.njop_bangunan) || 0);
  const harga_per_m2 = Math.max(0, Number(input.harga_per_m2) || 0);
  const npoptkp = Math.max(0, Number(input.npoptkp) || 0);
  const nilai_transaksi_override =
    input.nilai_transaksi_override !== null &&
    input.nilai_transaksi_override !== undefined
      ? Math.max(0, Number(input.nilai_transaksi_override) || 0)
      : null;

  const pajak = calculateLandTaxes({
    luas_m2,
    harga_per_m2,
    njop_bumi_m2,
    njop_bangunan,
    npoptkp,
    layanan,
    nilai_transaksi_override,
  });

  return {
    ...record,
    layanan,
    tanah: {
      ...record.tanah,
      luas_m2,
    },
    pbb: {
      ...record.pbb,
      njop_bumi_m2,
      njop_bangunan,
    },
    akta: {
      ...record.akta,
      harga_per_m2,
      nilai_transaksi_override,
      npoptkp,
    },
    pajak,
    updated_at: updatedAt,
  };
}

export function runClientAuditWarnings(
  paraPihak: ParaPihak[],
  tanah: DataTanah,
  options?: { butuhPersetujuanPasangan?: boolean; layanan?: ServiceType }
): AuditWarning[] {
  const warnings: AuditWarning[] = [];
  const isAkta = options?.layanan ? isAktaService(options.layanan) : true;

  // 1. Cek NIK ganda
  const nikMap = new Map<string, string[]>();
  for (const p of paraPihak) {
    const cleanNik = (p.nik || '').replace(/\D/g, '');
    if (cleanNik.length > 0) {
      const existing = nikMap.get(cleanNik) ?? [];
      existing.push(p.nama || 'Tanpa Nama');
      nikMap.set(cleanNik, existing);
    }
  }
  for (const [nik, names] of nikMap.entries()) {
    if (names.length > 1) {
      warnings.push({
        id: `dup-nik-${nik}`,
        tingkat: 'tinggi',
        judul: `NIK Ganda Terdeteksi (${nik})`,
        deskripsi: `NIK yang sama muncul pada ${names.length} pihak: ${names.join(', ')}. Pastikan tidak ada duplikasi entri KTP/KK.`,
      });
    }
  }

  // 2. Cek Pihak I berstatus KAWIN memerlukan Persetujuan Pasangan (khusus berkas Akta PPATS)
  const hasSpouseConsent = paraPihak.some((p) => p.peran === 'Persetujuan Pasangan');
  const pihakSatuKawin = paraPihak.filter((p) => {
    const st = (p.status_perkawinan || '').trim().toUpperCase();
    return (
      p.peran === 'Pihak I' &&
      st.includes('KAWIN') &&
      !st.includes('BELUM') &&
      !st.includes('CERAI')
    );
  });

  if (
    isAkta &&
    pihakSatuKawin.length > 0 &&
    !hasSpouseConsent &&
    options?.butuhPersetujuanPasangan !== false
  ) {
    const namesList = pihakSatuKawin
      .map((p) => p.nama.trim() || 'Pihak I (nama belum diisi)')
      .join(', ');
    warnings.push({
      id: 'need-spouse-consent',
      tingkat: 'tinggi',
      judul: 'Memerlukan Persetujuan Pasangan (Suami/Istri)',
      deskripsi: `${namesList} berstatus KAWIN sebagai Pihak I. Tambahkan pasangan sebagai "Persetujuan Pasangan" atau pilih opsi "Tidak Memerlukan Persetujuan" jika tanah berasal dari warisan/harta bawaan.`,
    });
  }

  // 3. Cek nama Pihak I vs pemegang hak sertipikat
  const certHolders = (tanah.pemegang_hak_sertipikat || [])
    .map((n) => n.trim().toUpperCase())
    .filter(Boolean);
  if (certHolders.length > 0 && options?.layanan !== 'Turun Waris') {
    const pihakSatuList = paraPihak.filter((p) => p.peran === 'Pihak I');
    for (const p1 of pihakSatuList) {
      const p1Name = (p1.nama || '').trim().toUpperCase();
      if (p1Name) {
        const matched = certHolders.some(
          (holder) => holder.includes(p1Name) || p1Name.includes(holder)
        );
        if (!matched) {
          warnings.push({
            id: `mismatch-cert-${p1.id}`,
            tingkat: 'sedang',
            judul: `Nama Pihak I (${p1.nama}) Tidak Tertera di Sertipikat`,
            deskripsi: `Pemegang hak pada sertipikat tercatat: ${tanah.pemegang_hak_sertipikat.join(', ')}. Jika terjadi perbedaan penulisan nama atau pewarisan, lengkapi Surat Pernyataan Beda Nama atau SKW.`,
          });
        }
      }
    }
  }

  // 4. Cek selisih wilayah atau RT/RW antar pihak dengan satu alamat keluarga / lokasi tanah
  const rts = new Set(
    paraPihak
      .filter((p) => p.peran === 'Pihak I' || p.peran === 'Persetujuan Pasangan')
      .map((p) => (p.rt_rw || '').trim())
      .filter(Boolean)
  );
  if (rts.size > 1) {
    warnings.push({
      id: 'diff-rtrw-pihak1',
      tingkat: 'info',
      judul: 'Perbedaan RT/RW pada Pihak I & Pasangan',
      deskripsi: `Ditemukan lebih dari satu RT/RW (${Array.from(rts).join(' vs ')}) di kelompok pelepas hak/pasangan. Mohon verifikasi dengan KK terbaru.`,
    });
  }

  return warnings;
}

export function suggestNextAktaNumber(
  records: BerkasRecord[],
  dateInput: string
): { noUrut: number; nomorAkta: string } {
  const parsed = extractMonthYearFromDate(dateInput);
  const year = parsed?.year ?? new Date().getFullYear();
  const month = parsed?.month ?? new Date().getMonth() + 1;

  const recordYear = (r: BerkasRecord) =>
    extractMonthYearFromDate(r.akta?.tanggal_akta || r.tanggal_masuk || '')?.year;

  const maxUrutYear = records
    .filter((r) => recordYear(r) === year)
    .reduce((max, r) => Math.max(max, r.no_urut || 0), 0);

  const aktaCountYear = records.filter(
    (r) => isAktaService(r.layanan) && recordYear(r) === year
  ).length;

  const nextUrut = maxUrutYear + 1;
  const nextAktaSeq = aktaCountYear + 1;
  const paddedAkta = String(nextAktaSeq).padStart(3, '0');
  const roman = toRomanMonth(month);

  return {
    noUrut: nextUrut,
    nomorAkta: `${paddedAkta}/PPATS/${roman}/${year}`,
  };
}
