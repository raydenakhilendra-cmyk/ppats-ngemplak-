import * as XLSX from 'xlsx';
import { BerkasRecord, BerkasStatus, OfficeSettings, ServiceType } from '../types/ppats';
import {
  BULAN_ID,
  extractMonthYearFromDate,
  formatTanggalDDMMYYYY,
} from './indonesianFormatter';
import { isAktaService } from './taxAndAudit';

export function filterRecordsByMonthYear(
  records: BerkasRecord[],
  monthOneBased: number,
  year: number
): BerkasRecord[] {
  return records
    .filter((r) => {
      if (!isAktaService(r.layanan)) return false;
      const my = extractMonthYearFromDate(r.akta?.tanggal_akta || r.tanggal_masuk || '');
      return my !== null && my.year === year && my.month === monthOneBased;
    })
    .sort((a, b) => a.no_urut - b.no_urut);
}

export interface NonAktaFilterOptions {
  layanan?: 'SEMUA' | ServiceType;
  status?: 'SEMUA' | BerkasStatus;
  month?: number | 'SEMUA';
  year?: number | 'SEMUA';
  query?: string;
}

export function filterNonAktaRecords(
  records: BerkasRecord[],
  opts: NonAktaFilterOptions = {}
): BerkasRecord[] {
  const q = (opts.query || '').trim().toLowerCase();
  return records
    .filter((r) => {
      if (isAktaService(r.layanan)) return false;
      if (opts.layanan && opts.layanan !== 'SEMUA') {
        if (opts.layanan === 'Pemecahan') {
          if (r.layanan !== 'Pemecahan' && r.layanan !== 'Pemecahan/Ukur Ulang') return false;
        } else if (r.layanan !== opts.layanan) {
          return false;
        }
      }
      if (opts.status && opts.status !== 'SEMUA' && r.status !== opts.status) {
        return false;
      }
      if (
        (opts.month !== undefined && opts.month !== 'SEMUA') ||
        (opts.year !== undefined && opts.year !== 'SEMUA')
      ) {
        const my = extractMonthYearFromDate(r.akta?.tanggal_akta || r.tanggal_masuk || '');
        if (!my) return false;
        if (opts.year !== undefined && opts.year !== 'SEMUA' && my.year !== opts.year) {
          return false;
        }
        if (opts.month !== undefined && opts.month !== 'SEMUA' && my.month !== opts.month) {
          return false;
        }
      }
      if (!q) return true;
      const partyText = r.para_pihak.map((p) => `${p.nama} ${p.nik}`).join(' ');
      const hay = [
        r.kode_berkas,
        r.layanan,
        r.tanah.nib,
        r.tanah.nomor_hak_lama,
        r.tanah.desa,
        r.tanah.kecamatan,
        r.non_akta?.nama_pewaris || '',
        r.non_akta?.nama_di_sertipikat || '',
        r.non_akta?.nama_di_ktp_kk || '',
        r.akta.nama_kuasa_bpn || '',
        partyText,
      ]
        .join(' ')
        .toLowerCase();
      return hay.includes(q);
    })
    .sort((a, b) => b.no_urut - a.no_urut);
}

export interface NonAktaStatsSummary {
  total: number;
  turunWaris: number;
  pemecahan: number;
  ukurUlang: number;
  ralatData: number;
  pemberkasanAwal: number;
  prosesBpn: number;
  selesai: number;
}

export function summarizeNonAktaStats(records: BerkasRecord[]): NonAktaStatsSummary {
  const nonAkta = records.filter((r) => !isAktaService(r.layanan));
  return {
    total: nonAkta.length,
    turunWaris: nonAkta.filter((r) => r.layanan === 'Turun Waris').length,
    pemecahan: nonAkta.filter(
      (r) => r.layanan === 'Pemecahan' || r.layanan === 'Pemecahan/Ukur Ulang'
    ).length,
    ukurUlang: nonAkta.filter((r) => r.layanan === 'Ukur Ulang').length,
    ralatData: nonAkta.filter((r) => r.layanan === 'Ralat Data').length,
    pemberkasanAwal: nonAkta.filter(
      (r) =>
        r.status === '1. Pemberkasan' ||
        r.status === '2. Validasi Pajak' ||
        r.status === '3. TTD Akta'
    ).length,
    prosesBpn: nonAkta.filter((r) => r.status === '4. Proses BPN').length,
    selesai: nonAkta.filter((r) => r.status === '5. Selesai').length,
  };
}

export function exportMonthlyReportXlsx(params: {
  month: number;
  year: number;
  records: BerkasRecord[];
  settings: OfficeSettings;
}): string {
  const { month, year, records, settings } = params;
  const bulanNama = BULAN_ID[month - 1] ?? String(month);
  const filtered = filterRecordsByMonthYear(records, month, year);

  const rows: (string | number)[][] = [
    [`PEMERINTAH ${settings.nama_kabupaten.toUpperCase()}`],
    [`LAPORAN BULANAN PEMBUATAN AKTA PPATS (${settings.nama_kecamatan.toUpperCase()})`],
    [`Disampaikan Kepada Yth. Kepala Kantor Pertanahan (BPN) & Kepala Bapenda ${settings.nama_kabupaten}`],
    [],
    ['Periode Laporan', `: ${bulanNama} ${year}`],
    ['Nama PPATS / Camat', `: ${settings.nama_ppats}`],
    ['NIP', `: ${settings.nip_ppats}`],
    ['Alamat Kantor', `: ${settings.alamat_kantor}`],
    [],
    [
      'No. Urut',
      'Nomor Akta',
      'Tanggal Akta (DD-MM-YYYY)',
      'Bentuk Perbuatan Hukum',
      'Pihak Pertama (Melepas)',
      'Pihak Kedua (Menerima)',
      'Jenis & Nomor Hak / NIB',
      'Letak Tanah (Desa/Kec)',
      'Luas (m²)',
      'Nilai Transaksi (Rp)',
      'NJOP PBB (Rp)',
      'BPHTB (Rp)',
      'PPh Final (Rp)',
      'Status Proses',
    ],
  ];

  if (filtered.length === 0) {
    rows.push([
      '-',
      'NIHIL',
      '-',
      'Tidak ada akta yang diterbitkan pada bulan ini',
      '-',
      '-',
      '-',
      '-',
      0,
      0,
      0,
      0,
      0,
      'NIHIL',
    ]);
  } else {
    filtered.forEach((rec, idx) => {
      const p1 = rec.para_pihak
        .filter((p) => p.peran === 'Pihak I')
        .map((p) => p.nama)
        .join(', ');
      const p2 = rec.para_pihak
        .filter((p) => p.peran === 'Pihak II')
        .map((p) => p.nama)
        .join(', ');

      rows.push([
        idx + 1,
        rec.akta.nomor_akta || rec.kode_berkas,
        rec.akta.tanggal_akta
          ? formatTanggalDDMMYYYY(rec.akta.tanggal_akta)
          : formatTanggalDDMMYYYY(rec.tanggal_masuk),
        rec.layanan,
        p1 || '-',
        p2 || '-',
        rec.tanah.jenis_sertipikat === 'digital' || !rec.tanah.nomor_hak_lama
          ? `${rec.tanah.jenis_hak} Nibel ${rec.tanah.nib || '-'}`
          : `${rec.tanah.jenis_hak} No. ${rec.tanah.nomor_hak_lama} (NIB: ${rec.tanah.nib || '-'})`,
        `${rec.tanah.desa}, ${rec.tanah.kecamatan}`,
        rec.tanah.luas_m2,
        rec.pajak.nilai_kesepakatan,
        rec.pajak.njop_total,
        rec.pajak.bphtb,
        rec.pajak.pph,
        rec.status,
      ]);
    });
  }

  const totalLuas = filtered.reduce((acc, r) => acc + (r.tanah.luas_m2 || 0), 0);
  const totalNilai = filtered.reduce((acc, r) => acc + (r.pajak.nilai_kesepakatan || 0), 0);
  const totalNjop = filtered.reduce((acc, r) => acc + (r.pajak.njop_total || 0), 0);
  const totalBphtb = filtered.reduce((acc, r) => acc + (r.pajak.bphtb || 0), 0);
  const totalPph = filtered.reduce((acc, r) => acc + (r.pajak.pph || 0), 0);

  rows.push([
    '',
    'JUMLAH TOTAL',
    `${filtered.length} Akta`,
    '',
    '',
    '',
    '',
    '',
    totalLuas,
    totalNilai,
    totalNjop,
    totalBphtb,
    totalPph,
    '',
  ]);

  rows.push([]);
  rows.push([]);
  rows.push(['', '', '', '', '', '', '', '', '', '', `${settings.nama_kecamatan}, Akhir Bulan ${bulanNama} ${year}`]);
  rows.push(['', '', '', '', '', '', '', '', '', '', 'Pejabat Pembuat Akta Tanah Sementara (PPATS)']);
  rows.push(['', '', '', '', '', '', '', '', '', '', settings.jabatan_ppats]);
  rows.push([]);
  rows.push([]);
  rows.push(['', '', '', '', '', '', '', '', '', '', settings.nama_ppats]);
  rows.push(['', '', '', '', '', '', '', '', '', '', `NIP. ${settings.nip_ppats}`]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 13 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 13 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 13 } },
  ];
  ws['!cols'] = [
    { wch: 9 },
    { wch: 22 },
    { wch: 16 },
    { wch: 18 },
    { wch: 24 },
    { wch: 24 },
    { wch: 32 },
    { wch: 24 },
    { wch: 12 },
    { wch: 18 },
    { wch: 18 },
    { wch: 16 },
    { wch: 16 },
    { wch: 18 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Laporan_${bulanNama}_${year}`);

  const fileName = `Laporan_Bulanan_PPATS_${bulanNama}_${year}.xlsx`;
  XLSX.writeFile(wb, fileName);
  return fileName;
}

function describeNonAktaDetail(rec: BerkasRecord): string {
  const na = rec.non_akta;
  if (rec.layanan === 'Turun Waris') {
    return na?.nama_pewaris ? `Pewaris: ${na.nama_pewaris}` : 'Permohonan Turun Waris';
  }
  if (rec.layanan === 'Pemecahan' || rec.layanan === 'Pemecahan/Ukur Ulang') {
    const count = na?.bidang_pecahan?.length || 0;
    return count > 0 ? `Pecah ${count} bidang` : 'Pemecahan Bidang';
  }
  if (rec.layanan === 'Ukur Ulang') {
    return na?.koordinat_gps ? `GPS: ${na.koordinat_gps}` : 'Pengukuran Ulang / Batas';
  }
  if (rec.layanan === 'Ralat Data') {
    if (na?.nama_di_sertipikat || na?.nama_di_ktp_kk) {
      return `${na.nama_di_sertipikat || '-'} -> ${na.nama_di_ktp_kk || '-'}`;
    }
    return 'Ralat Data Sertipikat';
  }
  return '-';
}

export function exportNonAktaRekapXlsx(params: {
  records: BerkasRecord[];
  settings: OfficeSettings;
  periodeLabel?: string;
}): string {
  const { records, settings, periodeLabel = 'Semua Periode' } = params;
  const rows: (string | number)[][] = [
    [`PEMERINTAH ${settings.nama_kabupaten.toUpperCase()}`],
    [`REKAP BERKAS & TRACKING PROGRES LAYANAN NON-AKTA (${settings.nama_kecamatan.toUpperCase()})`],
    [`Catatan: Berkas Non-Akta dikecualikan dari Laporan Bulanan PPATS`],
    [],
    ['Periode / Filter', `: ${periodeLabel}`],
    ['Jumlah Berkas', `: ${records.length} Berkas Non-Akta`],
    [],
    [
      'No.',
      'Kode Berkas',
      'Tgl Masuk (DD-MM-YYYY)',
      'Tgl Surat/Berkas (DD-MM-YYYY)',
      'Jenis Layanan Non-Akta',
      'Pemohon / Ahli Waris',
      'Keterangan Khusus',
      'Kuasa BPN',
      'Hak & NIB/NIBel',
      'Letak Tanah (Desa/Kec)',
      'Luas (m²)',
      'Status Tracking',
    ],
  ];

  if (records.length === 0) {
    rows.push(['-', 'NIHIL', '-', '-', '-', '-', '-', '-', '-', '-', 0, '-']);
  } else {
    records.forEach((rec, idx) => {
      const pemohon =
        rec.para_pihak
          .filter((p) => p.peran.startsWith('Pihak '))
          .map((p) => p.nama)
          .filter(Boolean)
          .join(', ') || '-';
      rows.push([
        idx + 1,
        rec.kode_berkas,
        formatTanggalDDMMYYYY(rec.tanggal_masuk),
        rec.akta.tanggal_akta ? formatTanggalDDMMYYYY(rec.akta.tanggal_akta) : 'Belum bertanggal',
        rec.layanan,
        pemohon,
        describeNonAktaDetail(rec),
        rec.akta.nama_kuasa_bpn || '-',
        rec.tanah.jenis_sertipikat === 'digital' || !rec.tanah.nomor_hak_lama
          ? `${rec.tanah.jenis_hak} Nibel ${rec.tanah.nib || '-'}`
          : `${rec.tanah.jenis_hak} No. ${rec.tanah.nomor_hak_lama} (NIB: ${rec.tanah.nib || '-'})`,
        `${rec.tanah.desa}, ${rec.tanah.kecamatan}`,
        rec.tanah.luas_m2,
        rec.status === '3. TTD Akta' ? '3. TTD / Pengesahan' : rec.status,
      ]);
    });
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 11 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 11 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 11 } },
  ];
  ws['!cols'] = [
    { wch: 6 },
    { wch: 16 },
    { wch: 20 },
    { wch: 24 },
    { wch: 18 },
    { wch: 26 },
    { wch: 28 },
    { wch: 22 },
    { wch: 28 },
    { wch: 24 },
    { wch: 12 },
    { wch: 20 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Rekap_Non_Akta');
  const safePeriod = periodeLabel.replace(/[^a-zA-Z0-9_-]+/g, '_');
  const fileName = `Rekap_Tracking_Non_Akta_${safePeriod}.xlsx`;
  XLSX.writeFile(wb, fileName);
  return fileName;
}

