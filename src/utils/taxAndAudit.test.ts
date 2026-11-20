import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyTaxCalculationToRecord,
  calculateLandTaxes,
  getDefaultNpoptkp,
  isAktaService,
  runClientAuditWarnings,
  suggestNextAktaNumber,
  summarizeTaxBreakdown,
} from './taxAndAudit';
import {
  filterNonAktaRecords,
  filterRecordsByMonthYear,
  summarizeNonAktaStats,
} from './reportGenerator';
import { BerkasRecord, DataTanah, ParaPihak } from '../types/ppats';

test('getDefaultNpoptkp uses 80M for AJB/APHB/Hibah and 300M for Turun Waris', () => {
  assert.equal(getDefaultNpoptkp('AJB'), 80_000_000);
  assert.equal(getDefaultNpoptkp('APHB'), 80_000_000);
  assert.equal(getDefaultNpoptkp('Akta Hibah'), 80_000_000);
  assert.equal(getDefaultNpoptkp('Turun Waris'), 300_000_000);
});

test('calculateLandTaxes computes NPOP, BPHTB 5%, and PPh 2.5% accurately', () => {
  const result = calculateLandTaxes({
    luas_m2: 200,
    harga_per_m2: 1_000_000,
    njop_bumi_m2: 800_000,
    njop_bangunan: 0,
    npoptkp: 80_000_000,
    layanan: 'AJB',
  });

  assert.equal(result.nilai_kesepakatan, 200_000_000);
  assert.equal(result.njop_total, 160_000_000);
  assert.equal(result.npop, 200_000_000);
  assert.equal(result.npopkp, 120_000_000);
  assert.equal(result.bphtb, 6_000_000);
  assert.equal(result.pph, 5_000_000);

  const summary = summarizeTaxBreakdown(result);
  assert.equal(summary.total_pajak, 11_000_000);
  assert.equal(summary.dasar_npop_sumber, 'Kesepakatan');
  assert.equal(summary.selisih_kesepakatan_njop, 40_000_000);
});

test('calculateLandTaxes picks NJOP if higher than agreed price and clamps BPHTB to zero', () => {
  const result = calculateLandTaxes({
    luas_m2: 100,
    harga_per_m2: 500_000,
    njop_bumi_m2: 700_000,
    njop_bangunan: 0,
    npoptkp: 80_000_000,
    layanan: 'AJB',
  });

  assert.equal(result.npop, 70_000_000);
  assert.equal(result.npopkp, 0);
  assert.equal(result.bphtb, 0);
  assert.equal(result.pph, 1_750_000);

  const summary = summarizeTaxBreakdown(result);
  assert.equal(summary.total_pajak, 1_750_000);
  assert.equal(summary.dasar_npop_sumber, 'NJOP');
});

test('applyTaxCalculationToRecord immutably updates land/PBB/akta tax values and recalculates BPHTB/PPh', () => {
  const baseRecord = {
    id: 'rec-1',
    no_urut: 1,
    kode_berkas: 'REG-2026-001',
    tanggal_masuk: '2026-11-01',
    layanan: 'AJB',
    status: '1. Pemberkasan',
    para_pihak: [],
    tanah: {
      jenis_hak: 'Hak Milik',
      nib: '11.17.0001',
      nomor_hak_lama: '',
      luas_m2: 200,
      desa: 'Sindon',
      kecamatan: 'Ngemplak',
      kabupaten: 'Boyolali',
      dasar_perolehan: 'Warisan',
      pemegang_hak_sertipikat: [],
    },
    pbb: {
      nop: '33.09.001',
      tahun_pajak: '2026',
      njop_bumi_m2: 500_000,
      njop_bangunan: 0,
    },
    audit_warnings: [],
    akta: {
      nomor_akta: '001/PPATS/XI/2026',
      tanggal_akta: '2026-11-01',
      nama_ppats: 'MUCHSONI, S.Sos',
      jabatan_ppats: 'PPATS Camat Ngemplak',
      nip_ppats: '-',
      harga_per_m2: 500_000,
      nilai_transaksi_override: null,
      npoptkp: 80_000_000,
      nama_kuasa_bpn: 'SUWARDJO',
      saksi_1: 'SALIMIN',
      saksi_2: 'SUYANTO',
      catatan_loket: '',
    },
    pajak: {
      nilai_kesepakatan: 100_000_000,
      njop_bumi_total: 100_000_000,
      njop_total: 100_000_000,
      npop: 100_000_000,
      npoptkp: 80_000_000,
      npopkp: 20_000_000,
      bphtb: 1_000_000,
      pph: 2_500_000,
    },
    updated_at: '2026-11-01T00:00:00.000Z',
  } as unknown as BerkasRecord;

  const updated = applyTaxCalculationToRecord(
    baseRecord,
    {
      luas_m2: 250,
      njop_bumi_m2: 600_000,
      njop_bangunan: 50_000_000,
      harga_per_m2: 0,
      nilai_transaksi_override: 300_000_000,
      npoptkp: 80_000_000,
    },
    '2026-11-20T05:00:00.000Z'
  );

  assert.equal(baseRecord.pajak.bphtb, 1_000_000);
  assert.equal(updated.tanah.luas_m2, 250);
  assert.equal(updated.pbb.njop_bumi_m2, 600_000);
  assert.equal(updated.pbb.njop_bangunan, 50_000_000);
  assert.equal(updated.pajak.njop_total, 200_000_000);
  assert.equal(updated.pajak.nilai_kesepakatan, 300_000_000);
  assert.equal(updated.pajak.npop, 300_000_000);
  assert.equal(updated.pajak.npopkp, 220_000_000);
  assert.equal(updated.pajak.bphtb, 11_000_000);
  assert.equal(updated.pajak.pph, 7_500_000);
  assert.equal(updated.updated_at, '2026-11-20T05:00:00.000Z');
});

test('runClientAuditWarnings flags duplicate NIK, married seller without spouse consent, and cert mismatch, while skipping spouse/cert noise for Non-Akta', () => {
  const pihak: ParaPihak[] = [
    {
      id: '1',
      peran: 'Pihak I',
      nik: '3301010101800001',
      nama: 'Budi Santoso',
      tempat_lahir: 'Semarang',
      tanggal_lahir: '1980-01-01',
      jenis_kelamin: 'Laki-laki',
      alamat_jalan: 'Dusun Krajan',
      rt_rw: '001/002',
      desa_kelurahan: 'Sukamaju',
      kecamatan: 'Bergas',
      kabupaten: 'Semarang',
      pekerjaan: 'Wiraswasta',
      status_perkawinan: 'KAWIN',
    },
    {
      id: '2',
      peran: 'Pihak II',
      nik: '3301010101800001',
      nama: 'Andi Wijaya',
      tempat_lahir: 'Semarang',
      tanggal_lahir: '1985-05-05',
      jenis_kelamin: 'Laki-laki',
      alamat_jalan: 'Jl. Melati',
      rt_rw: '003/001',
      desa_kelurahan: 'Bergas Kidul',
      kecamatan: 'Bergas',
      kabupaten: 'Semarang',
      pekerjaan: 'Pegawai Swasta',
      status_perkawinan: 'BELUM KAWIN',
    },
  ];

  const tanah: DataTanah = {
    jenis_hak: 'Hak Milik',
    nib: '11.02.03.04.00123',
    nomor_hak_lama: 'SHM No. 456',
    luas_m2: 200,
    desa: 'Sukamaju',
    kecamatan: 'Bergas',
    kabupaten: 'Semarang',
    dasar_perolehan: 'Jual Beli Tahun 2010',
    pemegang_hak_sertipikat: ['Slamet Riyadi'],
  };

  const warnings = runClientAuditWarnings(pihak, tanah);
  assert.ok(warnings.some((w) => w.id.startsWith('dup-nik-')));
  assert.ok(warnings.some((w) => w.id === 'need-spouse-consent'));
  assert.ok(warnings.some((w) => w.id.startsWith('mismatch-cert-')));

  const warningsNoConsentNeeded = runClientAuditWarnings(pihak, tanah, {
    butuhPersetujuanPasangan: false,
  });
  assert.ok(!warningsNoConsentNeeded.some((w) => w.id === 'need-spouse-consent'));

  const turunWarisWarnings = runClientAuditWarnings(pihak, tanah, {
    layanan: 'Turun Waris',
  });
  assert.ok(!turunWarisWarnings.some((w) => w.id === 'need-spouse-consent'));
  assert.ok(!turunWarisWarnings.some((w) => w.id.startsWith('mismatch-cert-')));
});

test('isAktaService & filterRecordsByMonthYear strictly exclude Non-Akta records from Laporan Bulanan PPATS and Akta numbering', () => {
  assert.equal(isAktaService('AJB'), true);
  assert.equal(isAktaService('APHB'), true);
  assert.equal(isAktaService('Akta Hibah'), true);
  assert.equal(isAktaService('Turun Waris'), false);
  assert.equal(isAktaService('Pemecahan'), false);
  assert.equal(isAktaService('Ukur Ulang'), false);
  assert.equal(isAktaService('Ralat Data'), false);
  assert.equal(isAktaService('Pemecahan/Ukur Ulang'), false);

  const dummyRecords = [
    { id: '1', no_urut: 1, layanan: 'AJB', akta: { tanggal_akta: '2026-11-05' }, tanggal_masuk: '2026-11-05' },
    { id: '2', no_urut: 2, layanan: 'Turun Waris', akta: { tanggal_akta: '2026-11-08' }, tanggal_masuk: '2026-11-08' },
    { id: '3', no_urut: 3, layanan: 'Pemecahan', akta: { tanggal_akta: '2026-11-10' }, tanggal_masuk: '2026-11-10' },
    { id: '4', no_urut: 4, layanan: 'Ukur Ulang', akta: { tanggal_akta: '2026-11-12' }, tanggal_masuk: '2026-11-12' },
    { id: '5', no_urut: 5, layanan: 'Ralat Data', akta: { tanggal_akta: '2026-11-14' }, tanggal_masuk: '2026-11-14' },
    { id: '6', no_urut: 6, layanan: 'APHB', akta: { tanggal_akta: '2026-11-18' }, tanggal_masuk: '2026-11-18' },
  ] as unknown as BerkasRecord[];

  const monthlyAkta = filterRecordsByMonthYear(dummyRecords, 11, 2026);
  assert.deepEqual(
    monthlyAkta.map((r) => r.id),
    ['1', '6']
  );

  // suggestNextAktaNumber increments noUrut across all records (7) while nomorAkta follows Akta count (2 + 1 = 003)
  const suggestion = suggestNextAktaNumber(dummyRecords, '20-11-2026');
  assert.equal(suggestion.noUrut, 7);
  assert.equal(suggestion.nomorAkta, '003/PPATS/XI/2026');
});

test('filterNonAktaRecords and summarizeNonAktaStats manage Non-Akta rekap & tracking independently from Laporan Bulanan', () => {
  const sampleRecords = [
    {
      id: 'a1',
      no_urut: 1,
      kode_berkas: 'REG-2026-001',
      layanan: 'AJB',
      status: '5. Selesai',
      tanggal_masuk: '05-11-2026',
      akta: { tanggal_akta: '05-11-2026', nama_kuasa_bpn: 'SUWARDJO' },
      tanah: { nib: '11.17.01', nomor_hak_lama: '100', desa: 'Sindon', kecamatan: 'Ngemplak' },
      para_pihak: [{ peran: 'Pihak I', nama: 'Budi', nik: '3309' }],
    },
    {
      id: 'na1',
      no_urut: 2,
      kode_berkas: 'REG-2026-002',
      layanan: 'Turun Waris',
      status: '4. Proses BPN',
      tanggal_masuk: '08-11-2026',
      akta: { tanggal_akta: '', nama_kuasa_bpn: 'SUWARDJO' },
      tanah: { nib: '11.17.02', nomor_hak_lama: '', desa: 'Ngesrep', kecamatan: 'Ngemplak' },
      para_pihak: [{ peran: 'Pihak I', nama: 'Siti Aminah', nik: '33091' }],
      non_akta: { nama_pewaris: 'HADI SUWARNO' },
    },
    {
      id: 'na2',
      no_urut: 3,
      kode_berkas: 'REG-2026-003',
      layanan: 'Pemecahan',
      status: '1. Pemberkasan',
      tanggal_masuk: '10-11-2026',
      akta: { tanggal_akta: '12-11-2026', nama_kuasa_bpn: 'SUWARDJO' },
      tanah: { nib: '11.17.03', nomor_hak_lama: '220', desa: 'Gagaksipat', kecamatan: 'Ngemplak' },
      para_pihak: [{ peran: 'Pihak I', nama: 'Parmin', nik: '33092' }],
    },
    {
      id: 'na3',
      no_urut: 4,
      kode_berkas: 'REG-2026-004',
      layanan: 'Ralat Data',
      status: '5. Selesai',
      tanggal_masuk: '14-12-2026',
      akta: { tanggal_akta: '14-12-2026', nama_kuasa_bpn: 'SUWARDJO' },
      tanah: { nib: '11.17.04', nomor_hak_lama: '5755', desa: 'Kismoyoso', kecamatan: 'Ngemplak' },
      para_pihak: [{ peran: 'Pihak I', nama: 'Sarno', nik: '33093' }],
    },
  ] as unknown as BerkasRecord[];

  const allNonAkta = filterNonAktaRecords(sampleRecords);
  assert.deepEqual(
    allNonAkta.map((r) => r.id),
    ['na3', 'na2', 'na1']
  );

  // Supports DD-MM-YYYY month/year filtering even when akta.tanggal_akta is blank (uses tanggal_masuk)
  const novNonAkta = filterNonAktaRecords(sampleRecords, { month: 11, year: 2026 });
  assert.deepEqual(
    novNonAkta.map((r) => r.id),
    ['na2', 'na1']
  );

  const warisSearch = filterNonAktaRecords(sampleRecords, { query: 'hadi suwarno' });
  assert.deepEqual(
    warisSearch.map((r) => r.id),
    ['na1']
  );

  const stats = summarizeNonAktaStats(sampleRecords);
  assert.equal(stats.total, 3);
  assert.equal(stats.turunWaris, 1);
  assert.equal(stats.pemecahan, 1);
  assert.equal(stats.ralatData, 1);
  assert.equal(stats.pemberkasanAwal, 1);
  assert.equal(stats.prosesBpn, 1);
  assert.equal(stats.selesai, 1);

  // Confirm Laporan Bulanan excludes all Non-Akta even with DD-MM-YYYY dates
  const novAktaOnly = filterRecordsByMonthYear(sampleRecords, 11, 2026);
  assert.deepEqual(
    novAktaOnly.map((r) => r.id),
    ['a1']
  );
});

