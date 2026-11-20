import test from 'node:test';
import assert from 'node:assert/strict';
import {
  extractMonthYearFromDate,
  formatRupiah,
  formatTanggalDDMMYYYY,
  formatTanggalLengkapHuruf,
  parseFlexibleDate,
  terbilang,
  terbilangLuas,
  terbilangRupiah,
  toRomanMonth,
} from './indonesianFormatter';

test('terbilang converts numbers accurately in Bahasa Indonesia', () => {
  assert.equal(terbilang(0), 'nol');
  assert.equal(terbilang(11), 'sebelas');
  assert.equal(terbilang(15), 'lima belas');
  assert.equal(terbilang(100), 'seratus');
  assert.equal(terbilang(125), 'seratus dua puluh lima');
  assert.equal(terbilang(1500), 'seribu lima ratus');
  assert.equal(terbilang(80_000_000), 'delapan puluh juta');
  assert.equal(terbilang(300_000_000), 'tiga ratus juta');
  assert.equal(terbilang(1_250_000_000), 'satu miliar dua ratus lima puluh juta');
  assert.equal(terbilang(125.5), 'seratus dua puluh lima koma lima');
});

test('terbilangRupiah and terbilangLuas append proper units', () => {
  assert.equal(terbilangRupiah(80_000_000), 'delapan puluh juta rupiah');
  assert.equal(terbilangLuas(350), 'tiga ratus lima puluh meter persegi');
});

test('formatRupiah formats IDR without decimals', () => {
  const formatted = formatRupiah(150_000_000);
  assert.ok(formatted.includes('150.000.000'));
});

test('formatTanggalLengkapHuruf parses ISO and DD-MM-YYYY dates to Indonesian deed sentence', () => {
  const resIso = formatTanggalLengkapHuruf('2026-11-19');
  assert.equal(resIso.hari, 'Kamis');
  assert.equal(resIso.tanggal_huruf, 'sembilan belas');
  assert.equal(resIso.bulan, 'November');
  assert.equal(resIso.tahun_huruf, 'dua ribu dua puluh enam');
  assert.equal(resIso.tanggal_dash, '19-11-2026');
  assert.equal(
    resIso.kalimat_akta,
    'Hari ini, Kamis tanggal sembilan belas bulan November tahun dua ribu dua puluh enam (19-11-2026)'
  );

  const resDmy = formatTanggalLengkapHuruf('19-11-2026');
  assert.equal(resDmy.kalimat_akta, resIso.kalimat_akta);

  const resEmpty = formatTanggalLengkapHuruf('');
  assert.equal(resEmpty.kalimat_akta, '-');
  assert.equal(resEmpty.tanggal_dash, '-');
});

test('parseFlexibleDate, extractMonthYearFromDate, and formatTanggalDDMMYYYY support DD-MM-YYYY and legacy ISO', () => {
  assert.deepEqual(parseFlexibleDate('05-03-2026'), {
    year: 2026,
    month: 3,
    day: 5,
    iso: '2026-03-05',
    ddMmYyyy: '05-03-2026',
  });
  assert.equal(parseFlexibleDate('31-02-2026'), null);
  assert.equal(formatTanggalDDMMYYYY('2026-11-04'), '04-11-2026');
  assert.equal(formatTanggalDDMMYYYY('04-11-2026'), '04-11-2026');
  assert.equal(formatTanggalDDMMYYYY('19/10/2011'), '19-10-2011');
  assert.equal(formatTanggalDDMMYYYY(''), '-');

  assert.deepEqual(extractMonthYearFromDate('20-11-2026'), { year: 2026, month: 11 });
  assert.deepEqual(extractMonthYearFromDate('2026-11-20'), { year: 2026, month: 11 });
  assert.equal(extractMonthYearFromDate(''), null);
});

test('toRomanMonth returns Roman numerals for months', () => {
  assert.equal(toRomanMonth(1), 'I');
  assert.equal(toRomanMonth(11), 'XI');
  assert.equal(toRomanMonth(12), 'XII');
});
