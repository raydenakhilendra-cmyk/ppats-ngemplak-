import React, { useMemo, useState } from 'react';
import { ClipboardList, FileSpreadsheet, Printer } from 'lucide-react';
import { BerkasRecord, OfficeSettings } from '../types/ppats';
import {
  BULAN_ID,
  extractMonthYearFromDate,
  formatNumberId,
  formatRupiah,
  formatTanggalDDMMYYYY,
} from '../utils/indonesianFormatter';
import {
  exportMonthlyReportXlsx,
  filterRecordsByMonthYear,
} from '../utils/reportGenerator';

interface ReportModuleProps {
  records: BerkasRecord[];
  settings: OfficeSettings;
  onOpenNonAktaRekap?: () => void;
  onNotify: (msg: string, type?: 'success' | 'error') => void;
}

export const ReportModule: React.FC<ReportModuleProps> = ({
  records,
  settings,
  onOpenNonAktaRekap,
  onNotify,
}) => {
  const defaultPeriod = useMemo(() => {
    const now = new Date();
    return {
      month: now.getMonth() + 1,
      year: now.getFullYear(),
    };
  }, []);

  const [selectedMonth, setSelectedMonth] = useState<number>(defaultPeriod.month);
  const [selectedYear, setSelectedYear] = useState<number>(defaultPeriod.year);

  const yearOptions = useMemo(() => {
    const set = new Set<number>([defaultPeriod.year - 1, defaultPeriod.year, defaultPeriod.year + 1]);
    for (const r of records) {
      const my = extractMonthYearFromDate(r.akta.tanggal_akta || r.tanggal_masuk);
      if (my && my.year > 2000) set.add(my.year);
    }
    return Array.from(set).sort((a, b) => b - a);
  }, [records, defaultPeriod.year]);

  const filtered = useMemo(
    () => filterRecordsByMonthYear(records, selectedMonth, selectedYear),
    [records, selectedMonth, selectedYear]
  );

  const totals = useMemo(() => {
    return {
      luas: filtered.reduce((s, r) => s + (r.tanah.luas_m2 || 0), 0),
      nilai: filtered.reduce((s, r) => s + (r.pajak.nilai_kesepakatan || 0), 0),
      njop: filtered.reduce((s, r) => s + (r.pajak.njop_total || 0), 0),
      bphtb: filtered.reduce((s, r) => s + (r.pajak.bphtb || 0), 0),
      pph: filtered.reduce((s, r) => s + (r.pajak.pph || 0), 0),
    };
  }, [filtered]);

  const bulanLabel = BULAN_ID[selectedMonth - 1] ?? String(selectedMonth);

  const handleExportExcel = () => {
    try {
      const fileName = exportMonthlyReportXlsx({
        month: selectedMonth,
        year: selectedYear,
        records,
        settings,
      });
      onNotify(`Berhasil mengekspor ${fileName}`, 'success');
    } catch (err) {
      onNotify(
        err instanceof Error ? err.message : 'Gagal mengekspor laporan Excel.',
        'error'
      );
    }
  };

  return (
    <div className="space-y-6">
      {/* Header & Controls */}
      <div className="no-print space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
          <div>
            <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
              Laporan Bulanan PPATS
            </h1>
            <p className="text-sm text-slate-600 mt-1">
              Format resmi pelaporan pembuatan akta bulanan untuk Kepala Kantor Pertanahan (BPN) dan Bapenda. Hanya mencakup Akta PPATS (AJB, APHB, Akta Hibah); berkas Non-Akta (Turun Waris, Pemecahan, Ukur Ulang, Ralat Data) otomatis dikecualikan.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white text-sm font-medium transition-colors whitespace-nowrap"
            >
              <FileSpreadsheet className="w-4 h-4" />
              <span>Export Laporan Bulanan (.XLSX)</span>
            </button>

            <button
              type="button"
              onClick={() => window.print()}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium transition-colors whitespace-nowrap"
            >
              <Printer className="w-4 h-4" />
              <span>Cetak / Simpan PDF</span>
            </button>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl px-4 py-3 flex flex-wrap items-center gap-4">
          <span className="text-xs font-semibold text-slate-600">Periode Pelaporan:</span>
          <div className="flex items-center gap-2">
            <label htmlFor="report-month" className="text-xs font-medium text-slate-500">
              Bulan
            </label>
            <select
              id="report-month"
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(Number(e.target.value))}
              className="text-sm font-medium text-slate-900 border border-slate-200 rounded-lg px-3 py-1.5 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            >
              {BULAN_ID.map((name, idx) => (
                <option key={name} value={idx + 1}>
                  {name}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-2">
            <label htmlFor="report-year" className="text-xs font-medium text-slate-500">
              Tahun
            </label>
            <select
              id="report-year"
              value={selectedYear}
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              className="text-sm font-medium text-slate-900 border border-slate-200 rounded-lg px-3 py-1.5 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-sky-500 font-mono tabular-nums"
            >
              {yearOptions.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>

          {onOpenNonAktaRekap && (
            <button
              type="button"
              onClick={onOpenNonAktaRekap}
              className="ml-auto inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-sky-200 bg-sky-50 hover:bg-sky-100 text-xs font-medium text-sky-800 transition-colors whitespace-nowrap"
            >
              <ClipboardList className="w-3.5 h-3.5" />
              <span>Buka Rekap &amp; Tracking Non-Akta</span>
            </button>
          )}
        </div>
      </div>

      {/* Printable Official Preview Sheet */}
      <div className="bg-white border border-slate-200 rounded-xl p-6 sm:p-8 space-y-6">
        <div className="text-center border-b-2 border-slate-900 pb-4 space-y-1">
          <p className="text-xs font-semibold uppercase tracking-wide text-slate-600">
            Pemerintah {settings.nama_kabupaten}
          </p>
          <h2 className="font-display text-lg font-bold uppercase text-slate-900">
            Laporan Bulanan Pembuatan Akta Pejabat Pembuat Akta Tanah Sementara (PPATS)
          </h2>
          <p className="text-sm font-semibold text-slate-800">
            {settings.nama_kecamatan} · Periode: {bulanLabel} {selectedYear}
          </p>
          <p className="text-xs text-slate-500">
            Kepada Yth. Kepala Kantor Pertanahan (BPN) & Kepala Bapenda {settings.nama_kabupaten}
          </p>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs border-b border-slate-200 pb-4">
          <div>
            <span className="text-slate-500">Pejabat PPATS:</span>{' '}
            <span className="font-semibold text-slate-900">{settings.nama_ppats}</span>
          </div>
          <div>
            <span className="text-slate-500">NIP:</span>{' '}
            <span className="font-mono text-slate-900">{settings.nip_ppats}</span>
          </div>
          <div>
            <span className="text-slate-500">Jumlah Akta Terbit:</span>{' '}
            <span className="font-mono font-semibold text-slate-900 tabular-nums">
              {filtered.length === 0 ? 'NIHIL (0 Akta)' : `${filtered.length} Akta`}
            </span>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="border-y border-slate-300 bg-slate-50 font-semibold text-slate-800">
                <th className="py-2.5 px-2.5 border-r border-slate-200">No</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200">Nomor & Tgl Akta</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200">Perbuatan Hukum</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200">Pihak Pertama</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200">Pihak Kedua</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200">Hak & NIB</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200">Letak Tanah</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200 text-right">Luas (m²)</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200 text-right">Nilai Transaksi</th>
                <th className="py-2.5 px-2.5 border-r border-slate-200 text-right">BPHTB (5%)</th>
                <th className="py-2.5 px-2.5 text-right">PPh (2,5%)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filtered.length === 0 ? (
                <tr>
                  <td
                    colSpan={11}
                    className="py-10 px-4 text-center font-display font-bold text-base tracking-wide text-slate-700 bg-slate-50/70"
                  >
                    NIHIL — Tidak ada akta PPATS yang diterbitkan pada bulan {bulanLabel} {selectedYear}
                  </td>
                </tr>
              ) : (
                filtered.map((r, idx) => {
                  const p1 =
                    r.para_pihak
                      .filter((p) => p.peran === 'Pihak I')
                      .map((p) => p.nama)
                      .join(', ') || '-';
                  const p2 =
                    r.para_pihak
                      .filter((p) => p.peran === 'Pihak II')
                      .map((p) => p.nama)
                      .join(', ') || '-';
                  const tgl = r.akta.tanggal_akta
                    ? formatTanggalDDMMYYYY(r.akta.tanggal_akta)
                    : formatTanggalDDMMYYYY(r.tanggal_masuk);

                  return (
                    <tr key={r.id}>
                      <td className="py-2.5 px-2.5 border-r border-slate-200 font-mono tabular-nums">
                        {idx + 1}
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200">
                        <div className="font-mono font-medium text-slate-900">
                          {r.akta.nomor_akta || '-'}
                        </div>
                        <div className="font-mono text-slate-500 tabular-nums">{tgl}</div>
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200 font-medium">
                        {r.layanan}
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200">{p1}</td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200">{p2}</td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200">
                        <div>
                          {r.tanah.jenis_sertipikat === 'digital' || !r.tanah.nomor_hak_lama
                            ? `${r.tanah.jenis_hak} (Digital)`
                            : `${r.tanah.jenis_hak} No. ${r.tanah.nomor_hak_lama}`}
                        </div>
                        <div className="font-mono text-slate-500">
                          {r.tanah.jenis_sertipikat === 'digital' || !r.tanah.nomor_hak_lama ? 'NIBel' : 'NIB'} {r.tanah.nib || '-'}
                        </div>
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200">
                        {r.tanah.desa}, {r.tanah.kecamatan}
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200 text-right font-mono tabular-nums">
                        {formatNumberId(r.tanah.luas_m2)}
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200 text-right font-mono tabular-nums whitespace-nowrap">
                        {formatRupiah(r.pajak.nilai_kesepakatan)}
                      </td>
                      <td className="py-2.5 px-2.5 border-r border-slate-200 text-right font-mono tabular-nums whitespace-nowrap">
                        {formatRupiah(r.pajak.bphtb)}
                      </td>
                      <td className="py-2.5 px-2.5 text-right font-mono tabular-nums whitespace-nowrap">
                        {formatRupiah(r.pajak.pph)}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
            <tfoot>
              <tr className="border-t-2 border-slate-300 bg-slate-50 font-semibold text-slate-900">
                <td colSpan={7} className="py-2.5 px-2.5 border-r border-slate-200 text-right">
                  JUMLAH TOTAL ({filtered.length} Akta)
                </td>
                <td className="py-2.5 px-2.5 border-r border-slate-200 text-right font-mono tabular-nums">
                  {formatNumberId(totals.luas)}
                </td>
                <td className="py-2.5 px-2.5 border-r border-slate-200 text-right font-mono tabular-nums whitespace-nowrap">
                  {formatRupiah(totals.nilai)}
                </td>
                <td className="py-2.5 px-2.5 border-r border-slate-200 text-right font-mono tabular-nums whitespace-nowrap">
                  {formatRupiah(totals.bphtb)}
                </td>
                <td className="py-2.5 px-2.5 text-right font-mono tabular-nums whitespace-nowrap">
                  {formatRupiah(totals.pph)}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>

        {/* Official PPATS Signature Block */}
        <div className="pt-6 flex justify-end">
          <div className="text-center text-xs space-y-1 w-72">
            <p className="text-slate-600">
              {settings.nama_kecamatan}, Akhir Bulan {bulanLabel} {selectedYear}
            </p>
            <p className="font-semibold text-slate-900">
              Pejabat Pembuat Akta Tanah Sementara (PPATS)
            </p>
            <p className="text-slate-600">{settings.jabatan_ppats}</p>
            <div className="h-16" />
            <p className="font-semibold underline text-slate-900">{settings.nama_ppats}</p>
            <p className="font-mono text-slate-600">NIP. {settings.nip_ppats}</p>
          </div>
        </div>
      </div>
    </div>
  );
};
