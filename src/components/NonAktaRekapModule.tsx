import React, { useMemo, useState } from 'react';
import {
  CheckCircle2,
  ClipboardList,
  FilePlus2,
  FileSpreadsheet,
  Pencil,
  Printer,
  Search,
  Trash2,
} from 'lucide-react';
import {
  BerkasRecord,
  BerkasStatus,
  OfficeSettings,
  ServiceType,
} from '../types/ppats';
import { STATUS_OPTIONS } from '../utils/seedData';
import {
  BULAN_ID,
  extractMonthYearFromDate,
  formatNumberId,
  formatTanggalDDMMYYYY,
} from '../utils/indonesianFormatter';
import {
  exportNonAktaRekapXlsx,
  filterNonAktaRecords,
  summarizeNonAktaStats,
} from '../utils/reportGenerator';

interface NonAktaRekapModuleProps {
  records: BerkasRecord[];
  settings: OfficeSettings;
  onCreateNewNonAkta: () => void;
  onEditRecord: (record: BerkasRecord) => void;
  onChangeStatus: (recordId: string, nextStatus: BerkasStatus) => void;
  onPrintReceipt: (record: BerkasRecord) => void;
  onDeleteRecord: (recordId: string) => void;
  onNotify: (msg: string, type?: 'success' | 'error') => void;
}

const NON_AKTA_TABS: Array<{ id: 'SEMUA' | ServiceType; label: string }> = [
  { id: 'SEMUA', label: 'Semua Non-Akta' },
  { id: 'Turun Waris', label: 'Turun Waris' },
  { id: 'Pemecahan', label: 'Pemecahan' },
  { id: 'Ukur Ulang', label: 'Ukur Ulang' },
  { id: 'Ralat Data', label: 'Ralat Data' },
];

function getNonAktaStatusLabel(status: BerkasStatus): string {
  return status === '3. TTD Akta' ? '3. TTD & Pengesahan' : status;
}

function getPengesahanSummary(rec: BerkasRecord, settings: OfficeSettings): string {
  const na = rec.non_akta;
  const dualModes =
    rec.layanan === 'Turun Waris' ||
    rec.layanan === 'Ralat Data' ||
    Boolean(na?.sertakan_beda_nama);
  const tunggal =
    na?.pengesahan_mode === 'kades'
      ? `Kepala Desa ${rec.tanah.desa || '-'}`
      : na?.pengesahan_mode === 'camat'
      ? `Camat (${na?.pengesahan_nama?.trim() || settings.nama_ppats})`
      : na?.pengesahan_mode === 'manual'
      ? na?.pengesahan_jabatan?.trim() || 'Pejabat Lainnya'
      : `PPATS (${na?.pengesahan_nama?.trim() || settings.nama_ppats})`;
  return dualModes
    ? `${tunggal} + Camat & Kades`
    : tunggal;
}

function getSpecificSummary(rec: BerkasRecord): string {
  const na = rec.non_akta;
  if (rec.layanan === 'Turun Waris') {
    return na?.nama_pewaris ? `Pewaris: ${na.nama_pewaris}` : 'Permohonan Turun Waris';
  }
  if (rec.layanan === 'Pemecahan' || rec.layanan === 'Pemecahan/Ukur Ulang') {
    const count = na?.bidang_pecahan?.length ?? 0;
    return count > 0 ? `Rencana pecah ${count} bidang` : 'Pemecahan bidang tanah';
  }
  if (rec.layanan === 'Ukur Ulang') {
    return na?.koordinat_gps ? `GPS: ${na.koordinat_gps}` : 'Pengukuran ulang / batas';
  }
  if (rec.layanan === 'Ralat Data') {
    if (na?.nama_di_sertipikat || na?.nama_di_ktp_kk) {
      return `${na.nama_di_sertipikat || '-'} → ${na.nama_di_ktp_kk || '-'}`;
    }
    return 'Ralat data sertipikat';
  }
  return '-';
}

export const NonAktaRekapModule: React.FC<NonAktaRekapModuleProps> = ({
  records,
  settings,
  onCreateNewNonAkta,
  onEditRecord,
  onChangeStatus,
  onPrintReceipt,
  onDeleteRecord,
  onNotify,
}) => {
  const [layananFilter, setLayananFilter] = useState<'SEMUA' | ServiceType>('SEMUA');
  const [statusFilter, setStatusFilter] = useState<'SEMUA' | BerkasStatus>('SEMUA');
  const [monthFilter, setMonthFilter] = useState<'SEMUA' | number>('SEMUA');
  const [yearFilter, setYearFilter] = useState<'SEMUA' | number>('SEMUA');
  const [searchQuery, setSearchQuery] = useState('');

  const allNonAkta = useMemo(() => filterNonAktaRecords(records), [records]);
  const stats = useMemo(() => summarizeNonAktaStats(records), [records]);

  const availableYears = useMemo(() => {
    const set = new Set<number>([new Date().getFullYear()]);
    for (const r of allNonAkta) {
      const my = extractMonthYearFromDate(r.akta.tanggal_akta || r.tanggal_masuk);
      if (my && my.year > 2000) set.add(my.year);
    }
    return Array.from(set).sort((a, b) => b - a);
  }, [allNonAkta]);

  const filtered = useMemo(
    () =>
      filterNonAktaRecords(records, {
        layanan: layananFilter,
        status: statusFilter,
        month: monthFilter,
        year: yearFilter,
        query: searchQuery,
      }),
    [records, layananFilter, statusFilter, monthFilter, yearFilter, searchQuery]
  );

  const periodeLabel = useMemo(() => {
    const parts: string[] = [];
    if (layananFilter !== 'SEMUA') parts.push(layananFilter);
    if (monthFilter !== 'SEMUA') parts.push(BULAN_ID[monthFilter - 1] ?? String(monthFilter));
    if (yearFilter !== 'SEMUA') parts.push(String(yearFilter));
    return parts.length > 0 ? parts.join('_') : 'Semua_Periode';
  }, [layananFilter, monthFilter, yearFilter]);

  const handleExport = () => {
    try {
      const fileName = exportNonAktaRekapXlsx({
        records: filtered,
        settings,
        periodeLabel,
      });
      onNotify(`Berhasil mengekspor ${fileName}`, 'success');
    } catch (err) {
      onNotify(
        err instanceof Error ? err.message : 'Gagal mengekspor rekap Non-Akta.',
        'error'
      );
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="no-print flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
        <div>
          <div className="flex items-center gap-2">
            <ClipboardList className="w-5 h-5 text-sky-700 shrink-0" />
            <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
              Rekap Berkas &amp; Tracking Non-Akta
            </h1>
          </div>
          <p className="text-sm text-slate-600 mt-1">
            Buku rekapitulasi dan pelacakan tahapan khusus berkas Non-Akta (Turun Waris, Pemecahan, Ukur Ulang, Ralat Data) di luar Laporan Bulanan PPATS.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleExport}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-medium transition-colors whitespace-nowrap"
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Export Rekap (.XLSX)</span>
          </button>

          <button
            type="button"
            onClick={() => window.print()}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium transition-colors whitespace-nowrap"
          >
            <Printer className="w-4 h-4" />
            <span>Cetak Rekap</span>
          </button>

          <button
            type="button"
            onClick={onCreateNewNonAkta}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium transition-colors whitespace-nowrap"
          >
            <FilePlus2 className="w-4 h-4" />
            <span>Buat Berkas Non-Akta</span>
          </button>
        </div>
      </div>

      {/* KPI Summary Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Total Berkas Non-Akta</p>
          <p className="font-mono text-2xl font-semibold text-slate-900 mt-2 tabular-nums">
            {stats.total}
          </p>
          <p className="text-[11px] text-slate-500 mt-1">
            Waris {stats.turunWaris} · Pecah {stats.pemecahan} · Ukur {stats.ukurUlang} · Ralat {stats.ralatData}
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Pemberkasan &amp; Pengesahan</p>
          <p className="font-mono text-2xl font-semibold text-amber-700 mt-2 tabular-nums">
            {stats.pemberkasanAwal}
          </p>
          <p className="text-[11px] text-slate-500 mt-1">
            Menunggu tanda tangan, Kades, atau validasi pajak
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Sedang Proses di BPN</p>
          <p className="font-mono text-2xl font-semibold text-indigo-700 mt-2 tabular-nums">
            {stats.prosesBpn}
          </p>
          <p className="text-[11px] text-slate-500 mt-1">
            Berkas sudah masuk loket Kantor Pertanahan
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Selesai / Terbit</p>
          <p className="font-mono text-2xl font-semibold text-emerald-700 mt-2 tabular-nums">
            {stats.selesai}
          </p>
          <p className="text-[11px] text-slate-500 mt-1">
            Sertipikat / hasil layanan Non-Akta selesai
          </p>
        </div>
      </div>

      {/* Filters */}
      <div className="no-print bg-white border border-slate-200 rounded-xl p-4 space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg">
            {NON_AKTA_TABS.map((tab) => {
              const active = layananFilter === tab.id;
              return (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setLayananFilter(tab.id)}
                  className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors whitespace-nowrap ${
                    active
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </div>

          <div className="relative w-full sm:w-72">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="search"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Cari pemohon, pewaris, NIB, desa..."
              aria-label="Cari berkas Non-Akta"
              className="w-full pl-9 pr-3 py-2 text-xs border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3 text-xs pt-1 border-t border-slate-100">
          <div className="flex items-center gap-1.5">
            <label htmlFor="nonakta-filter-status" className="text-slate-500 font-medium">
              Status
            </label>
            <select
              id="nonakta-filter-status"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value as 'SEMUA' | BerkasStatus)}
              className="px-2.5 py-1.5 border border-slate-200 rounded-lg bg-white text-slate-800"
            >
              <option value="SEMUA">Semua Status</option>
              {STATUS_OPTIONS.map((st) => (
                <option key={st} value={st}>
                  {getNonAktaStatusLabel(st)}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-1.5">
            <label htmlFor="nonakta-filter-bulan" className="text-slate-500 font-medium">
              Bulan
            </label>
            <select
              id="nonakta-filter-bulan"
              value={monthFilter}
              onChange={(e) =>
                setMonthFilter(e.target.value === 'SEMUA' ? 'SEMUA' : Number(e.target.value))
              }
              className="px-2.5 py-1.5 border border-slate-200 rounded-lg bg-white text-slate-800"
            >
              <option value="SEMUA">Semua Bulan</option>
              {BULAN_ID.map((name, idx) => (
                <option key={name} value={idx + 1}>
                  {name}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-1.5">
            <label htmlFor="nonakta-filter-tahun" className="text-slate-500 font-medium">
              Tahun
            </label>
            <select
              id="nonakta-filter-tahun"
              value={yearFilter}
              onChange={(e) =>
                setYearFilter(e.target.value === 'SEMUA' ? 'SEMUA' : Number(e.target.value))
              }
              className="px-2.5 py-1.5 border border-slate-200 rounded-lg bg-white font-mono text-slate-800"
            >
              <option value="SEMUA">Semua Tahun</option>
              {availableYears.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>

          <span className="ml-auto text-slate-500 tabular-nums">
            Menampilkan <strong className="text-slate-800">{filtered.length}</strong> dari{' '}
            {allNonAkta.length} berkas Non-Akta
          </span>
        </div>
      </div>

      {/* Print-only Header */}
      <div className="hidden print:block text-center border-b-2 border-slate-900 pb-3 mb-4">
        <p className="text-xs font-semibold uppercase">{settings.nama_kabupaten}</p>
        <h2 className="text-base font-bold uppercase">
          REKAPITULASI &amp; TRACKING BERKAS LAYANAN NON-AKTA ({settings.nama_kecamatan})
        </h2>
        <p className="text-xs text-slate-600">
          PPATS: {settings.nama_ppats} · Format Tanggal: DD-MM-YYYY
        </p>
      </div>

      {/* Interactive Tracking Cards */}
      {filtered.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-xl p-12 text-center space-y-2">
          <p className="text-sm font-semibold text-slate-800">
            Belum ada berkas Non-Akta yang sesuai filter saat ini.
          </p>
          <p className="text-xs text-slate-500">
            Buat berkas Turun Waris, Pemecahan, Ukur Ulang, atau Ralat Data baru agar muncul di rekap dan pelacakan ini.
          </p>
          <button
            type="button"
            onClick={onCreateNewNonAkta}
            className="mt-3 inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium transition-colors"
          >
            <FilePlus2 className="w-3.5 h-3.5" />
            <span>Buat Berkas Non-Akta</span>
          </button>
        </div>
      ) : (
        <div className="space-y-3">
          {filtered.map((rec) => {
            const currentStageIdx = STATUS_OPTIONS.indexOf(rec.status);
            const pemohon =
              rec.para_pihak
                .filter((p) => p.peran.startsWith('Pihak '))
                .map((p) => p.nama)
                .filter(Boolean)
                .join(', ') || '-';
            const tglMasuk = formatTanggalDDMMYYYY(rec.tanggal_masuk);
            const tglSurat = rec.akta.tanggal_akta
              ? formatTanggalDDMMYYYY(rec.akta.tanggal_akta)
              : 'Belum bertanggal';

            return (
              <div
                key={rec.id}
                className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5 space-y-4 break-inside-avoid"
              >
                {/* Top Meta & Actions */}
                <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3 border-b border-slate-100 pb-3">
                  <div className="flex flex-wrap items-center gap-2.5">
                    <span className="font-mono text-xs font-semibold text-slate-900 bg-slate-100 px-2.5 py-1 rounded-md tabular-nums">
                      {rec.kode_berkas}
                    </span>
                    <span className="text-xs font-semibold text-sky-900 bg-sky-50 border border-sky-200 px-2.5 py-0.5 rounded-md">
                      {rec.layanan}
                    </span>
                    <span className="font-mono text-xs text-slate-600 tabular-nums">
                      Masuk: <strong>{tglMasuk}</strong>
                    </span>
                    <span className="text-slate-300">·</span>
                    <span className="font-mono text-xs text-slate-600 tabular-nums">
                      Tgl Berkas:{' '}
                      <strong className={!rec.akta.tanggal_akta ? 'text-amber-700 font-sans' : ''}>
                        {tglSurat}
                      </strong>
                    </span>
                  </div>

                  <div className="no-print flex flex-wrap items-center gap-1.5">
                    <select
                      aria-label={`Ubah status Non-Akta ${rec.kode_berkas}`}
                      value={rec.status}
                      onChange={(e) =>
                        onChangeStatus(rec.id, e.target.value as BerkasStatus)
                      }
                      className="text-xs font-medium rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-slate-800 focus:outline-none focus:ring-2 focus:ring-sky-500"
                    >
                      {STATUS_OPTIONS.map((opt) => (
                        <option key={opt} value={opt}>
                          {getNonAktaStatusLabel(opt)}
                        </option>
                      ))}
                    </select>

                    <button
                      type="button"
                      onClick={() => onEditRecord(rec)}
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700"
                    >
                      <Pencil className="w-3 h-3" />
                      <span>Edit</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => onPrintReceipt(rec)}
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700"
                    >
                      <Printer className="w-3 h-3" />
                      <span>Kwitansi</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => onDeleteRecord(rec.id)}
                      aria-label={`Hapus berkas ${rec.kode_berkas}`}
                      className="p-1.5 rounded-lg border border-transparent hover:border-rose-200 hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>

                {/* Details Grid */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                  <div>
                    <p className="text-slate-500 font-medium">
                      {rec.layanan === 'Turun Waris' ? 'Para Ahli Waris / Pemohon' : 'Pemohon'}
                    </p>
                    <p className="text-slate-900 font-semibold mt-0.5">{pemohon}</p>
                    <p className="text-slate-600 mt-0.5">{getSpecificSummary(rec)}</p>
                  </div>

                  <div>
                    <p className="text-slate-500 font-medium">Objek Tanah &amp; Sertipikat</p>
                    <p className="text-slate-900 font-medium mt-0.5">
                      {rec.tanah.jenis_sertipikat === 'digital' || !rec.tanah.nomor_hak_lama
                        ? `${rec.tanah.jenis_hak} Digital · NIBel ${rec.tanah.nib || '-'}`
                        : `${rec.tanah.jenis_hak} No. ${rec.tanah.nomor_hak_lama} · NIB ${rec.tanah.nib || '-'}`}
                    </p>
                    <p className="text-slate-600 mt-0.5">
                      Desa {rec.tanah.desa}, Kec. {rec.tanah.kecamatan} ·{' '}
                      <span className="font-mono tabular-nums">
                        {formatNumberId(rec.tanah.luas_m2)} m²
                      </span>
                    </p>
                  </div>

                  <div>
                    <p className="text-slate-500 font-medium">Kuasa BPN &amp; Pengesahan</p>
                    <p className="text-slate-900 mt-0.5">
                      Kuasa: <span className="font-medium">{rec.akta.nama_kuasa_bpn || '-'}</span>
                    </p>
                    <p className="text-slate-600 mt-0.5">
                      Pengesahan: {getPengesahanSummary(rec, settings)}
                    </p>
                  </div>
                </div>

                {/* 5-Stage Interactive Progress Tracking */}
                <div className="pt-2">
                  <p className="text-[11px] font-medium text-slate-500 mb-2">
                    Tracking Tahapan Berkas Non-Akta (klik tahap untuk memperbarui progres):
                  </p>
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                    {STATUS_OPTIONS.map((stage, sIdx) => {
                      const isDone = sIdx < currentStageIdx;
                      const isCurrent = sIdx === currentStageIdx;
                      return (
                        <button
                          key={stage}
                          type="button"
                          onClick={() => onChangeStatus(rec.id, stage)}
                          className={`flex items-center gap-1.5 px-2.5 py-2 rounded-lg border text-left text-[11px] font-medium transition-colors ${
                            isCurrent
                              ? 'border-sky-600 bg-sky-50 text-sky-950 ring-1 ring-sky-600/20'
                              : isDone
                              ? 'border-emerald-200 bg-emerald-50/70 text-emerald-900 hover:bg-emerald-100/70'
                              : 'border-slate-200 bg-slate-50/70 text-slate-600 hover:bg-slate-100'
                          }`}
                        >
                          <CheckCircle2
                            className={`w-3.5 h-3.5 shrink-0 ${
                              isCurrent
                                ? 'text-sky-600'
                                : isDone
                                ? 'text-emerald-600'
                                : 'text-slate-300'
                            }`}
                          />
                          <span className="truncate">{getNonAktaStatusLabel(stage)}</span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
