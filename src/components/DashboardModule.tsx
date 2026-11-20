import React, { useMemo, useState } from 'react';
import {
  FilePlus2,
  Printer,
  Pencil,
  Trash2,
  Search,
  CheckCircle2,
  Clock,
  FileCheck2,
  Building2,
  FolderOpen,
  Calculator,
} from 'lucide-react';
import { BerkasRecord, BerkasStatus } from '../types/ppats';
import { STATUS_OPTIONS } from '../utils/seedData';
import {
  extractMonthYearFromDate,
  formatNumberId,
  formatRupiah,
  formatTanggalDDMMYYYY,
} from '../utils/indonesianFormatter';
import { isAktaService } from '../utils/taxAndAudit';

interface DashboardModuleProps {
  records: BerkasRecord[];
  onCreateNew: () => void;
  onEditRecord: (record: BerkasRecord) => void;
  onCalculateTax: (record: BerkasRecord) => void;
  onChangeStatus: (recordId: string, nextStatus: BerkasStatus) => void;
  onPrintReceipt: (record: BerkasRecord) => void;
  onDeleteRecord: (recordId: string) => void;
}

function statusStyle(status: BerkasStatus): {
  bg: string;
  text: string;
  border: string;
  Icon: React.ComponentType<{ className?: string }>;
} {
  switch (status) {
    case '1. Pemberkasan':
      return {
        bg: 'bg-slate-100',
        text: 'text-slate-700',
        border: 'border-slate-300',
        Icon: FolderOpen,
      };
    case '2. Validasi Pajak':
      return {
        bg: 'bg-amber-50',
        text: 'text-amber-800',
        border: 'border-amber-300',
        Icon: Clock,
      };
    case '3. TTD Akta':
      return {
        bg: 'bg-sky-50',
        text: 'text-sky-800',
        border: 'border-sky-300',
        Icon: FileCheck2,
      };
    case '4. Proses BPN':
      return {
        bg: 'bg-indigo-50',
        text: 'text-indigo-800',
        border: 'border-indigo-300',
        Icon: Building2,
      };
    case '5. Selesai':
      return {
        bg: 'bg-emerald-50',
        text: 'text-emerald-800',
        border: 'border-emerald-300',
        Icon: CheckCircle2,
      };
  }
}

export const DashboardModule: React.FC<DashboardModuleProps> = ({
  records,
  onCreateNew,
  onEditRecord,
  onCalculateTax,
  onChangeStatus,
  onPrintReceipt,
  onDeleteRecord,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<'SEMUA' | BerkasStatus>('SEMUA');

  const kpis = useMemo(() => {
    const now = new Date();
    const currentYm = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const recordYms = records
      .map((r) => {
        const my = extractMonthYearFromDate(r.akta.tanggal_akta || r.tanggal_masuk);
        return my ? `${my.year}-${String(my.month).padStart(2, '0')}` : '';
      })
      .filter(Boolean);
    const latestYm = recordYms.length > 0 ? [...recordYms].sort().pop() || currentYm : currentYm;

    const activeYm = recordYms.includes(currentYm) ? currentYm : latestYm;
    const [activeY, activeM] = activeYm.split('-');
    const activeLabel = activeM && activeY ? `${activeM}-${activeY}` : activeYm;

    const berkasBulanIni = records.filter((r) => {
      const my = extractMonthYearFromDate(r.akta.tanggal_akta || r.tanggal_masuk);
      const ym = my ? `${my.year}-${String(my.month).padStart(2, '0')}` : '';
      return ym === activeYm;
    }).length;

    const aktaTerbit = records.filter(
      (r) =>
        isAktaService(r.layanan) &&
        (r.status === '3. TTD Akta' ||
          r.status === '4. Proses BPN' ||
          r.status === '5. Selesai')
    ).length;

    const totalNilaiTransaksi = records.reduce(
      (sum, r) => sum + (r.pajak?.nilai_kesepakatan || 0),
      0
    );

    const prosesBpn = records.filter((r) => r.status === '4. Proses BPN').length;

    return { berkasBulanIni, aktaTerbit, totalNilaiTransaksi, prosesBpn, activeLabel };
  }, [records]);

  const filteredRecords = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return [...records]
      .sort((a, b) => b.no_urut - a.no_urut)
      .filter((r) => {
        if (statusFilter !== 'SEMUA' && r.status !== statusFilter) return false;
        if (!q) return true;
        const partiesText = r.para_pihak
          .map((p) => `${p.nama} ${p.nik}`)
          .join(' ')
          .toLowerCase();
        const haystack = [
          r.kode_berkas,
          r.akta.nomor_akta,
          r.layanan,
          r.tanah.nib,
          r.tanah.nomor_hak_lama,
          r.tanah.desa,
          partiesText,
        ]
          .join(' ')
          .toLowerCase();
        return haystack.includes(q);
      });
  }, [records, searchQuery, statusFilter]);

  return (
    <div className="space-y-6">
      {/* Page Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
            Dashboard & Tracking Progres Berkas
          </h1>
          <p className="text-sm text-slate-600 mt-1">
            Pemantauan berkas loket PPATS, validasi pajak daerah, penandatanganan akta, hingga proses balik nama BPN.
          </p>
        </div>
        <button
          type="button"
          onClick={onCreateNew}
          className="inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium transition-colors whitespace-nowrap shrink-0"
        >
          <FilePlus2 className="w-4 h-4" />
          <span>Buat Berkas Baru</span>
        </button>
      </div>

      {/* 4 KPI Summary Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">
            Total Berkas Bulan Ini ({kpis.activeLabel})
          </p>
          <p className="font-mono text-2xl font-semibold text-slate-900 mt-2 tabular-nums">
            {kpis.berkasBulanIni}
          </p>
          <p className="text-xs text-slate-500 mt-1">
            Dari total {records.length} berkas terdaftar
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Jumlah Akta Terbit</p>
          <p className="font-mono text-2xl font-semibold text-slate-900 mt-2 tabular-nums">
            {kpis.aktaTerbit}
          </p>
          <p className="text-xs text-slate-500 mt-1">
            Khusus Akta PPATS (tidak termasuk Non-Akta)
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Total Nilai Transaksi (Rp)</p>
          <p className="font-mono text-2xl font-semibold text-slate-900 mt-2 tabular-nums truncate">
            {formatRupiah(kpis.totalNilaiTransaksi)}
          </p>
          <p className="text-xs text-slate-500 mt-1">
            Akumulasi seluruh nilai peralihan hak
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <p className="text-xs font-medium text-slate-500">Berkas Sedang Proses di BPN</p>
          <p className="font-mono text-2xl font-semibold text-indigo-700 mt-2 tabular-nums">
            {kpis.prosesBpn}
          </p>
          <p className="text-xs text-slate-500 mt-1">
            Menunggu penerbitan sertipikat balik nama
          </p>
        </div>
      </div>

      {/* Monitoring Berkas Table */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-4 sm:p-5 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          <div className="flex flex-wrap items-center gap-1.5 p-1 bg-slate-100 rounded-lg">
            {(['SEMUA', ...STATUS_OPTIONS] as const).map((st) => {
              const active = statusFilter === st;
              return (
                <button
                  key={st}
                  type="button"
                  onClick={() => setStatusFilter(st)}
                  className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors whitespace-nowrap ${
                    active
                      ? 'bg-white text-slate-900 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {st === 'SEMUA' ? 'Semua Status' : st}
                </button>
              );
            })}
          </div>

          <div className="relative w-full lg:w-72 shrink-0">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="search"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Cari nama, no. akta, NIB..."
              aria-label="Cari berkas di papan monitoring"
              className="w-full pl-9 pr-3 py-2 text-sm border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
        </div>

        {filteredRecords.length === 0 ? (
          <div className="p-12 text-center">
            <p className="text-sm font-medium text-slate-700">
              Tidak ada berkas yang sesuai dengan filter saat ini.
            </p>
            <p className="text-xs text-slate-500 mt-1">
              Ubah kata kunci pencarian atau buat berkas pertanahan baru.
            </p>
            <button
              type="button"
              onClick={onCreateNew}
              className="mt-4 inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-[#0f172a] text-white text-xs font-medium hover:bg-slate-800 transition-colors"
            >
              <FilePlus2 className="w-3.5 h-3.5" />
              <span>Buat Berkas Baru</span>
            </button>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-xs font-semibold text-slate-600">
                  <th className="py-3 px-2 whitespace-nowrap">Kode & No. Akta</th>
                  <th className="py-3 px-2 whitespace-nowrap">Layanan & Tgl</th>
                  <th className="py-3 px-2">Para Pihak</th>
                  <th className="py-3 px-2">Objek Tanah</th>
                  <th className="py-3 px-2 text-right whitespace-nowrap">Nilai Transaksi</th>
                  <th className="py-3 px-2 whitespace-nowrap">Status Proses</th>
                  <th className="py-3 px-2 text-right whitespace-nowrap">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredRecords.map((rec) => {
                  const p1Names =
                    rec.para_pihak
                      .filter((p) => p.peran === 'Pihak I')
                      .map((p) => p.nama)
                      .join(', ') || '-';
                  const p2Names =
                    rec.para_pihak
                      .filter((p) => p.peran === 'Pihak II')
                      .map((p) => p.nama)
                      .join(', ') || '-';
                  const isAkta = isAktaService(rec.layanan);
                  const allParties =
                    rec.para_pihak
                      .filter((p) => p.peran.startsWith('Pihak '))
                      .map((p) => p.nama)
                      .filter(Boolean)
                      .join(', ') || '-';
                  const st = statusStyle(rec.status);
                  const StIcon = st.Icon;

                  return (
                    <tr key={rec.id} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-3 px-2 align-top whitespace-nowrap">
                        <div className="font-mono text-xs font-semibold text-slate-900 tabular-nums">
                          {rec.kode_berkas}
                        </div>
                        <div className="font-mono text-[11px] text-slate-500 mt-0.5 tabular-nums">
                          {isAkta
                            ? rec.akta.nomor_akta || 'Belum bernomor'
                            : 'Non-Akta (F4)'}
                        </div>
                      </td>
                      <td className="py-3 px-2 align-top whitespace-nowrap">
                        <div className="text-xs font-semibold text-slate-900">{rec.layanan}</div>
                        <div className="font-mono text-[11px] text-slate-600 mt-0.5 tabular-nums">
                          {rec.akta.tanggal_akta
                            ? formatTanggalDDMMYYYY(rec.akta.tanggal_akta)
                            : 'Belum bertanggal'}
                        </div>
                        {!rec.akta.tanggal_akta && rec.tanggal_masuk && (
                          <div className="font-mono text-[10px] text-slate-400 tabular-nums">
                            Masuk: {formatTanggalDDMMYYYY(rec.tanggal_masuk)}
                          </div>
                        )}
                      </td>
                      <td className="py-3 px-2 align-top max-w-[140px]">
                        {isAkta || p2Names !== '-' ? (
                          <>
                            <div className="text-xs text-slate-800 truncate" title={p1Names}>
                              <span className="text-slate-500">I:</span> {p1Names}
                            </div>
                            <div className="text-xs text-slate-800 truncate mt-0.5" title={p2Names}>
                              <span className="text-slate-500">II:</span> {p2Names}
                            </div>
                          </>
                        ) : (
                          <>
                            <div className="text-xs text-slate-800 truncate" title={allParties}>
                              <span className="text-slate-500">Pemohon:</span> {allParties}
                            </div>
                            {rec.non_akta?.nama_pewaris ? (
                              <div className="text-[11px] text-slate-500 truncate mt-0.5" title={rec.non_akta.nama_pewaris}>
                                Pewaris: {rec.non_akta.nama_pewaris}
                              </div>
                            ) : (
                              <div className="text-[11px] text-slate-500 mt-0.5">
                                Berkas Non-Akta
                              </div>
                            )}
                          </>
                        )}
                      </td>
                      <td className="py-3 px-2 align-top max-w-[190px]">
                        <div className="text-xs font-medium text-slate-900 truncate">
                          {rec.tanah.jenis_hak} · {rec.tanah.desa}
                        </div>
                        <div className="font-mono text-[11px] text-slate-500 mt-0.5 tabular-nums truncate">
                          {rec.tanah.nib || '-'} · {formatNumberId(rec.tanah.luas_m2)} m²
                        </div>
                      </td>
                      <td className="py-3 px-2 align-top text-right font-mono text-xs font-medium text-slate-900 tabular-nums whitespace-nowrap">
                        {formatRupiah(rec.pajak.nilai_kesepakatan)}
                      </td>
                      <td className="py-3 px-2 align-top whitespace-nowrap">
                        <div className="inline-flex items-center gap-1">
                          <StIcon className={`w-3.5 h-3.5 shrink-0 ${st.text}`} />
                          <select
                            value={rec.status}
                            aria-label={`Ubah status berkas ${rec.kode_berkas}`}
                            onChange={(e) =>
                              onChangeStatus(rec.id, e.target.value as BerkasStatus)
                            }
                            className={`text-xs font-medium rounded-md border px-1.5 py-1 focus:outline-none focus:ring-2 focus:ring-sky-500 cursor-pointer ${st.bg} ${st.text} ${st.border}`}
                          >
                            {STATUS_OPTIONS.map((opt) => (
                              <option key={opt} value={opt} className="bg-white text-slate-900">
                                {opt}
                              </option>
                            ))}
                          </select>
                        </div>
                      </td>
                      <td className="py-3 px-2 align-top text-right whitespace-nowrap">
                        <div className="inline-flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => onEditRecord(rec)}
                            title="Edit / Buka Ulang Berkas"
                            aria-label={`Edit berkas ${rec.kode_berkas}`}
                            className="inline-flex items-center gap-1 px-1.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-xs font-medium text-slate-700 transition-colors"
                          >
                            <Pencil className="w-3 h-3" />
                            <span>Edit</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => onCalculateTax(rec)}
                            title="Hitung / Ubah Pajak di Menu Perhitungan Pajak"
                            aria-label={`Hitung pajak ${rec.kode_berkas}`}
                            className="inline-flex items-center gap-1 px-1.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-xs font-medium text-slate-700 transition-colors"
                          >
                            <Calculator className="w-3 h-3" />
                            <span>Pajak</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => onPrintReceipt(rec)}
                            title="Cetak Tanda Terima Pemohon"
                            aria-label={`Cetak tanda terima ${rec.kode_berkas}`}
                            className="inline-flex items-center gap-1 px-1.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-xs font-medium text-slate-700 transition-colors"
                          >
                            <Printer className="w-3 h-3" />
                            <span>Kwitansi</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => onDeleteRecord(rec.id)}
                            title="Hapus Berkas"
                            aria-label={`Hapus berkas ${rec.kode_berkas}`}
                            className="p-1 rounded-md border border-transparent hover:border-rose-200 hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
