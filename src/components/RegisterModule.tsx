import React, { useMemo, useState } from 'react';
import { Search, Pencil } from 'lucide-react';
import { BerkasRecord, ServiceType } from '../types/ppats';
import { SERVICE_TYPES } from '../utils/seedData';
import {
  extractMonthYearFromDate,
  formatNumberId,
  formatRupiah,
  formatTanggalDDMMYYYY,
} from '../utils/indonesianFormatter';
import { isAktaService } from '../utils/taxAndAudit';

interface RegisterModuleProps {
  records: BerkasRecord[];
  onEditRecord: (record: BerkasRecord) => void;
}

export const RegisterModule: React.FC<RegisterModuleProps> = ({
  records,
  onEditRecord,
}) => {
  const [query, setQuery] = useState('');
  const [filterKategori, setFilterKategori] = useState<'SEMUA' | 'AKTA' | 'NON_AKTA'>('SEMUA');
  const [filterLayanan, setFilterLayanan] = useState<'SEMUA' | ServiceType>('SEMUA');
  const [filterTahun, setFilterTahun] = useState<string>('SEMUA');

  const availableYears = useMemo(() => {
    const years = new Set<string>();
    for (const r of records) {
      const my = extractMonthYearFromDate(r.akta.tanggal_akta || r.tanggal_masuk);
      if (my) years.add(String(my.year));
    }
    return Array.from(years).sort().reverse();
  }, [records]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return [...records]
      .sort((a, b) => a.no_urut - b.no_urut)
      .filter((r) => {
        if (filterKategori === 'AKTA' && !isAktaService(r.layanan)) return false;
        if (filterKategori === 'NON_AKTA' && isAktaService(r.layanan)) return false;
        if (filterLayanan !== 'SEMUA' && r.layanan !== filterLayanan) return false;
        const my = extractMonthYearFromDate(r.akta.tanggal_akta || r.tanggal_masuk);
        const year = my ? String(my.year) : '';
        if (filterTahun !== 'SEMUA' && year !== filterTahun) return false;
        if (!q) return true;
        const partyText = r.para_pihak
          .map((p) => `${p.nama} ${p.nik}`)
          .join(' ')
          .toLowerCase();
        const hay = [
          r.akta.nomor_akta,
          r.kode_berkas,
          r.layanan,
          r.tanah.nib,
          r.tanah.nomor_hak_lama,
          r.tanah.desa,
          partyText,
        ]
          .join(' ')
          .toLowerCase();
        return hay.includes(q);
      });
  }, [records, query, filterKategori, filterLayanan, filterTahun]);

  return (
    <div className="space-y-6">
      <div>
        <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
          Buku Register Akta &amp; Berkas (Repertorium)
        </h1>
        <p className="text-sm text-slate-600 mt-1">
          Catatan kronologis seluruh Akta PPATS dan Berkas Non-Akta (Turun Waris, Pemecahan, Ukur Ulang, Ralat Data) beserta NIB, luas bidang, nilai transaksi, serta BPHTB &amp; PPh.
        </p>
      </div>

      {/* Filter Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col lg:flex-row gap-3 lg:items-center lg:justify-between">
        <div className="flex flex-wrap items-center gap-3 flex-1">
          <div className="inline-flex items-center p-1 bg-slate-100 rounded-lg shrink-0">
            {[
              { id: 'SEMUA' as const, label: 'Semua Berkas' },
              { id: 'AKTA' as const, label: 'Akta PPATS' },
              { id: 'NON_AKTA' as const, label: 'Non-Akta' },
            ].map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setFilterKategori(tab.id)}
                className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors whitespace-nowrap ${
                  filterKategori === tab.id
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          <div className="relative flex-1 min-w-[200px] max-w-md">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="search"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Cari Nama, NIK, NIB, atau Nomor Akta..."
              aria-label="Cari di Buku Register Akta"
              className="w-full pl-9 pr-3 py-2 text-sm border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center gap-2">
            <label htmlFor="filter-jenis-akta" className="text-xs font-medium text-slate-600">
              Jenis Layanan
            </label>
            <select
              id="filter-jenis-akta"
              value={filterLayanan}
              onChange={(e) => setFilterLayanan(e.target.value as 'SEMUA' | ServiceType)}
              className="text-sm border border-slate-200 rounded-lg px-3 py-1.5 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            >
              <option value="SEMUA">Semua Jenis</option>
              {SERVICE_TYPES.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-2">
            <label htmlFor="filter-tahun-akta" className="text-xs font-medium text-slate-600">
              Tahun
            </label>
            <select
              id="filter-tahun-akta"
              value={filterTahun}
              onChange={(e) => setFilterTahun(e.target.value)}
              className="text-sm border border-slate-200 rounded-lg px-3 py-1.5 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            >
              <option value="SEMUA">Semua Tahun</option>
              {availableYears.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* Repertorium Table */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        {filtered.length === 0 ? (
          <div className="p-10 text-center text-sm text-slate-500">
            Belum ada catatan akta yang cocok dengan pencarian atau filter.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-[11px]">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 font-semibold text-slate-700">
                  <th className="py-2.5 px-1.5 whitespace-nowrap">No</th>
                  <th className="py-2.5 px-1.5 whitespace-nowrap">No. & Tgl Akta</th>
                  <th className="py-2.5 px-1.5 whitespace-nowrap">Jenis</th>
                  <th className="py-2.5 px-1.5">Pihak I</th>
                  <th className="py-2.5 px-1.5">Pihak II</th>
                  <th className="py-2.5 px-1.5">Hak & NIB</th>
                  <th className="py-2.5 px-1.5">Letak Tanah</th>
                  <th className="py-2.5 px-1.5 text-right whitespace-nowrap">Luas</th>
                  <th className="py-2.5 px-1.5 text-right whitespace-nowrap">Nilai Transaksi</th>
                  <th className="py-2.5 px-1.5 text-right whitespace-nowrap">BPHTB & PPh</th>
                  <th className="py-2.5 px-1.5 whitespace-nowrap">Status</th>
                  <th className="py-2.5 px-1.5 text-right whitespace-nowrap">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filtered.map((r) => {
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
                    : 'Belum bertanggal';

                  return (
                    <tr key={r.id} className="hover:bg-slate-50 transition-colors">
                      <td className="py-2.5 px-1.5 font-mono font-semibold text-slate-900 tabular-nums">
                        {String(r.no_urut).padStart(3, '0')}
                      </td>
                      <td className="py-2.5 px-1.5 whitespace-nowrap">
                        <div className="font-mono font-medium text-slate-900">
                          {isAktaService(r.layanan)
                            ? r.akta.nomor_akta || '-'
                            : 'Non-Akta (F4)'}
                        </div>
                        <div className="font-mono text-slate-500 mt-0.5 tabular-nums">{tgl}</div>
                      </td>
                      <td className="py-2.5 px-1.5 font-medium text-slate-900 whitespace-nowrap">
                        {r.layanan}
                      </td>
                      <td className="py-2.5 px-1.5 text-slate-800 max-w-[95px] truncate" title={p1}>
                        {p1}
                      </td>
                      <td className="py-2.5 px-1.5 text-slate-800 max-w-[95px] truncate" title={p2}>
                        {p2}
                      </td>
                      <td className="py-2.5 px-1.5 max-w-[140px]">
                        <div className="text-slate-900 truncate" title={`${r.tanah.jenis_hak} ${r.tanah.nomor_hak_lama}`}>
                          {r.tanah.jenis_sertipikat === 'digital' || !r.tanah.nomor_hak_lama
                            ? `${r.tanah.jenis_hak} (Digital)`
                            : `${r.tanah.jenis_hak} No. ${r.tanah.nomor_hak_lama}`}
                        </div>
                        <div className="font-mono text-slate-500 mt-0.5 truncate">
                          {r.tanah.jenis_sertipikat === 'digital' || !r.tanah.nomor_hak_lama ? 'NIBel' : 'NIB'}: {r.tanah.nib || '-'}
                        </div>
                      </td>
                      <td className="py-2.5 px-1.5 text-slate-700 max-w-[100px] truncate" title={`${r.tanah.desa}, ${r.tanah.kecamatan}`}>
                        {r.tanah.desa}, {r.tanah.kecamatan}
                      </td>
                      <td className="py-2.5 px-1.5 text-right font-mono tabular-nums text-slate-900 whitespace-nowrap">
                        {formatNumberId(r.tanah.luas_m2)} m²
                      </td>
                      <td className="py-2.5 px-1.5 text-right font-mono tabular-nums text-slate-900 whitespace-nowrap">
                        {formatRupiah(r.pajak.nilai_kesepakatan)}
                      </td>
                      <td className="py-2.5 px-1.5 text-right font-mono tabular-nums whitespace-nowrap">
                        <div className="text-slate-900">
                          {formatRupiah(r.pajak.bphtb)}
                        </div>
                        <div className="text-slate-500 mt-0.5">
                          PPh {formatRupiah(r.pajak.pph)}
                        </div>
                      </td>
                      <td className="py-2.5 px-1.5 text-slate-700 whitespace-nowrap">
                        {r.status}
                      </td>
                      <td className="py-2.5 px-1.5 text-right whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => onEditRecord(r)}
                          className="inline-flex items-center gap-1 px-2 py-1 rounded border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 font-medium"
                        >
                          <Pencil className="w-3 h-3" />
                          <span>Buka</span>
                        </button>
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
