import React, { useEffect, useMemo, useState } from 'react';
import {
  Calculator,
  FilePlus2,
  FolderCheck,
  Pencil,
  Printer,
  RotateCcw,
  Save,
  Search,
} from 'lucide-react';
import {
  BerkasRecord,
  OfficeSettings,
  ServiceType,
  TaxDraftPrefill,
} from '../types/ppats';
import { SERVICE_TYPES } from '../utils/seedData';
import {
  formatNumberId,
  formatRupiah,
  formatTanggalDDMMYYYY,
  terbilangLuas,
  terbilangRupiah,
} from '../utils/indonesianFormatter';
import {
  applyTaxCalculationToRecord,
  calculateLandTaxes,
  getDefaultNpoptkp,
  summarizeTaxBreakdown,
} from '../utils/taxAndAudit';

interface TaxCalculatorModuleProps {
  records: BerkasRecord[];
  settings: OfficeSettings;
  initialRecordId?: string | null;
  onSaveRecord: (record: BerkasRecord) => Promise<void>;
  onCreateBerkasWithTax: (prefill: TaxDraftPrefill) => void;
  onEditRecord: (record: BerkasRecord) => void;
  onNotify: (msg: string, type?: 'success' | 'error') => void;
}

type TaxMode = 'standalone' | 'linked';

function buildDraftFromRecord(
  record: BerkasRecord,
  settings: OfficeSettings
): TaxDraftPrefill {
  return {
    layanan: record.layanan,
    luas_m2: record.tanah?.luas_m2 ?? 0,
    njop_bumi_m2: record.pbb?.njop_bumi_m2 ?? 0,
    njop_bangunan: record.pbb?.njop_bangunan ?? 0,
    harga_per_m2: record.akta?.harga_per_m2 ?? 0,
    nilai_transaksi_override: record.akta?.nilai_transaksi_override ?? null,
    npoptkp:
      record.akta?.npoptkp ?? getDefaultNpoptkp(record.layanan, settings),
  };
}

function buildEmptyDraft(settings: OfficeSettings): TaxDraftPrefill {
  return {
    layanan: 'AJB',
    luas_m2: 0,
    njop_bumi_m2: 0,
    njop_bangunan: 0,
    harga_per_m2: 0,
    nilai_transaksi_override: null,
    npoptkp: getDefaultNpoptkp('AJB', settings),
  };
}

export const TaxCalculatorModule: React.FC<TaxCalculatorModuleProps> = ({
  records,
  settings,
  initialRecordId,
  onSaveRecord,
  onCreateBerkasWithTax,
  onEditRecord,
  onNotify,
}) => {
  const initialRecord = useMemo(
    () =>
      initialRecordId
        ? records.find((r) => r.id === initialRecordId) ?? null
        : null,
    [initialRecordId, records]
  );

  const [mode, setMode] = useState<TaxMode>(
    initialRecord ? 'linked' : 'standalone'
  );
  const [selectedRecordId, setSelectedRecordId] = useState<string>(
    initialRecord?.id ?? records[0]?.id ?? ''
  );
  const [form, setForm] = useState<TaxDraftPrefill>(() =>
    initialRecord
      ? buildDraftFromRecord(initialRecord, settings)
      : buildEmptyDraft(settings)
  );
  const [saving, setSaving] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  useEffect(() => {
    if (initialRecordId) {
      const target = records.find((r) => r.id === initialRecordId);
      if (target) {
        setMode('linked');
        setSelectedRecordId(target.id);
        setForm(buildDraftFromRecord(target, settings));
      }
    }
  }, [initialRecordId, records, settings]);

  const selectedRecord = useMemo(
    () => records.find((r) => r.id === selectedRecordId) ?? null,
    [records, selectedRecordId]
  );

  const taxResult = useMemo(
    () =>
      calculateLandTaxes({
        luas_m2: form.luas_m2,
        harga_per_m2: form.harga_per_m2,
        njop_bumi_m2: form.njop_bumi_m2,
        njop_bangunan: form.njop_bangunan,
        npoptkp: form.npoptkp,
        layanan: form.layanan,
        nilai_transaksi_override: form.nilai_transaksi_override,
      }),
    [form]
  );

  const taxSummary = useMemo(
    () => summarizeTaxBreakdown(taxResult),
    [taxResult]
  );

  const handleModeChange = (nextMode: TaxMode) => {
    setMode(nextMode);
    if (nextMode === 'linked') {
      const target = selectedRecord ?? records[0] ?? null;
      if (target) {
        setSelectedRecordId(target.id);
        setForm(buildDraftFromRecord(target, settings));
      }
    }
  };

  const handleSelectLinkedRecord = (recordId: string) => {
    setSelectedRecordId(recordId);
    const target = records.find((r) => r.id === recordId);
    if (target) {
      setMode('linked');
      setForm(buildDraftFromRecord(target, settings));
    }
  };

  const handleLoadRecordFromTable = (record: BerkasRecord) => {
    setMode('linked');
    setSelectedRecordId(record.id);
    setForm(buildDraftFromRecord(record, settings));
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleServiceChange = (nextLayanan: ServiceType) => {
    setForm((prev) => ({
      ...prev,
      layanan: nextLayanan,
      npoptkp: getDefaultNpoptkp(nextLayanan, settings),
    }));
  };

  const handleResetStandalone = () => {
    setMode('standalone');
    setForm(buildEmptyDraft(settings));
  };

  const handleSaveTaxToRecord = async () => {
    if (!selectedRecord) {
      onNotify('Pilih berkas register terlebih dahulu untuk menyimpan pajak.', 'error');
      return;
    }
    setSaving(true);
    try {
      const updated = applyTaxCalculationToRecord(selectedRecord, form);
      await onSaveRecord(updated);
      onNotify(
        `Perhitungan pajak berkas ${updated.kode_berkas} berhasil diperbarui (BPHTB ${formatRupiah(updated.pajak.bphtb)}, PPh ${formatRupiah(updated.pajak.pph)}).`,
        'success'
      );
    } catch (err) {
      onNotify(
        err instanceof Error
          ? err.message
          : 'Gagal menyimpan perhitungan pajak ke berkas.',
        'error'
      );
    } finally {
      setSaving(false);
    }
  };

  const filteredRecords = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return [...records]
      .sort((a, b) => b.no_urut - a.no_urut)
      .filter((r) => {
        if (!q) return true;
        const parties = r.para_pihak.map((p) => `${p.nama} ${p.nik}`).join(' ');
        const hay = [
          r.kode_berkas,
          r.akta.nomor_akta,
          r.layanan,
          r.tanah.desa,
          r.tanah.nib,
          parties,
        ]
          .join(' ')
          .toLowerCase();
        return hay.includes(q);
      });
  }, [records, searchQuery]);

  const todayFormatted = useMemo(
    () => formatTanggalDDMMYYYY(new Date().toISOString().slice(0, 10)),
    []
  );

  return (
    <div className="space-y-6">
      {/* Screen Header & Controls */}
      <div className="no-print flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
        <div>
          <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
            Perhitungan Pajak (BPHTB &amp; PPh Final)
          </h1>
          <p className="text-sm text-slate-600 mt-1">
            Jalur khusus perhitungan BPHTB (5%) dan PPh Final (2,5%), terpisah dari alur pembuatan berkas akta.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleResetStandalone}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 whitespace-nowrap"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Reset Simulasi</span>
          </button>
          <button
            type="button"
            onClick={() => window.print()}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 whitespace-nowrap"
          >
            <Printer className="w-3.5 h-3.5" />
            <span>Cetak Rincian Pajak</span>
          </button>
        </div>
      </div>

      {/* Mode Switcher Bar */}
      <div className="no-print bg-white border border-slate-200 rounded-xl p-4 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
        <div
          role="tablist"
          aria-label="Mode jalur perhitungan pajak"
          className="inline-flex items-center p-1 bg-slate-100 rounded-lg w-fit"
        >
          <button
            type="button"
            role="tab"
            aria-selected={mode === 'standalone'}
            onClick={() => handleModeChange('standalone')}
            className={`px-3.5 py-1.5 rounded-md text-xs font-semibold transition-colors whitespace-nowrap ${
              mode === 'standalone'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Simulasi Mandiri (Baru)
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mode === 'linked'}
            onClick={() => handleModeChange('linked')}
            disabled={records.length === 0}
            className={`px-3.5 py-1.5 rounded-md text-xs font-semibold transition-colors whitespace-nowrap disabled:opacity-50 ${
              mode === 'linked'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Kaitkan ke Berkas Terdaftar ({records.length})
          </button>
        </div>

        {mode === 'linked' && records.length > 0 && (
          <div className="flex flex-col sm:flex-row sm:items-center gap-2 flex-1 lg:max-w-xl">
            <label
              htmlFor="tax-linked-record"
              className="text-xs font-medium text-slate-600 whitespace-nowrap"
            >
              Pilih Berkas:
            </label>
            <select
              id="tax-linked-record"
              value={selectedRecordId}
              onChange={(e) => handleSelectLinkedRecord(e.target.value)}
              className="w-full px-3 py-2 border border-slate-200 rounded-lg bg-white text-xs font-medium text-slate-900 focus:outline-none focus:ring-2 focus:ring-sky-500"
            >
              {records
                .slice()
                .sort((a, b) => b.no_urut - a.no_urut)
                .map((r) => {
                  const p1 =
                    r.para_pihak
                      .filter((p) => p.peran === 'Pihak I')
                      .map((p) => p.nama)
                      .join(', ') || '-';
                  return (
                    <option key={r.id} value={r.id}>
                      {r.kode_berkas} · {r.layanan} · {p1} ({formatNumberId(r.tanah.luas_m2)} m²)
                    </option>
                  );
                })}
            </select>
          </div>
        )}
      </div>

      {/* Main 2-Column Calculator Workspace */}
      <div className="no-print grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Left: Input Parameter Tanah, NJOP, Kesepakatan, NPOPTKP */}
        <div className="lg:col-span-6 bg-white border border-slate-200 rounded-xl p-6 space-y-5">
          <div className="flex items-center justify-between border-b border-slate-200 pb-3">
            <div className="flex items-center gap-2">
              <Calculator className="w-4 h-4 text-slate-700" />
              <h2 className="font-display text-base font-semibold text-slate-900">
                Parameter Objek &amp; Nilai Transaksi
              </h2>
            </div>
            {mode === 'linked' && selectedRecord && (
              <span className="font-mono text-xs font-semibold text-sky-800 bg-sky-50 border border-sky-200 px-2.5 py-1 rounded-md">
                {selectedRecord.kode_berkas}
              </span>
            )}
          </div>

          <div className="space-y-4 text-xs">
            <div>
              <label className="block text-slate-600 font-medium mb-1.5">
                Jenis Layanan Peralihan Hak
              </label>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {SERVICE_TYPES.map((type) => {
                  const active = form.layanan === type;
                  return (
                    <button
                      key={type}
                      type="button"
                      onClick={() => handleServiceChange(type)}
                      className={`px-3 py-2 rounded-lg border text-xs font-semibold text-left transition-colors ${
                        active
                          ? 'border-sky-600 bg-sky-50 text-sky-900'
                          : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      {type}
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label
                  htmlFor="tax-luas"
                  className="block text-slate-600 font-medium mb-1"
                >
                  Luas Bidang Tanah (m²)
                </label>
                <input
                  id="tax-luas"
                  type="number"
                  min="0"
                  value={form.luas_m2}
                  onChange={(e) =>
                    setForm({ ...form, luas_m2: Number(e.target.value) || 0 })
                  }
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-sky-500"
                />
                <p className="text-[11px] text-slate-500 mt-1 italic">
                  {terbilangLuas(form.luas_m2)}
                </p>
              </div>

              <div>
                <label
                  htmlFor="tax-njop-bumi-m2"
                  className="block text-slate-600 font-medium mb-1"
                >
                  NJOP Bumi per m² (Rp)
                </label>
                <input
                  id="tax-njop-bumi-m2"
                  type="number"
                  min="0"
                  value={form.njop_bumi_m2}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      njop_bumi_m2: Number(e.target.value) || 0,
                    })
                  }
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-sky-500"
                />
              </div>

              <div>
                <label
                  htmlFor="tax-njop-bangunan"
                  className="block text-slate-600 font-medium mb-1"
                >
                  Total NJOP Bangunan (Rp)
                </label>
                <input
                  id="tax-njop-bangunan"
                  type="number"
                  min="0"
                  value={form.njop_bangunan}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      njop_bangunan: Number(e.target.value) || 0,
                    })
                  }
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-sky-500"
                />
              </div>

              <div>
                <label
                  htmlFor="tax-harga-m2"
                  className="block text-slate-600 font-medium mb-1"
                >
                  Harga Kesepakatan per m² (Rp)
                </label>
                <input
                  id="tax-harga-m2"
                  type="number"
                  min="0"
                  value={form.harga_per_m2}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      harga_per_m2: Number(e.target.value) || 0,
                    })
                  }
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-sky-500"
                />
              </div>

              <div className="sm:col-span-2">
                <label
                  htmlFor="tax-override"
                  className="block text-slate-600 font-medium mb-1"
                >
                  Total Kesepakatan Manual (Opsional · kosongkan untuk otomatis Luas × Harga/m²)
                </label>
                <input
                  id="tax-override"
                  type="number"
                  min="0"
                  value={form.nilai_transaksi_override ?? ''}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      nilai_transaksi_override:
                        e.target.value === '' ? null : Number(e.target.value),
                    })
                  }
                  placeholder="Otomatis dari Luas × Harga per m²"
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-sky-500"
                />
              </div>

              <div className="sm:col-span-2">
                <div className="flex items-center justify-between mb-1">
                  <label
                    htmlFor="tax-npoptkp"
                    className="text-slate-600 font-medium"
                  >
                    NPOPTKP (Rp · Default{' '}
                    {form.layanan === 'Turun Waris'
                      ? `Waris ${formatRupiah(settings.npoptkp_waris)}`
                      : `Umum ${formatRupiah(settings.npoptkp_umum)}`}
                    )
                  </label>
                  <button
                    type="button"
                    onClick={() =>
                      setForm({
                        ...form,
                        npoptkp: getDefaultNpoptkp(form.layanan, settings),
                      })
                    }
                    className="text-sky-700 hover:underline font-medium"
                  >
                    Reset Default
                  </button>
                </div>
                <input
                  id="tax-npoptkp"
                  type="number"
                  min="0"
                  value={form.npoptkp}
                  onChange={(e) =>
                    setForm({ ...form, npoptkp: Number(e.target.value) || 0 })
                  }
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono tabular-nums focus:outline-none focus:ring-2 focus:ring-sky-500"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Right: Hasil Perhitungan BPHTB & PPh Final */}
        <div className="lg:col-span-6 bg-white border border-slate-200 rounded-xl p-6 flex flex-col justify-between space-y-5">
          <div className="space-y-4">
            <div className="flex items-center justify-between border-b border-slate-200 pb-3">
              <h2 className="font-display text-base font-semibold text-slate-900">
                Rincian BPHTB &amp; PPh Final
              </h2>
              <span className="text-xs font-medium text-slate-500">
                Basis NPOP: <strong className="text-slate-800">{taxSummary.dasar_npop_sumber}</strong>
              </span>
            </div>

            <div className="divide-y divide-slate-200 text-xs border border-slate-200 rounded-xl overflow-hidden">
              <div className="px-4 py-2.5 bg-slate-50/70 flex items-center justify-between gap-2">
                <span className="text-slate-600">NJOP Bumi Total ({formatNumberId(form.luas_m2)} m² × {formatRupiah(form.njop_bumi_m2)})</span>
                <span className="font-mono font-medium text-slate-900 tabular-nums whitespace-nowrap">
                  {formatRupiah(taxResult.njop_bumi_total)}
                </span>
              </div>
              <div className="px-4 py-2.5 flex items-center justify-between gap-2">
                <span className="text-slate-600">Total NJOP PBB (Bumi + Bangunan)</span>
                <span className="font-mono font-medium text-slate-900 tabular-nums whitespace-nowrap">
                  {formatRupiah(taxResult.njop_total)}
                </span>
              </div>
              <div className="px-4 py-2.5 bg-slate-50/70 space-y-0.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-slate-600">Nilai Kesepakatan Transaksi</span>
                  <span className="font-mono font-semibold text-slate-900 tabular-nums whitespace-nowrap">
                    {formatRupiah(taxResult.nilai_kesepakatan)}
                  </span>
                </div>
                <p className="text-[11px] text-slate-500 italic">
                  Terbilang: {terbilangRupiah(taxResult.nilai_kesepakatan)}
                </p>
              </div>
              <div className="px-4 py-2.5 flex items-center justify-between gap-2">
                <span className="font-semibold text-slate-800">
                  NPOP (Nilai Tertinggi Kesepakatan vs NJOP)
                </span>
                <span className="font-mono font-bold text-slate-900 tabular-nums whitespace-nowrap">
                  {formatRupiah(taxResult.npop)}
                </span>
              </div>
              <div className="px-4 py-2.5 bg-slate-50/70 flex items-center justify-between gap-2">
                <span className="text-slate-600">Pengurang NPOPTKP</span>
                <span className="font-mono font-medium text-slate-700 tabular-nums whitespace-nowrap">
                  − {formatRupiah(taxResult.npoptkp)}
                </span>
              </div>
              <div className="px-4 py-2.5 flex items-center justify-between gap-2">
                <span className="font-semibold text-slate-800">
                  NPOPKP (Dasar Hitung BPHTB)
                </span>
                <span className="font-mono font-semibold text-slate-900 tabular-nums whitespace-nowrap">
                  {formatRupiah(taxResult.npopkp)}
                </span>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div className="p-4 rounded-xl border border-sky-200 bg-sky-50/70 space-y-1.5">
                <p className="font-semibold text-sky-900">
                  BPHTB Terutang (5% × NPOPKP)
                </p>
                <p className="font-mono text-lg font-bold text-sky-950 tabular-nums">
                  {formatRupiah(taxResult.bphtb)}
                </p>
                <p className="text-[11px] text-sky-800 italic">
                  {terbilangRupiah(taxResult.bphtb)}
                </p>
              </div>

              <div className="p-4 rounded-xl border border-amber-200 bg-amber-50/70 space-y-1.5">
                <p className="font-semibold text-amber-900">
                  PPh Final (2,5% × NPOP)
                </p>
                <p className="font-mono text-lg font-bold text-amber-950 tabular-nums">
                  {formatRupiah(taxResult.pph)}
                </p>
                <p className="text-[11px] text-amber-800 italic">
                  {terbilangRupiah(taxResult.pph)}
                </p>
              </div>
            </div>

            <div className="px-4 py-3 rounded-xl border border-slate-200 bg-slate-50 flex items-center justify-between text-xs">
              <span className="font-semibold text-slate-700">
                Total Estimasi Pajak (BPHTB + PPh Final)
              </span>
              <span className="font-mono text-base font-bold text-slate-900 tabular-nums">
                {formatRupiah(taxSummary.total_pajak)}
              </span>
            </div>
          </div>

          {/* Workflow Bridge Actions */}
          <div className="pt-3 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2.5">
            {mode === 'linked' && selectedRecord ? (
              <>
                <button
                  type="button"
                  onClick={() => onEditRecord(selectedRecord)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 whitespace-nowrap"
                >
                  <Pencil className="w-3.5 h-3.5" />
                  <span>Buka Berkas di Pembuatan Akta</span>
                </button>
                <button
                  type="button"
                  onClick={handleSaveTaxToRecord}
                  disabled={saving}
                  className="inline-flex items-center gap-2 px-4 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-semibold transition-colors disabled:opacity-60 whitespace-nowrap"
                >
                  <Save className="w-3.5 h-3.5 text-amber-400" />
                  <span>
                    {saving
                      ? 'Menyimpan...'
                      : `Simpan Pajak ke Berkas ${selectedRecord.kode_berkas}`}
                  </span>
                </button>
              </>
            ) : (
              <>
                <span className="text-[11px] text-slate-500">
                  Hasil simulasi ini dapat langsung dipakai untuk membuka berkas baru.
                </span>
                <button
                  type="button"
                  onClick={() => onCreateBerkasWithTax(form)}
                  className="inline-flex items-center gap-2 px-4 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-semibold transition-colors whitespace-nowrap"
                >
                  <FilePlus2 className="w-3.5 h-3.5 text-amber-400" />
                  <span>Buat Berkas Baru dengan Hasil Pajak Ini</span>
                </button>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Bottom Registered Records Quick Tax Table */}
      <div className="no-print bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-4 sm:p-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div className="flex items-center gap-2">
            <FolderCheck className="w-4 h-4 text-slate-700" />
            <div>
              <h2 className="font-display text-sm font-semibold text-slate-900">
                Daftar Pajak Berkas Terdaftar
              </h2>
              <p className="text-xs text-slate-500">
                Klik &ldquo;Hitung / Ubah&rdquo; pada berkas mana saja untuk memuat ulang dan memperbarui pajaknya.
              </p>
            </div>
          </div>

          <div className="relative w-full sm:w-64">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="search"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Cari kode berkas, pihak, desa..."
              aria-label="Cari daftar pajak berkas"
              className="w-full pl-9 pr-3 py-1.5 text-xs border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
        </div>

        {filteredRecords.length === 0 ? (
          <div className="p-8 text-center text-xs text-slate-500">
            Tidak ada berkas yang cocok dengan pencarian Anda.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 font-semibold text-slate-600">
                  <th className="py-2.5 px-3 whitespace-nowrap">Kode &amp; Layanan</th>
                  <th className="py-2.5 px-3">Para Pihak</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Luas</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">NPOP</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">BPHTB (5%)</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">PPh (2,5%)</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredRecords.map((r) => {
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
                  const isSelected = mode === 'linked' && selectedRecordId === r.id;
                  return (
                    <tr
                      key={r.id}
                      className={
                        isSelected
                          ? 'bg-sky-50/70'
                          : 'hover:bg-slate-50/80 transition-colors'
                      }
                    >
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        <div className="font-mono font-semibold text-slate-900">
                          {r.kode_berkas}
                        </div>
                        <div className="text-[11px] text-slate-500">
                          {r.layanan}
                        </div>
                      </td>
                      <td className="py-2.5 px-3 max-w-[200px]">
                        <div className="truncate text-slate-800" title={p1}>
                          I: {p1}
                        </div>
                        <div className="truncate text-slate-500" title={p2}>
                          II: {p2}
                        </div>
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums whitespace-nowrap">
                        {formatNumberId(r.tanah.luas_m2)} m²
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums whitespace-nowrap text-slate-900">
                        {formatRupiah(r.pajak.npop)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums whitespace-nowrap font-medium text-sky-900">
                        {formatRupiah(r.pajak.bphtb)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums whitespace-nowrap font-medium text-amber-900">
                        {formatRupiah(r.pajak.pph)}
                      </td>
                      <td className="py-2.5 px-3 text-right whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => handleLoadRecordFromTable(r)}
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-xs font-medium text-slate-700"
                        >
                          <Calculator className="w-3.5 h-3.5" />
                          <span>Hitung / Ubah</span>
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

      {/* Print-Only Official Summary Sheet */}
      <div className="hidden print:block p-6 text-slate-900 text-xs space-y-4">
        <div className="border-b-2 border-slate-900 pb-3 text-center">
          <p className="font-bold text-sm uppercase">
            KANTOR PEJABAT PEMBUAT AKTA TANAH SEMENTARA (PPATS)
          </p>
          <p className="font-bold text-sm uppercase">{settings.nama_kecamatan}</p>
          <p>{settings.alamat_kantor} · Telp. {settings.telepon_kantor}</p>
        </div>
        <div className="text-center">
          <h2 className="font-bold text-sm uppercase underline">
            LEMBAR PERHITUNGAN ESTIMASI PAJAK PERTANAHAN (BPHTB &amp; PPh FINAL)
          </h2>
          {mode === 'linked' && selectedRecord && (
            <p className="mt-1 font-mono">
              Berkas: {selectedRecord.kode_berkas} · Layanan: {form.layanan}
            </p>
          )}
        </div>
        <table className="w-full border-collapse border border-slate-400">
          <tbody>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-medium w-1/2">
                Jenis Layanan
              </td>
              <td className="border border-slate-400 px-3 py-2">{form.layanan}</td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-medium">
                Luas Bidang Tanah
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono">
                {formatNumberId(form.luas_m2)} m² ({terbilangLuas(form.luas_m2)})
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-medium">
                Total NJOP PBB (Bumi + Bangunan)
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono">
                {formatRupiah(taxResult.njop_total)}
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-medium">
                Nilai Kesepakatan Transaksi
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono">
                {formatRupiah(taxResult.nilai_kesepakatan)}
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-semibold">
                NPOP (Nilai Tertinggi)
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono font-semibold">
                {formatRupiah(taxResult.npop)}
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-medium">
                NPOPTKP
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono">
                {formatRupiah(taxResult.npoptkp)}
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-medium">
                NPOPKP
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono">
                {formatRupiah(taxResult.npopkp)}
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-bold">
                BPHTB Terutang (5% × NPOPKP)
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono font-bold">
                {formatRupiah(taxResult.bphtb)} ({terbilangRupiah(taxResult.bphtb)})
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-bold">
                PPh Final (2,5% × NPOP)
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono font-bold">
                {formatRupiah(taxResult.pph)} ({terbilangRupiah(taxResult.pph)})
              </td>
            </tr>
            <tr>
              <td className="border border-slate-400 px-3 py-2 font-bold">
                Total Pajak (BPHTB + PPh)
              </td>
              <td className="border border-slate-400 px-3 py-2 font-mono font-bold">
                {formatRupiah(taxSummary.total_pajak)}
              </td>
            </tr>
          </tbody>
        </table>
        <div className="pt-6 text-right">
          <p>{settings.nama_kecamatan}, {todayFormatted}</p>
          <p className="mt-12 font-semibold">{settings.nama_ppats}</p>
        </div>
      </div>
    </div>
  );
};
