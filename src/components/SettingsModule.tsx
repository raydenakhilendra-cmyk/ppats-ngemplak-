import React, { useState } from 'react';
import {
  Save,
  Upload,
  Download,
  Trash2,
  FileText,
  Database,
  CheckCircle2,
  Plus,
  X,
} from 'lucide-react';
import {
  CustomTemplateEntry,
  OfficeSettings,
  PersonBasicInfo,
  TemplateDocKey,
} from '../types/ppats';
import { DEFAULT_OFFICE_SETTINGS, TEMPLATE_DOCS_META } from '../utils/seedData';
import {
  downloadDefaultTemplateDocx,
  uint8ArrayToBase64,
} from '../utils/bundleGenerator';
import { SUPPORTED_TEMPLATE_TAGS } from '../utils/docxTemplates';

interface SettingsModuleProps {
  settings: OfficeSettings;
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>;
  onSaveSettings: (next: OfficeSettings) => Promise<void>;
  onUploadTemplate: (entry: CustomTemplateEntry) => Promise<void>;
  onDeleteTemplate: (key: TemplateDocKey) => Promise<void>;
  onExportBackup: () => void;
  onRestoreBackup: (file: File) => void;
  onNotify: (msg: string, type?: 'success' | 'error') => void;
}

function syncProfilesWithNames(
  names: string[],
  existing?: PersonBasicInfo[],
  fallback?: PersonBasicInfo[]
): PersonBasicInfo[] {
  const pool = [...(existing ?? []), ...(fallback ?? [])];
  return names.map((nama) => {
    const match = pool.find(
      (p) => p.nama.trim().toLowerCase() === nama.trim().toLowerCase()
    );
    return match ? { ...match, nama } : { nama };
  });
}

export const SettingsModule: React.FC<SettingsModuleProps> = ({
  settings,
  customTemplates,
  onSaveSettings,
  onUploadTemplate,
  onDeleteTemplate,
  onExportBackup,
  onRestoreBackup,
  onNotify,
}) => {
  const [form, setForm] = useState<OfficeSettings>(() => ({
    ...DEFAULT_OFFICE_SETTINGS,
    ...settings,
    daftar_saksi: Array.isArray(settings.daftar_saksi) ? settings.daftar_saksi : [],
    daftar_kuasa_bpn: Array.isArray(settings.daftar_kuasa_bpn) ? settings.daftar_kuasa_bpn : [],
    profil_saksi: syncProfilesWithNames(
      Array.isArray(settings.daftar_saksi) ? settings.daftar_saksi : [],
      settings.profil_saksi,
      DEFAULT_OFFICE_SETTINGS.profil_saksi
    ),
    profil_kuasa: syncProfilesWithNames(
      Array.isArray(settings.daftar_kuasa_bpn) ? settings.daftar_kuasa_bpn : [],
      settings.profil_kuasa,
      DEFAULT_OFFICE_SETTINGS.profil_kuasa
    ),
  }));
  const [newSaksi, setNewSaksi] = useState('');
  const [newKuasa, setNewKuasa] = useState('');
  const [expandedSaksiIdx, setExpandedSaksiIdx] = useState<number | null>(null);
  const [expandedKuasaIdx, setExpandedKuasaIdx] = useState<number | null>(null);
  const [saving, setSaving] = useState(false);

  React.useEffect(() => {
    const saksiNames = Array.isArray(settings.daftar_saksi) ? settings.daftar_saksi : [];
    const kuasaNames = Array.isArray(settings.daftar_kuasa_bpn) ? settings.daftar_kuasa_bpn : [];
    setForm({
      ...DEFAULT_OFFICE_SETTINGS,
      ...settings,
      daftar_saksi: saksiNames,
      daftar_kuasa_bpn: kuasaNames,
      profil_saksi: syncProfilesWithNames(
        saksiNames,
        settings.profil_saksi,
        DEFAULT_OFFICE_SETTINGS.profil_saksi
      ),
      profil_kuasa: syncProfilesWithNames(
        kuasaNames,
        settings.profil_kuasa,
        DEFAULT_OFFICE_SETTINGS.profil_kuasa
      ),
    });
  }, [settings]);

  const restoreInputRef = React.useRef<HTMLInputElement | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setSaving(true);
    try {
      await onSaveSettings(form);
      onNotify('Pengaturan profil kantor PPATS berhasil disimpan.', 'success');
    } catch (err) {
      onNotify(
        err instanceof Error ? err.message : 'Gagal menyimpan pengaturan.',
        'error'
      );
    } finally {
      setSaving(false);
    }
  };

  const handleTemplateUpload = async (
    key: TemplateDocKey,
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';

    if (!file.name.toLowerCase().endsWith('.docx')) {
      onNotify('Harap unggah file Microsoft Word berformat .docx.', 'error');
      return;
    }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const base64 = uint8ArrayToBase64(new Uint8Array(arrayBuffer));
      await onUploadTemplate({
        key,
        fileName: file.name,
        updatedAt: new Date().toISOString(),
        dataBase64: base64,
      });
      onNotify(`Template kustom "${file.name}" tersimpan.`, 'success');
    } catch (err) {
      onNotify(
        err instanceof Error ? err.message : 'Gagal membaca file .docx.',
        'error'
      );
    }
  };

  const addSaksi = async () => {
    const nextIdx = form.daftar_saksi.length;
    const val = newSaksi.trim() || `Saksi Baru ${nextIdx + 1}`;
    const nextForm: OfficeSettings = {
      ...form,
      daftar_saksi: [...form.daftar_saksi, val],
      profil_saksi: [...(form.profil_saksi ?? []), { nama: val }],
    };
    setForm(nextForm);
    setExpandedSaksiIdx(nextIdx);
    setNewSaksi('');
    try {
      await onSaveSettings(nextForm);
      onNotify(`Saksi "${val}" ditambahkan. Silakan lengkapi detailnya.`, 'success');
    } catch (err) {
      onNotify(err instanceof Error ? err.message : 'Gagal menambahkan saksi.', 'error');
    }
  };

  const removeSaksi = async (idx: number) => {
    const removedName = form.daftar_saksi[idx];
    const nextForm: OfficeSettings = {
      ...form,
      daftar_saksi: form.daftar_saksi.filter((_, i) => i !== idx),
      profil_saksi: (form.profil_saksi ?? []).filter((_, i) => i !== idx),
    };
    setForm(nextForm);
    if (expandedSaksiIdx === idx) setExpandedSaksiIdx(null);
    try {
      await onSaveSettings(nextForm);
      if (removedName) onNotify(`Saksi "${removedName}" dihapus.`, 'success');
    } catch (err) {
      onNotify(err instanceof Error ? err.message : 'Gagal menghapus saksi.', 'error');
    }
  };

  const updateSaksiProfile = (idx: number, patch: Partial<PersonBasicInfo>) => {
    setForm((prev) => {
      const profiles = [...(prev.profil_saksi ?? [])];
      const current = profiles[idx] ?? { nama: prev.daftar_saksi[idx] ?? '' };
      const updated = { ...current, ...patch };
      profiles[idx] = updated;
      const names = [...prev.daftar_saksi];
      if (patch.nama !== undefined) names[idx] = patch.nama;
      return { ...prev, daftar_saksi: names, profil_saksi: profiles };
    });
  };

  const addKuasa = async () => {
    const nextIdx = form.daftar_kuasa_bpn.length;
    const val = newKuasa.trim() || `Staf Kuasa ${nextIdx + 1}`;
    const nextForm: OfficeSettings = {
      ...form,
      daftar_kuasa_bpn: [...form.daftar_kuasa_bpn, val],
      profil_kuasa: [...(form.profil_kuasa ?? []), { nama: val }],
    };
    setForm(nextForm);
    setExpandedKuasaIdx(nextIdx);
    setNewKuasa('');
    try {
      await onSaveSettings(nextForm);
      onNotify(`Staf kuasa BPN "${val}" ditambahkan. Silakan lengkapi detailnya.`, 'success');
    } catch (err) {
      onNotify(err instanceof Error ? err.message : 'Gagal menambahkan staf kuasa BPN.', 'error');
    }
  };

  const removeKuasa = async (idx: number) => {
    const removedName = form.daftar_kuasa_bpn[idx];
    const nextForm: OfficeSettings = {
      ...form,
      daftar_kuasa_bpn: form.daftar_kuasa_bpn.filter((_, i) => i !== idx),
      profil_kuasa: (form.profil_kuasa ?? []).filter((_, i) => i !== idx),
    };
    setForm(nextForm);
    if (expandedKuasaIdx === idx) setExpandedKuasaIdx(null);
    try {
      await onSaveSettings(nextForm);
      if (removedName) onNotify(`Staf kuasa "${removedName}" dihapus.`, 'success');
    } catch (err) {
      onNotify(err instanceof Error ? err.message : 'Gagal menghapus staf kuasa.', 'error');
    }
  };

  const updateKuasaProfile = (idx: number, patch: Partial<PersonBasicInfo>) => {
    setForm((prev) => {
      const profiles = [...(prev.profil_kuasa ?? [])];
      const current = profiles[idx] ?? { nama: prev.daftar_kuasa_bpn[idx] ?? '' };
      const updated = { ...current, ...patch };
      profiles[idx] = updated;
      const names = [...prev.daftar_kuasa_bpn];
      if (patch.nama !== undefined) names[idx] = patch.nama;
      return { ...prev, daftar_kuasa_bpn: names, profil_kuasa: profiles };
    });
  };

  return (
    <div className="space-y-8">
      <div>
        <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
          Pengaturan Kantor & Master Template .DOCX
        </h1>
        <p className="text-sm text-slate-600 mt-1">
          Kelola identitas pejabat PPATS, daftar saksi & kuasa tetap, master template Word kantor, serta backup & restore database.
        </p>
      </div>

      {/* Section 1: Profil Kantor */}
      <form
        onSubmit={handleSubmit}
        className="bg-white border border-slate-200 rounded-xl p-6 space-y-6"
      >
        <div className="flex items-center justify-between border-b border-slate-200 pb-4">
          <div>
            <h2 className="font-display text-base font-semibold text-slate-900">
              Profil Kantor & Pejabat PPATS Default
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Otomatis digunakan sebagai nilai bawaan pada Wizard pembuatan akta dan kop laporan bulanan.
            </p>
          </div>
          <button
            type="submit"
            disabled={saving}
            className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium transition-colors whitespace-nowrap disabled:opacity-60"
          >
            <Save className="w-4 h-4" />
            <span>{saving ? 'Menyimpan...' : 'Simpan Pengaturan'}</span>
          </button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label htmlFor="set-kecamatan" className="block text-xs font-medium text-slate-700 mb-1">
              Nama Kecamatan
            </label>
            <input
              id="set-kecamatan"
              type="text"
              value={form.nama_kecamatan}
              onChange={(e) => setForm({ ...form, nama_kecamatan: e.target.value })}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-kabupaten" className="block text-xs font-medium text-slate-700 mb-1">
              Nama Kabupaten / Kota
            </label>
            <input
              id="set-kabupaten"
              type="text"
              value={form.nama_kabupaten}
              onChange={(e) => setForm({ ...form, nama_kabupaten: e.target.value })}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-ppats" className="block text-xs font-medium text-slate-700 mb-1">
              Nama Camat / PPATS Menjabat
            </label>
            <input
              id="set-ppats"
              type="text"
              value={form.nama_ppats}
              onChange={(e) => setForm({ ...form, nama_ppats: e.target.value })}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-jabatan" className="block text-xs font-medium text-slate-700 mb-1">
              Jabatan PPATS
            </label>
            <input
              id="set-jabatan"
              type="text"
              value={form.jabatan_ppats}
              onChange={(e) => setForm({ ...form, jabatan_ppats: e.target.value })}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-nip" className="block text-xs font-medium text-slate-700 mb-1">
              NIP PPATS
            </label>
            <input
              id="set-nip"
              type="text"
              value={form.nip_ppats}
              onChange={(e) => setForm({ ...form, nip_ppats: e.target.value })}
              className="w-full px-3 py-2 text-sm font-mono border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-provinsi" className="block text-xs font-medium text-slate-700 mb-1">
              Nama Provinsi
            </label>
            <input
              id="set-provinsi"
              type="text"
              value={form.nama_provinsi ?? ''}
              onChange={(e) => setForm({ ...form, nama_provinsi: e.target.value })}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-telp" className="block text-xs font-medium text-slate-700 mb-1">
              Telepon Kantor
            </label>
            <input
              id="set-telp"
              type="text"
              value={form.telepon_kantor}
              onChange={(e) => setForm({ ...form, telepon_kantor: e.target.value })}
              className="w-full px-3 py-2 text-sm font-mono border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-sk-nomor" className="block text-xs font-medium text-slate-700 mb-1">
              Nomor SK Pengangkatan PPATS
            </label>
            <input
              id="set-sk-nomor"
              type="text"
              value={form.sk_ppats_nomor ?? ''}
              onChange={(e) => setForm({ ...form, sk_ppats_nomor: e.target.value })}
              placeholder="65/SK-33.HP.03.04.PPAT/VI/2026"
              className="w-full px-3 py-2 text-sm font-mono border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-sk-tanggal" className="block text-xs font-medium text-slate-700 mb-1">
              Tanggal SK Pengangkatan PPATS
            </label>
            <input
              id="set-sk-tanggal"
              type="text"
              value={form.sk_ppats_tanggal ?? ''}
              onChange={(e) => setForm({ ...form, sk_ppats_tanggal: e.target.value })}
              placeholder="30 Juni 2026"
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div className="md:col-span-2">
            <label htmlFor="set-sk-penerbit" className="block text-xs font-medium text-slate-700 mb-1">
              Pejabat Penerbit SK PPATS
            </label>
            <input
              id="set-sk-penerbit"
              type="text"
              value={form.sk_ppats_penerbit ?? ''}
              onChange={(e) => setForm({ ...form, sk_ppats_penerbit: e.target.value })}
              placeholder="Kakanwil Pertanahan Nasional Propinsi Jawa Tengah"
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div className="md:col-span-2">
            <label htmlFor="set-alamat" className="block text-xs font-medium text-slate-700 mb-1">
              Alamat Lengkap Kantor Kecamatan / PPATS
            </label>
            <input
              id="set-alamat"
              type="text"
              value={form.alamat_kantor}
              onChange={(e) => setForm({ ...form, alamat_kantor: e.target.value })}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-npoptkp-umum" className="block text-xs font-medium text-slate-700 mb-1">
              Default NPOPTKP Umum (AJB / APHB / Hibah) (Rp)
            </label>
            <input
              id="set-npoptkp-umum"
              type="number"
              value={form.npoptkp_umum}
              onChange={(e) =>
                setForm({ ...form, npoptkp_umum: Number(e.target.value) || 0 })
              }
              className="w-full px-3 py-2 text-sm font-mono tabular-nums border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
          <div>
            <label htmlFor="set-npoptkp-waris" className="block text-xs font-medium text-slate-700 mb-1">
              Default NPOPTKP Turun Waris (Rp)
            </label>
            <input
              id="set-npoptkp-waris"
              type="number"
              value={form.npoptkp_waris}
              onChange={(e) =>
                setForm({ ...form, npoptkp_waris: Number(e.target.value) || 0 })
              }
              className="w-full px-3 py-2 text-sm font-mono tabular-nums border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
            />
          </div>
        </div>

        {/* Saksi & Kuasa BPN Lists */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2 border-t border-slate-200">
          <div className="space-y-3">
            <h3 className="text-sm font-semibold text-slate-900">
              Daftar Saksi Akta Tetap
            </h3>
            <div className="flex gap-2">
              <input
                type="text"
                value={newSaksi}
                onChange={(e) => setNewSaksi(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    void addSaksi();
                  }
                }}
                placeholder="Tambah nama saksi..."
                aria-label="Tambah nama saksi"
                className="flex-1 px-3 py-1.5 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
              />
              <button
                type="button"
                onClick={() => void addSaksi()}
                aria-label="Tambah saksi baru"
                className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-medium text-slate-800 whitespace-nowrap"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Tambah</span>
              </button>
            </div>
            {form.daftar_saksi.length === 0 ? (
              <div className="p-3 rounded-lg border border-dashed border-slate-200 text-xs text-slate-500 text-center">
                Belum ada daftar saksi tetap. Ketik nama atau langsung klik <strong>+ Tambah</strong> untuk mengisi data saksi baru.
              </div>
            ) : (
            <ul className="divide-y divide-slate-200 border border-slate-200 rounded-lg">
              {form.daftar_saksi.map((s, idx) => {
                const prof = form.profil_saksi?.[idx] ?? { nama: s };
                const isOpen = expandedSaksiIdx === idx;
                return (
                  <li key={`saksi-${idx}`} className="px-3 py-2 text-sm text-slate-800 space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <div className="min-w-0">
                        <span className="font-medium text-slate-900">{s}</span>
                        {prof.nik && (
                          <span className="text-xs text-slate-500 font-mono ml-2">
                            · {prof.nik}
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          type="button"
                          onClick={() => setExpandedSaksiIdx(isOpen ? null : idx)}
                          className="text-xs text-sky-700 hover:underline font-medium"
                        >
                          {isOpen ? 'Tutup Detail' : 'Detail Akta'}
                        </button>
                        <button
                          type="button"
                          onClick={() => removeSaksi(idx)}
                          aria-label={`Hapus saksi ${s}`}
                          className="text-slate-400 hover:text-rose-600 p-1"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                    {isOpen && (
                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 pt-2 border-t border-slate-100 text-xs">
                        <div className="col-span-2 sm:col-span-1">
                          <label className="block text-slate-500 mb-0.5">Nama Lengkap</label>
                          <input
                            type="text"
                            value={prof.nama ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { nama: e.target.value })}
                            placeholder="Nama sesuai KTP"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">NIK (16 digit)</label>
                          <input
                            type="text"
                            value={prof.nik ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { nik: e.target.value })}
                            placeholder="3309..."
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Pekerjaan</label>
                          <input
                            type="text"
                            value={prof.pekerjaan ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { pekerjaan: e.target.value })}
                            placeholder="Perangkat Desa / Wiraswasta"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Tempat Lahir</label>
                          <input
                            type="text"
                            value={prof.tempat_lahir ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { tempat_lahir: e.target.value })}
                            placeholder="Boyolali"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Tanggal Lahir</label>
                          <input
                            type="text"
                            value={prof.tanggal_lahir ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { tanggal_lahir: e.target.value })}
                            placeholder="DD-MM-YYYY"
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Dukuh / Jalan</label>
                          <input
                            type="text"
                            value={prof.alamat_jalan ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { alamat_jalan: e.target.value })}
                            placeholder="Gunungan / Dukuhrejo"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">RT</label>
                          <input
                            type="text"
                            value={prof.rt ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { rt: e.target.value })}
                            placeholder="005"
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">RW</label>
                          <input
                            type="text"
                            value={prof.rw ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { rw: e.target.value })}
                            placeholder="008"
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Desa / Kelurahan</label>
                          <input
                            type="text"
                            value={prof.desa_kelurahan ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { desa_kelurahan: e.target.value })}
                            placeholder="Giriroto"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Kecamatan</label>
                          <input
                            type="text"
                            value={prof.kecamatan ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { kecamatan: e.target.value })}
                            placeholder="Ngemplak"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div className="col-span-2 sm:col-span-2">
                          <label className="block text-slate-500 mb-0.5">Kabupaten / Kota</label>
                          <input
                            type="text"
                            value={prof.kabupaten ?? ''}
                            onChange={(e) => updateSaksiProfile(idx, { kabupaten: e.target.value })}
                            placeholder="Kabupaten Boyolali"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div className="col-span-2 sm:col-span-3 flex justify-end pt-1">
                          <button
                            type="submit"
                            disabled={saving}
                            className="inline-flex items-center gap-1.5 px-3 py-1 rounded-md bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium"
                          >
                            <Save className="w-3 h-3" />
                            <span>Simpan Detail Saksi</span>
                          </button>
                        </div>
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>
            )}
          </div>

          <div className="space-y-3">
            <h3 className="text-sm font-semibold text-slate-900">
              Daftar Staf Kuasa Pengurusan BPN
            </h3>
            <div className="flex gap-2">
              <input
                type="text"
                value={newKuasa}
                onChange={(e) => setNewKuasa(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    void addKuasa();
                  }
                }}
                placeholder="Tambah nama staf kuasa BPN..."
                aria-label="Tambah nama staf kuasa BPN"
                className="flex-1 px-3 py-1.5 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-500"
              />
              <button
                type="button"
                onClick={() => void addKuasa()}
                aria-label="Tambah staf kuasa BPN baru"
                className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-medium text-slate-800 whitespace-nowrap"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Tambah</span>
              </button>
            </div>
            {form.daftar_kuasa_bpn.length === 0 ? (
              <div className="p-3 rounded-lg border border-dashed border-slate-200 text-xs text-slate-500 text-center">
                Belum ada daftar staf kuasa BPN. Ketik nama atau langsung klik <strong>+ Tambah</strong> untuk mengisi data staf baru.
              </div>
            ) : (
            <ul className="divide-y divide-slate-200 border border-slate-200 rounded-lg">
              {form.daftar_kuasa_bpn.map((k, idx) => {
                const prof = form.profil_kuasa?.[idx] ?? { nama: k };
                const isOpen = expandedKuasaIdx === idx;
                return (
                  <li key={`kuasa-${idx}`} className="px-3 py-2 text-sm text-slate-800 space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <div className="min-w-0">
                        <span className="font-medium text-slate-900">{k}</span>
                        {prof.nik && (
                          <span className="text-xs text-slate-500 font-mono ml-2">
                            · {prof.nik}
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          type="button"
                          onClick={() => setExpandedKuasaIdx(isOpen ? null : idx)}
                          className="text-xs text-sky-700 hover:underline font-medium"
                        >
                          {isOpen ? 'Tutup Detail' : 'Detail Kuasa'}
                        </button>
                        <button
                          type="button"
                          onClick={() => removeKuasa(idx)}
                          aria-label={`Hapus kuasa BPN ${k}`}
                          className="text-slate-400 hover:text-rose-600 p-1"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                    {isOpen && (
                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 pt-2 border-t border-slate-100 text-xs">
                        <div className="col-span-2 sm:col-span-1">
                          <label className="block text-slate-500 mb-0.5">Nama Lengkap</label>
                          <input
                            type="text"
                            value={prof.nama ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { nama: e.target.value })}
                            placeholder="Nama sesuai KTP"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">NIK (16 digit)</label>
                          <input
                            type="text"
                            value={prof.nik ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { nik: e.target.value })}
                            placeholder="3309..."
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Pekerjaan</label>
                          <input
                            type="text"
                            value={prof.pekerjaan ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { pekerjaan: e.target.value })}
                            placeholder="PENSIUNAN / Staf"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Tempat Lahir</label>
                          <input
                            type="text"
                            value={prof.tempat_lahir ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { tempat_lahir: e.target.value })}
                            placeholder="Boyolali"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Tanggal Lahir</label>
                          <input
                            type="text"
                            value={prof.tanggal_lahir ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { tanggal_lahir: e.target.value })}
                            placeholder="DD-MM-YYYY"
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Dukuh / Jalan</label>
                          <input
                            type="text"
                            value={prof.alamat_jalan ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { alamat_jalan: e.target.value })}
                            placeholder="Dukuhrejo"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">RT</label>
                          <input
                            type="text"
                            value={prof.rt ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { rt: e.target.value })}
                            placeholder="01"
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">RW</label>
                          <input
                            type="text"
                            value={prof.rw ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { rw: e.target.value })}
                            placeholder="02"
                            className="w-full px-2 py-1 border border-slate-200 rounded font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Desa / Kelurahan</label>
                          <input
                            type="text"
                            value={prof.desa_kelurahan ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { desa_kelurahan: e.target.value })}
                            placeholder="Donohudan"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-500 mb-0.5">Kecamatan</label>
                          <input
                            type="text"
                            value={prof.kecamatan ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { kecamatan: e.target.value })}
                            placeholder="Ngemplak"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div className="col-span-2 sm:col-span-2">
                          <label className="block text-slate-500 mb-0.5">Kabupaten / Kota</label>
                          <input
                            type="text"
                            value={prof.kabupaten ?? ''}
                            onChange={(e) => updateKuasaProfile(idx, { kabupaten: e.target.value })}
                            placeholder="Boyolali"
                            className="w-full px-2 py-1 border border-slate-200 rounded"
                          />
                        </div>
                        <div className="col-span-2 sm:col-span-3 flex justify-end pt-1">
                          <button
                            type="submit"
                            disabled={saving}
                            className="inline-flex items-center gap-1.5 px-3 py-1 rounded-md bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium"
                          >
                            <Save className="w-3 h-3" />
                            <span>Simpan Detail Kuasa BPN</span>
                          </button>
                        </div>
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>
            )}
          </div>
        </div>
      </form>

      {/* Section 2: Manajemen Master Template .DOCX */}
      <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-5">
        <div>
          <h2 className="font-display text-base font-semibold text-slate-900">
            Manajemen Master Template .DOCX (Akta A4 &amp; 1 File Bundel Lampiran F4)
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Unggah file Microsoft Word (.docx) dengan kop surat, margin, dan tabel asli kantor Anda. Template bawaan mengikuti format resmi PPATS (Akta ukuran A4 210 × 297 mm; seluruh lampiran bundel digabung dalam 1 file ukuran F4 215 × 330 mm).
          </p>
        </div>

        <div className="divide-y divide-slate-200 border border-slate-200 rounded-xl overflow-hidden">
          {TEMPLATE_DOCS_META.map((meta) => {
            const custom = customTemplates[meta.key];
            const paperLabel =
              meta.paperSize === 'A4'
                ? 'Kertas A4 (210 × 297 mm)'
                : 'Kertas F4 (215 × 330 mm)';
            return (
              <div
                key={meta.key}
                className="p-4 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3 bg-white hover:bg-slate-50/70"
              >
                <div className="flex items-start gap-3">
                  <FileText className="w-5 h-5 text-slate-500 shrink-0 mt-0.5" />
                  <div>
                    <div className="font-medium text-sm text-slate-900">
                      {meta.label}{' '}
                      <span className="font-mono text-xs text-slate-500">
                        ({meta.fileName})
                      </span>
                      <span className="text-xs font-medium text-sky-800 ml-1.5">
                        · {paperLabel}
                      </span>
                    </div>
                    <p className="text-xs text-slate-500 mt-0.5">{meta.description}</p>
                    <div className="text-xs mt-1 flex items-center gap-1.5">
                      {custom ? (
                        <span className="inline-flex items-center gap-1 text-emerald-700 font-medium">
                          <CheckCircle2 className="w-3.5 h-3.5" />
                          Template Kustom Aktif: {custom.fileName}
                        </span>
                      ) : (
                        <span className="text-slate-500">
                          Menggunakan Template Bawaan Sistem ({meta.paperSize})
                        </span>
                      )}
                    </div>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  <button
                    type="button"
                    onClick={() => downloadDefaultTemplateDocx(meta.key, meta.fileName)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-xs font-medium text-slate-700 transition-colors whitespace-nowrap"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Unduh Contoh Template .DOCX Bawaan</span>
                  </button>

                  <label className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium transition-colors cursor-pointer whitespace-nowrap">
                    <Upload className="w-3.5 h-3.5" />
                    <span>Upload .DOCX Kustom</span>
                    <input
                      type="file"
                      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                      onChange={(e) => handleTemplateUpload(meta.key, e)}
                      className="hidden"
                    />
                  </label>

                  {custom && (
                    <button
                      type="button"
                      onClick={() => onDeleteTemplate(meta.key)}
                      aria-label={`Hapus template kustom ${meta.fileName}`}
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-rose-200 text-rose-700 hover:bg-rose-50 text-xs font-medium"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Reset Bawaan</span>
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        {/* Supported Tags Reference Table */}
        <div className="pt-2 space-y-3">
          <h3 className="text-sm font-semibold text-slate-900">
            Panduan Lengkap Tag Variabel {'{{...}}'} yang Didukung
          </h3>
          <div className="border border-slate-200 rounded-xl overflow-hidden max-h-80 overflow-y-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="bg-slate-50 border-b border-slate-200 sticky top-0">
                <tr>
                  <th className="py-2.5 px-3 font-semibold text-slate-700 w-72">
                    Tag Template
                  </th>
                  <th className="py-2.5 px-3 font-semibold text-slate-700">
                    Keterangan Variabel
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {SUPPORTED_TEMPLATE_TAGS.map((item) => (
                  <tr key={item.tag} className="hover:bg-slate-50">
                    <td className="py-2 px-3 font-mono text-sky-800 font-medium">
                      {item.tag}
                    </td>
                    <td className="py-2 px-3 text-slate-700">{item.desc}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Section 3: Backup & Restore Database */}
      <div className="bg-white border border-slate-200 rounded-xl p-6 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div className="flex items-start gap-3">
          <Database className="w-5 h-5 text-slate-700 shrink-0 mt-0.5" />
          <div>
            <h2 className="font-display text-base font-semibold text-slate-900">
              Backup & Restore Database Kantor (.JSON)
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Seluruh data register berkas, pengaturan pejabat PPATS, dan file master template .docx disimpan di IndexedDB browser. Ekspor file .JSON secara berkala untuk pemindahan antar-komputer.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5 shrink-0">
          <button
            type="button"
            onClick={onExportBackup}
            className="inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-sm font-medium text-slate-800 transition-colors whitespace-nowrap"
          >
            <Download className="w-4 h-4" />
            <span>Export Backup (.JSON)</span>
          </button>
          <button
            type="button"
            onClick={() => restoreInputRef.current?.click()}
            className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium transition-colors whitespace-nowrap"
          >
            <Upload className="w-4 h-4" />
            <span>Restore Database (.JSON)</span>
          </button>
          <input
            ref={restoreInputRef}
            type="file"
            accept="application/json,.json"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) {
                onRestoreBackup(f);
                e.target.value = '';
              }
            }}
            className="hidden"
          />
        </div>
      </div>
    </div>
  );
};
