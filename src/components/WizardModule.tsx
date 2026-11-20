import React, { useMemo, useState } from 'react';
import {
  UploadCloud,
  Sparkles,
  AlertTriangle,
  CheckCircle2,
  Plus,
  Trash2,
  ArrowRight,
  ArrowLeft,
  FileArchive,
  Save,
  Printer,
  FileText,
  Download,
  Eye,
  X,
  Loader2,
  Calculator,
} from 'lucide-react';
import { DocumentReviewPanel } from './DocumentReviewPanel';
import { DateInputId } from './DateInputId';
import {
  AuditWarning,
  BerkasRecord,
  BerkasStatus,
  CustomTemplateEntry,
  DataPBB,
  DataTanah,
  ExtractionApiResponse,
  NonAktaDetail,
  OfficeSettings,
  ParaPihak,
  ParameterAkta,
  PemecahanBidangItem,
  PengesahanMode,
  PersonBasicInfo,
  PihakRole,
  ServiceType,
  TaxDraftPrefill,
  TemplateDocKey,
  UploadedDocFile,
} from '../types/ppats';
import {
  STATUS_OPTIONS,
  TEMPLATE_DOCS_META,
} from '../utils/seedData';
import {
  extractMonthYearFromDate,
  formatRupiah,
  formatTanggalDDMMYYYY,
  formatTanggalLengkapHuruf,
  todayDdMmYyyy,
} from '../utils/indonesianFormatter';
import {
  AKTA_SERVICE_TYPES,
  NON_AKTA_SERVICE_TYPES,
  calculateLandTaxes,
  getDefaultNpoptkp,
  isAktaService,
  runClientAuditWarnings,
  suggestNextAktaNumber,
} from '../utils/taxAndAudit';
import {
  formatPersonFullAddress,
  formatPersonShortAddress,
  generateAndDownloadBundleZip,
  generateAndDownloadSingleDocx,
  getNonAktaBundleFileName,
  numberToRoman,
  parseAdditionalPihakMeta,
  resolveJenisSertipikat,
  resolvePersonDetail,
} from '../utils/bundleGenerator';

interface WizardModuleProps {
  existingRecords: BerkasRecord[];
  editingRecord: BerkasRecord | null;
  taxPrefill?: TaxDraftPrefill | null;
  initialLayanan?: ServiceType;
  settings: OfficeSettings;
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>;
  onSaveRecord: (record: BerkasRecord) => Promise<void>;
  onCancelEdit: () => void;
  onPrintReceipt: (record: BerkasRecord) => void;
  onOpenTaxCalculator: (record: BerkasRecord) => Promise<void> | void;
  onNotify: (msg: string, type?: 'success' | 'error') => void;
}

function toPartyRoman(num: number): string {
  return numberToRoman(num);
}

function buildRoleOptions(paraPihak: ParaPihak[], layanan: ServiceType): PihakRole[] {
  const isTwoPartyDeed = layanan === 'AJB' || layanan === 'Akta Hibah';
  const options: PihakRole[] = [];

  if (isTwoPartyDeed) {
    options.push('Pihak I', 'Pihak II', 'Persetujuan Pasangan', 'Lampiran Lain');
  } else {
    const highestOrder = paraPihak.reduce((max, p) => {
      if (!p.peran.startsWith('Pihak ')) return max;
      const ord = parseAdditionalPihakMeta(p.peran).order;
      return ord < 900 ? Math.max(max, ord) : max;
    }, 0);
    const maxNumbered = Math.max(20, highestOrder + 1, paraPihak.length + 1);
    for (let i = 1; i <= maxNumbered; i++) {
      options.push(`Pihak ${toPartyRoman(i)}` as PihakRole);
    }
    options.push('Persetujuan Pasangan', 'Lampiran Lain');
  }

  for (const p of paraPihak) {
    if (p.peran && !options.includes(p.peran)) {
      options.push(p.peran);
    }
  }
  return options;
}

function suggestNextPartyRole(paraPihak: ParaPihak[], layanan: ServiceType): PihakRole {
  const existing = new Set(paraPihak.map((p) => p.peran));
  if (layanan === 'AJB' || layanan === 'Akta Hibah') {
    if (!existing.has('Pihak I')) return 'Pihak I';
    if (!existing.has('Pihak II')) return 'Pihak II';
    return 'Pihak I';
  }
  const limit = Math.max(20, paraPihak.length + 2);
  for (let i = 1; i <= limit; i++) {
    const candidate = `Pihak ${toPartyRoman(i)}` as PihakRole;
    if (!existing.has(candidate)) {
      return candidate;
    }
  }
  return `Pihak ${toPartyRoman(paraPihak.length + 1)}` as PihakRole;
}

function mapSuggestedRole(raw?: string): PihakRole {
  const s = (raw || '').trim().toLowerCase();
  if (s.includes('pasangan') || s.includes('suami') || s.includes('istri')) {
    return 'Persetujuan Pasangan';
  }
  if (s.includes('lampiran') || s.includes('saksi')) {
    return 'Lampiran Lain';
  }
  const romanMatch = /^pihak\s+(x{0,2}(?:ix|iv|v?i{0,3}))$/i.exec(s);
  if (romanMatch?.[1]) {
    return `Pihak ${romanMatch[1].toUpperCase()}` as PihakRole;
  }
  const numMatch = /^pihak\s+(\d+)$/i.exec(s);
  if (numMatch?.[1]) {
    return `Pihak ${toPartyRoman(Number(numMatch[1]))}` as PihakRole;
  }
  if (s.includes('iii') || s.includes('ketiga')) {
    return 'Pihak III';
  }
  if (s.includes('iv') || s.includes('keempat')) {
    return 'Pihak IV';
  }
  if (s.includes('ii') || s.includes('kedua') || s.includes('pembeli') || s.includes('penerima')) {
    return 'Pihak II';
  }
  return 'Pihak I';
}

function createEmptyPihak(peran: PihakRole = 'Pihak I'): ParaPihak {
  return {
    id: `p-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    peran,
    nik: '',
    nama: '',
    tempat_lahir: '',
    tanggal_lahir: '',
    jenis_kelamin: 'Laki-laki',
    alamat_jalan: '',
    rt_rw: '',
    desa_kelurahan: '',
    kecamatan: '',
    kabupaten: '',
    pekerjaan: '',
    status_perkawinan: 'KAWIN',
  };
}

const DEFAULT_TANPA_PERSETUJUAN_TEXT =
  'Menurut keterangan para pihak, dalam melakukan perbuatan hukum ini tanpa memerlukan persetujuan dari siapapun, dikarenakan tanah tersebut berasal dari Warisan.';

function findPresetPerson(
  name: string,
  list?: PersonBasicInfo[]
): PersonBasicInfo {
  const match = (list ?? []).find(
    (item) => item.nama.trim().toLowerCase() === name.trim().toLowerCase()
  );
  return match ? { ...match } : { nama: name };
}

function createDefaultNonAkta(
  svc: ServiceType,
  _kecName: string,
  _desaName: string,
  settings: OfficeSettings
): NonAktaDetail {
  return {
    pengesahan_mode: 'ppats',
    pengesahan_jabatan: '',
    pengesahan_nama: settings.nama_ppats,
    pengesahan_nip: settings.nip_ppats,
    camat_pengesahan_jabatan: '',
    camat_pengesahan_nama: settings.nama_ppats,
    camat_pengesahan_nip: settings.nip_ppats,
    camat_pengesahan_nomor: '',
    kades_pengesahan_jabatan: '',
    kades_pengesahan_nama: '',
    kades_pengesahan_nomor: '',
    sebutan_pewaris: 'Almarhum',
    nama_pewaris: '',
    tanggal_meninggal: '',
    tempat_meninggal: '',
    hubungan_pasangan_pewaris: 'Istrinya',
    nama_pasangan_pewaris: '',
    tgl_lahir_pasangan_pewaris: '',
    jumlah_anak_waris: undefined,
    catatan_skw: '',
    sub_permohonan_ukur: svc === 'Pemecahan' ? 'Pemecahan' : 'Pengukuran',
    bidang_pecahan: [],
    sertakan_pelepasan_jalan: false,
    pelepasan_jalan_luas_m2: 0,
    pelepasan_jalan_keterangan: '',
    sertakan_beda_nama: false,
    koordinat_gps: '',
    link_maps: '',
    nama_di_sertipikat: '',
    nama_di_ktp_kk: '',
    nomor_kk: '',
    jenis_dokumen_lain: '',
    nama_di_dokumen_lain: '',
    tgl_lahir_baku: '',
    keperluan_ralat: 'Ralat Data Sertipikat',
    uraian_kronologi: '',
  };
}

export const WizardModule: React.FC<WizardModuleProps> = ({
  existingRecords,
  editingRecord,
  taxPrefill,
  initialLayanan,
  settings,
  customTemplates,
  onSaveRecord,
  onCancelEdit,
  onPrintReceipt,
  onOpenTaxCalculator,
  onNotify,
}) => {
  const today = useMemo(() => todayDdMmYyyy(), []);
  const nextSuggestion = useMemo(
    () => suggestNextAktaNumber(existingRecords, today),
    [existingRecords, today]
  );

  const [step, setStep] = useState<1 | 2 | 3 | 4>(editingRecord ? 2 : 1);
  const [layanan, setLayanan] = useState<ServiceType>(
    editingRecord?.layanan ?? taxPrefill?.layanan ?? initialLayanan ?? 'AJB'
  );
  const [status, setStatus] = useState<BerkasStatus>(
    editingRecord?.status ?? '1. Pemberkasan'
  );
  const [uploadedFiles, setUploadedFiles] = useState<UploadedDocFile[]>([]);
  const [extracting, setExtracting] = useState(false);
  const [extractError, setExtractError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  const [paraPihak, setParaPihak] = useState<ParaPihak[]>(
    editingRecord?.para_pihak?.length
      ? editingRecord.para_pihak
      : [createEmptyPihak('Pihak I'), createEmptyPihak('Pihak II')]
  );

  const [tanah, setTanah] = useState<DataTanah>(
    editingRecord?.tanah
      ? {
          jenis_sertipikat: resolveJenisSertipikat(editingRecord.tanah),
          surat_ukur_tanggal: '',
          surat_ukur_nomor: '',
          provinsi: settings.nama_provinsi ?? 'Jawa Tengah',
          batas_utara: '',
          batas_timur: '',
          batas_selatan: '',
          batas_barat: '',
          penggunaan_tanah: 'Pekarangan',
          status_tanah: 'Non-Pertanian',
          tanggal_pernyataan_pertanian: '',
          menguasai_sejak_tahun: '',
          ...editingRecord.tanah,
        }
      : {
          jenis_sertipikat: 'digital',
          jenis_hak: 'Hak Milik',
          nib: '',
          nomor_hak_lama: '',
          surat_ukur_tanggal: '',
          surat_ukur_nomor: '',
          luas_m2: taxPrefill?.luas_m2 ?? 0,
          desa: '',
          kecamatan: settings.nama_kecamatan.replace(/^Kecamatan\s+/i, ''),
          kabupaten: settings.nama_kabupaten,
          provinsi: settings.nama_provinsi ?? 'Jawa Tengah',
          batas_utara: '',
          batas_timur: '',
          batas_selatan: '',
          batas_barat: '',
          penggunaan_tanah: 'Pekarangan',
          status_tanah: 'Non-Pertanian',
          tanggal_pernyataan_pertanian: '',
          menguasai_sejak_tahun: '',
          dasar_perolehan: '',
          pemegang_hak_sertipikat: [],
        }
  );
  const [pemegangHakInput, setPemegangHakInput] = useState<string>(
    (editingRecord?.tanah.pemegang_hak_sertipikat ?? []).join(', ')
  );

  const [pbb, setPbb] = useState<DataPBB>(
    editingRecord?.pbb ?? {
      nop: '',
      tahun_pajak: String(new Date().getFullYear()),
      njop_bumi_m2: taxPrefill?.njop_bumi_m2 ?? 0,
      njop_bangunan: taxPrefill?.njop_bangunan ?? 0,
    }
  );

  const [aiWarnings, setAiWarnings] = useState<AuditWarning[]>(
    (editingRecord?.audit_warnings ?? []).filter((w) => w.id.startsWith('ai-'))
  );

  const [akta, setAkta] = useState<ParameterAkta>(() => {
    if (editingRecord?.akta) {
      const a = editingRecord.akta;
      return {
        nomor_pengantar: '',
        butuh_persetujuan_pasangan:
          a.butuh_persetujuan_pasangan ??
          editingRecord.para_pihak.some((p) => p.peran === 'Persetujuan Pasangan'),
        keterangan_tanpa_persetujuan:
          a.keterangan_tanpa_persetujuan ?? DEFAULT_TANPA_PERSETUJUAN_TEXT,
        pihak_penerima_aphb: a.pihak_penerima_aphb ?? 'Pihak II',
        kuasa_bpn_detail:
          a.kuasa_bpn_detail ?? findPresetPerson(a.nama_kuasa_bpn, settings.profil_kuasa),
        saksi_1_detail:
          a.saksi_1_detail ?? findPresetPerson(a.saksi_1, settings.profil_saksi),
        saksi_2_detail:
          a.saksi_2_detail ?? findPresetPerson(a.saksi_2, settings.profil_saksi),
        ...a,
      };
    }
    const defaultKuasa = settings.daftar_kuasa_bpn[0] ?? '';
    const defaultSaksi1 = settings.daftar_saksi[0] ?? '';
    const defaultSaksi2 = settings.daftar_saksi[1] ?? '';
    const initialSvc = taxPrefill?.layanan ?? initialLayanan ?? 'AJB';
    const pengantarYear =
      extractMonthYearFromDate(today)?.year ?? new Date().getFullYear();
    return {
      nomor_akta: nextSuggestion.nomorAkta,
      tanggal_akta: today,
      nomor_pengantar: `${String(nextSuggestion.noUrut).padStart(2, '0')}/${pengantarYear}`,
      nama_ppats: settings.nama_ppats,
      jabatan_ppats: settings.jabatan_ppats,
      nip_ppats: settings.nip_ppats,
      harga_per_m2: taxPrefill?.harga_per_m2 ?? 0,
      nilai_transaksi_override: taxPrefill?.nilai_transaksi_override ?? null,
      npoptkp: taxPrefill?.npoptkp ?? getDefaultNpoptkp(initialSvc, settings),
      butuh_persetujuan_pasangan: true,
      keterangan_tanpa_persetujuan: DEFAULT_TANPA_PERSETUJUAN_TEXT,
      pihak_penerima_aphb: 'Pihak II',
      nama_kuasa_bpn: defaultKuasa,
      kuasa_bpn_detail: findPresetPerson(defaultKuasa, settings.profil_kuasa),
      saksi_1: defaultSaksi1,
      saksi_1_detail: findPresetPerson(defaultSaksi1, settings.profil_saksi),
      saksi_2: defaultSaksi2,
      saksi_2_detail: findPresetPerson(defaultSaksi2, settings.profil_saksi),
      catatan_loket: '',
    };
  });

  const [nonAkta, setNonAkta] = useState<NonAktaDetail>(() => {
    const base = createDefaultNonAkta(
      editingRecord?.layanan ?? taxPrefill?.layanan ?? initialLayanan ?? 'AJB',
      editingRecord?.tanah?.kecamatan || settings.nama_kecamatan,
      editingRecord?.tanah?.desa || '',
      settings
    );
    return editingRecord?.non_akta ? { ...base, ...editingRecord.non_akta } : base;
  });

  const [generatingBundle, setGeneratingBundle] = useState(false);
  const [savedRecordId, setSavedRecordId] = useState<string | null>(
    editingRecord?.id ?? null
  );
  const [showReview, setShowReview] = useState(true);
  const [reviewDocKey, setReviewDocKey] = useState<'01_akta' | '02_lampiran_bundel'>(() =>
    editingRecord && !isAktaService(editingRecord.layanan)
      ? '02_lampiran_bundel'
      : '01_akta'
  );

  const isAkta = isAktaService(layanan);

  const handleServiceChange = (nextLayanan: ServiceType) => {
    setLayanan(nextLayanan);
    if (!isAktaService(nextLayanan)) {
      setReviewDocKey('02_lampiran_bundel');
      setNonAkta((prev) => ({
        ...prev,
        sub_permohonan_ukur:
          nextLayanan === 'Pemecahan'
            ? 'Pemecahan'
            : nextLayanan === 'Ukur Ulang' || nextLayanan === 'Pemecahan/Ukur Ulang'
            ? 'Pengukuran'
            : prev.sub_permohonan_ukur,
      }));
    } else {
      setReviewDocKey('01_akta');
    }
    if (!editingRecord) {
      setAkta((prev) => ({
        ...prev,
        npoptkp: getDefaultNpoptkp(nextLayanan, settings),
      }));
    }
  };

  const handleFilesSelected = async (filesList: FileList | null) => {
    if (!filesList || filesList.length === 0) return;
    const added: UploadedDocFile[] = [];

    for (const file of Array.from(filesList)) {
      const arrayBuf = await file.arrayBuffer();
      const bytes = new Uint8Array(arrayBuf);
      let binary = '';
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
      }
      const dataBase64 = btoa(binary);
      added.push({
        id: `doc-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        name: file.name,
        mimeType: file.type || 'application/octet-stream',
        size: file.size,
        dataBase64,
      });
    }

    setUploadedFiles((prev) => [...prev, ...added]);
    setExtractError(null);
  };

  const handleExtractAi = async () => {
    if (uploadedFiles.length === 0) {
      setExtractError('Unggah minimal 1 dokumen (KTP, KK, Sertipikat, SPPT PBB, atau SKW) sebelum menjalankan ekstraksi AI.');
      return;
    }

    setExtracting(true);
    setExtractError(null);

    try {
      const res = await fetch('/api/extract-documents', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          layanan,
          files: uploadedFiles.map((f) => ({
            name: f.name,
            mimeType: f.mimeType,
            dataBase64: f.dataBase64,
          })),
        }),
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data?.error || 'Ekstraksi dokumen gagal.');
      }

      const parsed = data as ExtractionApiResponse;

      if (Array.isArray(parsed.para_pihak) && parsed.para_pihak.length > 0) {
        setParaPihak(
          parsed.para_pihak.map((p, idx) => ({
            id: `p-ai-${Date.now()}-${idx}`,
            peran: mapSuggestedRole(p.peran_saran),
            nik: p.nik || '',
            nama: p.nama || '',
            tempat_lahir: p.tempat_lahir || '',
            tanggal_lahir: p.tanggal_lahir?.trim()
              ? formatTanggalDDMMYYYY(p.tanggal_lahir)
              : '',
            jenis_kelamin: p.jenis_kelamin || 'Laki-laki',
            alamat_jalan: p.alamat_jalan || '',
            rt_rw: p.rt_rw || '',
            desa_kelurahan: p.desa_kelurahan || '',
            kecamatan: p.kecamatan || '',
            kabupaten: p.kabupaten || '',
            pekerjaan: p.pekerjaan || '',
            status_perkawinan: p.status_perkawinan || '',
          }))
        );
      }

      if (parsed.tanah) {
        const holders = Array.isArray(parsed.tanah.pemegang_hak_sertipikat)
          ? parsed.tanah.pemegang_hak_sertipikat
          : [];
        const detectedCertType = resolveJenisSertipikat({
          jenis_sertipikat:
            parsed.tanah.jenis_sertipikat === 'digital' ||
            parsed.tanah.jenis_sertipikat === 'analog'
              ? parsed.tanah.jenis_sertipikat
              : undefined,
          surat_ukur_nomor: parsed.tanah.surat_ukur_nomor,
          nomor_hak_lama: parsed.tanah.nomor_hak_lama,
        });
        setTanah((prev) => ({
          ...prev,
          jenis_sertipikat: detectedCertType,
          jenis_hak: parsed.tanah.jenis_hak || 'Hak Milik',
          nib: parsed.tanah.nib || '',
          nomor_hak_lama:
            detectedCertType === 'analog' ? parsed.tanah.nomor_hak_lama || '' : '',
          surat_ukur_tanggal: parsed.tanah.surat_ukur_tanggal || '',
          surat_ukur_nomor: parsed.tanah.surat_ukur_nomor || '',
          luas_m2: Number(parsed.tanah.luas_m2) || 0,
          desa: parsed.tanah.desa || '',
          kecamatan: parsed.tanah.kecamatan || '',
          kabupaten: parsed.tanah.kabupaten || settings.nama_kabupaten,
          provinsi: parsed.tanah.provinsi || settings.nama_provinsi || 'Jawa Tengah',
          batas_utara: parsed.tanah.batas_utara || '',
          batas_timur: parsed.tanah.batas_timur || '',
          batas_selatan: parsed.tanah.batas_selatan || '',
          batas_barat: parsed.tanah.batas_barat || '',
          dasar_perolehan: parsed.tanah.dasar_perolehan || '',
          pemegang_hak_sertipikat: holders,
        }));
        setPemegangHakInput(holders.join(', '));
      }

      if (parsed.pbb) {
        setPbb({
          nop: parsed.pbb.nop || '',
          tahun_pajak: parsed.pbb.tahun_pajak || String(new Date().getFullYear()),
          njop_bumi_m2: Number(parsed.pbb.njop_bumi_m2) || 0,
          njop_bangunan: Number(parsed.pbb.njop_bangunan) || 0,
        });
      }

      if (Array.isArray(parsed.audit_warnings)) {
        setAiWarnings(
          parsed.audit_warnings.map((w, idx) => ({
            id: `ai-w-${idx}`,
            tingkat:
              w.tingkat === 'tinggi' || w.tingkat === 'sedang' ? w.tingkat : 'info',
            judul: w.judul || 'Temuan Verifikasi AI',
            deskripsi: w.deskripsi || '',
          }))
        );
      }

      onNotify('Ekstraksi dokumen AI selesai. Silakan verifikasi hasil pada Langkah 2.', 'success');
      setStep(2);
    } catch (err) {
      const msg =
        err instanceof Error ? err.message : 'Terjadi kesalahan saat mengekstrak dokumen.';
      setExtractError(msg);
      onNotify(msg, 'error');
    } finally {
      setExtracting(false);
    }
  };

  const combinedAuditWarnings = useMemo(() => {
    const tanahWithHolders: DataTanah = {
      ...tanah,
      pemegang_hak_sertipikat: pemegangHakInput
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean),
    };
    const clientWarnings = runClientAuditWarnings(paraPihak, tanahWithHolders, {
      butuhPersetujuanPasangan: akta.butuh_persetujuan_pasangan ?? true,
      layanan,
    });
    const seen = new Set<string>();
    const merged: AuditWarning[] = [];
    for (const w of [...aiWarnings, ...clientWarnings]) {
      if (!seen.has(w.id)) {
        seen.add(w.id);
        merged.push(w);
      }
    }
    return merged;
  }, [aiWarnings, paraPihak, tanah, pemegangHakInput, akta.butuh_persetujuan_pasangan, layanan]);

  const taxCalc = useMemo(
    () =>
      calculateLandTaxes({
        luas_m2: tanah.luas_m2,
        harga_per_m2: akta.harga_per_m2,
        njop_bumi_m2: pbb.njop_bumi_m2,
        njop_bangunan: pbb.njop_bangunan,
        npoptkp: akta.npoptkp,
        layanan,
        nilai_transaksi_override: akta.nilai_transaksi_override,
      }),
    [tanah.luas_m2, akta, pbb, layanan]
  );

  const formattedAktaDate = useMemo(
    () => formatTanggalLengkapHuruf(akta.tanggal_akta),
    [akta.tanggal_akta]
  );

  const buildCurrentRecord = (): BerkasRecord => {
    const id =
      savedRecordId ||
      editingRecord?.id ||
      `reg-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
    const noUrut = editingRecord?.no_urut ?? nextSuggestion.noUrut;
    const year = String(
      extractMonthYearFromDate(akta.tanggal_akta || today)?.year ??
        new Date().getFullYear()
    );
    const kodeBerkas =
      editingRecord?.kode_berkas ||
      `REG-${year}-${String(noUrut).padStart(3, '0')}`;

    const finalTanah: DataTanah = {
      ...tanah,
      pemegang_hak_sertipikat: pemegangHakInput
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean),
    };

    return {
      id,
      no_urut: noUrut,
      kode_berkas: kodeBerkas,
      tanggal_masuk: editingRecord?.tanggal_masuk || today,
      layanan,
      status,
      para_pihak: paraPihak,
      tanah: finalTanah,
      pbb,
      audit_warnings: combinedAuditWarnings,
      akta,
      pajak: taxCalc,
      non_akta: isAktaService(layanan) ? undefined : nonAkta,
      updated_at: new Date().toISOString(),
    };
  };

  const reviewRecord = useMemo<BerkasRecord>(() => {
    const id = savedRecordId || editingRecord?.id || 'preview-draft';
    const noUrut = editingRecord?.no_urut ?? nextSuggestion.noUrut;
    const year = String(
      extractMonthYearFromDate(akta.tanggal_akta || today)?.year ??
        new Date().getFullYear()
    );
    const kodeBerkas =
      editingRecord?.kode_berkas ||
      `REG-${year}-${String(noUrut).padStart(3, '0')}`;
    const finalTanah: DataTanah = {
      ...tanah,
      pemegang_hak_sertipikat: pemegangHakInput
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean),
    };
    return {
      id,
      no_urut: noUrut,
      kode_berkas: kodeBerkas,
      tanggal_masuk: editingRecord?.tanggal_masuk || today,
      layanan,
      status,
      para_pihak: paraPihak,
      tanah: finalTanah,
      pbb,
      audit_warnings: combinedAuditWarnings,
      akta,
      pajak: taxCalc,
      non_akta: isAktaService(layanan) ? undefined : nonAkta,
      updated_at: editingRecord?.updated_at || today,
    };
  }, [
    savedRecordId,
    editingRecord,
    nextSuggestion.noUrut,
    akta,
    today,
    tanah,
    pemegangHakInput,
    layanan,
    status,
    paraPihak,
    pbb,
    combinedAuditWarnings,
    taxCalc,
    nonAkta,
  ]);

  const handleSaveOnly = async () => {
    const rec = buildCurrentRecord();
    await onSaveRecord(rec);
    setSavedRecordId(rec.id);
    onNotify(`Berkas ${rec.kode_berkas} berhasil disimpan ke Buku Register Akta.`, 'success');
  };

  const handleGenerateZipAndSave = async () => {
    setGeneratingBundle(true);
    try {
      const rec = buildCurrentRecord();
      await onSaveRecord(rec);
      setSavedRecordId(rec.id);
      const result = await generateAndDownloadBundleZip(rec, settings, customTemplates);
      onNotify(
        isAktaService(rec.layanan)
          ? `Berhasil membuat ${result.fileCount} dokumen Word (Akta A4 & Lampiran F4) dalam ${result.zipFileName} dan memperbarui Buku Register.`
          : `Berhasil membuat bundel Non-Akta F4 (${result.fileCount} file) dalam ${result.zipFileName} tanpa Akta A4.`,
        'success'
      );
    } catch (err) {
      onNotify(
        err instanceof Error
          ? `Gagal generate bundel Word: ${err.message}`
          : 'Gagal generate bundel .DOCX.',
        'error'
      );
    } finally {
      setGeneratingBundle(false);
    }
  };

  const handleDownloadSingle = async (docKey: TemplateDocKey) => {
    setGeneratingBundle(true);
    try {
      const rec = buildCurrentRecord();
      await onSaveRecord(rec);
      setSavedRecordId(rec.id);
      const result = await generateAndDownloadSingleDocx(rec, settings, customTemplates, docKey);
      onNotify(`Berhasil mengunduh ${result.fileName} dan menyimpan berkas ke Buku Register.`, 'success');
    } catch (err) {
      onNotify(
        err instanceof Error ? `Gagal mengunduh dokumen: ${err.message}` : 'Gagal mengunduh .DOCX.',
        'error'
      );
    } finally {
      setGeneratingBundle(false);
    }
  };

  const updateParty = (id: string, patch: Partial<ParaPihak>) => {
    setParaPihak((prev) =>
      prev.map((p) => (p.id === id ? { ...p, ...patch } : p))
    );
  };

  const removeParty = (id: string) => {
    setParaPihak((prev) => prev.filter((p) => p.id !== id));
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h1 className="font-display text-2xl font-bold tracking-tight text-slate-900">
            {editingRecord
              ? `Edit / Revisi Berkas — ${editingRecord.kode_berkas}`
              : 'Buat Berkas Baru (Wizard 4 Langkah)'}
          </h1>
          <p className="text-sm text-slate-600 mt-1">
            Alur terpandu dari unggah berkas &amp; ekstraksi AI, verifikasi audit, parameter akta &amp; saksi, hingga dokumen Word (.DOCX · Akta A4 &amp; 1 File Lampiran F4). Perhitungan pajak dilakukan di menu terpisah.
          </p>
        </div>

        {editingRecord && (
          <button
            type="button"
            onClick={onCancelEdit}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-xs font-medium text-slate-700 whitespace-nowrap"
          >
            <X className="w-3.5 h-3.5" />
            <span>Batal Edit & Buat Baru</span>
          </button>
        )}
      </div>

      {/* Stepper */}
      <nav aria-label="Tahapan Wizard" className="bg-white border border-slate-200 rounded-xl p-2">
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-2">
          {[
            { num: 1 as const, label: '1. Layanan & Upload AI' },
            { num: 2 as const, label: '2. Meja Verifikasi & Audit' },
            {
              num: 3 as const,
              label: isAkta
                ? '3. Parameter Akta & Saksi'
                : '3. Pengesahan & Detail Berkas',
            },
            {
              num: 4 as const,
              label: isAkta ? '4. Generate Bundel .ZIP' : '4. Review & Unduh F4',
            },
          ].map((item) => {
            const active = step === item.num;
            const completed = step > item.num;
            return (
              <button
                key={item.num}
                type="button"
                onClick={() => setStep(item.num)}
                className={`px-3.5 py-2.5 rounded-lg text-xs font-semibold text-left transition-colors whitespace-nowrap truncate ${
                  active
                    ? 'bg-[#0f172a] text-white'
                    : completed
                    ? 'bg-emerald-50 text-emerald-800 hover:bg-emerald-100'
                    : 'bg-slate-50 text-slate-600 hover:bg-slate-100'
                }`}
              >
                {item.label}
              </button>
            );
          })}
        </div>
      </nav>

      {/* STEP 1 */}
      {step === 1 && (
        <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-6">
          <div className="space-y-4">
            <h2 className="font-display text-base font-semibold text-slate-900">
              Pilih Jenis Layanan Pertanahan
            </h2>

            <div className="space-y-2">
              <p className="text-xs font-semibold text-slate-600">
                1. Layanan Akta PPATS (Blangko Akta A4 + Lampiran F4 · Masuk Laporan Bulanan PPATS)
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                {AKTA_SERVICE_TYPES.map((type) => {
                  const selected = layanan === type;
                  return (
                    <button
                      key={type}
                      type="button"
                      onClick={() => handleServiceChange(type)}
                      className={`px-4 py-3 rounded-xl border text-left transition-colors ${
                        selected
                          ? 'border-sky-600 bg-sky-50 text-sky-900'
                          : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <span className="block text-sm font-semibold">{type}</span>
                      <span className="block text-[11px] opacity-75 mt-0.5">
                        Akta A4 + Lampiran F4
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="space-y-2">
              <p className="text-xs font-semibold text-slate-600">
                2. Berkas Non-Akta (Hanya Bundel F4 · Tidak Masuk Laporan Bulanan PPATS)
              </p>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                {NON_AKTA_SERVICE_TYPES.map((type) => {
                  const selected = layanan === type;
                  return (
                    <button
                      key={type}
                      type="button"
                      onClick={() => handleServiceChange(type)}
                      className={`px-4 py-3 rounded-xl border text-left transition-colors ${
                        selected
                          ? 'border-sky-600 bg-sky-50 text-sky-900'
                          : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <span className="block text-sm font-semibold">{type}</span>
                      <span className="block text-[11px] opacity-75 mt-0.5">
                        Non-Akta · Bundel F4
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {!isAkta && (
              <div className="p-3.5 rounded-lg border border-sky-200 bg-sky-50/70 text-xs text-sky-950">
                <strong>Catatan Berkas Non-Akta ({layanan}):</strong> Berkas ini{' '}
                <strong>tidak menerbitkan Akta A4</strong>,{' '}
                <strong>tidak masuk dalam Laporan Bulanan PPATS</strong>, dan hanya menghasilkan 1 file bundel Word ukuran F4. Pengesahan dokumen dapat diatur di Langkah 3 (SKW &amp; Beda Nama otomatis 2 pengesahan Camat + Kepala Desa asal sertipikat; dokumen lainnya default PPATS atau pilih Kepala Desa / isi manual).
              </div>
            )}
          </div>

          <div className="space-y-3">
            <h2 className="font-display text-base font-semibold text-slate-900">
              Unggah Dokumen Pendukung (KTP, KK, Sertipikat Tanah/Elektronik, SPPT PBB, SKW)
            </h2>

            <div
              onDragOver={(e) => {
                e.preventDefault();
                setIsDragging(true);
              }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setIsDragging(false);
                handleFilesSelected(e.dataTransfer.files);
              }}
              className={`border-2 border-dashed rounded-xl p-8 text-center transition-colors ${
                isDragging
                  ? 'border-sky-500 bg-sky-50/70'
                  : 'border-slate-300 bg-slate-50/50 hover:bg-slate-50'
              }`}
            >
              <UploadCloud className="w-8 h-8 text-slate-400 mx-auto mb-3" />
              <p className="text-sm font-medium text-slate-800">
                Tarik dan lepaskan multi-file PDF atau Foto/Scan (JPG, PNG, WebP) di sini
              </p>
              <p className="text-xs text-slate-500 mt-1">
                Gemini AI (gemini-4-hydrogen) akan mengekstrak seluruh identitas para pihak, NIB sertipikat, SPPT PBB, dan mendeteksi kejanggalan berkas secara otomatis.
              </p>
              <label className="mt-4 inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-xs font-semibold text-slate-800 cursor-pointer">
                <span>Pilih File Dokumen</span>
                <input
                  type="file"
                  multiple
                  accept="application/pdf,image/*"
                  onChange={(e) => handleFilesSelected(e.target.files)}
                  className="hidden"
                />
              </label>
            </div>

            {uploadedFiles.length > 0 && (
              <div className="border border-slate-200 rounded-xl divide-y divide-slate-200">
                {uploadedFiles.map((file) => (
                  <div
                    key={file.id}
                    className="flex items-center justify-between px-4 py-2.5 text-sm"
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      <FileText className="w-4 h-4 text-sky-600 shrink-0" />
                      <span className="font-medium text-slate-800 truncate">
                        {file.name}
                      </span>
                      <span className="font-mono text-xs text-slate-500 shrink-0">
                        {(file.size / 1024).toFixed(0)} KB
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() =>
                        setUploadedFiles((prev) =>
                          prev.filter((f) => f.id !== file.id)
                        )
                      }
                      aria-label={`Hapus file ${file.name}`}
                      className="text-slate-400 hover:text-rose-600 p-1"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ))}
              </div>
            )}

            {extractError && (
              <div
                role="alert"
                className="p-3.5 rounded-lg border border-rose-200 bg-rose-50 text-xs text-rose-800 flex items-start gap-2.5"
              >
                <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5 text-rose-600" />
                <div>
                  <p className="font-semibold">Ekstraksi AI Tidak Dapat Dilanjutkan</p>
                  <p className="mt-0.5">{extractError}</p>
                </div>
              </div>
            )}
          </div>

          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 pt-4 border-t border-slate-200">
            <button
              type="button"
              onClick={() => setStep(2)}
              className="inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-sm font-medium text-slate-700"
            >
              <span>Lanjut Isi / Verifikasi Manual</span>
              <ArrowRight className="w-4 h-4" />
            </button>

            <button
              type="button"
              onClick={handleExtractAi}
              disabled={extracting}
              className="inline-flex items-center justify-center gap-2 px-5 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium transition-colors disabled:opacity-60"
            >
              {extracting ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>Mengekstrak Dokumen via Gemini AI...</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4 text-amber-400" />
                  <span>Ekstrak Dokumen dengan AI & Lanjut</span>
                </>
              )}
            </button>
          </div>
        </div>
      )}

      {/* STEP 2 */}
      {step === 2 && (
        <div className="space-y-6">
          {/* Audit Warnings Panel */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-3">
            <div className="flex items-center justify-between">
              <h2 className="font-display text-base font-semibold text-slate-900">
                Hasil Audit & Peringatan Verifikasi Berkas ({combinedAuditWarnings.length})
              </h2>
            </div>

            {combinedAuditWarnings.length === 0 ? (
              <div className="p-3.5 rounded-lg border border-emerald-200 bg-emerald-50 text-xs text-emerald-800 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                <span>
                  Tidak ditemukan kejanggalan data (NIK ganda, selisih nama sertipikat, atau persetujuan pasangan). Tetap lakukan pengecekan fisik dokumen asli.
                </span>
              </div>
            ) : (
              <div className="space-y-2.5">
                {combinedAuditWarnings.map((w) => {
                  const tone =
                    w.tingkat === 'tinggi'
                      ? 'border-rose-200 bg-rose-50 text-rose-900'
                      : w.tingkat === 'sedang'
                      ? 'border-amber-200 bg-amber-50 text-amber-900'
                      : 'border-sky-200 bg-sky-50 text-sky-900';
                  const label =
                    w.tingkat === 'tinggi'
                      ? 'Perhatian Tinggi'
                      : w.tingkat === 'sedang'
                      ? 'Perlu Cek'
                      : 'Informasi';

                  return (
                    <div
                      key={w.id}
                      className={`p-3.5 rounded-lg border text-xs flex items-start gap-2.5 ${tone}`}
                    >
                      <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" />
                      <div>
                        <p className="font-semibold">
                          [{label}] {w.judul}
                        </p>
                        <p className="mt-0.5 opacity-90">{w.deskripsi}</p>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Para Pihak Editable List */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-3">
              <div>
                <h2 className="font-display text-base font-semibold text-slate-900">
                  {isAkta
                    ? `Data Para Pihak (${paraPihak.length})`
                    : layanan === 'Turun Waris'
                    ? `Daftar Ahli Waris / Pemohon (${paraPihak.length})`
                    : `Data Pemohon / Pemilik Sertipikat (${paraPihak.length})`}
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  {layanan === 'AJB' || layanan === 'Akta Hibah'
                    ? `Akta ${layanan} hanya terdiri dari Pihak I (${layanan === 'AJB' ? 'Penjual' : 'Pemberi Hibah'}) dan Pihak II (${layanan === 'AJB' ? 'Pembeli' : 'Penerima Hibah'}). Jika pemilik lebih dari satu orang, tambahkan sebagai Pihak I agar tergabung dalam PIHAK PERTAMA.`
                    : layanan === 'APHB'
                    ? 'Untuk APHB, setiap pemilik hak menjadi pihak terpisah secara berurutan (Pihak I, Pihak II, Pihak III, ... s/d Pihak XX atau lebih). Keterangan persetujuan / tidak memerlukan persetujuan otomatis ditulis setelah pihak terakhir.'
                    : layanan === 'Turun Waris'
                    ? 'Tambahkan seluruh ahli waris yang berhak sebagai Pihak I, Pihak II, dst. Data pewaris (almarhum/almarhumah) diisi pada Langkah 3.'
                    : 'Tambahkan seluruh pemohon / pemegang hak pada sertipikat (Pihak I, Pihak II, dst.) untuk injeksi ke surat permohonan, kuasa, dan pernyataan BPN.'}
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2 shrink-0">
                {layanan === 'AJB' || layanan === 'Akta Hibah' ? (
                  <>
                    <button
                      type="button"
                      onClick={() =>
                        setParaPihak((prev) => [...prev, createEmptyPihak('Pihak I')])
                      }
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium whitespace-nowrap"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Tambah Pemilik (Pihak I)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        setParaPihak((prev) => [...prev, createEmptyPihak('Pihak II')])
                      }
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 text-xs font-medium whitespace-nowrap"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Tambah Pihak II</span>
                    </button>
                  </>
                ) : (
                  <button
                    type="button"
                    onClick={() =>
                      setParaPihak((prev) => [
                        ...prev,
                        createEmptyPihak(suggestNextPartyRole(prev, layanan)),
                      ])
                    }
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium whitespace-nowrap"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>
                      {isAkta
                        ? `Tambah Pihak (${suggestNextPartyRole(paraPihak, layanan)})`
                        : layanan === 'Turun Waris'
                        ? `Tambah Ahli Waris (${suggestNextPartyRole(paraPihak, layanan)})`
                        : `Tambah Pemohon (${suggestNextPartyRole(paraPihak, layanan)})`}
                    </span>
                  </button>
                )}
              </div>
            </div>

            {(layanan === 'AJB' || layanan === 'Akta Hibah') &&
              paraPihak.some(
                (p) => p.peran.startsWith('Pihak ') && p.peran !== 'Pihak I' && p.peran !== 'Pihak II'
              ) && (
                <div className="p-3 rounded-lg border border-amber-200 bg-amber-50 text-xs text-amber-900 flex flex-wrap items-center justify-between gap-2">
                  <span>
                    Ditemukan peran <strong>Pihak III+</strong> pada layanan {layanan}. Dalam {layanan}, seluruh pemilik seharusnya berada di <strong>Pihak I</strong> atau <strong>Pihak II</strong>.
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      setParaPihak((prev) =>
                        prev.map((p) =>
                          p.peran.startsWith('Pihak ') &&
                          p.peran !== 'Pihak I' &&
                          p.peran !== 'Pihak II'
                            ? { ...p, peran: 'Pihak I' }
                            : p
                        )
                      )
                    }
                    className="px-2.5 py-1 rounded-md bg-amber-900 text-white font-medium hover:bg-amber-800"
                  >
                    Gabungkan Pihak III+ ke Pihak I
                  </button>
                </div>
              )}

            {layanan === 'APHB' &&
              (() => {
                const partyRoles = paraPihak
                  .filter((p) => p.peran.startsWith('Pihak '))
                  .map((p) => p.peran);
                return new Set(partyRoles).size < partyRoles.length;
              })() && (
                <div className="p-3 rounded-lg border border-sky-200 bg-sky-50 text-xs text-sky-900 flex flex-wrap items-center justify-between gap-2">
                  <span>
                    Pada APHB, setiap pemilik sebaiknya menjadi pihak yang terpisah (Pihak I, Pihak II, Pihak III, dst.).
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      setParaPihak((prev) => {
                        let seq = 0;
                        return prev.map((p) => {
                          if (!p.peran.startsWith('Pihak ')) return p;
                          seq += 1;
                          return { ...p, peran: `Pihak ${toPartyRoman(seq)}` as PihakRole };
                        });
                      })
                    }
                    className="px-2.5 py-1 rounded-md bg-sky-800 text-white font-medium hover:bg-sky-700"
                  >
                    Pisahkan Jadi Pihak I, II, III...
                  </button>
                </div>
              )}

            <div className="space-y-4">
              {paraPihak.map((p, idx) => {
                const sameRoleCount = paraPihak.filter((x) => x.peran === p.peran).length;
                const sameRoleIdx =
                  paraPihak
                    .filter((x) => x.peran === p.peran)
                    .findIndex((x) => x.id === p.id) + 1;
                return (
                <div
                  key={p.id}
                  className="border border-slate-200 rounded-xl p-4 bg-slate-50/40 space-y-3"
                >
                  <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-slate-200">
                    <div className="flex flex-wrap items-center gap-2.5">
                      <span className="font-mono text-xs font-semibold text-slate-500">
                        Subjek #{idx + 1}
                      </span>
                      <select
                        aria-label={`Peran subjek ${idx + 1}`}
                        value={p.peran}
                        onChange={(e) =>
                          updateParty(p.id, { peran: e.target.value as PihakRole })
                        }
                        className="text-xs font-semibold border border-slate-300 rounded-md px-2.5 py-1 bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-sky-500"
                      >
                        {buildRoleOptions(paraPihak, layanan).map((r) => (
                          <option key={r} value={r}>
                            {r}
                          </option>
                        ))}
                      </select>
                      {sameRoleCount > 1 && p.peran.startsWith('Pihak ') && (
                        <span className="text-[11px] font-medium text-sky-800 bg-sky-50 border border-sky-200 px-2 py-0.5 rounded">
                          Kepemilikan Bersama #{sameRoleIdx} ({p.peran})
                        </span>
                      )}
                    </div>
                    <button
                      type="button"
                      onClick={() => removeParty(p.id)}
                      aria-label={`Hapus pihak ${p.nama || idx + 1}`}
                      className="inline-flex items-center gap-1 text-xs text-rose-600 hover:text-rose-800 font-medium"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      <span>Hapus</span>
                    </button>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
                    <div>
                      <label className="block text-slate-600 mb-1">NIK (16 Digit)</label>
                      <input
                        type="text"
                        value={p.nik}
                        onChange={(e) => updateParty(p.id, { nik: e.target.value })}
                        placeholder="3322..."
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                      />
                    </div>
                    <div className="lg:col-span-2">
                      <label className="block text-slate-600 mb-1">Nama Lengkap</label>
                      <input
                        type="text"
                        value={p.nama}
                        onChange={(e) => updateParty(p.id, { nama: e.target.value })}
                        placeholder="Nama sesuai KTP"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-medium"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Status Perkawinan</label>
                      <input
                        type="text"
                        value={p.status_perkawinan}
                        onChange={(e) =>
                          updateParty(p.id, { status_perkawinan: e.target.value })
                        }
                        placeholder="KAWIN / BELUM KAWIN"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Tempat Lahir</label>
                      <input
                        type="text"
                        value={p.tempat_lahir}
                        onChange={(e) =>
                          updateParty(p.id, { tempat_lahir: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Tanggal Lahir</label>
                      <DateInputId
                        value={p.tanggal_lahir}
                        onChange={(val) => updateParty(p.id, { tanggal_lahir: val })}
                        placeholder="DD-MM-YYYY"
                        ariaLabel={`Tanggal lahir ${p.nama || p.peran}`}
                        className="py-1.5 px-2.5 rounded-md"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Jenis Kelamin</label>
                      <select
                        value={p.jenis_kelamin}
                        onChange={(e) =>
                          updateParty(p.id, { jenis_kelamin: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      >
                        <option value="Laki-laki">Laki-laki</option>
                        <option value="Perempuan">Perempuan</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Pekerjaan</label>
                      <input
                        type="text"
                        value={p.pekerjaan}
                        onChange={(e) =>
                          updateParty(p.id, { pekerjaan: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block text-slate-600 mb-1">Alamat Jalan / Dukuh</label>
                      <input
                        type="text"
                        value={p.alamat_jalan}
                        onChange={(e) =>
                          updateParty(p.id, { alamat_jalan: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">RT / RW</label>
                      <input
                        type="text"
                        value={p.rt_rw}
                        onChange={(e) => updateParty(p.id, { rt_rw: e.target.value })}
                        placeholder="001/002"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Desa / Kelurahan</label>
                      <input
                        type="text"
                        value={p.desa_kelurahan}
                        onChange={(e) =>
                          updateParty(p.id, { desa_kelurahan: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Kecamatan</label>
                      <input
                        type="text"
                        value={p.kecamatan}
                        onChange={(e) =>
                          updateParty(p.id, { kecamatan: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">Kabupaten / Kota</label>
                      <input
                        type="text"
                        value={p.kabupaten}
                        onChange={(e) =>
                          updateParty(p.id, { kabupaten: e.target.value })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                      />
                    </div>
                  </div>
                </div>
                );
              })}
            </div>

            {/* Pengaturan Persetujuan Pasangan: hanya untuk Akta PPATS */}
            {isAkta && (
              <div className="border-t border-slate-200 pt-4 space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                  <div>
                    <h3 className="text-sm font-semibold text-slate-900">
                      Persetujuan Pasangan / Pihak Terkait
                    </h3>
                    <p className="text-xs text-slate-500">
                      Pilih apakah perbuatan hukum ini memerlukan persetujuan pasangan (suami/istri) atau tidak.
                    </p>
                  </div>
                  <div className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100">
                    <button
                      type="button"
                      onClick={() =>
                        setAkta((prev) => ({ ...prev, butuh_persetujuan_pasangan: true }))
                      }
                      className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-colors ${
                        akta.butuh_persetujuan_pasangan !== false
                          ? 'bg-[#0f172a] text-white'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Ada Persetujuan
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        setAkta((prev) => ({
                          ...prev,
                          butuh_persetujuan_pasangan: false,
                          keterangan_tanpa_persetujuan:
                            prev.keterangan_tanpa_persetujuan || DEFAULT_TANPA_PERSETUJUAN_TEXT,
                        }))
                      }
                      className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-colors ${
                        akta.butuh_persetujuan_pasangan === false
                          ? 'bg-[#0f172a] text-white'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Tidak Memerlukan Persetujuan
                    </button>
                  </div>
                </div>

                {akta.butuh_persetujuan_pasangan === false && (
                  <div className="p-4 rounded-xl border border-amber-200 bg-amber-50/50 space-y-2.5 text-xs">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <label
                        htmlFor="keterangan-tanpa-persetujuan"
                        className="font-semibold text-slate-800"
                      >
                        Kolom Keterangan Tanpa Persetujuan (akan dicetak dalam Akta):
                      </label>
                      <div className="flex flex-wrap gap-1.5">
                        <button
                          type="button"
                          onClick={() =>
                            setAkta((prev) => ({
                              ...prev,
                              keterangan_tanpa_persetujuan:
                                'Menurut keterangan para pihak, dalam melakukan perbuatan hukum ini tanpa memerlukan persetujuan dari siapapun, dikarenakan tanah tersebut berasal dari Warisan.',
                            }))
                          }
                          className="px-2 py-1 rounded border border-slate-200 bg-white hover:bg-slate-50 text-slate-700"
                        >
                          Template Warisan
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            setAkta((prev) => ({
                              ...prev,
                              keterangan_tanpa_persetujuan:
                                'Menurut keterangan para pihak, dalam melakukan perbuatan hukum ini tanpa memerlukan persetujuan dari siapapun, dikarenakan yang bersangkutan berstatus Belum Kawin / Cerai.',
                            }))
                          }
                          className="px-2 py-1 rounded border border-slate-200 bg-white hover:bg-slate-50 text-slate-700"
                        >
                          Template Belum Kawin / Cerai
                        </button>
                      </div>
                    </div>
                    <textarea
                      id="keterangan-tanpa-persetujuan"
                      rows={2}
                      value={akta.keterangan_tanpa_persetujuan ?? ''}
                      onChange={(e) =>
                        setAkta((prev) => ({
                          ...prev,
                          keterangan_tanpa_persetujuan: e.target.value,
                        }))
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-sky-500"
                    />
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Data Tanah & PBB */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2.5">
                <h2 className="font-display text-base font-semibold text-slate-900">
                  Data Objek Tanah (Sertipikat)
                </h2>
                <div
                  role="radiogroup"
                  aria-label="Pilih Jenis Sertipikat"
                  className="inline-flex rounded-lg border border-slate-200 p-0.5 bg-slate-100 self-start sm:self-auto"
                >
                  <button
                    type="button"
                    role="radio"
                    aria-checked={resolveJenisSertipikat(tanah) === 'digital'}
                    onClick={() =>
                      setTanah((prev) => ({ ...prev, jenis_sertipikat: 'digital' }))
                    }
                    className={`px-3 py-1 rounded-md text-xs font-semibold transition-colors whitespace-nowrap ${
                      resolveJenisSertipikat(tanah) === 'digital'
                        ? 'bg-[#0f172a] text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Digital (Sertipikat Elektronik)
                  </button>
                  <button
                    type="button"
                    role="radio"
                    aria-checked={resolveJenisSertipikat(tanah) === 'analog'}
                    onClick={() =>
                      setTanah((prev) => ({ ...prev, jenis_sertipikat: 'analog' }))
                    }
                    className={`px-3 py-1 rounded-md text-xs font-semibold transition-colors whitespace-nowrap ${
                      resolveJenisSertipikat(tanah) === 'analog'
                        ? 'bg-[#0f172a] text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Analog (SHM / Fisik)
                  </button>
                </div>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label htmlFor="tanah-jenis-hak" className="block text-slate-600 mb-1">
                    Jenis Hak
                  </label>
                  <input
                    id="tanah-jenis-hak"
                    type="text"
                    value={tanah.jenis_hak}
                    onChange={(e) => setTanah({ ...tanah, jenis_hak: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                {resolveJenisSertipikat(tanah) === 'digital' ? (
                  <div>
                    <label htmlFor="tanah-nib" className="block text-slate-600 mb-1">
                      NIBel (Nomor Identifikasi Bidang Elektronik)
                    </label>
                    <input
                      id="tanah-nib"
                      type="text"
                      value={tanah.nib}
                      onChange={(e) => setTanah({ ...tanah, nib: e.target.value })}
                      placeholder="Contoh: 11.17.000023383.0"
                      className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                    />
                    <p className="text-[11px] text-slate-500 mt-1">
                      Pada akta hanya ditulis NIBel (tanpa Nomor Hak atau Surat Ukur).
                    </p>
                  </div>
                ) : (
                  <>
                    <div>
                      <label htmlFor="tanah-no-hak" className="block text-slate-600 mb-1">
                        Nomor SHM / Sertipikat
                      </label>
                      <input
                        id="tanah-no-hak"
                        type="text"
                        value={tanah.nomor_hak_lama}
                        onChange={(e) =>
                          setTanah({ ...tanah, nomor_hak_lama: e.target.value })
                        }
                        placeholder="Contoh: 3354"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                      />
                    </div>
                    <div>
                      <label htmlFor="tanah-nib" className="block text-slate-600 mb-1">
                        NIB (Nomor Identifikasi Bidang)
                      </label>
                      <input
                        id="tanah-nib"
                        type="text"
                        value={tanah.nib}
                        onChange={(e) => setTanah({ ...tanah, nib: e.target.value })}
                        placeholder="Contoh: 11.17.11.10.01450"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                      />
                    </div>
                    <div>
                      <label htmlFor="tanah-su-tanggal" className="block text-slate-600 mb-1">
                        Tanggal Surat Ukur
                      </label>
                      <input
                        id="tanah-su-tanggal"
                        type="text"
                        value={tanah.surat_ukur_tanggal ?? ''}
                        onChange={(e) =>
                          setTanah({ ...tanah, surat_ukur_tanggal: e.target.value })
                        }
                        placeholder="DD-MM-YYYY (atau 19/10/2011)"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                      />
                    </div>
                    <div>
                      <label htmlFor="tanah-su-nomor" className="block text-slate-600 mb-1">
                        Nomor Surat Ukur
                      </label>
                      <input
                        id="tanah-su-nomor"
                        type="text"
                        value={tanah.surat_ukur_nomor ?? ''}
                        onChange={(e) =>
                          setTanah({ ...tanah, surat_ukur_nomor: e.target.value })
                        }
                        placeholder="Contoh: 00143/Sindon/2011"
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                      />
                    </div>
                  </>
                )}
                <div>
                  <label htmlFor="tanah-luas" className="block text-slate-600 mb-1">
                    Luas Tanah (m²)
                  </label>
                  <input
                    id="tanah-luas"
                    type="number"
                    min="0"
                    step="0.01"
                    value={tanah.luas_m2}
                    onChange={(e) =>
                      setTanah({ ...tanah, luas_m2: Number(e.target.value) || 0 })
                    }
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono tabular-nums"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-desa" className="block text-slate-600 mb-1">
                    Desa / Kelurahan
                  </label>
                  <input
                    id="tanah-desa"
                    type="text"
                    value={tanah.desa}
                    onChange={(e) => setTanah({ ...tanah, desa: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-kec" className="block text-slate-600 mb-1">
                    Kecamatan
                  </label>
                  <input
                    id="tanah-kec"
                    type="text"
                    value={tanah.kecamatan}
                    onChange={(e) => setTanah({ ...tanah, kecamatan: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-kab" className="block text-slate-600 mb-1">
                    Kabupaten / Kota
                  </label>
                  <input
                    id="tanah-kab"
                    type="text"
                    value={tanah.kabupaten}
                    onChange={(e) => setTanah({ ...tanah, kabupaten: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-provinsi" className="block text-slate-600 mb-1">
                    Provinsi
                  </label>
                  <input
                    id="tanah-provinsi"
                    type="text"
                    value={tanah.provinsi ?? ''}
                    onChange={(e) => setTanah({ ...tanah, provinsi: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-dasar" className="block text-slate-600 mb-1">
                    Dasar Perolehan
                  </label>
                  <input
                    id="tanah-dasar"
                    type="text"
                    value={tanah.dasar_perolehan}
                    onChange={(e) =>
                      setTanah({ ...tanah, dasar_perolehan: e.target.value })
                    }
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-penggunaan" className="block text-slate-600 mb-1">
                    Penggunaan Tanah (untuk Akta & Lampiran F4)
                  </label>
                  <input
                    id="tanah-penggunaan"
                    type="text"
                    value={tanah.penggunaan_tanah ?? ''}
                    onChange={(e) =>
                      setTanah({ ...tanah, penggunaan_tanah: e.target.value })
                    }
                    placeholder="Misal: Pekarangan / Tanah Pertanian"
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-status" className="block text-slate-600 mb-1">
                    Status Jenis Tanah (Pasal Akta)
                  </label>
                  <select
                    id="tanah-status"
                    value={tanah.status_tanah ?? 'Non-Pertanian'}
                    onChange={(e) =>
                      setTanah({
                        ...tanah,
                        status_tanah: e.target.value as 'Non-Pertanian' | 'Pertanian',
                      })
                    }
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                  >
                    <option value="Non-Pertanian">Non-Pertanian (Pekarangan / Bangunan)</option>
                    <option value="Pertanian">Tanah Pertanian (Sisipkan Pasal 3 Maksimum Tanah)</option>
                  </select>
                </div>
                {tanah.status_tanah === 'Pertanian' && (
                  <div className="sm:col-span-2 p-2.5 rounded-lg border border-emerald-200 bg-emerald-50/70 space-y-1.5">
                    <label
                      htmlFor="tanah-tgl-pertanian"
                      className="block font-medium text-emerald-900"
                    >
                      Tanggal Surat Pernyataan Tidak Melanggar Batas Maksimum Tanah Pertanian
                    </label>
                    <DateInputId
                      id="tanah-tgl-pertanian"
                      allowEmpty
                      value={tanah.tanggal_pernyataan_pertanian ?? ''}
                      onChange={(val) =>
                        setTanah({ ...tanah, tanggal_pernyataan_pertanian: val })
                      }
                      placeholder="DD-MM-YYYY (kosongkan jika ikuti tanggal akta)"
                      className="border-emerald-300 rounded-md py-1.5"
                    />
                    <p className="text-[11px] text-emerald-800">
                      Pasal 3 khusus ketentuan maksimum penguasaan tanah pertanian akan otomatis disisipkan dan nomor pasal berikutnya bergeser (+1).
                    </p>
                  </div>
                )}
                <div>
                  <label htmlFor="tanah-tahun-kuasa" className="block text-slate-600 mb-1">
                    Dikuasai Sejak Tahun
                  </label>
                  <input
                    id="tanah-tahun-kuasa"
                    type="text"
                    value={tanah.menguasai_sejak_tahun ?? ''}
                    onChange={(e) =>
                      setTanah({ ...tanah, menguasai_sejak_tahun: e.target.value })
                    }
                    placeholder="Misal: 2012"
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-batas-utara" className="block text-slate-600 mb-1">
                    Batas Sebelah Utara
                  </label>
                  <input
                    id="tanah-batas-utara"
                    type="text"
                    value={tanah.batas_utara ?? ''}
                    onChange={(e) => setTanah({ ...tanah, batas_utara: e.target.value })}
                    placeholder="Misal: Jalan Kampung / Tanah Sdr. ..."
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-batas-timur" className="block text-slate-600 mb-1">
                    Batas Sebelah Timur
                  </label>
                  <input
                    id="tanah-batas-timur"
                    type="text"
                    value={tanah.batas_timur ?? ''}
                    onChange={(e) => setTanah({ ...tanah, batas_timur: e.target.value })}
                    placeholder="Misal: Tanah Sdr. ..."
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-batas-selatan" className="block text-slate-600 mb-1">
                    Batas Sebelah Selatan
                  </label>
                  <input
                    id="tanah-batas-selatan"
                    type="text"
                    value={tanah.batas_selatan ?? ''}
                    onChange={(e) => setTanah({ ...tanah, batas_selatan: e.target.value })}
                    placeholder="Misal: Tanah Sdr. ..."
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div>
                  <label htmlFor="tanah-batas-barat" className="block text-slate-600 mb-1">
                    Batas Sebelah Barat
                  </label>
                  <input
                    id="tanah-batas-barat"
                    type="text"
                    value={tanah.batas_barat ?? ''}
                    onChange={(e) => setTanah({ ...tanah, batas_barat: e.target.value })}
                    placeholder="Misal: Selokan / Jalan ..."
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
                <div className="sm:col-span-2">
                  <label htmlFor="tanah-holders" className="block text-slate-600 mb-1">
                    Pemegang Hak pada Sertipikat (pisahkan koma jika lebih dari 1)
                  </label>
                  <input
                    id="tanah-holders"
                    type="text"
                    value={pemegangHakInput}
                    onChange={(e) => setPemegangHakInput(e.target.value)}
                    placeholder="Contoh: Slamet Raharjo"
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                  />
                </div>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
              <h2 className="font-display text-base font-semibold text-slate-900">
                Data SPPT PBB-P2
              </h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div className="sm:col-span-2">
                  <label htmlFor="pbb-nop" className="block text-slate-600 mb-1">
                    Nomor Objek Pajak (NOP)
                  </label>
                  <input
                    id="pbb-nop"
                    type="text"
                    value={pbb.nop}
                    onChange={(e) => setPbb({ ...pbb, nop: e.target.value })}
                    placeholder="33.22.140.xxx.xxx-xxxx.0"
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                  />
                </div>
                <div>
                  <label htmlFor="pbb-tahun" className="block text-slate-600 mb-1">
                    Tahun Pajak
                  </label>
                  <input
                    id="pbb-tahun"
                    type="text"
                    value={pbb.tahun_pajak}
                    onChange={(e) => setPbb({ ...pbb, tahun_pajak: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                  />
                </div>
                <div>
                  <label htmlFor="pbb-njop-bumi" className="block text-slate-600 mb-1">
                    NJOP Bumi / m² (Rp)
                  </label>
                  <input
                    id="pbb-njop-bumi"
                    type="number"
                    min="0"
                    value={pbb.njop_bumi_m2}
                    onChange={(e) =>
                      setPbb({ ...pbb, njop_bumi_m2: Number(e.target.value) || 0 })
                    }
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono tabular-nums"
                  />
                </div>
                <div className="sm:col-span-2">
                  <label htmlFor="pbb-njop-bangunan" className="block text-slate-600 mb-1">
                    Total NJOP Bangunan (Rp)
                  </label>
                  <input
                    id="pbb-njop-bangunan"
                    type="number"
                    min="0"
                    value={pbb.njop_bangunan}
                    onChange={(e) =>
                      setPbb({ ...pbb, njop_bangunan: Number(e.target.value) || 0 })
                    }
                    className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono tabular-nums"
                  />
                </div>
              </div>
            </div>
          </div>

          <div className="flex items-center justify-between pt-2">
            <button
              type="button"
              onClick={() => setStep(1)}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-sm font-medium text-slate-700"
            >
              <ArrowLeft className="w-4 h-4" />
              <span>Kembali ke Langkah 1</span>
            </button>
            <button
              type="button"
              onClick={() => setStep(3)}
              className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium"
            >
              <span>
                {isAkta
                  ? 'Lanjut ke Parameter Akta & Saksi'
                  : 'Lanjut ke Pengesahan & Detail Berkas'}
              </span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* STEP 3 */}
      {step === 3 && (
        <div className="space-y-6">
          <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-5">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 border-b border-slate-200 pb-4">
              <div>
                <h2 className="font-display text-base font-semibold text-slate-900">
                  {isAkta
                    ? 'Parameter Akta & Penandatanganan'
                    : `Pengesahan, Detail Berkas Non-Akta (${layanan}) & Saksi`}
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  {isAkta
                    ? 'Lengkapi nomor akta, pejabat PPATS, kuasa BPN, saksi-saksi, dan nilai transaksi pada akta.'
                    : 'Atur pejabat pengesahan dokumen, detail khusus layanan Non-Akta, kuasa BPN, dan saksi-saksi.'}
                </p>
              </div>
              <button
                type="button"
                onClick={async () => {
                  const rec = buildCurrentRecord();
                  setSavedRecordId(rec.id);
                  await onOpenTaxCalculator(rec);
                }}
                className="inline-flex items-center gap-2 px-3.5 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-medium text-slate-800 whitespace-nowrap shrink-0"
              >
                <Calculator className="w-3.5 h-3.5 text-sky-700" />
                <span>Buka Jalur Perhitungan Pajak (BPHTB/PPh)</span>
              </button>
            </div>

            {!isAkta && (
              <div className="p-3.5 rounded-lg border border-slate-200 bg-slate-50 text-xs text-slate-700">
                <strong>Berkas Non-Akta ({layanan}):</strong> Tidak menggunakan Nomor Akta PPATS, tidak menerbitkan Blangko Akta A4, dan otomatis dikecualikan dari Laporan Bulanan PPATS.
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
              {isAkta && (
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label htmlFor="akta-nomor" className="text-slate-600 font-medium">
                      Nomor Akta PPATS
                    </label>
                    <button
                      type="button"
                      onClick={() =>
                        setAkta({
                          ...akta,
                          nomor_akta: suggestNextAktaNumber(existingRecords, akta.tanggal_akta)
                            .nomorAkta,
                        })
                      }
                      className="text-sky-700 hover:underline font-medium"
                    >
                      Pakai Saran ({nextSuggestion.nomorAkta})
                    </button>
                  </div>
                  <input
                    id="akta-nomor"
                    type="text"
                    value={akta.nomor_akta}
                    onChange={(e) => setAkta({ ...akta, nomor_akta: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono"
                  />
                </div>
              )}

              <div className={!isAkta ? 'sm:col-span-2 grid grid-cols-1 sm:grid-cols-2 gap-4' : ''}>
                <div>
                  <div className="flex items-center justify-between gap-2 mb-1">
                    <label htmlFor="akta-tanggal" className="text-slate-600 font-medium">
                      {isAkta ? 'Tanggal Akta (Opsional)' : 'Tanggal Surat / Berkas (Opsional)'}
                    </label>
                    {akta.tanggal_akta.trim() !== '' && (
                      <button
                        type="button"
                        onClick={() => setAkta({ ...akta, tanggal_akta: '' })}
                        className="text-[11px] font-medium text-rose-600 hover:underline"
                      >
                        Kosongkan Tanggal
                      </button>
                    )}
                  </div>
                  <DateInputId
                    id="akta-tanggal"
                    allowEmpty
                    value={akta.tanggal_akta}
                    onChange={(val) => setAkta({ ...akta, tanggal_akta: val })}
                    placeholder="DD-MM-YYYY (boleh dikosongi)"
                    ariaLabel={isAkta ? 'Tanggal Akta (format DD-MM-YYYY)' : 'Tanggal Surat Berkas (format DD-MM-YYYY)'}
                  />
                  <p className="text-[11px] text-slate-500 mt-1">
                    Format hari-bulan-tahun (DD-MM-YYYY). Boleh dikosongkan untuk pengisian blangko manual.
                  </p>
                </div>
                {!isAkta && (
                  <div>
                    <label htmlFor="akta-status-nonakta" className="block text-slate-600 font-medium mb-1">
                      Status Progres Berkas
                    </label>
                    <select
                      id="akta-status-nonakta"
                      value={status}
                      onChange={(e) => setStatus(e.target.value as BerkasStatus)}
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg bg-white"
                    >
                      {STATUS_OPTIONS.map((opt) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  </div>
                )}
              </div>

              {!isAkta && !akta.tanggal_akta.trim() && (
                <div className="sm:col-span-2 p-3 rounded-lg bg-amber-50/70 border border-amber-200 text-amber-900">
                  <span className="font-semibold">Tanggal Berkas Dikosongkan:</span>{' '}
                  Seluruh titik tanggal pada bundel Non-Akta (Permohonan, Surat Kuasa, Pernyataan, dan Pengesahan) akan dicetak kosong (<span className="font-mono">......................</span>) untuk diisi manual.
                </div>
              )}

              {isAkta && (
                <>
                  <div className="sm:col-span-2 p-3 rounded-lg bg-slate-50 border border-slate-200 text-slate-700">
                    <span className="font-semibold text-slate-900">Pratinjau Kalimat Tanggal Akta:</span>{' '}
                    {akta.tanggal_akta.trim() && formattedAktaDate.kalimat_akta !== '-'
                      ? formattedAktaDate.kalimat_akta
                      : 'Tanggal dikosongkan — blangko akta & lampiran akan dicetak dengan titik-titik (..........) untuk pengisian manual.'}
                  </div>

                  <div>
                    <label htmlFor="akta-ppats" className="block text-slate-600 font-medium mb-1">
                      Nama Pejabat PPATS
                    </label>
                    <input
                      id="akta-ppats"
                      type="text"
                      value={akta.nama_ppats}
                      onChange={(e) => setAkta({ ...akta, nama_ppats: e.target.value })}
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div>
                    <label htmlFor="akta-jabatan" className="block text-slate-600 font-medium mb-1">
                      Jabatan PPATS
                    </label>
                    <input
                      id="akta-jabatan"
                      type="text"
                      value={akta.jabatan_ppats}
                      onChange={(e) => setAkta({ ...akta, jabatan_ppats: e.target.value })}
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div className="sm:col-span-2 p-3.5 rounded-lg border border-slate-200 bg-slate-50/70 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                    <div className="space-y-0.5">
                      <p className="font-semibold text-slate-800">
                        Ringkasan Nilai Transaksi &amp; Pajak Tersimpan
                      </p>
                      <p className="text-[11px] text-slate-600 font-mono tabular-nums">
                        Nilai Kesepakatan: {formatRupiah(taxCalc.nilai_kesepakatan)} · BPHTB: {formatRupiah(taxCalc.bphtb)} · PPh Final: {formatRupiah(taxCalc.pph)}
                      </p>
                    </div>
                    <div className="sm:w-64">
                      <label htmlFor="akta-override" className="block text-slate-600 font-medium mb-1">
                        Nilai Transaksi Akta (Rp, opsional)
                      </label>
                      <input
                        id="akta-override"
                        type="number"
                        min="0"
                        value={akta.nilai_transaksi_override ?? ''}
                        onChange={(e) =>
                          setAkta({
                            ...akta,
                            nilai_transaksi_override:
                              e.target.value === '' ? null : Number(e.target.value),
                          })
                        }
                        placeholder="Atur via Perhitungan Pajak"
                        className="w-full px-3 py-1.5 border border-slate-200 rounded-lg bg-white font-mono tabular-nums"
                      />
                    </div>
                  </div>
                </>
              )}

              {layanan === 'APHB' && (
                <div className="sm:col-span-2 p-3 rounded-lg border border-sky-200 bg-sky-50/60">
                  <label htmlFor="akta-penerima-aphb" className="block text-slate-800 font-semibold mb-1">
                    Pihak Penerima Tunggal Hak Bersama (APHB)
                  </label>
                  <select
                    id="akta-penerima-aphb"
                    value={akta.pihak_penerima_aphb ?? 'Pihak II'}
                    onChange={(e) =>
                      setAkta({ ...akta, pihak_penerima_aphb: e.target.value })
                    }
                    className="w-full px-3 py-2 border border-sky-200 rounded-lg bg-white text-slate-900"
                  >
                    {paraPihak
                      .filter((p) => p.peran.startsWith('Pihak '))
                      .sort(
                        (a, b) =>
                          parseAdditionalPihakMeta(a.peran).order -
                          parseAdditionalPihakMeta(b.peran).order
                      )
                      .map((p) => (
                        <option key={p.id} value={p.peran}>
                          {p.peran} — {p.nama || '(Nama belum diisi)'}
                        </option>
                      ))}
                  </select>
                </div>
              )}

              {isAkta && (
                <>
                  <div>
                    <label htmlFor="akta-status" className="block text-slate-600 font-medium mb-1">
                      Status Progres Berkas
                    </label>
                    <select
                      id="akta-status"
                      value={status}
                      onChange={(e) => setStatus(e.target.value as BerkasStatus)}
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg bg-white"
                    >
                      {STATUS_OPTIONS.map((opt) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="hidden sm:block" />
                </>
              )}

              {/* PANEL PENGESAHAN & DETAIL NON-AKTA */}
              {!isAkta && (
                <>
                  {/* 1. Pengesahan Tunggal ("Mengetahui") - Default PPATS */}
                  <div className="sm:col-span-2 border border-slate-200 rounded-xl p-4 bg-slate-50/50 space-y-3">
                    <div className="border-b border-slate-200 pb-2">
                      <h3 className="text-xs font-semibold text-slate-900">
                        Pengesahan Dokumen Non-Akta Lainnya (&ldquo;Mengetahui&rdquo;) — Default PPATS
                      </h3>
                      <p className="text-[11px] text-slate-500 mt-0.5">
                        Berlaku untuk Surat Kuasa, Penguasaan Fisik, SPPT/PBB, BPHTB Waris, Permohonan Pemecahan/Ukur, Tanda Batas, Sket, Pelepasan Jalan, dan Penutupan PNBP. Secara default menggunakan PPATS, atau pilih Kepala Desa / Camat / Isi Manual sesuai kebutuhan.
                      </p>
                    </div>

                    <div
                      role="radiogroup"
                      aria-label="Pilih Pejabat Pengesahan Dokumen Lainnya"
                      className="grid grid-cols-2 lg:grid-cols-4 gap-2"
                    >
                      {(
                        [
                          { id: 'ppats' as const, label: 'Default PPATS' },
                          { id: 'kades' as const, label: 'Kepala Desa Asal' },
                          { id: 'camat' as const, label: 'Camat' },
                          { id: 'manual' as const, label: 'Isi Manual / Lainnya' },
                        ] satisfies Array<{ id: PengesahanMode; label: string }>
                      ).map((opt) => {
                        const active = (nonAkta.pengesahan_mode ?? 'ppats') === opt.id;
                        return (
                          <button
                            key={opt.id}
                            type="button"
                            role="radio"
                            aria-checked={active}
                            onClick={() => {
                              setNonAkta((prev) => {
                                if (opt.id === 'ppats') {
                                  return {
                                    ...prev,
                                    pengesahan_mode: 'ppats',
                                    pengesahan_jabatan: '',
                                    pengesahan_nama: akta.nama_ppats || settings.nama_ppats,
                                    pengesahan_nip: akta.nip_ppats || settings.nip_ppats,
                                  };
                                }
                                if (opt.id === 'kades') {
                                  return {
                                    ...prev,
                                    pengesahan_mode: 'kades',
                                    pengesahan_jabatan: '',
                                    pengesahan_nama: prev.kades_pengesahan_nama || '',
                                    pengesahan_nip: '',
                                  };
                                }
                                if (opt.id === 'camat') {
                                  return {
                                    ...prev,
                                    pengesahan_mode: 'camat',
                                    pengesahan_jabatan: '',
                                    pengesahan_nama: prev.camat_pengesahan_nama || akta.nama_ppats || settings.nama_ppats,
                                    pengesahan_nip: prev.camat_pengesahan_nip || akta.nip_ppats || settings.nip_ppats,
                                  };
                                }
                                return {
                                  ...prev,
                                  pengesahan_mode: 'manual',
                                  pengesahan_jabatan:
                                    prev.pengesahan_mode === 'manual' && prev.pengesahan_jabatan
                                      ? prev.pengesahan_jabatan
                                      : '',
                                  pengesahan_nama:
                                    prev.pengesahan_mode === 'manual' ? prev.pengesahan_nama : '',
                                  pengesahan_nip:
                                    prev.pengesahan_mode === 'manual' ? prev.pengesahan_nip : '',
                                };
                              });
                            }}
                            className={`px-3 py-2 rounded-lg border text-xs font-semibold text-left transition-colors ${
                              active
                                ? 'border-[#0f172a] bg-[#0f172a] text-white'
                                : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                            }`}
                          >
                            {opt.label}
                          </button>
                        );
                      })}
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                      <div>
                        <label htmlFor="na-pengesahan-jabatan" className="block text-slate-600 mb-1">
                          Jabatan Pengesahan
                        </label>
                        <input
                          id="na-pengesahan-jabatan"
                          type="text"
                          value={nonAkta.pengesahan_jabatan ?? ''}
                          onChange={(e) =>
                            setNonAkta({ ...nonAkta, pengesahan_jabatan: e.target.value })
                          }
                          placeholder={
                            nonAkta.pengesahan_mode === 'kades'
                              ? `Kepala Desa ${tanah.desa || '-'}`
                              : nonAkta.pengesahan_mode === 'camat'
                                ? `Camat ${(tanah.kecamatan || settings.nama_kecamatan || '-').replace(/^Kecamatan\s+/i, '')}`
                                : nonAkta.pengesahan_mode === 'ppats'
                                  ? `PPATS / Camat ${(tanah.kecamatan || settings.nama_kecamatan || '-').replace(/^Kecamatan\s+/i, '')}`
                                  : 'Jabatan pejabat pengesahan'
                          }
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                      <div>
                        <label htmlFor="na-pengesahan-nama" className="block text-slate-600 mb-1">
                          Nama Pejabat Pengesahan
                        </label>
                        <input
                          id="na-pengesahan-nama"
                          type="text"
                          value={nonAkta.pengesahan_nama ?? ''}
                          onChange={(e) =>
                            setNonAkta({ ...nonAkta, pengesahan_nama: e.target.value })
                          }
                          placeholder="Nama & gelar pejabat"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-medium"
                        />
                      </div>
                      <div>
                        <label htmlFor="na-pengesahan-nip" className="block text-slate-600 mb-1">
                          NIP (Opsional)
                        </label>
                        <input
                          id="na-pengesahan-nip"
                          type="text"
                          value={nonAkta.pengesahan_nip ?? ''}
                          onChange={(e) =>
                            setNonAkta({ ...nonAkta, pengesahan_nip: e.target.value })
                          }
                          placeholder="Kosongkan jika bukan ASN"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                        />
                      </div>
                    </div>
                  </div>

                  {/* 2. Detail Khusus Turun Waris */}
                  {layanan === 'Turun Waris' && (
                    <div className="sm:col-span-2 border border-slate-200 rounded-xl p-4 bg-white space-y-3">
                      <h3 className="text-xs font-semibold text-slate-900 border-b border-slate-200 pb-2">
                        Data Pewaris &amp; Surat Keterangan Warisan (SKW)
                      </h3>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        <div>
                          <label htmlFor="tw-sebutan" className="block text-slate-600 mb-1">
                            Sebutan Pewaris
                          </label>
                          <select
                            id="tw-sebutan"
                            value={nonAkta.sebutan_pewaris ?? 'Almarhum'}
                            onChange={(e) =>
                              setNonAkta({
                                ...nonAkta,
                                sebutan_pewaris: e.target.value as 'Almarhum' | 'Almarhumah',
                              })
                            }
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          >
                            <option value="Almarhum">Almarhum</option>
                            <option value="Almarhumah">Almarhumah</option>
                          </select>
                        </div>
                        <div className="sm:col-span-2">
                          <label htmlFor="tw-pewaris" className="block text-slate-600 mb-1">
                            Nama Pewaris (Pemegang Hak Meninggal)
                          </label>
                          <input
                            id="tw-pewaris"
                            type="text"
                            value={nonAkta.nama_pewaris ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, nama_pewaris: e.target.value })
                            }
                            placeholder="Contoh: HADI SUWARNO"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-medium"
                          />
                        </div>
                        <div>
                          <label htmlFor="tw-tgl-meninggal" className="block text-slate-600 mb-1">
                            Tanggal Meninggal (Opsional)
                          </label>
                          <DateInputId
                            id="tw-tgl-meninggal"
                            allowEmpty
                            value={nonAkta.tanggal_meninggal ?? ''}
                            onChange={(val) =>
                              setNonAkta({ ...nonAkta, tanggal_meninggal: val })
                            }
                            placeholder="DD-MM-YYYY"
                            className="py-1.5 px-2.5 rounded-md"
                          />
                        </div>
                        <div className="sm:col-span-2">
                          <label htmlFor="tw-tempat-meninggal" className="block text-slate-600 mb-1">
                            Tempat Meninggal / Akta Kematian
                          </label>
                          <input
                            id="tw-tempat-meninggal"
                            type="text"
                            value={nonAkta.tempat_meninggal ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, tempat_meninggal: e.target.value })
                            }
                            placeholder="Contoh: Desa Sindon, Kec. Ngemplak, Boyolali"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                        <div>
                          <label htmlFor="tw-hubungan" className="block text-slate-600 mb-1">
                            Hubungan Pasangan Pewaris
                          </label>
                          <select
                            id="tw-hubungan"
                            value={nonAkta.hubungan_pasangan_pewaris ?? 'Istrinya'}
                            onChange={(e) =>
                              setNonAkta({
                                ...nonAkta,
                                hubungan_pasangan_pewaris: e.target.value as
                                  | 'Istrinya'
                                  | 'Suaminya'
                                  | 'Pasangannya',
                              })
                            }
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          >
                            <option value="Istrinya">Istrinya</option>
                            <option value="Suaminya">Suaminya</option>
                            <option value="Pasangannya">Pasangannya / Lainnya</option>
                          </select>
                        </div>
                        <div>
                          <label htmlFor="tw-pasangan" className="block text-slate-600 mb-1">
                            Nama Pasangan Pewaris
                          </label>
                          <input
                            id="tw-pasangan"
                            type="text"
                            value={nonAkta.nama_pasangan_pewaris ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, nama_pasangan_pewaris: e.target.value })
                            }
                            placeholder="Nama janda/duda pewaris"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                        <div>
                          <label htmlFor="tw-tgl-pasangan" className="block text-slate-600 mb-1">
                            Tgl Lahir Pasangan
                          </label>
                          <DateInputId
                            id="tw-tgl-pasangan"
                            allowEmpty
                            value={nonAkta.tgl_lahir_pasangan_pewaris ?? ''}
                            onChange={(val) =>
                              setNonAkta({
                                ...nonAkta,
                                tgl_lahir_pasangan_pewaris: val,
                              })
                            }
                            placeholder="DD-MM-YYYY"
                            className="py-1.5 px-2.5 rounded-md"
                          />
                        </div>
                        <div className="sm:col-span-3">
                          <label htmlFor="tw-catatan" className="block text-slate-600 mb-1">
                            Catatan Tambahan SKW (Opsional)
                          </label>
                          <input
                            id="tw-catatan"
                            type="text"
                            value={nonAkta.catatan_skw ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, catatan_skw: e.target.value })
                            }
                            placeholder="Misal: Pernikahan tercatat pada KUA Ngemplak Nomor ..."
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  {/* 3. Detail Khusus Pemecahan & Ukur Ulang */}
                  {(layanan === 'Pemecahan' ||
                    layanan === 'Ukur Ulang' ||
                    layanan === 'Pemecahan/Ukur Ulang') && (
                    <div className="sm:col-span-2 border border-slate-200 rounded-xl p-4 bg-white space-y-3">
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2">
                        <h3 className="text-xs font-semibold text-slate-900">
                          Detail Permohonan {layanan}, Bidang Pecahan &amp; Sket Tanah
                        </h3>
                        <select
                          aria-label="Jenis permohonan pengukuran atau pemecahan"
                          value={nonAkta.sub_permohonan_ukur ?? 'Pemecahan'}
                          onChange={(e) =>
                            setNonAkta({
                              ...nonAkta,
                              sub_permohonan_ukur: e.target.value as
                                | 'Pemecahan'
                                | 'Pemecahan & Validasi'
                                | 'Pengukuran'
                                | 'Pengukuran & Validasi',
                            })
                          }
                          className="px-2.5 py-1 border border-slate-300 rounded-md bg-white text-xs font-medium"
                        >
                          <option value="Pemecahan">Permohonan Pemecahan Bidang</option>
                          <option value="Pemecahan & Validasi">Pemecahan &amp; Validasi</option>
                          <option value="Pengukuran">Permohonan Pengukuran Ulang</option>
                          <option value="Pengukuran & Validasi">Pengukuran &amp; Validasi</option>
                        </select>
                      </div>

                      <div className="space-y-2">
                        <div className="flex items-center justify-between">
                          <span className="font-semibold text-slate-800">
                            Rincian Bidang Pecahan / Pengukuran ({(nonAkta.bidang_pecahan ?? []).length})
                          </span>
                          <button
                            type="button"
                            onClick={() => {
                              const nextNum = (nonAkta.bidang_pecahan ?? []).length + 1;
                              const item: PemecahanBidangItem = {
                                id: `bd-${Date.now()}-${nextNum}`,
                                label: `Bagian ${nextNum}`,
                                luas_m2: 0,
                              };
                              setNonAkta((prev) => ({
                                ...prev,
                                bidang_pecahan: [...(prev.bidang_pecahan ?? []), item],
                              }));
                            }}
                            className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md bg-[#0f172a] hover:bg-slate-800 text-white font-medium"
                          >
                            <Plus className="w-3 h-3" />
                            <span>Tambah Bidang Pecahan</span>
                          </button>
                        </div>

                        {(nonAkta.bidang_pecahan ?? []).map((b, bIdx) => (
                          <div
                            key={b.id}
                            className="p-3 rounded-lg border border-slate-200 bg-slate-50/70 grid grid-cols-1 sm:grid-cols-4 gap-2 items-end"
                          >
                            <div className="sm:col-span-2">
                              <label className="block text-slate-600 mb-1">
                                Nama / Label Bidang #{bIdx + 1}
                              </label>
                              <input
                                type="text"
                                value={b.label}
                                onChange={(e) =>
                                  setNonAkta((prev) => ({
                                    ...prev,
                                    bidang_pecahan: (prev.bidang_pecahan ?? []).map((x) =>
                                      x.id === b.id ? { ...x, label: e.target.value } : x
                                    ),
                                  }))
                                }
                                placeholder="Misal: Bagian 1 (Sdr. ...)"
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                              />
                            </div>
                            <div>
                              <label className="block text-slate-600 mb-1">Luas (m²)</label>
                              <input
                                type="number"
                                min="0"
                                step="0.01"
                                value={b.luas_m2}
                                onChange={(e) =>
                                  setNonAkta((prev) => ({
                                    ...prev,
                                    bidang_pecahan: (prev.bidang_pecahan ?? []).map((x) =>
                                      x.id === b.id
                                        ? { ...x, luas_m2: Number(e.target.value) || 0 }
                                        : x
                                    ),
                                  }))
                                }
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono tabular-nums"
                              />
                            </div>
                            <div className="flex justify-end">
                              <button
                                type="button"
                                onClick={() =>
                                  setNonAkta((prev) => ({
                                    ...prev,
                                    bidang_pecahan: (prev.bidang_pecahan ?? []).filter(
                                      (x) => x.id !== b.id
                                    ),
                                  }))
                                }
                                className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md border border-rose-200 text-rose-700 hover:bg-rose-50"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                                <span>Hapus</span>
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                        <label className="flex items-center gap-2 p-2.5 rounded-lg border border-slate-200 bg-slate-50/50 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={Boolean(nonAkta.sertakan_pelepasan_jalan)}
                            onChange={(e) =>
                              setNonAkta({
                                ...nonAkta,
                                sertakan_pelepasan_jalan: e.target.checked,
                              })
                            }
                            className="rounded border-slate-300"
                          />
                          <span className="font-medium text-slate-800">
                            Sertakan Surat Pelepasan Hak Tanah untuk Jalan
                          </span>
                        </label>
                        <label className="flex items-center gap-2 p-2.5 rounded-lg border border-slate-200 bg-slate-50/50 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={Boolean(nonAkta.sertakan_beda_nama)}
                            onChange={(e) =>
                              setNonAkta({
                                ...nonAkta,
                                sertakan_beda_nama: e.target.checked,
                              })
                            }
                            className="rounded border-slate-300"
                          />
                          <span className="font-medium text-slate-800">
                            Sertakan Surat Pernyataan Beda Nama (Pengesahan Camat &amp; Kades)
                          </span>
                        </label>
                      </div>

                      {nonAkta.sertakan_pelepasan_jalan && (
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 p-3 rounded-lg border border-slate-200 bg-slate-50/40">
                          <div>
                            <label className="block text-slate-600 mb-1">
                              Luas Tanah Dilepas untuk Jalan (m²)
                            </label>
                            <input
                              type="number"
                              min="0"
                              value={nonAkta.pelepasan_jalan_luas_m2 ?? 0}
                              onChange={(e) =>
                                setNonAkta({
                                  ...nonAkta,
                                  pelepasan_jalan_luas_m2: Number(e.target.value) || 0,
                                })
                              }
                              className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                            />
                          </div>
                          <div>
                            <label className="block text-slate-600 mb-1">
                              Keterangan Jalan (Opsional)
                            </label>
                            <input
                              type="text"
                              value={nonAkta.pelepasan_jalan_keterangan ?? ''}
                              onChange={(e) =>
                                setNonAkta({
                                  ...nonAkta,
                                  pelepasan_jalan_keterangan: e.target.value,
                                })
                              }
                              placeholder="Misal: Akses jalan umum RT 001/002"
                              className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                            />
                          </div>
                        </div>
                      )}

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div>
                          <label className="block text-slate-600 mb-1">
                            Koordinat GPS Lokasi Bidang (Opsional)
                          </label>
                          <input
                            type="text"
                            value={nonAkta.koordinat_gps ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, koordinat_gps: e.target.value })
                            }
                            placeholder="Contoh: -7.5194, 110.7798"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">
                            Link Google Maps / Poligon (Opsional)
                          </label>
                          <input
                            type="text"
                            value={nonAkta.link_maps ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, link_maps: e.target.value })
                            }
                            placeholder="https://maps.google.com/..."
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  {/* 4. Detail Khusus Ralat Data */}
                  {layanan === 'Ralat Data' && (
                    <div className="sm:col-span-2 border border-slate-200 rounded-xl p-4 bg-white space-y-3">
                      <h3 className="text-xs font-semibold text-slate-900 border-b border-slate-200 pb-2">
                        Data Ralat Sertipikat, Beda Nama &amp; Kronologi
                      </h3>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div>
                          <label htmlFor="rd-keperluan" className="block text-slate-600 mb-1">
                            Keperluan Permohonan
                          </label>
                          <input
                            id="rd-keperluan"
                            type="text"
                            value={nonAkta.keperluan_ralat ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, keperluan_ralat: e.target.value })
                            }
                            placeholder="Ralat Data Sertipikat"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                        <div>
                          <label htmlFor="rd-kk" className="block text-slate-600 mb-1">
                            Nomor Kartu Keluarga (KK)
                          </label>
                          <input
                            id="rd-kk"
                            type="text"
                            value={nonAkta.nomor_kk ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, nomor_kk: e.target.value })
                            }
                            placeholder="3309..."
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                          />
                        </div>
                        <div>
                          <label htmlFor="rd-nama-cert" className="block text-slate-600 mb-1">
                            Nama Tertulis di Sertipikat
                          </label>
                          <input
                            id="rd-nama-cert"
                            type="text"
                            value={nonAkta.nama_di_sertipikat ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, nama_di_sertipikat: e.target.value })
                            }
                            placeholder="Contoh: SARNO / WAGINEM"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-medium"
                          />
                        </div>
                        <div>
                          <label htmlFor="rd-nama-ktp" className="block text-slate-600 mb-1">
                            Nama Benar Sesuai KTP &amp; KK
                          </label>
                          <input
                            id="rd-nama-ktp"
                            type="text"
                            value={nonAkta.nama_di_ktp_kk ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, nama_di_ktp_kk: e.target.value })
                            }
                            placeholder="Contoh: SUWARNO"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-medium"
                          />
                        </div>
                        <div>
                          <label htmlFor="rd-dok-lain" className="block text-slate-600 mb-1">
                            Jenis Dokumen Lain (Opsional)
                          </label>
                          <input
                            id="rd-dok-lain"
                            type="text"
                            value={nonAkta.jenis_dokumen_lain ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, jenis_dokumen_lain: e.target.value })
                            }
                            placeholder="Akta Kelahiran / Buku Nikah / Ijazah"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                        <div>
                          <label htmlFor="rd-nama-dok-lain" className="block text-slate-600 mb-1">
                            Nama pada Dokumen Lain (Opsional)
                          </label>
                          <input
                            id="rd-nama-dok-lain"
                            type="text"
                            value={nonAkta.nama_di_dokumen_lain ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, nama_di_dokumen_lain: e.target.value })
                            }
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                        <div className="sm:col-span-2">
                          <label htmlFor="rd-kronologi" className="block text-slate-600 mb-1">
                            Uraian Kronologi Tambahan (Opsional)
                          </label>
                          <textarea
                            id="rd-kronologi"
                            rows={2}
                            value={nonAkta.uraian_kronologi ?? ''}
                            onChange={(e) =>
                              setNonAkta({ ...nonAkta, uraian_kronologi: e.target.value })
                            }
                            placeholder="Kosongkan untuk menggunakan uraian kronologi standar..."
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  {/* 5. Pengesahan Ganda Camat & Kepala Desa Asal Sertipikat */}
                  {(layanan === 'Turun Waris' ||
                    layanan === 'Ralat Data' ||
                    nonAkta.sertakan_beda_nama) && (
                    <div className="sm:col-span-2 border border-amber-200 rounded-xl p-4 bg-amber-50/40 space-y-3">
                      <div className="border-b border-amber-200 pb-2">
                        <h3 className="text-xs font-semibold text-slate-900">
                          Pengesahan Wajib Camat &amp; Kepala Desa Asal Sertipikat (SKW / Beda Nama / Kronologi)
                        </h3>
                        <p className="text-[11px] text-slate-600 mt-0.5">
                          Khusus Surat Keterangan Warisan, Surat Pernyataan Beda Nama, dan Kronologi Ralat Data wajib disahkan oleh kedua pejabat berikut: Camat dan Kepala Desa letak/asal sertipikat.
                        </p>
                      </div>

                      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                        <div className="p-3 rounded-lg border border-slate-200 bg-white space-y-2.5">
                          <p className="font-semibold text-slate-800">1. Pengesahan Camat</p>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div>
                              <label className="block text-slate-600 mb-1">Jabatan Camat</label>
                              <input
                                type="text"
                                value={nonAkta.camat_pengesahan_jabatan ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    camat_pengesahan_jabatan: e.target.value,
                                  })
                                }
                                placeholder={`Camat ${tanah.kecamatan || '-'}`}
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                              />
                            </div>
                            <div>
                              <label className="block text-slate-600 mb-1">Nama Camat</label>
                              <input
                                type="text"
                                value={nonAkta.camat_pengesahan_nama ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    camat_pengesahan_nama: e.target.value,
                                  })
                                }
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-medium"
                              />
                            </div>
                            <div>
                              <label className="block text-slate-600 mb-1">NIP Camat</label>
                              <input
                                type="text"
                                value={nonAkta.camat_pengesahan_nip ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    camat_pengesahan_nip: e.target.value,
                                  })
                                }
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                              />
                            </div>
                            <div>
                              <label className="block text-slate-600 mb-1">
                                No. Register Kec (Opsional)
                              </label>
                              <input
                                type="text"
                                value={nonAkta.camat_pengesahan_nomor ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    camat_pengesahan_nomor: e.target.value,
                                  })
                                }
                                placeholder="593/..."
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                              />
                            </div>
                          </div>
                        </div>

                        <div className="p-3 rounded-lg border border-slate-200 bg-white space-y-2.5">
                          <p className="font-semibold text-slate-800">
                            2. Pengesahan Kepala Desa Asal Sertipikat ({tanah.desa || '-'})
                          </p>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div className="sm:col-span-2">
                              <label className="block text-slate-600 mb-1">Jabatan Kepala Desa</label>
                              <input
                                type="text"
                                value={nonAkta.kades_pengesahan_jabatan ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    kades_pengesahan_jabatan: e.target.value,
                                  })
                                }
                                placeholder={`Kepala Desa ${tanah.desa || '-'}`}
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md"
                              />
                            </div>
                            <div>
                              <label className="block text-slate-600 mb-1">Nama Kepala Desa</label>
                              <input
                                type="text"
                                value={nonAkta.kades_pengesahan_nama ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    kades_pengesahan_nama: e.target.value,
                                  })
                                }
                                placeholder="Nama lengkap Kepala Desa"
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-medium"
                              />
                            </div>
                            <div>
                              <label className="block text-slate-600 mb-1">
                                No. Register Desa (Opsional)
                              </label>
                              <input
                                type="text"
                                value={nonAkta.kades_pengesahan_nomor ?? ''}
                                onChange={(e) =>
                                  setNonAkta({
                                    ...nonAkta,
                                    kades_pengesahan_nomor: e.target.value,
                                  })
                                }
                                placeholder="474/..."
                                className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md font-mono"
                              />
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </>
              )}

              {/* Form Lengkap Data Kuasa BPN */}
              <div className="sm:col-span-2 border border-slate-200 rounded-xl p-4 bg-slate-50/50 space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 border-b border-slate-200 pb-2">
                  <div>
                    <h3 className="text-xs font-semibold text-slate-900">
                      Data Penerima Kuasa Substitusi BPN
                    </h3>
                    <p className="text-[11px] text-slate-500">
                      Pilih staf kuasa dari daftar kantor atau isi/ubah detail lengkap untuk Surat Kuasa & Permohonan BPN.
                    </p>
                  </div>
                  <select
                    aria-label="Pilih preset Kuasa BPN"
                    value={
                      settings.daftar_kuasa_bpn.includes(akta.nama_kuasa_bpn)
                        ? akta.nama_kuasa_bpn
                        : '__manual__'
                    }
                    onChange={(e) => {
                      const val = e.target.value;
                      if (val === '__manual__') return;
                      const preset = findPresetPerson(val, settings.profil_kuasa);
                      setAkta((prev) => ({
                        ...prev,
                        nama_kuasa_bpn: val,
                        kuasa_bpn_detail: preset,
                      }));
                    }}
                    className="px-2.5 py-1.5 border border-slate-300 rounded-lg bg-white text-xs font-medium"
                  >
                    <option value="__manual__">-- Isi Manual / Custom --</option>
                    {settings.daftar_kuasa_bpn.map((k) => (
                      <option key={k} value={k}>
                        Preset: {k}
                      </option>
                    ))}
                  </select>
                </div>

                {(() => {
                  const kd = akta.kuasa_bpn_detail ?? { nama: akta.nama_kuasa_bpn };
                  const updateKuasaDetail = (patch: Partial<PersonBasicInfo>) => {
                    const nextDetail: PersonBasicInfo = {
                      ...kd,
                      ...patch,
                      nama: patch.nama !== undefined ? patch.nama : kd.nama || akta.nama_kuasa_bpn,
                    };
                    setAkta((prev) => ({
                      ...prev,
                      nama_kuasa_bpn: nextDetail.nama,
                      kuasa_bpn_detail: nextDetail,
                    }));
                  };
                  return (
                    <>
                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                        <div>
                          <label htmlFor="akta-kuasa" className="block text-slate-600 mb-1">
                            Nama Kuasa BPN
                          </label>
                          <input
                            id="akta-kuasa"
                            type="text"
                            value={akta.nama_kuasa_bpn}
                            onChange={(e) => updateKuasaDetail({ nama: e.target.value })}
                            placeholder="SUWARDJO"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-medium"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Nomor KTP (NIK)</label>
                          <input
                            type="text"
                            value={kd.nik ?? ''}
                            onChange={(e) => updateKuasaDetail({ nik: e.target.value })}
                            placeholder="3309111011570003"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Pekerjaan</label>
                          <input
                            type="text"
                            value={kd.pekerjaan ?? ''}
                            onChange={(e) => updateKuasaDetail({ pekerjaan: e.target.value })}
                            placeholder="PENSIUNAN"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Tempat Lahir</label>
                          <input
                            type="text"
                            value={kd.tempat_lahir ?? ''}
                            onChange={(e) => updateKuasaDetail({ tempat_lahir: e.target.value })}
                            placeholder="Boyolali"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Tanggal Lahir</label>
                          <DateInputId
                            allowEmpty
                            value={kd.tanggal_lahir ?? ''}
                            onChange={(val) => updateKuasaDetail({ tanggal_lahir: val })}
                            placeholder="DD-MM-YYYY"
                            ariaLabel="Tanggal lahir kuasa BPN"
                            className="py-1.5 px-2.5 rounded-md"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Dukuh / Jalan</label>
                          <input
                            type="text"
                            value={kd.alamat_jalan ?? ''}
                            onChange={(e) => updateKuasaDetail({ alamat_jalan: e.target.value })}
                            placeholder="Dukuhrejo"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          />
                        </div>
                        <div className="grid grid-cols-2 gap-1.5">
                          <div>
                            <label className="block text-slate-600 mb-1">RT</label>
                            <input
                              type="text"
                              value={kd.rt ?? ''}
                              onChange={(e) => updateKuasaDetail({ rt: e.target.value })}
                              placeholder="01"
                              className="w-full px-2 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                            />
                          </div>
                          <div>
                            <label className="block text-slate-600 mb-1">RW</label>
                            <input
                              type="text"
                              value={kd.rw ?? ''}
                              onChange={(e) => updateKuasaDetail({ rw: e.target.value })}
                              placeholder="02"
                              className="w-full px-2 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                            />
                          </div>
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Desa / Kelurahan</label>
                          <input
                            type="text"
                            value={kd.desa_kelurahan ?? ''}
                            onChange={(e) => updateKuasaDetail({ desa_kelurahan: e.target.value })}
                            placeholder="Donohudan"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">Kecamatan</label>
                          <input
                            type="text"
                            value={kd.kecamatan ?? ''}
                            onChange={(e) => updateKuasaDetail({ kecamatan: e.target.value })}
                            placeholder="Ngemplak"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          />
                        </div>
                        <div className="col-span-2 sm:col-span-3">
                          <label className="block text-slate-600 mb-1">Kabupaten / Kota</label>
                          <input
                            type="text"
                            value={kd.kabupaten ?? ''}
                            onChange={(e) => updateKuasaDetail({ kabupaten: e.target.value })}
                            placeholder="Boyolali"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                          />
                        </div>
                      </div>
                      <div className="p-2.5 rounded-lg bg-white border border-slate-200 text-[11px] text-slate-600">
                        <span className="font-semibold text-slate-800">Pratinjau Surat Kuasa:</span>{' '}
                        Nama: <strong>{akta.nama_kuasa_bpn || '-'}</strong> · Tempat/Tgl.lahir:{' '}
                        {[kd.tempat_lahir, formatTanggalDDMMYYYY(kd.tanggal_lahir)].filter((x) => x && x !== '-').join(', ') || '-'} · Pekerjaan:{' '}
                        {kd.pekerjaan || '-'} · KTP: {kd.nik || '-'} · Alamat:{' '}
                        {formatPersonShortAddress(kd)}
                      </div>
                    </>
                  );
                })()}
              </div>

              {/* Form Lengkap Data Saksi 1 & Saksi 2 */}
              {[
                {
                  num: 1 as const,
                  inputId: 'akta-saksi1',
                  label: 'Saksi 1 Akta',
                  nameVal: akta.saksi_1,
                  detailVal: akta.saksi_1_detail,
                  setSaksi: (nextName: string, nextDetail: PersonBasicInfo) =>
                    setAkta((prev) => ({
                      ...prev,
                      saksi_1: nextName,
                      saksi_1_detail: nextDetail,
                    })),
                },
                {
                  num: 2 as const,
                  inputId: 'akta-saksi2',
                  label: 'Saksi 2 Akta',
                  nameVal: akta.saksi_2,
                  detailVal: akta.saksi_2_detail,
                  setSaksi: (nextName: string, nextDetail: PersonBasicInfo) =>
                    setAkta((prev) => ({
                      ...prev,
                      saksi_2: nextName,
                      saksi_2_detail: nextDetail,
                    })),
                },
              ].map((sCfg) => {
                const sd = sCfg.detailVal ?? { nama: sCfg.nameVal };
                const resolved = resolvePersonDetail(
                  sCfg.nameVal,
                  settings.profil_saksi,
                  sd
                );
                const updateSaksiField = (patch: Partial<PersonBasicInfo>) => {
                  const nextDetail: PersonBasicInfo = {
                    ...sd,
                    ...patch,
                    nama: patch.nama !== undefined ? patch.nama : sd.nama || sCfg.nameVal,
                  };
                  sCfg.setSaksi(nextDetail.nama, nextDetail);
                };
                return (
                  <div
                    key={sCfg.num}
                    className="sm:col-span-2 border border-slate-200 rounded-xl p-4 bg-slate-50/50 space-y-3"
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 border-b border-slate-200 pb-2">
                      <h3 className="text-xs font-semibold text-slate-900">
                        Data {sCfg.label}
                      </h3>
                      <select
                        aria-label={`Pilih preset ${sCfg.label}`}
                        value={
                          settings.daftar_saksi.includes(sCfg.nameVal)
                            ? sCfg.nameVal
                            : '__manual__'
                        }
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val === '__manual__') return;
                          sCfg.setSaksi(val, findPresetPerson(val, settings.profil_saksi));
                        }}
                        className="px-2.5 py-1.5 border border-slate-300 rounded-lg bg-white text-xs font-medium"
                      >
                        <option value="__manual__">-- Isi Manual / Custom --</option>
                        {settings.daftar_saksi.map((s) => (
                          <option key={s} value={s}>
                            Preset: {s}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                      <div>
                        <label htmlFor={sCfg.inputId} className="block text-slate-600 mb-1">
                          Nama Lengkap {sCfg.label}
                        </label>
                        <input
                          id={sCfg.inputId}
                          type="text"
                          value={sCfg.nameVal}
                          onChange={(e) => updateSaksiField({ nama: e.target.value })}
                          placeholder="SALIMIN"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-medium"
                        />
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Nomor KTP (NIK)</label>
                        <input
                          type="text"
                          value={sd.nik ?? ''}
                          onChange={(e) => updateSaksiField({ nik: e.target.value })}
                          placeholder="3309110307730005"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                        />
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Pekerjaan</label>
                        <input
                          type="text"
                          value={sd.pekerjaan ?? ''}
                          onChange={(e) => updateSaksiField({ pekerjaan: e.target.value })}
                          placeholder="Perangkat Desa"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Tempat Lahir</label>
                        <input
                          type="text"
                          value={sd.tempat_lahir ?? ''}
                          onChange={(e) => updateSaksiField({ tempat_lahir: e.target.value })}
                          placeholder="Boyolali"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Tanggal Lahir</label>
                        <DateInputId
                          allowEmpty
                          value={sd.tanggal_lahir ?? ''}
                          onChange={(val) => updateSaksiField({ tanggal_lahir: val })}
                          placeholder="DD-MM-YYYY"
                          ariaLabel={`Tanggal lahir ${sCfg.label}`}
                          className="py-1.5 px-2.5 rounded-md"
                        />
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Dukuh / Jalan</label>
                        <input
                          type="text"
                          value={sd.alamat_jalan ?? ''}
                          onChange={(e) => updateSaksiField({ alamat_jalan: e.target.value })}
                          placeholder="Gunungan"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                      <div className="grid grid-cols-2 gap-1.5">
                        <div>
                          <label className="block text-slate-600 mb-1">RT</label>
                          <input
                            type="text"
                            value={sd.rt ?? ''}
                            onChange={(e) => updateSaksiField({ rt: e.target.value })}
                            placeholder="005"
                            className="w-full px-2 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                          />
                        </div>
                        <div>
                          <label className="block text-slate-600 mb-1">RW</label>
                          <input
                            type="text"
                            value={sd.rw ?? ''}
                            onChange={(e) => updateSaksiField({ rw: e.target.value })}
                            placeholder="008"
                            className="w-full px-2 py-1.5 border border-slate-200 rounded-md bg-white font-mono"
                          />
                        </div>
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Desa / Kelurahan</label>
                        <input
                          type="text"
                          value={sd.desa_kelurahan ?? ''}
                          onChange={(e) => updateSaksiField({ desa_kelurahan: e.target.value })}
                          placeholder="Giriroto"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                      <div>
                        <label className="block text-slate-600 mb-1">Kecamatan</label>
                        <input
                          type="text"
                          value={sd.kecamatan ?? ''}
                          onChange={(e) => updateSaksiField({ kecamatan: e.target.value })}
                          placeholder="Ngemplak"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                      <div className="col-span-2 sm:col-span-3">
                        <label className="block text-slate-600 mb-1">Kabupaten / Kota</label>
                        <input
                          type="text"
                          value={sd.kabupaten ?? ''}
                          onChange={(e) => updateSaksiField({ kabupaten: e.target.value })}
                          placeholder="Kabupaten Boyolali"
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md bg-white"
                        />
                      </div>
                    </div>
                    <div className="p-2.5 rounded-lg bg-white border border-slate-200 text-[11px] text-slate-600 leading-relaxed">
                      <span className="font-semibold text-slate-800">Pratinjau Kalimat Akta:</span>{' '}
                      {sCfg.num}. <strong>{resolved.nama}</strong>, Warga Negara Indonesia, lahir di{' '}
                      {resolved.tempat_lahir}, tanggal {formatTanggalDDMMYYYY(resolved.tanggal_lahir)}, Pekerjaan{' '}
                      {resolved.pekerjaan}, bertempat tinggal di{' '}
                      {formatPersonFullAddress(sd) !== '-' ? formatPersonFullAddress(sd) : resolved.alamat}. Pemegang Kartu Tanda Penduduk Nomor : {resolved.nik}.
                    </div>
                  </div>
                );
              })}

              <div className="sm:col-span-2 grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label htmlFor="akta-pengantar" className="block text-slate-600 font-medium mb-1">
                    Nomor Surat Pengantar BPN (Lampiran F4)
                  </label>
                  <input
                    id="akta-pengantar"
                    type="text"
                    value={akta.nomor_pengantar ?? ''}
                    onChange={(e) =>
                      setAkta({ ...akta, nomor_pengantar: e.target.value })
                    }
                    placeholder="Misal: 01/XI/2026"
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label htmlFor="akta-catatan" className="block text-slate-600 font-medium mb-1">
                    Catatan Loket / Kelengkapan
                  </label>
                  <input
                    id="akta-catatan"
                    type="text"
                    value={akta.catatan_loket}
                    onChange={(e) => setAkta({ ...akta, catatan_loket: e.target.value })}
                    placeholder="Misal: Sertipikat asli sudah diverifikasi."
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg"
                  />
                </div>
              </div>
            </div>
          </div>

          <div className="flex items-center justify-between pt-2">
            <button
              type="button"
              onClick={() => setStep(2)}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-sm font-medium text-slate-700"
            >
              <ArrowLeft className="w-4 h-4" />
              <span>Kembali ke Langkah 2</span>
            </button>
            <button
              type="button"
              onClick={() => setStep(4)}
              className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium"
            >
              <span>Lanjut ke Generate Bundel .DOCX</span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* STEP 4 */}
      {step === 4 && (
        <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 border-b border-slate-200 pb-4">
            <div>
              <h2 className="font-display text-lg font-semibold text-slate-900">
                {isAkta
                  ? 'Siap Generate Dokumen Akta (A4) & Bundel Lampiran Lengkap (F4)'
                  : `Siap Review & Unduh Bundel Non-Akta (${layanan} · Kertas F4)`}
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                {isAkta
                  ? 'Blangko Akta utama dicetak dalam format A4 (210 × 297 mm) dan seluruh lampiran bundel BPN (Hal. 1–9) digabung utuh dalam satu file F4 (215 × 330 mm). Data otomatis disimpan ke Buku Register Akta.'
                  : 'Berkas Non-Akta hanya menerbitkan 1 file bundel Word ukuran F4 (215 × 330 mm) tanpa Akta A4 dan tidak dicatat ke dalam Laporan Bulanan PPATS.'}
              </p>
            </div>
            <div className="font-mono text-xs text-slate-700 bg-slate-100 px-3 py-1.5 rounded-lg">
              {isAkta
                ? `${akta.nomor_akta || 'No. Akta Belum Diisi'} · ${layanan}`
                : `Non-Akta (F4 Only) · ${layanan}`}
            </div>
          </div>

          {/* Document Cards: 2 files for Akta PPATS, 1 file (F4) for Non-Akta */}
          <div className={`grid grid-cols-1 ${isAkta ? 'md:grid-cols-2' : ''} gap-3`}>
            {TEMPLATE_DOCS_META.filter((m) => isAkta || m.key === '02_lampiran_bundel').map((meta) => {
              const hasCustom = Boolean(customTemplates[meta.key]);
              const displayFileName = isAkta
                ? meta.fileName
                : getNonAktaBundleFileName(layanan);
              const displayDesc = isAkta
                ? meta.description
                : `Bundel lengkap berkas Non-Akta ${layanan} dalam 1 file F4 dengan format pengesahan sesuai jenis dokumen.`;
              const paperInfo =
                meta.paperSize === 'A4'
                  ? 'Kertas A4 (210 × 297 mm)'
                  : 'Kertas F4 (215 × 330 mm)';
              return (
                <div
                  key={meta.key}
                  className="p-4 rounded-xl border border-slate-200 flex flex-col justify-between gap-3"
                >
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <p className="font-mono text-xs font-semibold text-slate-900">
                        {displayFileName}{' '}
                        <span className="font-sans font-medium text-sky-800">
                          · {paperInfo}
                        </span>
                      </p>
                      <p className="text-xs text-slate-500 mt-1">{displayDesc}</p>
                    </div>
                    <span
                      className={`text-xs font-medium shrink-0 ${
                        hasCustom ? 'text-emerald-700' : 'text-slate-500'
                      }`}
                    >
                      {hasCustom ? 'Template Kustom' : 'Template Bawaan'}
                    </span>
                  </div>
                  <div className="flex flex-wrap items-center justify-end gap-2">
                    <button
                      type="button"
                      onClick={() => {
                        setReviewDocKey(
                          isAkta && meta.key === '01_akta' ? '01_akta' : '02_lampiran_bundel'
                        );
                        setShowReview(true);
                      }}
                      className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-xs font-medium transition-colors whitespace-nowrap ${
                        showReview && (reviewDocKey === meta.key || !isAkta)
                          ? 'border-[#0f172a] bg-[#0f172a] text-white'
                          : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <Eye className="w-3.5 h-3.5" />
                      <span>Review {meta.paperSize}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDownloadSingle(meta.key)}
                      disabled={generatingBundle}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-medium text-slate-800 disabled:opacity-60 whitespace-nowrap"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Unduh {meta.paperSize} (.DOCX)</span>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Toggle & Live Review Dokumen Sebelum Generate */}
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <p className="text-xs text-slate-600">
                {isAkta
                  ? 'Periksa kembali isi naskah Akta (A4) dan seluruh halaman Lampiran Bundel (F4) sebelum mengunduh file Word.'
                  : `Periksa kembali seluruh halaman bundel Non-Akta (${layanan}) ukuran F4 beserta pengesahannya sebelum mengunduh file Word.`}
              </p>
              <button
                type="button"
                onClick={() => setShowReview((v) => !v)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 whitespace-nowrap shrink-0"
              >
                <Eye className="w-3.5 h-3.5" />
                <span>{showReview ? 'Sembunyikan Review' : 'Tampilkan Review Dokumen'}</span>
              </button>
            </div>

            {showReview && (
              <DocumentReviewPanel
                record={reviewRecord}
                settings={settings}
                customTemplates={customTemplates}
                activeDocKey={isAkta ? reviewDocKey : '02_lampiran_bundel'}
                onSelectDocKey={setReviewDocKey}
                onEditStep={(targetStep) => setStep(targetStep)}
              />
            )}
          </div>

          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 pt-4 border-t border-slate-200">
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => setStep(3)}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-sm font-medium text-slate-700"
              >
                <ArrowLeft className="w-4 h-4" />
                <span>Kembali</span>
              </button>
              <button
                type="button"
                onClick={() => onPrintReceipt(buildCurrentRecord())}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-sm font-medium text-slate-700"
              >
                <Printer className="w-4 h-4" />
                <span>Cetak Tanda Terima</span>
              </button>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <button
                type="button"
                onClick={handleSaveOnly}
                className="inline-flex items-center gap-2 px-4 py-2.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-sm font-medium text-slate-800"
              >
                <Save className="w-4 h-4" />
                <span>Simpan ke Register Saja</span>
              </button>

              <button
                type="button"
                onClick={handleGenerateZipAndSave}
                disabled={generatingBundle}
                className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-sm font-medium transition-colors disabled:opacity-60"
              >
                {generatingBundle ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>Menyusun File .DOCX...</span>
                  </>
                ) : (
                  <>
                    <FileArchive className="w-4 h-4 text-amber-400" />
                    <span>Generate Bundel .ZIP &amp; Simpan Register</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
