export type ServiceType =
  | 'AJB'
  | 'APHB'
  | 'Akta Hibah'
  | 'Turun Waris'
  | 'Pemecahan'
  | 'Ukur Ulang'
  | 'Ralat Data'
  | 'Pemecahan/Ukur Ulang';

export type PengesahanMode = 'ppats' | 'kades' | 'camat' | 'manual';

export interface PemecahanBidangItem {
  id: string;
  label: string;
  luas_m2: number;
  atas_nama?: string;
  keterangan?: string;
}

export interface NonAktaDetail {
  pengesahan_mode?: PengesahanMode;
  pengesahan_jabatan?: string;
  pengesahan_nama?: string;
  pengesahan_nip?: string;

  camat_pengesahan_jabatan?: string;
  camat_pengesahan_nama?: string;
  camat_pengesahan_nip?: string;
  camat_pengesahan_nomor?: string;
  kades_pengesahan_jabatan?: string;
  kades_pengesahan_nama?: string;
  kades_pengesahan_nomor?: string;

  sebutan_pewaris?: 'Almarhum' | 'Almarhumah';
  nama_pewaris?: string;
  tempat_meninggal?: string;
  tanggal_meninggal?: string;
  hubungan_pasangan_pewaris?: 'Suaminya' | 'Istrinya' | 'Pasangannya';
  nama_pasangan_pewaris?: string;
  tgl_lahir_pasangan_pewaris?: string;
  jumlah_anak_waris?: number;
  catatan_skw?: string;

  sub_permohonan_ukur?:
    | 'Pengukuran'
    | 'Pemecahan'
    | 'Pengukuran & Validasi'
    | 'Pemecahan & Validasi';
  bidang_pecahan?: PemecahanBidangItem[];
  sertakan_pelepasan_jalan?: boolean;
  pelepasan_jalan_luas_m2?: number;
  pelepasan_jalan_keterangan?: string;
  sertakan_beda_nama?: boolean;
  koordinat_gps?: string;
  link_maps?: string;

  nama_di_sertipikat?: string;
  nama_di_ktp_kk?: string;
  nomor_kk?: string;
  jenis_dokumen_lain?: string;
  nama_di_dokumen_lain?: string;
  tgl_lahir_baku?: string;
  keperluan_ralat?: string;
  uraian_kronologi?: string;
}

export type BerkasStatus =
  | '1. Pemberkasan'
  | '2. Validasi Pajak'
  | '3. TTD Akta'
  | '4. Proses BPN'
  | '5. Selesai';

export type PihakRole =
  | 'Pihak I'
  | 'Pihak II'
  | 'Pihak III'
  | 'Pihak IV'
  | 'Pihak V'
  | 'Pihak VI'
  | 'Persetujuan Pasangan'
  | 'Lampiran Lain'
  | `Pihak ${string}`;

export interface ParaPihak {
  id: string;
  peran: PihakRole;
  nik: string;
  nama: string;
  tempat_lahir: string;
  tanggal_lahir: string;
  jenis_kelamin: string;
  alamat_jalan: string;
  rt_rw: string;
  desa_kelurahan: string;
  kecamatan: string;
  kabupaten: string;
  pekerjaan: string;
  status_perkawinan: string;
}

export interface PersonBasicInfo {
  nama: string;
  tempat_lahir?: string;
  tanggal_lahir?: string;
  pekerjaan?: string;
  nik?: string;
  alamat?: string;
  alamat_jalan?: string;
  rt?: string;
  rw?: string;
  desa_kelurahan?: string;
  kecamatan?: string;
  kabupaten?: string;
}

export interface DataTanah {
  jenis_sertipikat?: 'digital' | 'analog';
  jenis_hak: string;
  nib: string;
  nomor_hak_lama: string;
  surat_ukur_tanggal?: string;
  surat_ukur_nomor?: string;
  luas_m2: number;
  desa: string;
  kecamatan: string;
  kabupaten: string;
  provinsi?: string;
  batas_utara?: string;
  batas_timur?: string;
  batas_selatan?: string;
  batas_barat?: string;
  penggunaan_tanah?: string;
  status_tanah?: 'Non-Pertanian' | 'Pertanian';
  tanggal_pernyataan_pertanian?: string;
  menguasai_sejak_tahun?: string;
  dasar_perolehan: string;
  pemegang_hak_sertipikat: string[];
}

export interface DataPBB {
  nop: string;
  tahun_pajak: string;
  njop_bumi_m2: number;
  njop_bangunan: number;
}

export interface AuditWarning {
  id: string;
  tingkat: 'tinggi' | 'sedang' | 'info';
  judul: string;
  deskripsi: string;
}

export interface ParameterAkta {
  nomor_akta: string;
  tanggal_akta: string;
  nomor_pengantar?: string;
  nama_ppats: string;
  jabatan_ppats: string;
  nip_ppats: string;
  harga_per_m2: number;
  nilai_transaksi_override: number | null;
  npoptkp: number;
  butuh_persetujuan_pasangan?: boolean;
  keterangan_tanpa_persetujuan?: string;
  pihak_penerima_aphb?: string;
  nama_kuasa_bpn: string;
  kuasa_bpn_detail?: PersonBasicInfo;
  saksi_1: string;
  saksi_1_detail?: PersonBasicInfo;
  saksi_2: string;
  saksi_2_detail?: PersonBasicInfo;
  catatan_loket: string;
}

export interface TaxCalculationResult {
  nilai_kesepakatan: number;
  njop_bumi_total: number;
  njop_total: number;
  npop: number;
  npoptkp: number;
  npopkp: number;
  bphtb: number;
  pph: number;
}

export interface TaxDraftPrefill {
  layanan: ServiceType;
  luas_m2: number;
  njop_bumi_m2: number;
  njop_bangunan: number;
  harga_per_m2: number;
  nilai_transaksi_override: number | null;
  npoptkp: number;
}

export interface BerkasRecord {
  id: string;
  no_urut: number;
  kode_berkas: string;
  tanggal_masuk: string;
  layanan: ServiceType;
  status: BerkasStatus;
  para_pihak: ParaPihak[];
  tanah: DataTanah;
  pbb: DataPBB;
  audit_warnings: AuditWarning[];
  akta: ParameterAkta;
  pajak: TaxCalculationResult;
  non_akta?: NonAktaDetail;
  updated_at: string;
}

export interface OfficeSettings {
  nama_kecamatan: string;
  nama_kabupaten: string;
  nama_provinsi?: string;
  nama_ppats: string;
  jabatan_ppats: string;
  nip_ppats: string;
  sk_ppats_nomor?: string;
  sk_ppats_tanggal?: string;
  sk_ppats_penerbit?: string;
  alamat_kantor: string;
  telepon_kantor: string;
  daftar_saksi: string[];
  daftar_kuasa_bpn: string[];
  profil_saksi?: PersonBasicInfo[];
  profil_kuasa?: PersonBasicInfo[];
  npoptkp_umum: number;
  npoptkp_waris: number;
}

export type PaperSizeType = 'A4' | 'F4';

export type TemplateDocKey =
  | '01_akta'
  | '02_lampiran_bundel'
  | '02_permohonan_bpn'
  | '03_kuasa_bpn'
  | '04_tidak_sengketa'
  | '05_sppt_pbb'
  | '06_rincian_pajak'
  | '07_tanda_terima'
  | '08_keabsahan_data_pribadi';

export interface TemplateDocMeta {
  key: TemplateDocKey;
  fileName: string;
  label: string;
  paperSize: PaperSizeType;
  description: string;
}

export interface CustomTemplateEntry {
  key: TemplateDocKey;
  fileName: string;
  updatedAt: string;
  dataBase64: string;
}

export interface UploadedDocFile {
  id: string;
  name: string;
  mimeType: string;
  size: number;
  dataBase64: string;
  previewUrl?: string;
}

export interface ExtractionApiResponse {
  para_pihak: Array<{
    peran_saran?: string;
    nik: string;
    nama: string;
    tempat_lahir: string;
    tanggal_lahir: string;
    jenis_kelamin: string;
    alamat_jalan: string;
    rt_rw: string;
    desa_kelurahan: string;
    kecamatan: string;
    kabupaten: string;
    pekerjaan: string;
    status_perkawinan: string;
  }>;
  tanah: {
    jenis_sertipikat?: 'digital' | 'analog';
    jenis_hak: string;
    nib: string;
    nomor_hak_lama: string;
    surat_ukur_tanggal?: string;
    surat_ukur_nomor?: string;
    luas_m2: number;
    desa: string;
    kecamatan: string;
    kabupaten: string;
    provinsi?: string;
    batas_utara?: string;
    batas_timur?: string;
    batas_selatan?: string;
    batas_barat?: string;
    dasar_perolehan: string;
    pemegang_hak_sertipikat: string[];
  };
  pbb: {
    nop: string;
    tahun_pajak: string;
    njop_bumi_m2: number;
    njop_bangunan: number;
  };
  audit_warnings: Array<{
    tingkat: 'tinggi' | 'sedang' | 'info';
    judul: string;
    deskripsi: string;
  }>;
}
