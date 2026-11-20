import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalMergeType,
  WidthType,
} from 'docx';
import { TemplateDocKey } from '../types/ppats';

export interface TemplateTagDocItem {
  tag: string;
  desc: string;
}

// Konversi ukuran kertas standar ke twips (1 mm = 56.692913 twips)
export const PAPER_SIZE_A4_TWIPS = {
  width: 11906, // 210 mm
  height: 16838, // 297 mm
};

export const PAPER_SIZE_F4_TWIPS = {
  width: 12190, // 215 mm
  height: 18709, // 330 mm (Folio / F4 Indonesia)
};

export const SUPPORTED_TEMPLATE_TAGS: TemplateTagDocItem[] = [
  { tag: '{{nomor_akta}}', desc: 'Nomor Akta PPATS (mis. 001/PPATS/XI/2026)' },
  { tag: '{{nomor_pengantar}}', desc: 'Nomor Surat Pengantar berkas ke ATR/BPN' },
  { tag: '{{kalimat_akta}}', desc: 'Kalimat pembuka tanggal akta lengkap (Hari, tanggal huruf, bulan, tahun huruf)' },
  { tag: '{{hari_akta}} / {{tanggal_angka}} / {{tanggal_akta_huruf}} / {{bulan_akta}} / {{tahun_angka}} / {{tahun_akta_huruf}}', desc: 'Komponen tanggal akta terpisah untuk naskah blangko Akta' },
  { tag: '{{tanggal_akta}}', desc: 'Tanggal akta format normal (mis. 19 November 2026)' },
  { tag: '{{layanan}} / {{layanan_upper}} / {{judul_akta}}', desc: 'Jenis layanan & judul resmi akta (mis. AKTA JUAL BELI)' },
  { tag: '{{kode_berkas}}', desc: 'Nomor register berkas loket (mis. REG-2026-001)' },
  { tag: '{{#pihak_pertama}} ... {{/pihak_pertama}}', desc: 'Loop Pihak I. Tersedia: {{sebutan}}, {{nama}}, {{nik}}, {{tempat_lahir}}, {{tanggal_lahir}}, {{pekerjaan}}, {{alamat_lengkap}}, {{rt}}, {{rw}}, {{desa_kelurahan}}, {{kecamatan}}, {{kabupaten}}, {{status_perkawinan}}' },
  { tag: '{{#pihak_kedua}} ... {{/pihak_kedua}}', desc: 'Loop Pihak II (Pembeli/Penerima). Field sama dengan Pihak I' },
  { tag: '{{#pihak_lanjutan}} ... {{/pihak_lanjutan}}', desc: 'Loop Pihak III, IV, V, dan selanjutnya. Tersedia {{roman_numeral}}, {{peran_label_upper}}, serta seluruh field identitas pihak' },
  { tag: '{{#persetujuan_pasangan}} ... {{/persetujuan_pasangan}}', desc: 'Loop persetujuan suami/istri dari Pihak I' },
  { tag: '{{pihak_pertama_ringkas}} / {{pihak_kedua_ringkas}}', desc: 'Gabungan nama seluruh Pihak I / Pihak II' },
  { tag: '{{p1_nama}} / {{p1_ttl}} / {{p1_pekerjaan}} / {{p1_nik}} / {{p1_alamat}}', desc: 'Data identitas Pihak I utama' },
  { tag: '{{p2_nama}} / {{p2_ttl}} / {{p2_pekerjaan}} / {{p2_nik}} / {{p2_alamat}}', desc: 'Data identitas Pihak II utama' },
  { tag: '{{penerima_nama}} / {{penerima_ttl}} / {{penerima_nik}} / {{penerima_alamat}} / {{penerima_ringkas}}', desc: 'Pihak Pembeli / Penerima Hibah / Penerima Hak APHB untuk Kuasa, Permohonan, BPHTB, Fisik & Keabsahan' },
  { tag: '{{pemilik_awal_nama}} / {{pemilik_awal_ttl}} / {{pemilik_awal_nik}} / {{pemilik_awal_alamat}}', desc: 'Pemilik awal sertipikat untuk Surat Pernyataan SPPT/PBB' },
  { tag: '{{ket_pihak1_label}} / {{ket_pihak1_names}} / {{ket_pihak2_label}} / {{ket_pihak2_names}}', desc: 'Label Penjual/Pembeli (AJB), Pemberi/Penerima Hak (APHB), atau Pemberi/Penerima Hibah (Hibah) di Surat Pengantar' },
  { tag: '{{#pihak_data_pribadi_ttd}} ... {{/pihak_data_pribadi_ttd}}', desc: 'Daftar penandatangan Pemrosesan Data Pribadi (seluruh pihak termasuk pembeli/penerima, tanpa pihak menyetujui)' },
  { tag: '{{pasangan_nama}}', desc: 'Nama pasangan yang menyetujui (jika ada)' },
  { tag: '{{jenis_hak}} / {{nib}} / {{nomor_hak_lama}} / {{surat_ukur_tanggal}} / {{surat_ukur_nomor}}', desc: 'Jenis hak, NIB/NIBel, Nomor Hak Analog, serta Tanggal & Nomor Surat Ukur' },
  { tag: '{{luas_m2}} / {{luas_terbilang}}', desc: 'Luas bidang tanah angka dan terbilang' },
  { tag: '{{batas_utara}} / {{batas_timur}} / {{batas_selatan}} / {{batas_barat}}', desc: 'Batas-batas bidang tanah pada blangko Akta' },
  { tag: '{{penggunaan_tanah}} / {{menguasai_sejak_tahun}}', desc: 'Penggunaan tanah (mis. Pekarangan) & tahun mulai penguasaan fisik' },
  { tag: '{{desa_tanah}} / {{kecamatan_tanah}} / {{kabupaten_tanah}} / {{provinsi_tanah}}', desc: 'Lokasi administratif bidang tanah' },
  { tag: '{{dasar_perolehan}}', desc: 'Dasar perolehan hak sebelumnya' },
  { tag: '{{nop}} / {{tahun_pajak}} / {{njop_bumi_m2_rp}} / {{njop_bangunan_rp}} / {{njop_total_rp}}', desc: 'Data SPPT PBB-P2 dan NJOP' },
  { tag: '{{nilai_transaksi_rp}} / {{nilai_transaksi_terbilang}}', desc: 'Nilai transaksi/kesepakatan angka & terbilang' },
  { tag: '{{npop_rp}} / {{npoptkp_rp}} / {{bphtb_rp}} / {{bphtb_terbilang}} / {{pph_rp}} / {{pph_terbilang}}', desc: 'Rincian perhitungan BPHTB (5%) & PPh Final (2,5%)' },
  { tag: '{{nama_ppats}} / {{jabatan_ppats}} / {{nip_ppats}}', desc: 'Identitas Pejabat PPATS' },
  { tag: '{{sk_ppats_nomor}} / {{sk_ppats_tanggal}} / {{sk_ppats_penerbit}}', desc: 'SK Pengangkatan PPATS dari Kakanwil BPN' },
  { tag: '{{nama_kecamatan}} / {{nama_kabupaten}} / {{nama_provinsi}} / {{alamat_kantor}} / {{telepon_kantor}}', desc: 'Profil wilayah kerja dan alamat kantor PPATS' },
  { tag: '{{kuasa_bpn_nama}} / {{kuasa_bpn_ttl}} / {{kuasa_bpn_pekerjaan}} / {{kuasa_bpn_nik}} / {{kuasa_bpn_alamat}}', desc: 'Identitas lengkap staf kuasa pengurusan di BPN' },
  { tag: '{{saksi_1_nama}} / {{saksi_1_ttl}} / {{saksi_1_pekerjaan}} / {{saksi_1_nik}} / {{saksi_1_alamat}}', desc: 'Identitas lengkap Saksi 1 Akta & Lampiran' },
  { tag: '{{saksi_2_nama}} / {{saksi_2_ttl}} / {{saksi_2_pekerjaan}} / {{saksi_2_nik}} / {{saksi_2_alamat}}', desc: 'Identitas lengkap Saksi 2 Akta & Lampiran' },
];

const AKTA_FONT = 'Bookman Old Style';
const LAMPIRAN_FONT = 'Bookman Old Style';
const BASIC_FONT_SIZE = 24; // 12 pt
const LINE_SPACING_1_5 = 360; // 1.5 lines (240 * 1.5)
const PARAGRAPH_AFTER_1_5 = 120; // 6 pt

// Pengaturan margin standar sesuai Page Setup:
// Top 2,54 cm (1440), Bottom 2,54 cm (1440), Left 4,5 cm (2551), Right 1 cm (567),
// Gutter 0 cm, Header from Top 0 cm, Footer from Bottom 2 cm (1134)
export const STANDARD_PAGE_MARGINS = {
  top: 1440,
  bottom: 1440,
  left: 2551,
  right: 567,
  gutter: 0,
  header: 0,
  footer: 1134,
};

function createDocxStyles(font = AKTA_FONT) {
  return {
    default: {
      document: {
        run: {
          font,
          size: BASIC_FONT_SIZE,
        },
        paragraph: {
          alignment: AlignmentType.JUSTIFIED,
          spacing: {
            line: LINE_SPACING_1_5,
            after: PARAGRAPH_AFTER_1_5,
          },
        },
      },
    },
  };
}

function p(
  text: string,
  opts?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    size?: number;
    spaceAfter?: number;
    spaceBefore?: number;
    line?: number;
    font?: string;
    indentLeft?: number;
    pageBreakBefore?: boolean;
  }
): Paragraph {
  return new Paragraph({
    alignment: opts?.align ?? AlignmentType.JUSTIFIED,
    pageBreakBefore: opts?.pageBreakBefore ?? false,
    indent: opts?.indentLeft ? { left: opts.indentLeft } : undefined,
    spacing: {
      before: opts?.spaceBefore ?? 0,
      after: opts?.spaceAfter ?? PARAGRAPH_AFTER_1_5,
      line: opts?.line ?? LINE_SPACING_1_5,
    },
    children: [
      new TextRun({
        text,
        bold: opts?.bold ?? false,
        italics: opts?.italic ?? false,
        underline: opts?.underline ? {} : undefined,
        size: opts?.size ?? BASIC_FONT_SIZE,
        font: opts?.font ?? AKTA_FONT,
      }),
    ],
  });
}

function pRuns(
  runs: Array<{ text: string; bold?: boolean; italic?: boolean; underline?: boolean; size?: number }>,
  opts?: {
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    spaceAfter?: number;
    line?: number;
    indentLeft?: number;
    font?: string;
  }
): Paragraph {
  return new Paragraph({
    alignment: opts?.align ?? AlignmentType.JUSTIFIED,
    indent: opts?.indentLeft ? { left: opts.indentLeft } : undefined,
    spacing: {
      after: opts?.spaceAfter ?? PARAGRAPH_AFTER_1_5,
      line: opts?.line ?? LINE_SPACING_1_5,
    },
    children: runs.map(
      (r) =>
        new TextRun({
          text: r.text,
          bold: r.bold ?? false,
          italics: r.italic ?? false,
          underline: r.underline ? {} : undefined,
          size: r.size ?? BASIC_FONT_SIZE,
          font: opts?.font ?? AKTA_FONT,
        })
    ),
  });
}

const NO_BORDER = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

const THIN_BORDER = {
  top: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
  bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
  left: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
  right: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
};

function makeBorderlessRow(leftText: string, rightText: string, leftPct = 42, font = LAMPIRAN_FONT, size = BASIC_FONT_SIZE): TableRow {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: leftPct, type: WidthType.PERCENTAGE },
        borders: NO_BORDER,
        children: [p(leftText, { align: AlignmentType.LEFT, spaceAfter: 40, line: LINE_SPACING_1_5, font, size })],
      }),
      new TableCell({
        width: { size: 100 - leftPct, type: WidthType.PERCENTAGE },
        borders: NO_BORDER,
        children: [p(rightText, { align: AlignmentType.LEFT, spaceAfter: 40, line: LINE_SPACING_1_5, font, size })],
      }),
    ],
  });
}

function makeSimpleKvTable(rows: Array<[string, string]>, font = LAMPIRAN_FONT, size = BASIC_FONT_SIZE): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(([k, v]) => makeBorderlessRow(k, `: ${v}`, 32, font, size)),
  });
}

function aktaKopSurat(): Paragraph[] {
  return [
    p('PEJABAT PEMBUAT AKTA TANAH SEMENTARA', {
      bold: true,
      align: AlignmentType.CENTER,
      size: 32, // 16 pt
      spaceAfter: 20,
      line: 240,
    }),
    p('(PPATS)', {
      bold: true,
      align: AlignmentType.CENTER,
      size: 32, // 16 pt
      spaceAfter: 20,
      line: 240,
    }),
    p('{{jabatan_ppats_upper}}', {
      bold: true,
      align: AlignmentType.CENTER,
      size: 32, // 16 pt
      spaceAfter: 20,
      line: 240,
    }),
    p('{{nama_ppats}}', {
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      size: 28, // 14 pt
      spaceAfter: 20,
      line: 240,
    }),
    p('DAERAH KERJA : {{nama_kecamatan_upper}}, {{nama_kabupaten_upper}}', {
      bold: true,
      align: AlignmentType.CENTER,
      size: 20,
      spaceAfter: 20,
      line: 240,
    }),
    p('Surat Keputusan {{sk_ppats_penerbit}}', {
      bold: true,
      align: AlignmentType.CENTER,
      size: 19,
      spaceAfter: 20,
      line: 240,
    }),
    p('Nomor : {{sk_ppats_nomor}} Tanggal {{sk_ppats_tanggal}}', {
      bold: true,
      align: AlignmentType.CENTER,
      size: 19,
      spaceAfter: 20,
      line: 240,
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200, line: 240 },
      border: {
        bottom: { style: BorderStyle.DOUBLE, size: 12, color: '000000' },
      },
      children: [
        new TextRun({
          text: '{{alamat_kantor}}, Telp. {{telepon_kantor}}',
          bold: true,
          size: 19,
          font: AKTA_FONT,
        }),
      ],
    }),
  ];
}

const FOOTER_TOP_BORDER = {
  top: { style: BorderStyle.SINGLE, size: 6, color: '64748B' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

function aktaFooter(): Footer {
  return new Footer({
    children: [
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 60, type: WidthType.PERCENTAGE },
                borders: FOOTER_TOP_BORDER,
                children: [
                  p('{{judul_akta}}', { align: AlignmentType.LEFT, size: BASIC_FONT_SIZE, spaceBefore: 60, spaceAfter: 0, line: 260 }),
                  p('{{nama_ppats}}', { align: AlignmentType.LEFT, bold: true, italic: true, size: BASIC_FONT_SIZE, spaceAfter: 0, line: 260 }),
                  p('Daerah Kerja : {{nama_kecamatan}}.', { align: AlignmentType.LEFT, size: BASIC_FONT_SIZE, spaceAfter: 0, line: 260 }),
                ],
              }),
              new TableCell({
                width: { size: 40, type: WidthType.PERCENTAGE },
                borders: FOOTER_TOP_BORDER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    spacing: { before: 60, after: 0, line: 260 },
                    children: [
                      new TextRun({ text: 'Halaman ', size: BASIC_FONT_SIZE, font: AKTA_FONT }),
                      new TextRun({ children: [PageNumber.CURRENT], size: BASIC_FONT_SIZE, font: AKTA_FONT }),
                      new TextRun({ text: ' dari ', size: BASIC_FONT_SIZE, font: AKTA_FONT }),
                      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: BASIC_FONT_SIZE, font: AKTA_FONT }),
                      new TextRun({ text: ' halaman', size: BASIC_FONT_SIZE, font: AKTA_FONT }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function pertanianPasal3Block(): Paragraph[] {
  return [
    p('{{#is_pertanian}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
    p('------------------------------------------ Pasal 3 ---------------------------------------', {
      align: AlignmentType.CENTER,
      spaceAfter: 60,
    }),
    p(
      'Pihak Kedua dengan ini menyatakan bahwa perjanjian {{perbuatan_pertanian}} ini kepemilikan tanahnya tidak melanggar ketentuan maksimum penguasaan tanah menurut peraturan perundang-undangan yang berlaku sebagaimana tercantum dalam pernyataan tanggal : {{tanggal_pernyataan_pertanian}}.-----------------------------------------',
      { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
    ),
    p('{{/is_pertanian}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
  ];
}

function buildAktaDocument(): Document {
  return new Document({
    styles: createDocxStyles(AKTA_FONT),
    sections: [
      {
        properties: {
          page: {
            size: PAPER_SIZE_A4_TWIPS,
            margin: STANDARD_PAGE_MARGINS,
          },
        },
        footers: {
          default: aktaFooter(),
        },
        children: [
          ...aktaKopSurat(),
          p('{{judul_akta}}', { bold: true, align: AlignmentType.CENTER, size: BASIC_FONT_SIZE, spaceAfter: 60 }),
          p('Nomor : {{nomor_akta}}', { bold: true, align: AlignmentType.CENTER, size: BASIC_FONT_SIZE, spaceAfter: 60 }),
          p('Lembar Pertama', { bold: true, align: AlignmentType.CENTER, size: BASIC_FONT_SIZE, spaceAfter: 180 }),

          pRuns(
            [
              {
                text: 'Pada hari ini, {{hari_akta}}, Tanggal {{tanggal_angka}} ({{tanggal_akta_huruf}}) {{bulan_akta}} Tahun {{tahun_angka}} ({{tahun_akta_huruf}}), hadir dihadapan Saya, ',
              },
              { text: '{{nama_ppats}}', bold: true },
              { text: ' yang berdasarkan ' },
              { text: 'SK. Kepala Badan Pertanahan Nasional {{sk_ppats_penerbit}}, Nomor: {{sk_ppats_nomor}} Tanggal {{sk_ppats_tanggal}}', bold: true },
              {
                text: ' diangkat sebagai Pejabat Pembuat Akta Tanah Sementara yang dimaksud dalam Pasal 7 Peraturan Pemerintah Nomor 24 Tahun 1997 tentang Pendaftaran Tanah, dengan daerah ',
              },
              { text: '{{nama_kecamatan}}, {{nama_kabupaten}}', bold: true },
              { text: ' dan berkantor di ' },
              { text: '{{nama_kecamatan}}, {{nama_kabupaten}}', bold: true },
              {
                text: ' dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini : -------------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 140, line: LINE_SPACING_1_5 }
          ),

          // ==================== PARA PIHAK (AJB / HIBAH: Pihak I & II; APHB: Terpisah I..XX+) ====================
          p('{{#is_not_aphb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#pihak_pertama}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          pRuns(
            [
              { text: '{{no_urut_pihak}}. ', bold: true },
              { text: '{{sebutan}} {{nama}}, ', bold: true },
              {
                text: 'Warga Negara Indonesia, lahir di {{tempat_lahir}} pada tanggal {{tanggal_lahir}}, Pekerjaan {{pekerjaan}}, bertempat tinggal di {{alamat_lengkap}}. Pemegang Kartu Penduduk Nomor : {{nik}}.-------------------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/pihak_pertama}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),

          p('Selaku {{pihak1_selaku}}, untuk selanjutnya dalam akta ini disebut : ---------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 360,
            spaceAfter: 60,
          }),
          p('--------------------------------------- PIHAK PERTAMA -------------------------------------', {
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 140,
          }),

          // PERSETUJUAN PASANGAN / TANPA PERSETUJUAN (untuk AJB/Hibah ditulis setelah Pihak I seperti contoh)
          p('{{#ada_persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          p(
            'Menurut keterangannya dalam melakukan perbuatan hukum ini telah mendapatkan persetujuan dari Pasangannya yang bernama {{sebutan}} {{nama}}, Warga Negara Indonesia, lahir di {{tempat_lahir}} pada tanggal {{tanggal_lahir}}, Pekerjaan {{pekerjaan}}, bertempat tinggal di {{alamat_lengkap}}. Pemegang Kartu Penduduk Nomor : {{nik}}.-----------',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          p('{{/ada_persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          p('{{#tanpa_persetujuan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p(
            '{{keterangan_tanpa_persetujuan}}-----------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/tanpa_persetujuan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // PIHAK KEDUA
          p('{{#pihak_kedua}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          pRuns(
            [
              { text: '{{no_urut_pihak}}. ', bold: true },
              { text: '{{sebutan}} {{nama}}, ', bold: true },
              {
                text: 'Warga Negara Indonesia, lahir di {{tempat_lahir}} pada tanggal {{tanggal_lahir}}, Pekerjaan {{pekerjaan}}, bertempat tinggal di {{alamat_lengkap}}. Pemegang Kartu Penduduk Nomor : {{nik}}.---------------------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/pihak_kedua}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),

          p('Selaku {{pihak2_selaku}}, untuk selanjutnya dalam akta ini disebut : --------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 360,
            spaceAfter: 60,
          }),
          p('--------------------------------------- PIHAK KEDUA ----------------------------------------', {
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 140,
          }),

          // Fallback jika ada Pihak III+ pada layanan non-APHB
          p('{{#pihak_lanjutan}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          pRuns(
            [
              { text: '{{no_urut_pihak}}. ', bold: true },
              { text: '{{sebutan}} {{nama}}, ', bold: true },
              {
                text: 'Warga Negara Indonesia, lahir di {{tempat_lahir}} pada tanggal {{tanggal_lahir}}, Pekerjaan {{pekerjaan}}, bertempat tinggal di {{alamat_lengkap}}. Pemegang Kartu Penduduk Nomor : {{nik}}.---------------------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('-------------------------------------- {{peran_label_upper}} ---------------------------------------', {
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 140,
          }),
          p('{{/pihak_lanjutan}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          p('{{/is_not_aphb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // KHUSUS APHB: Semua pemilik menjadi pihak terpisah (Pihak I, II, III, ..., XX+)
          // dan persetujuan / tanpa persetujuan ditulis SETELAH pihak terakhir
          p('{{#is_aphb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#pihak_terpisah}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          pRuns(
            [
              { text: '{{no_urut_pihak}}. ', bold: true },
              { text: '{{sebutan}} {{nama}}, ', bold: true },
              {
                text: 'Warga Negara Indonesia, lahir di {{tempat_lahir}} pada tanggal {{tanggal_lahir}}, Pekerjaan {{pekerjaan}}, bertempat tinggal di {{alamat_lengkap}}. Pemegang Kartu Penduduk Nomor : {{nik}}.-------------------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 80, line: LINE_SPACING_1_5 }
          ),
          p('Untuk selanjutnya dalam akta ini disebut : ------------------------------------------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 360,
            spaceAfter: 60,
          }),
          p('-------------------------------------- {{peran_label_upper}} ---------------------------------------', {
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 120,
          }),
          p('{{/pihak_terpisah}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),

          p('{{#ada_persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          p(
            'Menurut keterangannya dalam melakukan perbuatan hukum ini telah mendapatkan persetujuan dari Pasangannya yang bernama {{sebutan}} {{nama}}, Warga Negara Indonesia, lahir di {{tempat_lahir}} pada tanggal {{tanggal_lahir}}, Pekerjaan {{pekerjaan}}, bertempat tinggal di {{alamat_lengkap}}. Pemegang Kartu Penduduk Nomor : {{nik}}.-----------',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 20, line: 240 }),
          p('{{/ada_persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          p('{{#tanpa_persetujuan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p(
            '{{keterangan_tanpa_persetujuan}}-----------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/tanpa_persetujuan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{/is_aphb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          p(
            'Para penghadap dikenal oleh saya, Pejabat Pembuat Akta Tanah Sementara {{jabatan_ppats}}. {{kalimat_perbuatan_pihak}}-----------------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 120, line: LINE_SPACING_1_5 }
          ),
          p('{{#is_sertipikat_digital}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p(
            '{{jenis_hak}} Nibel : {{nib}}/Desa {{desa_tanah}}, Kecamatan {{kecamatan_tanah}}, {{kabupaten_tanah}} atas sebidang tanah seluas : {{luas_m2}} M2 ({{luas_terbilang}}) ---------------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_sertipikat_digital}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#is_sertipikat_analog}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p(
            '{{jenis_hak}} Nomor : {{nomor_hak_analog}} /Desa {{desa_tanah}}, Kecamatan {{kecamatan_tanah}}, {{kabupaten_tanah}} atas sebidang tanah sebagaimana diuraikan dalam Surat Ukur tanggal : {{surat_ukur_tanggal}} , Nomor: {{surat_ukur_nomor}}, seluas : {{luas_m2}} M2 ({{luas_terbilang}}) Dengan Nomor Identifikasi Bidang tanah (NIB) : {{nib}}---------------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_sertipikat_analog}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('Dengan batas-batas :-----------------------------------------------------------------', {
            align: AlignmentType.LEFT,
            spaceAfter: 60,
          }),
          p('  -   Utara     : {{batas_utara}} ----------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Timur     : {{batas_timur}} ----------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Selatan   : {{batas_selatan}} --------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Barat     : {{batas_barat}} ----------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 100 }),

          p('Berdasarkan alat-alat bukti berupa : -------------------------------------------', {
            align: AlignmentType.LEFT,
            spaceAfter: 60,
          }),
          p('{{#is_sertipikat_digital}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('1. Sertipikat {{jenis_hak}} Nibel : {{nib}} Desa {{desa_tanah}}------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 200,
            spaceAfter: 40,
          }),
          p('{{/is_sertipikat_digital}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#is_sertipikat_analog}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('1. Sertipikat {{jenis_hak}} Nomor : {{nomor_hak_analog}} NIB : {{nib}} Surat Ukur Tgl. {{surat_ukur_tanggal}} Nomor : {{surat_ukur_nomor}} Desa {{desa_tanah}}------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 200,
            spaceAfter: 40,
          }),
          p('{{/is_sertipikat_analog}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('2. Foto copy Kartu Tanda Penduduk dan Kartu Keluarga Para Penghadap.', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 200,
            spaceAfter: 40,
          }),
          p('3. Foto copy SPPT PBB NOP: {{nop}} Tahun {{tahun_pajak}}.-----------------------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 200,
            spaceAfter: 40,
          }),
          p('4. SSPD BPHTB, SSP PPh.--------------------------------------------------------------------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 200,
            spaceAfter: 100,
          }),

          p('Terletak di : -------------------------------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Propinsi  : {{provinsi_tanah}} --------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Kabupaten : {{kabupaten_tanah}} ------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Kecamatan : {{kecamatan_tanah}} ------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 40 }),
          p('  -   Desa      : {{desa_tanah}} -----------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 100 }),

          // ==================== 1. KLAUSUL & PASAL KHUSUS APHB ====================
          p('{{#is_aphb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p(
            'Selanjutnya dalam akta ini disebut “Hak Bersama”. Para pihak selanjutnya menerangkan bahwa mereka telah sepakat untuk mengakhiri pemilikan Hak Bersama tersebut, dan untuk dengan ini menyepakati pembagian Hak Bersama tersebut sebagai berikut : ------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{#is_sertipikat_digital}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          pRuns(
            [
              { text: '{{penerima_aphb_label}} atas nama {{penerima_aphb_nama}}', bold: true },
              {
                text: ', memperoleh dan menjadi pemegang tunggal dari {{jenis_hak}} atas sebidang tanah Nibel : {{nib}} / Desa {{desa_tanah}}----------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_sertipikat_digital}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#is_sertipikat_analog}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          pRuns(
            [
              { text: '{{penerima_aphb_label}} atas nama {{penerima_aphb_nama}}', bold: true },
              {
                text: ', memperoleh dan menjadi pemegang tunggal dari {{jenis_hak}} Nomor : {{nomor_hak_analog}} / NIB : {{nib}} / Desa {{desa_tanah}}----------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_sertipikat_analog}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('sebagaimana telah diuraikan di atas .---------------------------------------------', {
            align: AlignmentType.JUSTIFIED,
            spaceAfter: 100,
            line: LINE_SPACING_1_5,
          }),
          p('Pembagian Hak Bersama diatas juga meliputi : --------------------------------', {
            align: AlignmentType.LEFT,
            spaceAfter: 60,
          }),
          p(
            'Tanah beserta segala sesuatu yang tumbuh, berdiri dan tertanam diatas tanah tersebut tanpa terkecuali. ---------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('Selanjutnya para pihak menerangkan bahwa : ---------------------------------', {
            align: AlignmentType.LEFT,
            spaceAfter: 60,
          }),
          p(
            'Dalam pembagian hak bersama ini tidak terdapat kelebihan nilai yang diperoleh oleh salah satu pihak. --------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('Pembagian hak bersama ini dilakukan dengan syarat-syarat lebih lanjut sebagai berikut : ----------------------------------------------------------------------', {
            align: AlignmentType.JUSTIFIED,
            spaceAfter: 100,
            line: LINE_SPACING_1_5,
          }),

          p('------------------------------------------Pasal 1----------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Mulai hari ini yang diuraikan dalam akta ini telah menjadi milik masing-masing pihak yang memperolehnya dan karenanya segala keuntungan yang didapat dari, dan segala kerugian / beban atas hak tersebut diatas menjadi hak / beban pihak yang memperoleh hak.-----------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------Pasal 2.---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Hak tersebut diterima oleh masing-masing pihak yang memperolehnya menurut keadaan sebagaimana didapatinya pada hari ini dan masing-masing pihak dengan ini menyatakan tidak akan mengadakan segala tuntutan mengenai kerusakan dan / atau cacat yang tampak dan / atau tidak tampak. -------------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          ...pertanianPasal3Block(),
          p('------------------------------------------Pasal {{no_pasal_pengukuran}}.---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Dalam hal terdapat perbedaan luas tanah yang diuraikan dalam pembagian hak bersama ini dengan hasil pengukuran oleh instansi Badan Pertanahan {{kabupaten_tanah}}, maka para pihak akan menerima hasil pengukuran instansi Badan Pertanahan {{kabupaten_tanah}} tersebut. ---------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------Pasal {{no_pasal_keabsahan}}.--------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Semua pihak menjamin bahwa surat tanda bukti hak atas tanahnya adalah satu-satunya yang sah, tidak pernah dipalsukan dan tidak pernah dibuat duplikatnya oleh instansi yang berwenang atas permintaannya. Para pihak menjamin bahwa pembagian tersebut di atas adalah kesepakatan bersama. Sehubungan dengan hal tersebut maka semua pihak menyatakan tegas membebaskan Pejabat Pembuat Akta Tanah Sementara dari segala tuntutan berupa apapun juga mengenai hal-hal tersebut. --------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------Pasal {{no_pasal_domisili}}.---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Para pihak dalam hal ini dengan segala akibatnya memilih tempat kediaman hukum yang umum dan tidak berubah pada Kantor Panitera Pengadilan Negeri {{kabupaten_tanah}}.---------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------Pasal {{no_pasal_biaya}}.---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Biaya pembuatan akta ini, uang saksi dan segala biaya mengenai peralihan hak ini, dibayar oleh pihak penerima.---------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_aphb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // ==================== 2. KLAUSUL & PASAL KHUSUS JUAL BELI (AJB) ====================
          p('{{#is_ajb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('Jual beli ini meliputi pula : ------------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 60 }),
          p(
            'Tanah dan bangunan, berikut segala sesuatu yang terdapat diatas bidang tanah hak tersebut yang menurut sifat, peruntukannya atau menurut penetapan Undang-Undang dapat dianggap sebagai barang tidak bergerak. Selanjutnya semua yang diuraikan di atas dalam akta ini disebut “Obyek Jual Beli” .-----------------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 120, line: LINE_SPACING_1_5 }
          ),
          p('Pihak Pertama dan Pihak Kedua menerangkan bahwa : -----------------------', { align: AlignmentType.LEFT, spaceAfter: 80 }),
          p(
            'a. Jual beli ini dilakukan dengan harga {{nilai_transaksi_rp}} ({{nilai_transaksi_terbilang}})----------------------------------------------------------------',
            { indentLeft: 200, align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p(
            'b. Pihak Pertama mengaku telah menerima sepenuhnya uang tersebut diatas dari Pihak Kedua dan untuk penerimaan uang tersebut, akta ini berlaku pula sebagai tanda penerimaan yang sah (kwitansi).--------------',
            { indentLeft: 200, align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p('c. Jual beli ini dilakukan dengan syarat-syarat sebagai berikut : ------------', {
            align: AlignmentType.JUSTIFIED,
            indentLeft: 200,
            spaceAfter: 100,
          }),

          p('------------------------------------------- Pasal 1 ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Mulai hari ini obyek Jual Beli yang diuraikan dalam akta ini telah menjadi hak milik Pihak Kedua dan karenanya segala keuntungan yang didapat dari, dan segala kerugian/beban atas obyek Jual Beli tersebut di atas menjadi hak/beban Pihak Kedua.---------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('----------------------------------------- Pasal 2 ----------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Pihak Pertama Menjamin, bahwa obyek Jual Beli tersebut diatas tidak tersangkut dalam satu sengketa, bebas dari sitaan, tidak terikat sebagai jaminan untuk sesuatu utang yang tidak tercatat dalam sertipikat, dan bebas dari beban-beban lain--------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          ...pertanianPasal3Block(),
          p('------------------------------------------ Pasal {{no_pasal_pengukuran}} ----------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Dalam hal terdapat perbedaan luas tanah yang menjadi obyek Jual Beli dalam akta ini dengan hasil pengukuran oleh instansi Badan Pertanahan {{kabupaten_tanah}}, maka para pihak akan menerima hasil pengukuran instansi Badan Pertanahan {{kabupaten_tanah}} tersebut dengan tidak mempermasalahkan dan tidak akan saling mengadakan gugatan.-----------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------ Pasal {{no_pasal_keabsahan}} --------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            '- Pihak pertama menjamin bahwa surat tanda bukti hak atas tanahnya adalah satu-satunya yang sah, tidak pernah dipalsukan dan tidak pernah dibuat duplikatnya oleh instansi yang berwenang atas permintaannya .--',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p(
            '- Pihak Kedua telah melihat dan mengetahui keadaan tanah serta bukti surat kepemilikan atas tanah tersebut dan menerangkan dapat menerima sebagaimana mestinya .--------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p(
            '- Menurut keterangan para Pihak dokumen-dokumen penyertanya adalah asli dan tidak dipalsukan.----------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: LINE_SPACING_1_5 }
          ),
          p(
            '- Sehubungan dengan hal tersebut maka para pihak menyatakan dengan tegas membebaskan Pejabat Pembuat Akta Tanah dari segala Tuntutan berupa apapun mengenai hal-hal tersebut.----------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------- Pasal {{no_pasal_domisili}} --------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Kedua belah pihak dalam hal ini dengan segala akibatnya memilih tempat kediaman hukum yang umum dan tidak berubah pada Kantor Panitera Pengadilan Negeri {{kabupaten_tanah}} .-------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('-------------------------------------------- Pasal {{no_pasal_biaya}} -------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Biaya pembuatan akta ini, uang saksi dan segala biaya peralihan hak ini dibayar oleh Pihak Penerima .-------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_ajb}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // ==================== 3. KLAUSUL & PASAL KHUSUS AKTA HIBAH ====================
          p('{{#is_hibah}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('Hibah ini meliputi pula : ------------------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 60 }),
          p(
            'Berikut segala sesuatu yang terdapat diatas bidang tanah hak tersebut yang menurut sifat, peruntukannya atau menurut penetapan Undang-Undang dapat dianggap sebagai barang tidak bergerak. -----------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 80, line: LINE_SPACING_1_5 }
          ),
          p('Selanjutnya semua yang diuraikan di atas dalam akta ini disebut “Obyek Hibah”-----------------------------------------------------------------------------------', {
            align: AlignmentType.JUSTIFIED,
            spaceAfter: 80,
            line: LINE_SPACING_1_5,
          }),
          p('Pihak Pertama dan Pihak Kedua menerangkan bahwa hibah ini dilakukan dengan syarat-syarat sebagai berikut : -------------------------------------------', {
            align: AlignmentType.JUSTIFIED,
            spaceAfter: 100,
            line: LINE_SPACING_1_5,
          }),

          p('-------------------------------------------Pasal 1--------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Mulai hari ini obyek Hibah yang diuraikan dalam akta ini telah menjadi hak milik Pihak Kedua dan karenanya segala keuntungan yang didapat dari, dan segala kerugian/ beban atas obyek Hibah tersebut di atas menjadi hak/ beban Pihak Kedua.------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('-----------------------------------------Pasal 2-----------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Pihak Pertama Menjamin, bahwa obyek Hibah tersebut diatas tidak tersangkut dalam satu sengketa, bebas dari sitaan, tidak terikat sebagai jaminan untuk sesuatu utang yang tidak tercatat dalam sertipikat, dan bebas dari beban-beban lainnya yang berupa apapun.-------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          ...pertanianPasal3Block(),
          p('------------------------------------------Pasal {{no_pasal_pengukuran}}---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Dalam hal terdapat perbedaan luas tanah yang menjadi obyek Hibah dalam akta ini dengan hasil pengukuran oleh instansi Badan Pertanahan {{kabupaten_tanah}}, maka para pihak akan menerima hasil pengukuran instansi Badan Pertanahan {{kabupaten_tanah}} tersebut dengan tidak mempermasalahkan dan tidak akan saling mengadakan gugatan.-----------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('-------------------------------------------Pasal {{no_pasal_penerimaan_hibah}}---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Pihak Kedua mengaku telah menerima {{penggunaan_tanah}} yang menjadi obyek Hibah tersebut dari Pihak Pertama.------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('----------------------------------------Pasal {{no_pasal_penjaminan_hibah}}-----------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Para pihak dengan ini menjamin kebenaran semua data-data atau dokumen-dokumen yang disampaikan kepada saya selaku Pejabat Pembuat Akta Tanah Sementara, bilamana dikemudian hari ternyata timbul masalah atau ada tuntutan hukum dari pihak yang berkepentingan maka para pihak akan bertanggung jawab atas segala resiko yang timbul dari saya selaku Pejabat Pembuat Akta Tanah Sementara dibebaskan dari segala tuntutan hukum.--------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('--------------------------------------------Pasal {{no_pasal_domisili}}--------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Kedua belah pihak dalam hal ini dengan segala akibatnya memilih tempat kediaman hukum yang umum dan tidak berubah pada Kantor Panitera Pengadilan Negeri {{kabupaten_tanah}}.----------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('-------------------------------------------Pasal {{no_pasal_biaya}}---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Biaya pembuatan akta ini, uang saksi dan segala biaya peralihan hak ini dibayar oleh Pihak Penerima.-------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_hibah}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // ==================== 4. KLAUSUL & PASAL LAYANAN LAINNYA (WARIS / PEMECAHAN) ====================
          p('{{#is_lainnya}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('Peralihan hak ini meliputi pula : ------------------------------------------------------', { align: AlignmentType.LEFT, spaceAfter: 60 }),
          p(
            'Tanah dan bangunan, berikut segala sesuatu yang terdapat diatas bidang tanah hak tersebut yang menurut sifat, peruntukannya atau menurut penetapan Undang-Undang dapat dianggap sebagai barang tidak bergerak. Selanjutnya semua yang diuraikan di atas dalam akta ini disebut "Obyek Peralihan Hak".-----------------------------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 120, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------- Pasal 1 ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Mulai hari ini obyek yang diuraikan dalam akta ini telah menjadi hak milik Pihak Kedua dan karenanya segala keuntungan yang didapat dari, dan segala kerugian/beban atas obyek tersebut di atas menjadi hak/beban Pihak Kedua.---------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------- Pasal 2 ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Pihak Pertama menjamin, bahwa obyek tersebut diatas tidak tersangkut dalam satu sengketa, bebas dari sitaan, tidak terikat sebagai jaminan untuk sesuatu utang yang tidak tercatat dalam sertipikat, dan bebas dari beban-beban lain.--------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          ...pertanianPasal3Block(),
          p('------------------------------------------- Pasal {{no_pasal_pengukuran}} ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Dalam hal terdapat perbedaan luas tanah yang menjadi obyek dalam akta ini dengan hasil pengukuran oleh instansi Badan Pertanahan {{kabupaten_tanah}}, maka para pihak akan menerima hasil pengukuran instansi Badan Pertanahan {{kabupaten_tanah}} tersebut dengan tidak mempermasalahkan dan tidak akan saling mengadakan gugatan.-----------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------- Pasal {{no_pasal_keabsahan}} ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Para pihak menjamin bahwa surat tanda bukti hak atas tanah serta dokumen-dokumen penyertanya adalah asli, sah, dan membebaskan Pejabat Pembuat Akta Tanah Sementara dari segala tuntutan berupa apapun mengenai hal-hal tersebut.----------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------- Pasal {{no_pasal_domisili}} ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Kedua belah pihak dalam hal ini dengan segala akibatnya memilih tempat kediaman hukum yang umum dan tidak berubah pada Kantor Panitera Pengadilan Negeri {{kabupaten_tanah}}.-------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('------------------------------------------- Pasal {{no_pasal_biaya}} ---------------------------------------', {
            align: AlignmentType.CENTER,
            spaceAfter: 60,
          }),
          p(
            'Biaya pembuatan akta ini, uang saksi dan segala biaya peralihan hak ini dibayar oleh Pihak Penerima.-------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          p('{{/is_lainnya}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // ==================== SAKSI-SAKSI ====================
          p(
            'Akhirnya hadir juga dihadapan saya, dengan dihadiri oleh saksi-saksi yang sama dan disebutkan pada akhir akta ini.--------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          pRuns(
            [
              { text: '1. ', bold: true },
              { text: '{{saksi_1_nama}}, ', bold: true },
              {
                text: 'Warga Negara Indonesia, lahir di {{saksi_1_tempat_lahir}}, tanggal {{saksi_1_tanggal_lahir}}, Pekerjaan {{saksi_1_pekerjaan}}, bertempat tinggal di {{saksi_1_alamat}}. Pemegang Kartu Tanda Penduduk Nomor : {{saksi_1_nik}}--------------------------------------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, indentLeft: 200, spaceAfter: 100, line: LINE_SPACING_1_5 }
          ),
          pRuns(
            [
              { text: '2. ', bold: true },
              { text: '{{saksi_2_nama}}, ', bold: true },
              {
                text: 'Warga Negara Indonesia, lahir di {{saksi_2_tempat_lahir}}, tanggal {{saksi_2_tanggal_lahir}}, Pekerjaan {{saksi_2_pekerjaan}}, bertempat tinggal di {{saksi_2_alamat}}. Pemegang Kartu Tanda Penduduk Nomor : {{saksi_2_nik}}.-----------------------------------',
              },
            ],
            { align: AlignmentType.JUSTIFIED, indentLeft: 200, spaceAfter: 120, line: LINE_SPACING_1_5 }
          ),

          p(
            'Sebagai saksi-saksi dan setelah dibacakan serta dijelaskan, maka sebagai bukti kebenaran pernyataan yang dikemukakan oleh {{daftar_sebutan_para_pihak}} tersebut diatas, akta ini ditanda tangani/cap ibu jari oleh {{daftar_sebutan_para_pihak}}, para saksi dan saya PPATS, sebanyak 2 (dua) rangkap asli, yaitu 1 (satu) rangkap lembar pertama disimpan di kantor saya, dan 1 (satu) rangkap lembar kedua disampaikan kepada Kantor Pertanahan {{kabupaten_tanah}}, untuk keperluan pendaftaran peralihan hak dalam akta ini.--------------------------------------------------------',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 200, line: LINE_SPACING_1_5 }
          ),

          // ==================== HALAMAN TANDA TANGAN ====================
          new Paragraph({ children: [new PageBreak()] }),

          // Mode AJB / Hibah: 2 Kelompok Pihak (Pihak Pertama bisa multi-pemilik, Pihak Kedua)
          p('{{#sig_mode_grouped_2}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('Pihak Pertama', { bold: true, align: AlignmentType.CENTER, spaceAfter: 80 }),
                      p('{{#pihak_pertama_ttd}}', { align: AlignmentType.CENTER, spaceAfter: 0, line: 240 }),
                      p('{{#show_no}}{{no}}. {{/show_no}}{{nama}}', { align: AlignmentType.CENTER, spaceBefore: 640, spaceAfter: 160 }),
                      p('{{/pihak_pertama_ttd}}', { align: AlignmentType.CENTER, spaceAfter: 120, line: 240 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('Pihak Kedua', { bold: true, align: AlignmentType.CENTER, spaceAfter: 80 }),
                      p('{{#pihak_kedua_ttd}}', { align: AlignmentType.CENTER, spaceAfter: 0, line: 240 }),
                      p('{{#show_no}}{{no}}. {{/show_no}}{{nama}}', { align: AlignmentType.CENTER, spaceBefore: 640, spaceAfter: 160 }),
                      p('{{/pihak_kedua_ttd}}', { align: AlignmentType.CENTER, spaceAfter: 120, line: 240 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          p('{{/sig_mode_grouped_2}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // Mode 2 Pihak Terpisah (mis. APHB 2 pihak)
          p('{{#sig_mode_2}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('{{sig_c1_label}}', { bold: true, align: AlignmentType.CENTER, spaceAfter: 720 }),
                      p('{{sig_c1_nama}}', { align: AlignmentType.CENTER, spaceAfter: 280 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('{{sig_c2_label}}', { bold: true, align: AlignmentType.CENTER, spaceAfter: 720 }),
                      p('{{sig_c2_nama}}', { align: AlignmentType.CENTER, spaceAfter: 280 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          p('{{/sig_mode_2}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // Mode Grid 3 Kolom Berulang (APHB 3..20+ pihak terpisah)
          p('{{#sig_mode_grid}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{#sig_rows}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 33, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('{{c1_label}}', { bold: true, align: AlignmentType.CENTER, spaceAfter: 720 }),
                      p('{{c1_nama}}', { align: AlignmentType.CENTER, spaceAfter: 240 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 34, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('{{c2_label}}', { bold: true, align: AlignmentType.CENTER, spaceAfter: 720 }),
                      p('{{c2_nama}}', { align: AlignmentType.CENTER, spaceAfter: 240 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 33, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('{{c3_label}}', { bold: true, align: AlignmentType.CENTER, spaceAfter: 720 }),
                      p('{{c3_nama}}', { align: AlignmentType.CENTER, spaceAfter: 240 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          p('{{/sig_rows}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('{{/sig_mode_grid}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // Persetujuan Pasangan TTD (hanya tampil jika Ada Persetujuan Pasangan)
          p('{{#ada_persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),
          p('Disetujui Pasangan', { bold: true, align: AlignmentType.CENTER, spaceBefore: 120, spaceAfter: 720 }),
          p('{{pasangan_nama}}', { align: AlignmentType.CENTER, spaceAfter: 280 }),
          p('{{/ada_persetujuan_pasangan}}', { align: AlignmentType.LEFT, spaceAfter: 0, line: 240 }),

          // Saksi 1 & Saksi 2
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('Saksi', { bold: true, align: AlignmentType.CENTER, spaceBefore: 120, spaceAfter: 720 }),
                      p('{{saksi_1_nama}}', { align: AlignmentType.CENTER, spaceAfter: 280 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      p('Saksi', { bold: true, align: AlignmentType.CENTER, spaceBefore: 120, spaceAfter: 720 }),
                      p('{{saksi_2_nama}}', { align: AlignmentType.CENTER, spaceAfter: 280 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          p('Pejabat Pembuat Akta Tanah Sementara', { bold: true, align: AlignmentType.CENTER, spaceBefore: 240, spaceAfter: 40 }),
          p('{{jabatan_ppats}}', { bold: true, align: AlignmentType.CENTER, spaceAfter: 840 }),
          p('{{nama_ppats}}', { bold: true, italic: true, underline: true, align: AlignmentType.CENTER, spaceAfter: 40 }),
        ],
      },
    ],
  });
}

// Pengaturan margin & tipografi Lampiran Bundel Akta F4 sesuai contoh LAMPIRAN JUAL (Hal. 1-9)
export const LAMPIRAN_MARGINS = {
  top: 1247, // 2,2 cm
  bottom: 1134, // 2,0 cm
  left: 1440, // 2,54 cm
  right: 1440, // 2,54 cm
  gutter: 0,
  header: 0,
  footer: 0,
};

const LAMPIRAN_SIZE = 22; // 11 pt Bookman Old Style sesuai contoh PDF lampiran
const LAMPIRAN_LINE = 276; // 1.15 spasi agar setiap formulir tepat 1 halaman F4 (dan Hal. 8-9 tepat 2 halaman)
const LAMPIRAN_LINE_AIRY = 300; // 1.25 spasi untuk surat pernyataan ringkas

function lp(
  text: string,
  opts?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    size?: number;
    spaceBefore?: number;
    spaceAfter?: number;
    line?: number;
    indentLeft?: number;
    firstLine?: number;
    hanging?: number;
    pageBreakBefore?: boolean;
    borderBottom?: boolean;
  }
): Paragraph {
  return new Paragraph({
    alignment: opts?.align ?? AlignmentType.LEFT,
    pageBreakBefore: opts?.pageBreakBefore ?? false,
    indent:
      opts?.indentLeft || opts?.firstLine || opts?.hanging
        ? {
            left: opts.indentLeft,
            firstLine: opts.firstLine,
            hanging: opts.hanging,
          }
        : undefined,
    spacing: {
      before: opts?.spaceBefore ?? 0,
      after: opts?.spaceAfter ?? 20,
      line: opts?.line ?? LAMPIRAN_LINE,
    },
    border: opts?.borderBottom
      ? {
          bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
        }
      : undefined,
    children: [
      new TextRun({
        text,
        bold: opts?.bold ?? false,
        italics: opts?.italic ?? false,
        underline: opts?.underline ? {} : undefined,
        size: opts?.size ?? LAMPIRAN_SIZE,
        font: LAMPIRAN_FONT,
      }),
    ],
  });
}

function lpRuns(
  runs: Array<{ text: string; bold?: boolean; italic?: boolean; underline?: boolean; size?: number }>,
  opts?: {
    align?: (typeof AlignmentType)[keyof typeof AlignmentType];
    spaceBefore?: number;
    spaceAfter?: number;
    line?: number;
    indentLeft?: number;
    firstLine?: number;
    hanging?: number;
  }
): Paragraph {
  return new Paragraph({
    alignment: opts?.align ?? AlignmentType.LEFT,
    indent:
      opts?.indentLeft || opts?.firstLine || opts?.hanging
        ? {
            left: opts.indentLeft,
            firstLine: opts.firstLine,
            hanging: opts.hanging,
          }
        : undefined,
    spacing: {
      before: opts?.spaceBefore ?? 0,
      after: opts?.spaceAfter ?? 20,
      line: opts?.line ?? LAMPIRAN_LINE,
    },
    children: runs.map(
      (r) =>
        new TextRun({
          text: r.text,
          bold: r.bold ?? false,
          italics: r.italic ?? false,
          underline: r.underline ? {} : undefined,
          size: r.size ?? LAMPIRAN_SIZE,
          font: LAMPIRAN_FONT,
        })
    ),
  });
}

function makeLampiranKvTable(
  rows: Array<[string, string]>,
  opts?: { labelPct?: number; indentTwips?: number; spaceAfterRow?: number; line?: number }
): Table {
  const labelPct = opts?.labelPct ?? 26;
  const line = opts?.line ?? 260;
  const spaceAfterRow = opts?.spaceAfterRow ?? 10;
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    indent: opts?.indentTwips ? { size: opts.indentTwips, type: WidthType.DXA } : undefined,
    rows: rows.map(
      ([k, v]) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: labelPct, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [lp(k, { spaceAfter: spaceAfterRow, line })],
            }),
            new TableCell({
              width: { size: 100 - labelPct, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [lp(`: ${v}`, { spaceAfter: spaceAfterRow, line })],
            }),
          ],
        })
    ),
  });
}

function makeSaksiSignLines(spaceBefore = 200): Array<Paragraph | Table> {
  return [
    lp('Saksi – saksi :', { spaceBefore, spaceAfter: 140, line: LAMPIRAN_LINE_AIRY }),
    new Table({
      width: { size: 70, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 65, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [lp('1. {{saksi_1_nama}}', { spaceAfter: 180, line: LAMPIRAN_LINE_AIRY })],
            }),
            new TableCell({
              width: { size: 35, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [lp('(                    )', { spaceAfter: 180, line: LAMPIRAN_LINE_AIRY })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              width: { size: 65, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [lp('2. {{saksi_2_nama}}', { spaceAfter: 60, line: LAMPIRAN_LINE_AIRY })],
            }),
            new TableCell({
              width: { size: 35, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [lp('(                    )', { spaceAfter: 60, line: LAMPIRAN_LINE_AIRY })],
            }),
          ],
        }),
      ],
    }),
  ];
}

const LEFT_RIGHT_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
  right: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
};

const LEFT_RIGHT_BOTTOM_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
  left: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
  right: { style: BorderStyle.SINGLE, size: 6, color: '1E293B' },
};

function buildLampiranBundelDocument(): Document {
  // Tabel Checklist 13 Dokumen Hal. 1 (tanpa garis horizontal antarbaris, Keterangan merge vertikal)
  const checklistItems: Array<[string, string[], string]> = [
    [
      '1.',
      [
        'Sertipikat {{jenis_hak}} (Asli)',
        '{{sertipikat_ringkas_baris}}',
        'Desa {{desa_tanah}}',
        'An. {{sertipikat_an}}',
      ],
      '1 (satu)',
    ],
    [
      '2.',
      [
        '{{lampiran_akta_label}}:',
        'Nomor : {{nomor_akta}}',
        'Tanggal : {{tanggal_akta}}',
      ],
      '1 (satu)',
    ],
    ['3.', ['Pengecekan Sertipikat'], '1 (satu)'],
    ['4.', ['Surat Kuasa'], '1 (satu)'],
    ['5.', ['Foto Copy KTP Penerima Kuasa'], '1 (satu)'],
    ['6.', ['Surat penguasaan fisik tanah dan tidak sengketa'], '1 (satu)'],
    ['7.', ['Foto Copy KTP Para Pihak, Saksi'], '1 (satu)'],
    ['8.', ['Foto Copy Kartu Keluarga'], '1 (satu)'],
    ['9.', ['Surat pernyataan bphtb'], '1 (satu)'],
    ['10.', ['Surat Pernyataan pbb'], '1 (satu)'],
    ['11.', ['Foto Copy SPPD BPHTB, SPPT PBB, PPH'], '1 (satu)'],
    ['12.', ['Surat pernyataan keabsahan dan kebenaran dokumen'], '1 (satu)'],
    ['13.', ['Surat pernyataan pemrosesan data pribadi'], '1 (satu)'],
  ];

  const checklistTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9, type: WidthType.PERCENTAGE },
            borders: THIN_BORDER,
            children: [lp('Nomor', { bold: true, align: AlignmentType.CENTER, spaceBefore: 60, spaceAfter: 60, line: 240 })],
          }),
          new TableCell({
            width: { size: 48, type: WidthType.PERCENTAGE },
            borders: THIN_BORDER,
            children: [lp('Dokumen yang diserahkan', { bold: true, align: AlignmentType.CENTER, spaceBefore: 60, spaceAfter: 60, line: 240 })],
          }),
          new TableCell({
            width: { size: 14, type: WidthType.PERCENTAGE },
            borders: THIN_BORDER,
            children: [lp('Jumlah', { bold: true, align: AlignmentType.CENTER, spaceBefore: 60, spaceAfter: 60, line: 240 })],
          }),
          new TableCell({
            width: { size: 29, type: WidthType.PERCENTAGE },
            borders: THIN_BORDER,
            children: [lp('Keterangan', { bold: true, align: AlignmentType.CENTER, spaceBefore: 60, spaceAfter: 60, line: 240 })],
          }),
        ],
      }),
      ...checklistItems.map(([no, docLines, jml], idx) => {
        const isFirst = idx === 0;
        const isLast = idx === checklistItems.length - 1;
        const colBorder = isLast ? LEFT_RIGHT_BOTTOM_BORDERS : LEFT_RIGHT_BORDERS;
        return new TableRow({
          children: [
            new TableCell({
              width: { size: 9, type: WidthType.PERCENTAGE },
              borders: colBorder,
              children: [
                lp(no, {
                  align: AlignmentType.CENTER,
                  spaceBefore: isFirst ? 60 : 10,
                  spaceAfter: isLast ? 60 : 10,
                  line: 240,
                }),
              ],
            }),
            new TableCell({
              width: { size: 48, type: WidthType.PERCENTAGE },
              borders: colBorder,
              children: docLines.map((lineText, lIdx) =>
                lp(lineText, {
                  spaceBefore: isFirst && lIdx === 0 ? 60 : 0,
                  spaceAfter: isLast && lIdx === docLines.length - 1 ? 60 : lIdx === docLines.length - 1 ? 30 : 0,
                  line: 240,
                })
              ),
            }),
            new TableCell({
              width: { size: 14, type: WidthType.PERCENTAGE },
              borders: colBorder,
              children: [
                lp(jml, {
                  align: AlignmentType.CENTER,
                  spaceBefore: isFirst ? 60 : 10,
                  spaceAfter: isLast ? 60 : 10,
                  line: 240,
                }),
              ],
            }),
            new TableCell({
              width: { size: 29, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              verticalMerge: isFirst ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE,
              children: isFirst
                ? [
                    lp('{{ket_pihak1_label}}:', { spaceBefore: 240, spaceAfter: 10, line: 250 }),
                    lp('{{ket_pihak1_names}}', { spaceAfter: 200, line: 250 }),
                    lp('{{ket_pihak2_label}}:', { spaceAfter: 10, line: 250 }),
                    lp('{{ket_pihak2_names}}', { spaceAfter: 40, line: 250 }),
                  ]
                : [lp('', { spaceAfter: 0, line: 220 })],
            }),
          ],
        });
      }),
    ],
  });

  return new Document({
    styles: createDocxStyles(LAMPIRAN_FONT),
    sections: [
      {
        properties: {
          page: {
            size: PAPER_SIZE_F4_TWIPS,
            margin: LAMPIRAN_MARGINS,
          },
        },
        children: [
          // ==================== HALAMAN 1: SURAT PENGANTAR KE KEPALA KANTOR ATR/BPN ====================
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: {
                      ...NO_BORDER,
                      bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
                    },
                    children: [
                      lp('{{nama_ppats}}', { bold: true, spaceAfter: 0, line: 240 }),
                      lp('PPATS {{nama_kecamatan_upper}}', { bold: true, spaceAfter: 0, line: 240 }),
                      lp('Alamat : {{alamat_kantor}}', { spaceAfter: 0, line: 240 }),
                      lp('Telp. {{telepon_kantor}}', { spaceAfter: 40, line: 240 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('', { spaceAfter: 0, line: 240 }),
                      lp('', { spaceAfter: 0, line: 240 }),
                      lp('', { spaceAfter: 0, line: 240 }),
                      lp('{{kota_kabupaten}}, {{tanggal_akta}}', { indentLeft: 720, spaceAfter: 0, line: 240 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Nomor     : {{nomor_pengantar}}', { spaceBefore: 100, spaceAfter: 0, line: 250 }),
                      lp('Lampiran : 1 (satu) set', { spaceAfter: 0, line: 250 }),
                      lpRuns(
                        [
                          { text: 'Perihal    : ' },
                          { text: '{{perihal_pengantar}}', italic: true, underline: true },
                        ],
                        { spaceAfter: 40, line: 250 }
                      ),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Kepada', { indentLeft: 720, spaceBefore: 40, spaceAfter: 0, line: 250 }),
                      lp('Yth.Kepala Kantor ATR/BPN', { indentLeft: 720, spaceAfter: 0, line: 250 }),
                      lp('{{kabupaten_tanah}}', { indentLeft: 720, spaceAfter: 0, line: 250 }),
                      lp('Di {{kota_kabupaten}}', { indentLeft: 720, spaceAfter: 40, line: 250 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          lp('Dengan hormat,', { spaceBefore: 80, spaceAfter: 30, line: 250 }),
          lp(
            'Bersama ini kami sampaikan berkas pendaftaran peralihan hak karena {{perihal_kalimat}} dengan rincian Dokumen sebagai berikut :',
            { align: AlignmentType.JUSTIFIED, firstLine: 720, spaceAfter: 100, line: 260 }
          ),
          checklistTable,
          lp(
            'Mohon dilakukan pendaftaran peralihan hak karena {{perihal_kalimat}} atas sertipikat tersebut diatas, atas kerja sama yang baik saya ucapkan terima kasih.',
            { align: AlignmentType.JUSTIFIED, firstLine: 720, spaceBefore: 100, spaceAfter: 120, line: 260 }
          ),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 44, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Yang Menerima', { align: AlignmentType.CENTER, spaceAfter: 840, line: 250 }),
                      lp('(                    )', { align: AlignmentType.CENTER, spaceAfter: 20, line: 250 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 56, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Pejabat Pembuat Akta Tanah Sementara', { align: AlignmentType.CENTER, spaceAfter: 0, line: 250 }),
                      lp('{{jabatan_camat}}', { align: AlignmentType.CENTER, spaceAfter: 820, line: 250 }),
                      lp('({{nama_ppats}})', { bold: true, align: AlignmentType.CENTER, spaceAfter: 20, line: 250 }),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // ==================== HALAMAN 2: PERMOHONAN PENDAFTARAN PERALIHAN HAK ====================
          lp('Kepada Yth.', { pageBreakBefore: true, indentLeft: 4400, spaceAfter: 0, line: 255 }),
          lp('Kantor Kementrian Agraria dan Tata Ruang/', { indentLeft: 4400, spaceAfter: 0, line: 255 }),
          lp('Badan Pertanahan Nasional {{kabupaten_tanah}}', { indentLeft: 4400, spaceAfter: 0, line: 255 }),
          lp('di {{kota_kabupaten}}', { indentLeft: 4400, spaceAfter: 180, line: 255 }),

          lp('Dengan Hormat,', { spaceAfter: 20, line: 255 }),
          lp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: 255 }),
          makeLampiranKvTable(
            [
              ['Nama', '{{kuasa_bpn_nama}}'],
              ['Tempat/Tgl.lahir', '{{kuasa_bpn_ttl}}'],
              ['Nomor KTP', '{{kuasa_bpn_nik}}'],
              ['Alamat', '{{kuasa_bpn_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 25, spaceAfterRow: 0, line: 250 }
          ),
          lp('Dalam hal ini bertindak untuk selaku kuasa dari :', { spaceBefore: 120, spaceAfter: 60, line: 255 }),
          makeLampiranKvTable(
            [
              ['Nama', '{{penerima_nama}}'],
              ['Tempat/Tgl.lahir', '{{penerima_ttl}}'],
              ['Pekerjaan', '{{penerima_pekerjaan}}'],
              ['Nomor KTP', '{{penerima_nik}}'],
              ['Alamat', '{{penerima_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 25, spaceAfterRow: 0, line: 250 }
          ),
          lp('Berdasarkan surat kuasa Tanggal : {{tanggal_akta_dash}}', { spaceBefore: 120, spaceAfter: 20, line: 255 }),
          lp('Dengan ini mengajukan permohonan :', { spaceAfter: 40, line: 255 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            indent: { size: 420, type: WidthType.DXA },
            rows: [
              makeBorderlessRow('1. Pengukuran', '8. Pemecahan/Penggabungan Hak', 48, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('2. Konversi / Pendaftaran Hak', '9. Pendaftaran Hak Tanggungan', 48, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('3. Pendaftaran Hak Milik Sarusun', '10. Roya atas Hak Tanggungan', 48, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('4. Pendaftaran Tanah Wakaf', '11. Penerbitan Sertipikat Pengganti', 48, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 48, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [lp('5. Pendaftaran Peralihan Hak', { bold: true, spaceAfter: 10, line: 250 })],
                  }),
                  new TableCell({
                    width: { size: 52, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [lp('12. Surat Keterangan Pendaftaran Tanah', { spaceAfter: 10, line: 250 })],
                  }),
                ],
              }),
              makeBorderlessRow('6. Pendaftaran Pemindahan Hak', '13. Pengecekan Sertipikat', 48, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('7. Pendaftaran Perubahan Hak', '14. Pencatatan ……………………………', 48, LAMPIRAN_FONT, LAMPIRAN_SIZE),
            ],
          }),
          lp('Atas bidang tanah Hak / Tanah Negara :', { spaceBefore: 100, spaceAfter: 10, line: 255 }),
          lp('Terletak di', { indentLeft: 420, spaceAfter: 10, line: 250 }),
          makeLampiranKvTable(
            [
              ['Desa', '{{desa_tanah}}'],
              ['Kecamatan', '{{kecamatan_tanah}}'],
              ['Kabupaten', '{{kabupaten_tanah}}'],
              ['{{label_nomor_sertipikat}}', '{{nib_atau_nomor_hak}}'],
            ],
            { indentTwips: 720, labelPct: 24, spaceAfterRow: 0, line: 250 }
          ),
          lp('Untuk melengkapi permohonan dimaksud, bersama ini kami lampirkan ;', { spaceBefore: 80, spaceAfter: 30, line: 255 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            indent: { size: 420, type: WidthType.DXA },
            rows: [
              makeBorderlessRow('1. {{jenis_hak}} {{sertipikat_ringkas_baris}}', '9. Surat Pembayaran BPHTB', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('2. {{lampiran_akta_label_lower}} : NO {{nomor_akta}} Tgl {{tanggal_akta}}', '10. Surat pernyataan PBB', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('3. Pengecekan Sertipikat', '11. Fotokopi PBB terbaru, SSPD BPHTB, PPh', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('4. Surat Kuasa', '12. Surat keabsahan dokumen', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('5. Fotokopi KTP penerima Kuasa', '13. Surat pernyataan pemrosesan data pribadi', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('6. Surat Pernyataan penguasaan fisik', '', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('7. Fotokopi KTP, KK Para Pihak', '', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
              makeBorderlessRow('8. Fotokopi KTP Para Saksi', '', 54, LAMPIRAN_FONT, LAMPIRAN_SIZE),
            ],
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ width: { size: 45, type: WidthType.PERCENTAGE }, borders: NO_BORDER, children: [lp('', { spaceAfter: 0 })] }),
                  new TableCell({
                    width: { size: 55, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Hormat Kami,', { align: AlignmentType.CENTER, spaceBefore: 120, spaceAfter: 800, line: 255 }),
                      lp('{{kuasa_bpn_nama}}, qq. {{penerima_ringkas}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 255 }),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // ==================== HALAMAN 3: SURAT KUASA SUBSTITUSI ====================
          lp('SURAT KUASA', { pageBreakBefore: true, bold: true, underline: true, align: AlignmentType.CENTER, spaceAfter: 200, line: 260 }),
          lp('Yang bertanda tangan dibawah ini :', { spaceAfter: 80, line: 260 }),
          makeLampiranKvTable(
            [
              ['Nama', '{{penerima_nama}}'],
              ['Tempat/Tgl.lahir', '{{penerima_ttl}}'],
              ['Pekerjaan', '{{penerima_pekerjaan}}'],
              ['Nomor KTP', '{{penerima_nik}}'],
              ['Alamat', '{{penerima_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 255 }
          ),
          lp('Sehubungan dengan Sertipikat atas tanah {{jenis_hak}} :', { spaceBefore: 160, spaceAfter: 40, line: 260 }),
          makeLampiranKvTable(
            [
              ['{{label_nomor_sertipikat}}', '{{nib_atau_nomor_hak}}'],
              ['Desa', '{{desa_tanah}}'],
              ['Kecamatan', '{{kecamatan_tanah}}'],
              ['Kabupaten', '{{kabupaten_tanah}}'],
            ],
            { indentTwips: 720, labelPct: 22, spaceAfterRow: 0, line: 255 }
          ),
          lp('Maka dengan ini memberi kuasa Substitusi kepada :', { spaceBefore: 160, spaceAfter: 40, line: 260 }),
          makeLampiranKvTable(
            [
              ['Nama', '{{kuasa_bpn_nama}}'],
              ['Tempat/Tgl.lahir', '{{kuasa_bpn_ttl}}'],
              ['Pekerjaan', '{{kuasa_bpn_pekerjaan}}'],
              ['Nomor KTP', '{{kuasa_bpn_nik}}'],
              ['Alamat', '{{kuasa_bpn_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 255 }
          ),
          lpRuns(
            [
              { text: '----------------------------------------------' },
              { text: 'KHUSUS', bold: true, underline: true },
              { text: '---------------------------------------------' },
            ],
            { align: AlignmentType.CENTER, spaceBefore: 160, spaceAfter: 100, line: 260 }
          ),
          lp(
            '-  Untuk dan atas nama pemberi Kuasa untuk mendaftarkan permohonan Peralihan Hak Sertipikat ke Kantor Badan Pertanahan Nasional {{kabupaten_tanah}}.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 240, spaceAfter: 60, line: 276 }
          ),
          lp(
            '-  Guna keperluan tersebut penerima kuasa berhak menghadap kepada pegawai Kantor Pertanahan yang berwenang, membayar biaya pendaftaran, menandatangani surat-surat yang diperlukan dalam permohonan pendaftaran tersebut serta mengambil sertipikat tanah kelak dikemudian hari setelah selesai.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 240, spaceAfter: 160, line: 276 }
          ),
          lp('Demikian kuasa ini diberikan untuk dapat dipergunakan sebagaimana mestinya.', { spaceAfter: 200, line: 276 }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 48, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('', { spaceAfter: 0, line: 260 }),
                      lp('Penerima Kuasa', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{kuasa_bpn_nama}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 52, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('{{kota_kabupaten}}, {{tanggal_akta}}', { align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('Pemberi Kuasa', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{penerima_nama}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          lp('Mengetahui', { align: AlignmentType.CENTER, spaceBefore: 160, spaceAfter: 0, line: 260 }),
          lp('PPATS / {{jabatan_camat_upper}}', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
          lp('{{nama_ppats}}', { underline: true, align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
          lp('NIP. {{nip_ppats}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),

          // ==================== HALAMAN 4: SURAT PERNYATAAN BPHTB ====================
          lp('SURAT PERNYATAAN', { pageBreakBefore: true, bold: true, underline: true, align: AlignmentType.CENTER, spaceAfter: 220, line: 280 }),
          lp('Yang bertanda tangan dibawah ini :', { spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }),
          makeLampiranKvTable(
            [
              ['Nama', '{{penerima_nama}}'],
              ['Tempat/Tgl.lahir', '{{penerima_ttl}}'],
              ['Pekerjaan', '{{penerima_pekerjaan}}'],
              ['Nomor KTP', '{{penerima_nik}}'],
              ['Alamat', '{{penerima_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 10, line: 276 }
          ),
          lp('Dengan ini menyatakan :', { spaceBefore: 180, spaceAfter: 80, line: LAMPIRAN_LINE_AIRY }),
          lp(
            '-  Bahwa saya telah {{pernyataan_bphtb_perbuatan}} sebidang tanah tersebut {{sertipikat_ringkas_baris}} Desa {{desa_tanah}} dengan Luas {{luas_m2}} m2 ({{luas_terbilang}}) yang terletak di Desa {{desa_tanah}}, Kecamatan {{kecamatan_tanah}}, {{kabupaten_tanah}}.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 240, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            '-  Bahwa apabila dalam Proses Peralihan Hak atas tanah yang saya terima tersebut diatas, ternyata terjadi kekurangan pembayaran BPHTB. Maka saya sanggup untuk membayar kekurangan BPHTB tersebut sesuai dengan ketentuan peraturan perundang-undangan yang berlaku',
            { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 240, spaceAfter: 300, line: LAMPIRAN_LINE_AIRY }
          ),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Mengetahui', { align: AlignmentType.CENTER, spaceBefore: 120, spaceAfter: 0, line: 260 }),
                      lp('PPATS / {{jabatan_camat_upper}}', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{nama_ppats}}', { underline: true, align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('NIP. {{nip_ppats}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('{{kota_kabupaten}}, {{tanggal_akta}}', { align: AlignmentType.CENTER, spaceAfter: 80, line: 260 }),
                      lp('Yang menyatakan', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{penerima_nama}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          ...makeSaksiSignLines(360),

          // ==================== HALAMAN 5: SURAT PERNYATAAN PENGUASAAN FISIK TANAH DAN TIDAK SENGKETA ====================
          lp('SURAT PERNYATAAN PENGUASAAN FISIK TANAH DAN TIDAK SENGKETA', {
            pageBreakBefore: true,
            bold: true,
            underline: true,
            align: AlignmentType.CENTER,
            spaceAfter: 180,
            line: 260,
          }),
          lp('Yang bertanda tangan dibawah ini :', { spaceAfter: 40, line: 250 }),
          makeLampiranKvTable(
            [
              ['Nama', '{{penerima_nama}}'],
              ['Tempat/Tgl.lahir', '{{penerima_ttl}}'],
              ['Pekerjaan', '{{penerima_pekerjaan}}'],
              ['Nomor KTP', '{{penerima_nik}}'],
              ['Alamat', '{{penerima_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 245 }
          ),
          lp('Dengan disaksikan oleh :', { spaceBefore: 100, spaceAfter: 40, line: 250 }),
          makeLampiranKvTable(
            [
              ['1. Nama', '{{saksi_1_nama}}'],
              ['   Tempat/Tgl.lahir', '{{saksi_1_ttl}}'],
              ['   Pekerjaan', '{{saksi_1_pekerjaan}}'],
              ['   Nomor KTP', '{{saksi_1_nik}}'],
              ['   Alamat', '{{saksi_1_alamat}}'],
            ],
            { indentTwips: 240, labelPct: 27, spaceAfterRow: 0, line: 245 }
          ),
          lp('', { spaceAfter: 40, line: 180 }),
          makeLampiranKvTable(
            [
              ['2. Nama', '{{saksi_2_nama}}'],
              ['   Tempat/Tgl. Lahir', '{{saksi_2_ttl}}'],
              ['   Pekerjaan', '{{saksi_2_pekerjaan}}'],
              ['   Nomor KTP', '{{saksi_2_nik}}'],
              ['   Alamat', '{{saksi_2_alamat}}'],
            ],
            { indentTwips: 240, labelPct: 27, spaceAfterRow: 0, line: 245 }
          ),
          lp('Pada hari ini {{hari_akta}} tanggal {{tanggal_akta}} dengan ini menyatakan :', {
            spaceBefore: 120,
            spaceAfter: 60,
            line: 255,
          }),
          lp(
            '1. Bahwa saya adalah benar-benar sebagai pemilik sebidang tanah seluas {{luas_m2}} M2 ({{luas_terbilang}}) yang terletak di Desa {{desa_tanah}}, Kecamatan {{kecamatan_tanah}}, {{kabupaten_tanah}} yang sudah bersertipikat {{jenis_hak}} {{sertipikat_ringkas_baris}} Desa {{desa_tanah}}, yang tercatat atas nama {{sertipikat_an}} dan merupakan produk Kantor Pertanahan {{kabupaten_tanah}}.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 260 }
          ),
          lp(
            '2. Bahwa saya sebagai pemilik tanah tersebut diatas, benar-benar menguasai tanah secara fisik sejak tahun {{menguasai_sejak_tahun}} sampai sekarang dan saya pergunakan untuk {{penggunaan_tanah}}',
            { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 260 }
          ),
          lp(
            '3. Bahwa tanah tersebut tidak dalam keadaan sengketa, konflik, perkara, Sita jaminan atau pemblokiran dari pihak lain.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 60, line: 260 }
          ),
          lp(
            'Demikian surat pernyataan ini saya buat dengan sebenarnya dan saya bertanggung jawab sepenuhnya atas pernyataan ini. Jika diperlukan saya sanggup mengangkat sumpah menurut Agama/Kepercayaan saya',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: 260 }
          ),
          lp(
            'Apabila dikemudian hari terjadi perkara perdata maupun pidana atas tanah ini, maka perkara tersebut merupakan tanggung jawab pribadi saya dengan tidak melibatkan Pejabat/Kantor Badan Pertanahan Nasional {{kabupaten_tanah}}.',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 160, line: 260 }
          ),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 58, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Saksi – saksi :', { spaceAfter: 160, line: 250 }),
                      lp('1. {{saksi_1_nama}}          (                    )', { spaceAfter: 320, line: 250 }),
                      lp('2. {{saksi_2_nama}}          (                    )', { spaceAfter: 40, line: 250 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 42, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Yang membuat pernyataan', { align: AlignmentType.CENTER, spaceAfter: 720, line: 250 }),
                      lp('{{penerima_nama}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 250 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          lp('Mengetahui', { align: AlignmentType.CENTER, spaceBefore: 60, spaceAfter: 0, line: 250 }),
          lp('PPATS / {{jabatan_camat_upper}}', { align: AlignmentType.CENTER, spaceAfter: 800, line: 250 }),
          lp('{{nama_ppats}}', { underline: true, align: AlignmentType.CENTER, spaceAfter: 0, line: 250 }),
          lp('NIP. {{nip_ppats}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 250 }),

          // ==================== HALAMAN 6: SURAT PERNYATAAN SPPT / PBB (PEMILIK AWAL) ====================
          lp('SURAT PERNYATAAN', { pageBreakBefore: true, bold: true, underline: true, align: AlignmentType.CENTER, spaceAfter: 220, line: 280 }),
          lp('Yang bertanda tangan dibawah ini :', { spaceAfter: 80, line: LAMPIRAN_LINE_AIRY }),
          makeLampiranKvTable(
            [
              ['Nama', '{{pemilik_awal_nama}}'],
              ['Tempat/Tgl lahir', '{{pemilik_awal_ttl}}'],
              ['Pekerjaan', '{{pemilik_awal_pekerjaan}}'],
              ['Nomor KTP', '{{pemilik_awal_nik}}'],
              ['Alamat', '{{pemilik_awal_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 10, line: 276 }
          ),
          lp(
            'Dengan menyatakan sebenarnya bahwa Surat Pemberitahuan Pajak Terhutang Pajak Bumi dan Bangunan (SPPT/PBB) Nomor Obyek Pajak ( NOP ) : {{nop}}',
            { align: AlignmentType.JUSTIFIED, spaceBefore: 200, spaceAfter: 40, line: LAMPIRAN_LINE_AIRY }
          ),
          lp('Atas tanah saya bersertipikat {{jenis_hak}} :', { spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }),
          makeLampiranKvTable(
            [
              ['{{label_nomor_sertipikat}}', '{{nib_atau_nomor_hak}}'],
              ['Desa', '{{desa_tanah}}'],
              ['Kecamatan', '{{kecamatan_tanah}}'],
              ['Kabupaten', '{{kabupaten_tanah}}'],
              ['Atas nama', '{{sertipikat_an}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 10, line: 276 }
          ),
          lp('Masih atas nama pemilik lama karena belum dibalik nama atas nama saya.', {
            align: AlignmentType.JUSTIFIED,
            spaceBefore: 100,
            spaceAfter: 40,
            line: LAMPIRAN_LINE_AIRY,
          }),
          lp('Demikian surat pernyataan ini saya buat agar diperunakan sebagaimana mestinya.', {
            align: AlignmentType.JUSTIFIED,
            spaceAfter: 340,
            line: LAMPIRAN_LINE_AIRY,
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Mengetahui', { align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('PPATS / {{jabatan_camat_upper}}', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{nama_ppats}}', { underline: true, align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('NIP. {{nip_ppats}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('{{kota_kabupaten}}, {{tanggal_akta}}', { align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('Yang menyatakan', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{pemilik_awal_nama}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          ...makeSaksiSignLines(360),

          // ==================== HALAMAN 7: SURAT PERTANGGUNGJAWABAN KEABSAHAN DAN KEBENARAN DOKUMEN (TANPA SAKSI) ====================
          lp('SURAT PERTANGGUNGJAWABAN KEABSAHAN DAN', {
            pageBreakBefore: true,
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 0,
            line: 270,
          }),
          lp('KEBENARAN DOKUMEN', {
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 240,
            line: 270,
          }),
          lp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: 260 }),
          makeLampiranKvTable(
            [
              ['Nama', '{{penerima_nama}}'],
              ['Tempat/Tgl.lahir', '{{penerima_ttl}}'],
              ['Pekerjaan', '{{penerima_pekerjaan}}'],
              ['Nomor KTP', '{{penerima_nik}}'],
              ['Alamat', '{{penerima_alamat}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 255 }
          ),
          lp('Merupakan pemilik tanah dengan jenis tanah :', { spaceBefore: 160, spaceAfter: 40, line: 260 }),
          lp('{{jenis_hak}} {{sertipikat_ringkas_baris}}', { spaceAfter: 80, line: 260 }),
          lp('Letak Tanah :', { spaceAfter: 40, line: 260 }),
          makeLampiranKvTable(
            [
              ['a. Desa', '{{desa_tanah}}'],
              ['b. Kecamatan', '{{kecamatan_tanah}}'],
              ['c. Kabupaten', '{{kabupaten_tanah}}'],
              ['d. Provinsi', '{{provinsi_tanah}}'],
            ],
            { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 255 }
          ),
          lp('Dengan ini menyatakan :', { spaceBefore: 160, spaceAfter: 60, line: 260 }),
          lp(
            '1. Sertipikat tersebut adalah asli dan nama yang tercantum dalam sertipikat tersebut adalah nama, serta pemegang hak sebenarnya dan menguasai secara fisik bidang tanah tersebut',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 60, line: 276 }
          ),
          lp(
            '2. Seluruh dokumen yang diunggah pada sistem aplikasi layanan pengecekan sertipikat/ SKPT pada Kantor Pertanahan {{kabupaten_tanah}} adalah benar dan sesuai dengan aslinya',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 60, line: 276 }
          ),
          lp(
            '3. Beritikad baik dan bertanggungjawab sepenuhnya atas penggunaan data yang diakses. Apabila dikemudian hari data tersebut dinyatakan tidak benar, maka saya bersedia menerima sanksi administratif maupun hukum sesuai ketentuan peraturan perundang-undangan yang berlaku.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 120, line: 276 }
          ),
          lp('Demikian surat pernyataan ini dibuat dengan sebenarnya untuk dapat dipergunakan sebagaimana mestinya.', {
            align: AlignmentType.JUSTIFIED,
            spaceAfter: 360,
            line: 276,
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('Mengetahui', { align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('PPATS/ {{jabatan_camat_upper}}', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{nama_ppats}}', { underline: true, align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 50, type: WidthType.PERCENTAGE },
                    borders: NO_BORDER,
                    children: [
                      lp('{{kota_kabupaten}}, {{tanggal_akta}}', { align: AlignmentType.CENTER, spaceAfter: 0, line: 260 }),
                      lp('Yang menyatakan', { align: AlignmentType.CENTER, spaceAfter: 900, line: 260 }),
                      lp('{{penerima_nama}}', { align: AlignmentType.CENTER, spaceAfter: 20, line: 260 }),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // ==================== HALAMAN 8-9: SURAT PERNYATAAN PERSETUJUAN PEMROSESAN DATA PRIBADI ====================
          lp('SURAT PERNYATAAN PERSETUJUAN PEMROSESAN DATA PRIBADI', {
            pageBreakBefore: true,
            bold: true,
            align: AlignmentType.CENTER,
            spaceAfter: 200,
            line: 280,
          }),
          lp(
            'Saya, yang bertandatangan di bawah ini, selaku pengguna layanan Pertanahan Kementerian Agraria dan Tata Ruang/Badan Pertanahan Nasional (“Kementerian ATR/BPN”), dengan ini menyatakan sebagai berikut:',
            { align: AlignmentType.JUSTIFIED, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lpRuns(
            [
              { text: '1. ', bold: true, italic: true },
              {
                text: 'Bahwa saya bertindak dalam kapasitas sebagai ',
              },
              {
                text: 'Diri sendiri (perorangan), saya menyatakan memiliki kewenangan hukum yang sah dan bertanggung jawab sepenuhnya atas kewenangan tersebut, serta membebaskan Kementerian ATR/BPN dari segala tuntutan atau klaim apabila di kemudian hari kewenangan tersebut terbukti tidak sah.',
                italic: true,
              },
              {
                text: ' Dengan ini saya memberikan persetujuan kepada Kementerian ATR/BPN selaku Pengendali Data Pribadi untuk melakukan pemrosesan Data Pribadi saya dan/atau pihak yang saya wakili, termasuk namun tidak terbatas pada pengumpulan, pencatatan, penyimpanan, pengolahan, penggunaan, pengungkapan, pengiriman, pembaruan, perbaikan, pembatasan, penghapusan, dan/atau pemusnahan sesuai ketentuan peraturan perundang-undangan.',
              },
            ],
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            '2. Saya telah menerima dan memahami informasi yang diberikan oleh Kementerian ATR/BPN sehubungan dengan tujuan pemrosesan Data Pribadi dilakukan sebagai berikut:',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 60, line: LAMPIRAN_LINE_AIRY }
          ),
          lp('a. Proses administrasi layanan pertanahan;', { indentLeft: 720, hanging: 360, spaceAfter: 30, line: LAMPIRAN_LINE_AIRY }),
          lp('b. Pertukaran data dengan Kementerian/Lembaga lainnya.', { indentLeft: 720, hanging: 360, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }),
          lp(
            '3. Saya telah menerima dan memahami informasi yang diberikan oleh Kementerian ATR/BPN sehubungan dengan pemrosesan data pribadi yang saya ungkapkan terkait hal-hal sebagai berikut:',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 60, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            'a. Hak untuk menyampaikan permintaan penundaan, pembatasan, pengakhiran, dan/atau penarikan persetujuan pemrosesan data pribadi;',
            { align: AlignmentType.JUSTIFIED, indentLeft: 720, hanging: 360, spaceAfter: 30, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            'b. Hak untuk menyampaikan permintaan penghapusan dan/atau pemusnahan data pribadi;',
            { align: AlignmentType.JUSTIFIED, indentLeft: 720, hanging: 360, spaceAfter: 30, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            'c. Hak untuk melengkapi, memperbarui, dan/atau memperbaiki data pribadi.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 720, hanging: 360, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            '4. Saya memahami dan menyetujui bahwa Kementerian ATR/BPN akan mempergunakan dan menyimpan Data Pribadi sesuai dengan ketentuan peraturan perundang-undangan yang berlaku atau selama diperlukan.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            '5. Saya memahami dan setuju bahwa Kementerian ATR/BPN dapat mengungkapkan data pribadi saya kepada pihak ketiga lainnya, antara lain',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 60, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            'a. Aparat penegak hukum termasuk Kepolisian, Kejaksaan, Pusat Pelaporan dan Analisis Transaksi Keuangan (PPATK) dan Komisi Pemberantasan Korupsi (KPK);',
            { align: AlignmentType.JUSTIFIED, indentLeft: 720, hanging: 360, spaceAfter: 30, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            'b. Instansi pemerintah/Kementerian/Lembaga, Penyedia jasa Keuangan dan/atau Mitra resmi Kementerian ATR/BPN terkait layanan pertanahan melalui sistem elektronik atau pun sistem elektronik pemerintah yang terintegrasi.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 720, hanging: 360, spaceAfter: 60, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            'selama pengungkapan tersebut dilaksanakan sesuai dengan tujuan pemrosesan di atas dan/atau sesuai dengan ketentuan peraturan perundang-undangan yang berlaku.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            '6. Saya memahami bahwa kesalahan terhadap penyalahgunaan data pribadi oleh pihak ketiga sebagaimana dimaksud pada angka 5, sepenuhnya menjadi tanggung jawab pihak ketiga dan bukan menjadi kesalahan Kementerian ATR/BPN.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 100, line: LAMPIRAN_LINE_AIRY }
          ),
          lp(
            '7. Saya memahami dan menyetujui bahwa saya berkewajiban untuk memastikan Data Pribadi yang saya ungkapkan kepada Kementerian ATR/BPN adalah lengkap, akurat, benar dan terkini. Demikian persetujuan ini saya buat dan tandatangani dengan sebenar-benarnya, secara sadar, dan tanpa tekanan maupun paksaan dari Pihak manapun.',
            { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 160, line: LAMPIRAN_LINE_AIRY }
          ),
          lp('Demikian Surat Keterangan ini dibuat untuk dapat dipergunakan sebagaimana mestinya.', {
            align: AlignmentType.JUSTIFIED,
            firstLine: 720,
            spaceAfter: 360,
            line: LAMPIRAN_LINE_AIRY,
          }),
          lp('{{kota_kabupaten}}, {{tanggal_akta}}', { indentLeft: 5200, spaceAfter: 320, line: LAMPIRAN_LINE_AIRY }),
          lp('{{#pihak_data_pribadi_ttd}}', { indentLeft: 4200, spaceAfter: 0, line: 220 }),
          lp('{{no}}. {{nama}}..........................', { indentLeft: 4200, spaceAfter: 320, line: LAMPIRAN_LINE_AIRY }),
          lp('{{/pihak_data_pribadi_ttd}}', { indentLeft: 4200, spaceAfter: 160, line: 220 }),
          ...makeSaksiSignLines(240),
        ],
      },
    ],
  });
}

export async function generateDefaultDocxBuffer(key: TemplateDocKey): Promise<Uint8Array> {
  let doc: Document;
  switch (key) {
    case '01_akta':
      doc = buildAktaDocument();
      break;
    case '02_lampiran_bundel':
    case '02_permohonan_bpn':
    case '03_kuasa_bpn':
    case '04_tidak_sengketa':
    case '05_sppt_pbb':
    case '06_rincian_pajak':
    case '07_tanda_terima':
    case '08_keabsahan_data_pribadi':
      doc = buildLampiranBundelDocument();
      break;
  }
  const arrayBuffer = await Packer.toArrayBuffer(doc);
  return new Uint8Array(arrayBuffer);
}
