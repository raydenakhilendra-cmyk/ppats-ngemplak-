import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import {
  BerkasRecord,
  OfficeSettings,
  ParaPihak,
  PengesahanMode,
} from '../types/ppats';
import {
  formatNumberId,
  formatTanggalDDMMYYYY,
  formatTanggalLengkapHuruf,
  parseFlexibleDate,
  terbilangLuas,
} from './indonesianFormatter';
import { LAMPIRAN_MARGINS, PAPER_SIZE_F4_TWIPS } from './docxTemplates';

const F4_FONT = 'Bookman Old Style';
const F4_SIZE = 22; // 11 pt
const F4_LINE = 276; // 1.15 spasi
const F4_LINE_AIRY = 300; // 1.25 spasi

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

function fp(
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
      line: opts?.line ?? F4_LINE,
    },
    border: opts?.borderBottom
      ? { bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000' } }
      : undefined,
    children: [
      new TextRun({
        text,
        bold: opts?.bold ?? false,
        italics: opts?.italic ?? false,
        underline: opts?.underline ? {} : undefined,
        size: opts?.size ?? F4_SIZE,
        font: F4_FONT,
      }),
    ],
  });
}

function fpRuns(
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
      line: opts?.line ?? F4_LINE,
    },
    children: runs.map(
      (r) =>
        new TextRun({
          text: r.text,
          bold: r.bold ?? false,
          italics: r.italic ?? false,
          underline: r.underline ? {} : undefined,
          size: r.size ?? F4_SIZE,
          font: F4_FONT,
        })
    ),
  });
}

function kvTable(
  rows: Array<[string, string]>,
  opts?: { labelPct?: number; indentTwips?: number; spaceAfterRow?: number; line?: number }
): Table {
  const labelPct = opts?.labelPct ?? 26;
  const line = opts?.line ?? 255;
  const spaceAfterRow = opts?.spaceAfterRow ?? 0;
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
              children: [fp(k, { spaceAfter: spaceAfterRow, line })],
            }),
            new TableCell({
              width: { size: 100 - labelPct, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [fp(`: ${v}`, { spaceAfter: spaceAfterRow, line })],
            }),
          ],
        })
    ),
  });
}

export interface ResolvedPengesahanTunggal {
  mode: PengesahanMode;
  jabatanLines: string[];
  nama: string;
  nip: string;
}

export interface ResolvedPengesahanGanda {
  camatJabatan: string;
  camatNama: string;
  camatNip: string;
  camatNomor: string;
  kadesJabatan: string;
  kadesNama: string;
  kadesNomor: string;
}

export function resolvePengesahanTunggal(
  record: BerkasRecord,
  settings: OfficeSettings
): ResolvedPengesahanTunggal {
  const na = record.non_akta;
  const mode: PengesahanMode = na?.pengesahan_mode ?? 'ppats';
  const kecTanah = (record.tanah.kecamatan || settings.nama_kecamatan || '')
    .replace(/^Kecamatan\s+/i, '')
    .trim();
  const desaTanah = (record.tanah.desa || '-').trim();

  if (mode === 'kades') {
    const jab = na?.pengesahan_jabatan?.trim() || `Kepala Desa ${desaTanah}`;
    return {
      mode,
      jabatanLines: jab.split('\n').map((s) => s.trim()).filter(Boolean),
      nama: na?.pengesahan_nama?.trim() || na?.kades_pengesahan_nama?.trim() || '..............................',
      nip: na?.pengesahan_nip?.trim() || '',
    };
  }

  if (mode === 'camat') {
    const jab = na?.pengesahan_jabatan?.trim() || `Camat ${kecTanah || '-'}`;
    return {
      mode,
      jabatanLines: jab.split('\n').map((s) => s.trim()).filter(Boolean),
      nama: na?.pengesahan_nama?.trim() || record.akta.nama_ppats || settings.nama_ppats || '-',
      nip: na?.pengesahan_nip?.trim() || record.akta.nip_ppats || settings.nip_ppats || '',
    };
  }

  if (mode === 'manual') {
    const jab = na?.pengesahan_jabatan?.trim() || 'Pejabat Yang Berwenang';
    return {
      mode,
      jabatanLines: jab.split('\n').map((s) => s.trim()).filter(Boolean),
      nama: na?.pengesahan_nama?.trim() || '..............................',
      nip: na?.pengesahan_nip?.trim() || '',
    };
  }

  // Default: PPATS
  const defaultJabatan =
    na?.pengesahan_jabatan?.trim() ||
    `PPATS / Camat ${kecTanah || settings.nama_kecamatan.replace(/^Kecamatan\s+/i, '')}`;
  return {
    mode: 'ppats',
    jabatanLines: defaultJabatan.split('\n').map((s) => s.trim()).filter(Boolean),
    nama: na?.pengesahan_nama?.trim() || record.akta.nama_ppats || settings.nama_ppats || '-',
    nip:
      na?.pengesahan_nip !== undefined
        ? na.pengesahan_nip.trim()
        : record.akta.nip_ppats || settings.nip_ppats || '',
  };
}

export function resolvePengesahanGanda(
  record: BerkasRecord,
  settings: OfficeSettings
): ResolvedPengesahanGanda {
  const na = record.non_akta;
  const kecTanah = (record.tanah.kecamatan || settings.nama_kecamatan || '')
    .replace(/^Kecamatan\s+/i, '')
    .trim();
  const desaTanah = (record.tanah.desa || '-').trim();

  return {
    camatJabatan:
      na?.camat_pengesahan_jabatan?.trim() || `Camat ${kecTanah || '-'}`,
    camatNama:
      na?.camat_pengesahan_nama?.trim() ||
      record.akta.nama_ppats ||
      settings.nama_ppats ||
      '-',
    camatNip:
      na?.camat_pengesahan_nip?.trim() ||
      record.akta.nip_ppats ||
      settings.nip_ppats ||
      '-',
    camatNomor: na?.camat_pengesahan_nomor?.trim() || '..............................',
    kadesJabatan:
      na?.kades_pengesahan_jabatan?.trim() || `Kepala Desa ${desaTanah}`,
    kadesNama: na?.kades_pengesahan_nama?.trim() || '..............................',
    kadesNomor: na?.kades_pengesahan_nomor?.trim() || '..............................',
  };
}

interface NonAktaPartyView {
  no: number;
  nama: string;
  ttl: string;
  pekerjaan: string;
  nik: string;
  alamatLengkap: string;
  alamatSingkat: string;
}

function formatPartyAddress(p: ParaPihak, shortMode: boolean): string {
  const rtRw = (p.rt_rw || '').split('/').map((x) => x.trim());
  const rt = rtRw[0] || '-';
  const rw = rtRw[1] || '-';
  if (shortMode) {
    const dukuh = [p.alamat_jalan, p.rt_rw ? `Rt.${rt}/${rw}` : ''].filter(Boolean).join(' ');
    const desa = p.desa_kelurahan ? `Ds. ${p.desa_kelurahan}` : '';
    const kec = p.kecamatan ? p.kecamatan.replace(/^Kecamatan\s+/i, '') : '';
    const kab = p.kabupaten ? p.kabupaten.replace(/^Kabupaten\s+/i, '') : '';
    return [dukuh, desa, kec, kab].filter(Boolean).join(', ') || '-';
  }
  const parts = [
    p.alamat_jalan,
    p.rt_rw ? `Rukun Tetangga (RT) ${rt} Rukun Warga (RW) ${rw}` : '',
    p.desa_kelurahan ? `Desa ${p.desa_kelurahan}` : '',
    p.kecamatan ? `Kecamatan ${p.kecamatan}` : '',
    p.kabupaten,
  ].filter(Boolean);
  return parts.join(', ') || '-';
}

function collectNonAktaParties(record: BerkasRecord): NonAktaPartyView[] {
  const list = record.para_pihak.filter((p) => p.peran !== 'Persetujuan Pasangan');
  const source = list.length > 0 ? list : record.para_pihak;
  if (source.length === 0) {
    return [
      {
        no: 1,
        nama: '-',
        ttl: '-',
        pekerjaan: '-',
        nik: '-',
        alamatLengkap: '-',
        alamatSingkat: '-',
      },
    ];
  }
  return source.map((p, idx) => {
    const tgl = formatTanggalDDMMYYYY(p.tanggal_lahir);
    const ttl = [p.tempat_lahir, tgl !== '-' ? tgl : p.tanggal_lahir]
      .filter(Boolean)
      .join(', ') || '-';
    return {
      no: idx + 1,
      nama: p.nama || '-',
      ttl,
      pekerjaan: p.pekerjaan || '-',
      nik: p.nik || '-',
      alamatLengkap: formatPartyAddress(p, false),
      alamatSingkat: formatPartyAddress(p, true),
    };
  });
}

function resolvePersonCtx(
  name: string,
  list?: OfficeSettings['profil_saksi'],
  override?: BerkasRecord['akta']['saksi_1_detail']
) {
  const fromList = (list ?? []).find(
    (item) => item.nama.trim().toLowerCase() === name.trim().toLowerCase()
  );
  const merged = fromList || override ? { ...(fromList ?? { nama: name }), ...(override ?? {}) } : undefined;
  const nama = merged?.nama?.trim() || name || '-';
  const ttl =
    [merged?.tempat_lahir, formatTanggalDDMMYYYY(merged?.tanggal_lahir)]
      .filter((x) => x && x !== '-')
      .join(', ') || '-';
  const rtRw =
    merged?.rt || merged?.rw ? `Rt.${merged?.rt || '-'}/${merged?.rw || '-'}` : '';
  const alamatSingkat =
    [
      [merged?.alamat_jalan, rtRw].filter(Boolean).join(' '),
      merged?.desa_kelurahan ? `Ds. ${merged.desa_kelurahan}` : '',
      merged?.kecamatan ? merged.kecamatan.replace(/^Kecamatan\s+/i, '') : '',
      merged?.kabupaten ? merged.kabupaten.replace(/^Kabupaten\s+/i, '') : '',
    ]
      .filter(Boolean)
      .join(', ') ||
    merged?.alamat ||
    '-';

  return {
    nama,
    ttl,
    pekerjaan: merged?.pekerjaan || '-',
    nik: merged?.nik || '-',
    alamat: merged?.alamat || alamatSingkat,
    alamatSingkat,
  };
}

function buildSaksiBlocks(
  saksi1Name: string,
  saksi2Name: string,
  spaceBefore = 180
): Array<Paragraph | Table> {
  return [
    fp('Saksi – saksi :', { spaceBefore, spaceAfter: 120, line: F4_LINE_AIRY }),
    new Table({
      width: { size: 75, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 62, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [fp(`1. ${saksi1Name}`, { spaceAfter: 160, line: F4_LINE_AIRY })],
            }),
            new TableCell({
              width: { size: 38, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [fp('(                    )', { spaceAfter: 160, line: F4_LINE_AIRY })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              width: { size: 62, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [fp(`2. ${saksi2Name}`, { spaceAfter: 60, line: F4_LINE_AIRY })],
            }),
            new TableCell({
              width: { size: 38, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [fp('(                    )', { spaceAfter: 60, line: F4_LINE_AIRY })],
            }),
          ],
        }),
      ],
    }),
  ];
}

function buildSinglePengesahanCentered(pengesahan: ResolvedPengesahanTunggal, spaceBefore = 140): Paragraph[] {
  const out: Paragraph[] = [
    fp('Mengetahui,', { align: AlignmentType.CENTER, spaceBefore, spaceAfter: 0, line: 255 }),
    ...pengesahan.jabatanLines.map((lineText, idx) =>
      fp(lineText, {
        align: AlignmentType.CENTER,
        spaceAfter: idx === pengesahan.jabatanLines.length - 1 ? 820 : 0,
        line: 255,
      })
    ),
    fp(pengesahan.nama, {
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 0,
      line: 255,
    }),
  ];
  if (pengesahan.nip && pengesahan.nip !== '-') {
    out.push(fp(`NIP. ${pengesahan.nip}`, { align: AlignmentType.CENTER, spaceAfter: 20, line: 255 }));
  }
  return out;
}

function buildDualPengesahanTable(
  ganda: ResolvedPengesahanGanda,
  kotaTanggalLine: string,
  opts?: { withNumbers?: boolean; mode?: 'skw' | 'beda_nama' }
): Table {
  const withNumbers = opts?.withNumbers ?? false;
  const camatHeader =
    opts?.mode === 'skw'
      ? 'Dikuatkan oleh kami :'
      : 'Dikuatkan oleh';
  const kadesHeader =
    opts?.mode === 'skw'
      ? 'Disaksikan dan dibenarkan oleh kami:'
      : 'Disaksikan dan Dibenarkan';

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: NO_BORDER,
            children: [
              ...(withNumbers
                ? [fp(`Nomor : ${ganda.camatNomor}`, { spaceAfter: 10, line: 250 })]
                : []),
              fp(camatHeader, { align: AlignmentType.CENTER, spaceAfter: 0, line: 250 }),
              fp(ganda.camatJabatan, {
                align: AlignmentType.CENTER,
                spaceAfter: 820,
                line: 250,
              }),
              fp(ganda.camatNama, {
                bold: true,
                underline: true,
                align: AlignmentType.CENTER,
                spaceAfter: 0,
                line: 250,
              }),
              ...(ganda.camatNip && ganda.camatNip !== '-'
                ? [
                    fp(`NIP. ${ganda.camatNip}`, {
                      align: AlignmentType.CENTER,
                      spaceAfter: 20,
                      line: 250,
                    }),
                  ]
                : []),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: NO_BORDER,
            children: [
              ...(withNumbers
                ? [fp(`Nomor : ${ganda.kadesNomor}`, { spaceAfter: 10, line: 250 })]
                : []),
              fp(kotaTanggalLine, { align: AlignmentType.CENTER, spaceAfter: 0, line: 250 }),
              fp(kadesHeader, { align: AlignmentType.CENTER, spaceAfter: 0, line: 250 }),
              fp(ganda.kadesJabatan, {
                align: AlignmentType.CENTER,
                spaceAfter: 820,
                line: 250,
              }),
              fp(ganda.kadesNama, {
                bold: true,
                underline: true,
                align: AlignmentType.CENTER,
                spaceAfter: 20,
                line: 250,
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function buildNumberedPartiesKvBlocks(
  parties: NonAktaPartyView[],
  opts?: { shortAddress?: boolean; spaceAfterRow?: number }
): Array<Paragraph | Table> {
  const out: Array<Paragraph | Table> = [];
  if (parties.length === 1) {
    const pt = parties[0];
    out.push(
      kvTable(
        [
          ['Nama', pt.nama],
          ['Tempat/Tgl.lahir', pt.ttl],
          ['Pekerjaan', pt.pekerjaan],
          ['Nomor KTP', pt.nik],
          ['Alamat', opts?.shortAddress ? pt.alamatSingkat : pt.alamatLengkap],
        ],
        { indentTwips: 420, labelPct: 26, spaceAfterRow: opts?.spaceAfterRow ?? 0, line: 250 }
      )
    );
    return out;
  }

  for (const pt of parties) {
    out.push(
      kvTable(
        [
          [`${pt.no}. Nama`, pt.nama],
          ['   Tempat/Tgl.lahir', pt.ttl],
          ['   Pekerjaan', pt.pekerjaan],
          ['   Nomor KTP', pt.nik],
          ['   Alamat', opts?.shortAddress ? pt.alamatSingkat : pt.alamatLengkap],
        ],
        { indentTwips: 240, labelPct: 28, spaceAfterRow: opts?.spaceAfterRow ?? 0, line: 245 }
      )
    );
    out.push(fp('', { spaceAfter: 20, line: 160 }));
  }
  return out;
}

function buildNumberedSignerLines(parties: NonAktaPartyView[], indentLeft = 0): Paragraph[] {
  if (parties.length === 1) {
    return [fp(parties[0].nama, { align: AlignmentType.CENTER, spaceAfter: 20, line: 255 })];
  }
  return parties.map((pt) =>
    fp(`${pt.no}. ${pt.nama} (....................)`, {
      indentLeft,
      spaceAfter: 180,
      line: 260,
    })
  );
}

function resolveNonAktaDocDate(record: BerkasRecord): {
  tanggalText: string;
  tanggalDash: string;
  hariTanggal: string;
} {
  const parsed = parseFlexibleDate(record.akta.tanggal_akta);
  if (!parsed) {
    return {
      tanggalText: '......................',
      tanggalDash: '',
      hariTanggal: '............ tanggal ......................',
    };
  }
  const fmt = formatTanggalLengkapHuruf(parsed.ddMmYyyy);
  return {
    tanggalText: fmt.tanggal_biasa,
    tanggalDash: fmt.tanggal_dash,
    hariTanggal: `${fmt.hari} tanggal ${fmt.tanggal_biasa}`,
  };
}

function getSertipikatInfo(record: BerkasRecord) {
  const t = record.tanah;
  const isDigital =
    t.jenis_sertipikat === 'digital' ||
    (!t.nomor_hak_lama?.trim() && !t.surat_ukur_nomor?.trim());
  const nibClean = t.nib?.trim() || '-';
  const noHakClean = (t.nomor_hak_lama || '-')
    .replace(/^(?:SHM|SHGB|SHP|HP|Sertipikat)\s*(?:No\.?|Nomor)\s*[:.]?\s*/i, '')
    .trim();
  const labelNo = isDigital ? 'Nibel' : 'Nomor Hak / NIB';
  const nilaiNo = isDigital ? nibClean : `${noHakClean} (NIB : ${nibClean})`;
  const ringkas = isDigital
    ? `Nibel : ${nibClean}`
    : `Nomor : ${noHakClean} (NIB : ${nibClean})`;
  const suTgl = t.surat_ukur_tanggal?.trim()
    ? /^\d{4}-\d{2}-\d{2}$/.test(t.surat_ukur_tanggal.trim())
      ? formatTanggalDDMMYYYY(t.surat_ukur_tanggal)
      : t.surat_ukur_tanggal.trim()
    : '-';
  const suInfo = isDigital
    ? ''
    : t.surat_ukur_nomor?.trim()
    ? ` Surat Ukur tanggal ${suTgl} Nomor ${t.surat_ukur_nomor}`
    : '';
  return {
    isDigital,
    jenisHak: t.jenis_hak || 'Hak Milik',
    nibClean,
    noHakClean,
    labelNo,
    nilaiNo,
    ringkas,
    suInfo,
    luasStr: formatNumberId(t.luas_m2),
    luasTerbilang: terbilangLuas(t.luas_m2),
    desa: t.desa || '-',
    kecamatan: t.kecamatan || '-',
    kabupaten: t.kabupaten || '-',
    provinsi: t.provinsi || 'Jawa Tengah',
    pemegangHak:
      t.pemegang_hak_sertipikat?.filter(Boolean).join(', ') ||
      record.non_akta?.nama_pewaris ||
      '-',
    penggunaan: t.penggunaan_tanah || 'Pekarangan',
    sejakTahun: t.menguasai_sejak_tahun || '2020',
  };
}

function buildPermohonanPage(params: {
  record: BerkasRecord;
  parties: NonAktaPartyView[];
  kuasa: ReturnType<typeof resolvePersonCtx>;
  item14Text: string;
  boldIndices: number[];
  lampiranItems: string[];
  kotaKabupaten: string;
  tanggalText: string;
  tanggalDash?: string;
  pageBreakBefore?: boolean;
}): Array<Paragraph | Table> {
  const sert = getSertipikatInfo(params.record);
  const qqSignatureLine =
    params.parties.length === 1
      ? `${params.kuasa.nama} QQ ${params.parties[0].nama}`
      : params.parties.length <= 3
        ? `${params.kuasa.nama} qq. ${params.parties.map((p) => `${p.no}. ${p.nama}`).join(', ')}`
        : `${params.kuasa.nama} qq. ${params.parties[0].nama}, dkk (${params.parties.length} orang)`;

  const leftPermohonan: Array<[number, string]> = [
    [1, 'Pengukuran'],
    [2, 'Konversi / Pendaftaran Hak'],
    [3, 'Pendaftaran Hak Milik Sarusun'],
    [4, 'Pendaftaran Tanah Wakaf'],
    [5, 'Pendaftaran Peralihan Hak'],
    [6, 'Pendaftaran Pemindahan Hak'],
    [7, 'Pendaftaran Perubahan Hak'],
  ];
  const rightPermohonan: Array<[number, string]> = [
    [8, 'Pemecahan/Penggabungan Hak'],
    [9, 'Pendaftaran Hak Tanggungan'],
    [10, 'Roya atas Hak Tanggungan'],
    [11, 'Penerbitan Sertipikat Pengganti'],
    [12, 'Surat Keterangan Pendaftaran Tanah'],
    [13, 'Pengecekan Sertipikat'],
    [14, params.item14Text],
  ];

  const leftCount =
    params.lampiranItems.length <= 6
      ? Math.min(params.lampiranItems.length, 4)
      : Math.ceil(params.lampiranItems.length / 2);
  const lampLeft = params.lampiranItems.slice(0, leftCount);
  const lampRight = params.lampiranItems.slice(leftCount);

  return [
    fp('Kepada Yth.', {
      pageBreakBefore: params.pageBreakBefore ?? false,
      indentLeft: 4400,
      spaceAfter: 0,
      line: 245,
    }),
    fp('Kantor Kementerian Agraria dan Tata Ruang/', {
      indentLeft: 4400,
      spaceAfter: 0,
      line: 245,
    }),
    fp(`Badan Pertanahan Nasional ${sert.kabupaten}`, {
      indentLeft: 4400,
      spaceAfter: 0,
      line: 245,
    }),
    fp(`di ${params.kotaKabupaten}`, {
      indentLeft: 4400,
      spaceAfter: 160,
      line: 245,
    }),

    fp('Dengan Hormat,', { spaceAfter: 20, line: 245 }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 40, line: 245 }),
    kvTable(
      [
        ['Nama', params.kuasa.nama],
        ['Tempat/Tgl.lahir', params.kuasa.ttl],
        ['Pekerjaan', params.kuasa.pekerjaan],
        ['Nomor KTP', params.kuasa.nik],
        ['Alamat', params.kuasa.alamatSingkat],
      ],
      { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 240 }
    ),

    fp('Dalam hal ini bertindak untuk dan atas nama diri sendiri / selaku kuasa dari :', {
      spaceBefore: 100,
      spaceAfter: 40,
      line: 245,
    }),
    ...buildNumberedPartiesKvBlocks(params.parties, { shortAddress: true, spaceAfterRow: 0 }),

    fp(
      `Berdasarkan surat kuasa   Nomor : ...................   Tanggal  :${params.tanggalDash ? ` ${params.tanggalDash}` : ''}`,
      {
        spaceBefore: 80,
        spaceAfter: 20,
        line: 245,
      }
    ),
    fp('Dengan ini mengajukan permohonan :', { spaceAfter: 30, line: 245 }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      indent: { size: 420, type: WidthType.DXA },
      rows: leftPermohonan.map(([lNum, lText], idx) => {
        const [rNum, rText] = rightPermohonan[idx];
        const lBold = params.boldIndices.includes(lNum);
        const rBold = params.boldIndices.includes(rNum);
        return new TableRow({
          children: [
            new TableCell({
              width: { size: 48, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${lNum}. ${lText}`, {
                  bold: lBold,
                  spaceAfter: 10,
                  line: 240,
                }),
              ],
            }),
            new TableCell({
              width: { size: 52, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${rNum}. ${rText}`, {
                  bold: rBold,
                  spaceAfter: 10,
                  line: 240,
                }),
              ],
            }),
          ],
        });
      }),
    }),

    fp('Atas bidang tanah Hak / Tanah Negara :', {
      spaceBefore: 100,
      spaceAfter: 10,
      line: 245,
    }),
    fp('Terletak di', { indentLeft: 420, spaceAfter: 10, line: 245 }),
    kvTable(
      [
        ['Desa', sert.desa],
        ['Kecamatan', sert.kecamatan],
        ['Kabupaten', params.kotaKabupaten],
        [
          sert.isDigital ? 'Nibel' : `Nomor ${sert.jenisHak} No.`,
          sert.isDigital ? sert.nibClean : sert.noHakClean,
        ],
      ],
      { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 240 }
    ),

    fp('Untuk melengkapi permohonan dimaksud, bersama ini kami lampirkan ;', {
      spaceBefore: 100,
      spaceAfter: 30,
      line: 245,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      indent: { size: 420, type: WidthType.DXA },
      rows: lampLeft.map((leftText, rowIdx) => {
        const rightText = lampRight[rowIdx];
        const rightNum = leftCount + rowIdx + 1;
        return new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${rowIdx + 1}. ${leftText}`, {
                  spaceAfter: 10,
                  line: 240,
                }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(rightText ? `${rightNum}. ${rightText}` : '', {
                  spaceAfter: 10,
                  line: 240,
                }),
              ],
            }),
          ],
        });
      }),
    }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [fp('', { spaceAfter: 0 })],
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Hormat Kami,', {
                  align: AlignmentType.CENTER,
                  spaceBefore: 120,
                  spaceAfter: 760,
                  line: 245,
                }),
                fp(qqSignatureLine, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 20,
                  line: 245,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

function buildSuratKuasaNonAkta(params: {
  record: BerkasRecord;
  parties: NonAktaPartyView[];
  kuasa: ReturnType<typeof resolvePersonCtx>;
  pengesahan: ResolvedPengesahanTunggal;
  kotaKabupaten: string;
  tanggalText: string;
  khususBullets: string[];
  objekPrefix: string;
}): Array<Paragraph | Table> {
  const sert = getSertipikatInfo(params.record);
  return [
    fp('SURAT KUASA', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 160,
      line: 255,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 40, line: 250 }),
    ...buildNumberedPartiesKvBlocks(params.parties, { spaceAfterRow: 0 }),
    fp(`${params.objekPrefix} ${sert.jenisHak} :`, { spaceBefore: 100, spaceAfter: 40, line: 250 }),
    kvTable(
      [
        [sert.labelNo, sert.nilaiNo],
        ['Desa', sert.desa],
        ['Kecamatan', sert.kecamatan],
        ['Kabupaten', sert.kabupaten],
        ['Pemegang Hak Tercatat', sert.pemegangHak],
      ],
      { indentTwips: 540, labelPct: 28, spaceAfterRow: 0, line: 245 }
    ),
    fp('Maka dengan ini memberi kuasa kepada :', { spaceBefore: 100, spaceAfter: 40, line: 250 }),
    kvTable(
      [
        ['Nama', params.kuasa.nama],
        ['Tempat/Tgl.lahir', params.kuasa.ttl],
        ['Pekerjaan', params.kuasa.pekerjaan],
        ['Nomor KTP', params.kuasa.nik],
        ['Alamat', params.kuasa.alamatSingkat],
      ],
      { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 245 }
    ),
    fpRuns(
      [
        { text: '----------------------------------------------' },
        { text: 'KHUSUS', bold: true, underline: true },
        { text: '---------------------------------------------' },
      ],
      { align: AlignmentType.CENTER, spaceBefore: 120, spaceAfter: 80, line: 250 }
    ),
    ...params.khususBullets.map((b) =>
      fp(`-  ${b}`, {
        align: AlignmentType.JUSTIFIED,
        indentLeft: 320,
        hanging: 220,
        spaceAfter: 40,
        line: 260,
      })
    ),
    fp('Demikian kuasa ini diberikan untuk dapat dipergunakan sebagaimana mestinya.', {
      spaceBefore: 60,
      spaceAfter: 140,
      line: 255,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 44, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Penerima Kuasa', { align: AlignmentType.CENTER, spaceAfter: 820, line: 250 }),
                fp(params.kuasa.nama, { align: AlignmentType.CENTER, spaceAfter: 20, line: 250 }),
              ],
            }),
            new TableCell({
              width: { size: 56, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${params.kotaKabupaten}, ${params.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp(params.parties.length > 1 ? 'Para Pemberi Kuasa' : 'Pemberi Kuasa', {
                  align: AlignmentType.CENTER,
                  spaceAfter: params.parties.length > 1 ? 120 : 820,
                  line: 250,
                }),
                ...buildNumberedSignerLines(params.parties, 420),
              ],
            }),
          ],
        }),
      ],
    }),
    ...buildSinglePengesahanCentered(params.pengesahan, 140),
  ];
}

function buildPenguasaanFisikNonAkta(params: {
  record: BerkasRecord;
  parties: NonAktaPartyView[];
  saksi1: ReturnType<typeof resolvePersonCtx>;
  saksi2: ReturnType<typeof resolvePersonCtx>;
  pengesahan: ResolvedPengesahanTunggal;
  hariTanggal: string;
  butir1Text: string;
  catatanAlias?: string;
}): Array<Paragraph | Table> {
  const sert = getSertipikatInfo(params.record);
  const sayaAtauKami = params.parties.length > 1 ? 'kami' : 'saya';
  return [
    fp('SURAT PERNYATAAN PENGUASAAN FISIK TANAH DAN TIDAK SENGKETA', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 160,
      line: 255,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 40, line: 245 }),
    ...buildNumberedPartiesKvBlocks(params.parties, { spaceAfterRow: 0 }),
    fp('Dengan disaksikan oleh :', { spaceBefore: 80, spaceAfter: 40, line: 245 }),
    kvTable(
      [
        ['1. Nama', params.saksi1.nama],
        ['   Tempat/Tgl.lahir', params.saksi1.ttl],
        ['   Pekerjaan', params.saksi1.pekerjaan],
        ['   Nomor KTP', params.saksi1.nik],
        ['   Alamat', params.saksi1.alamat],
      ],
      { indentTwips: 240, labelPct: 27, spaceAfterRow: 0, line: 240 }
    ),
    fp('', { spaceAfter: 20, line: 150 }),
    kvTable(
      [
        ['2. Nama', params.saksi2.nama],
        ['   Tempat/Tgl.lahir', params.saksi2.ttl],
        ['   Pekerjaan', params.saksi2.pekerjaan],
        ['   Nomor KTP', params.saksi2.nik],
        ['   Alamat', params.saksi2.alamat],
      ],
      { indentTwips: 240, labelPct: 27, spaceAfterRow: 0, line: 240 }
    ),
    fp(`Pada hari ini ${params.hariTanggal} dengan ini menyatakan :`, {
      spaceBefore: 100,
      spaceAfter: 50,
      line: 250,
    }),
    fp(`1. ${params.butir1Text}`, {
      align: AlignmentType.JUSTIFIED,
      indentLeft: 320,
      hanging: 320,
      spaceAfter: 40,
      line: 255,
    }),
    fp(
      `2. Bahwa ${sayaAtauKami} benar-benar menguasai tanah secara fisik sejak tahun ${sert.sejakTahun} sampai sekarang dan dipergunakan untuk ${sert.penggunaan}.${params.catatanAlias ? ` (${params.catatanAlias})` : ''}`,
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 255 }
    ),
    fp(
      '3. Bahwa tanah tersebut tidak dalam keadaan sengketa, konflik, perkara, sita jaminan atau pemblokiran dari pihak lain.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 60, line: 255 }
    ),
    fp(
      `Demikian surat pernyataan ini ${sayaAtauKami} buat dengan sebenarnya dan ${sayaAtauKami} bertanggung jawab sepenuhnya atas pernyataan ini. Apabila dikemudian hari terjadi perkara perdata maupun pidana atas tanah ini, maka perkara tersebut merupakan tanggung jawab pribadi ${sayaAtauKami} dengan tidak melibatkan Pejabat/Kantor Badan Pertanahan Nasional ${sert.kabupaten}.`,
      { align: AlignmentType.JUSTIFIED, spaceAfter: 120, line: 255 }
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
                fp('Saksi – saksi :', { spaceAfter: 120, line: 245 }),
                fp(`1. ${params.saksi1.nama} (...............)`, { spaceAfter: 260, line: 245 }),
                fp(`2. ${params.saksi2.nama} (...............)`, { spaceAfter: 40, line: 245 }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Yang membuat pernyataan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: params.parties.length > 1 ? 120 : 720,
                  line: 245,
                }),
                ...buildNumberedSignerLines(params.parties, 360),
              ],
            }),
          ],
        }),
      ],
    }),
    ...buildSinglePengesahanCentered(params.pengesahan, 100),
  ];
}

function buildDataPribadiPage(params: {
  parties: NonAktaPartyView[];
  kotaKabupaten: string;
  tanggalText: string;
}): Array<Paragraph | Table> {
  return [
    fp('SURAT PERNYATAAN PERSETUJUAN PEMROSESAN DATA PRIBADI', {
      pageBreakBefore: true,
      bold: true,
      align: AlignmentType.CENTER,
      spaceAfter: 160,
      line: 270,
    }),
    fp(
      'Saya, yang bertandatangan di bawah ini, selaku pengguna layanan Pertanahan Kementerian Agraria dan Tata Ruang/Badan Pertanahan Nasional (“Kementerian ATR/BPN”), dengan ini menyatakan sebagai berikut:',
      { align: AlignmentType.JUSTIFIED, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fpRuns(
      [
        { text: '1. ', bold: true, italic: true },
        { text: 'Bahwa saya bertindak dalam kapasitas sebagai ' },
        {
          text: 'Diri sendiri (perorangan), saya menyatakan memiliki kewenangan hukum yang sah dan bertanggung jawab sepenuhnya atas kewenangan tersebut, serta membebaskan Kementerian ATR/BPN dari segala tuntutan atau klaim apabila di kemudian hari kewenangan tersebut terbukti tidak sah.',
          italic: true,
        },
        {
          text: ' Dengan ini saya memberikan persetujuan kepada Kementerian ATR/BPN selaku Pengendali Data Pribadi untuk melakukan pemrosesan Data Pribadi saya dan/atau pihak yang saya wakili sesuai ketentuan peraturan perundang-undangan.',
        },
      ],
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '2. Saya telah menerima dan memahami informasi yang diberikan oleh Kementerian ATR/BPN sehubungan dengan tujuan pemrosesan Data Pribadi dilakukan untuk proses administrasi layanan pertanahan dan pertukaran data dengan Kementerian/Lembaga lainnya.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '3. Saya memahami hak saya sebagai subjek data pribadi untuk menyampaikan permintaan penundaan, pembatasan, pengakhiran, penarikan persetujuan, penghapusan, maupun pembaruan data pribadi sesuai peraturan yang berlaku.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '4. Saya memahami dan menyetujui bahwa Kementerian ATR/BPN akan mempergunakan dan menyimpan Data Pribadi sesuai ketentuan peraturan perundang-undangan atau selama diperlukan.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '5. Saya memahami dan setuju bahwa Kementerian ATR/BPN dapat mengungkapkan data pribadi kepada aparat penegak hukum maupun instansi pemerintah/mitra resmi terkait layanan pertanahan melalui sistem elektronik yang terintegrasi.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '6. Saya memahami bahwa kesalahan terhadap penyalahgunaan data pribadi oleh pihak ketiga sepenuhnya menjadi tanggung jawab pihak ketiga dan bukan menjadi kesalahan Kementerian ATR/BPN.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '7. Saya memahami dan menyetujui bahwa saya berkewajiban untuk memastikan Data Pribadi yang saya ungkapkan kepada Kementerian ATR/BPN adalah lengkap, akurat, benar dan terkini. Demikian persetujuan ini dibuat dengan sebenar-benarnya tanpa paksaan dari pihak manapun.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 360, hanging: 360, spaceAfter: 200, line: F4_LINE_AIRY }
    ),
    fp(`${params.kotaKabupaten}, ${params.tanggalText}`, {
      indentLeft: 4800,
      spaceAfter: 240,
      line: F4_LINE_AIRY,
    }),
    ...params.parties.map((pt) =>
      fp(`${pt.no}. ${pt.nama} ..........................`, {
        indentLeft: 4200,
        spaceAfter: 260,
        line: F4_LINE_AIRY,
      })
    ),
  ];
}

function buildBedaNamaPage(params: {
  record: BerkasRecord;
  parties: NonAktaPartyView[];
  saksi1: ReturnType<typeof resolvePersonCtx>;
  saksi2: ReturnType<typeof resolvePersonCtx>;
  ganda: ResolvedPengesahanGanda;
  kotaKabupaten: string;
  tanggalText: string;
}): Array<Paragraph | Table> {
  const sert = getSertipikatInfo(params.record);
  const na = params.record.non_akta;
  const namaSertipikat =
    na?.nama_di_sertipikat?.trim() ||
    sert.pemegangHak ||
    params.parties[0]?.nama ||
    '-';
  const namaKtpKk =
    na?.nama_di_ktp_kk?.trim() || params.parties[0]?.nama || '-';
  const nomorKk = na?.nomor_kk?.trim() || '-';
  const dokLain = na?.jenis_dokumen_lain?.trim() || 'Kutipan Akta Nikah / Dokumen Pendukung';
  const namaDokLain = na?.nama_di_dokumen_lain?.trim() || namaSertipikat;
  const keperluan = na?.keperluan_ralat?.trim() || 'Ralat Data / Penyesuaian Data Sertipikat';
  const pemohon = params.parties[0];

  return [
    fp('SURAT PERNYATAAN BEDA NAMA', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 200,
      line: 260,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: 250 }),
    kvTable(
      [
        ['Nama', pemohon.nama],
        ['Tempat/Tgl. Lahir', pemohon.ttl],
        ['Pekerjaan', pemohon.pekerjaan],
        ['Nomor KTP', pemohon.nik],
        ['Nomor KK', nomorKk],
        ['Alamat', pemohon.alamatLengkap],
      ],
      { indentTwips: 420, labelPct: 28, spaceAfterRow: 0, line: 250 }
    ),
    fp('Dengan ini menerangkan bahwa :', { spaceBefore: 140, spaceAfter: 60, line: 250 }),
    fp(
      `1. Nama yang tercantum dalam Sertipikat ${sert.jenisHak} ${sert.ringkas} Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten} tertulis : ${namaSertipikat}`,
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 260 }
    ),
    fp(
      `2. Nama yang tercantum dalam KTP dan Kartu Keluarga (No. ${nomorKk}) tertulis : ${namaKtpKk}`,
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 260 }
    ),
    fp(
      `3. Nama yang tercantum dalam ${dokLain} tertulis : ${namaDokLain}`,
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 80, line: 260 }
    ),
    fp(
      `Bahwa nama-nama tersebut diatas adalah benar-benar SATU ORANG YANG SAMA. Adapun untuk selanjutnya, penulisan nama baku yang benar dan berlaku sah adalah ${namaKtpKk} sesuai KTP dan Kartu Keluarga yang berlaku, guna keperluan ${keperluan}.`,
      { align: AlignmentType.JUSTIFIED, spaceAfter: 80, line: 260 }
    ),
    fp(
      'Demikian Surat Pernyataan Beda Nama ini dibuat dengan sebenarnya dan apabila dikemudian hari terdapat permasalahan hukum menjadi tanggung jawab pribadi yang menyatakan sepenuhnya.',
      { align: AlignmentType.JUSTIFIED, spaceAfter: 160, line: 260 }
    ),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 52, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Saksi – saksi :', { spaceAfter: 120, line: 250 }),
                fp(`1. ${params.saksi1.nama} (...............)`, { spaceAfter: 260, line: 250 }),
                fp(`2. ${params.saksi2.nama} (...............)`, { spaceAfter: 20, line: 250 }),
              ],
            }),
            new TableCell({
              width: { size: 48, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${params.kotaKabupaten}, ${params.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp('Yang membuat pernyataan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: 720,
                  line: 250,
                }),
                fp(pemohon.nama, {
                  bold: true,
                  underline: true,
                  align: AlignmentType.CENTER,
                  spaceAfter: 20,
                  line: 250,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    fp('', { spaceAfter: 80, line: 160 }),
    buildDualPengesahanTable(
      params.ganda,
      `${sert.desa}, Tanggal ${params.tanggalText}`,
      { withNumbers: true, mode: 'beda_nama' }
    ),
  ];
}

function buildTurunWarisBundle(
  record: BerkasRecord,
  settings: OfficeSettings
): Array<Paragraph | Table> {
  const parties = collectNonAktaParties(record);
  const sert = getSertipikatInfo(record);
  const na = record.non_akta;
  const pengesahanTunggal = resolvePengesahanTunggal(record, settings);
  const ganda = resolvePengesahanGanda(record, settings);

  const saksi1 = resolvePersonCtx(record.akta.saksi_1, settings.profil_saksi, record.akta.saksi_1_detail);
  const saksi2 = resolvePersonCtx(record.akta.saksi_2, settings.profil_saksi, record.akta.saksi_2_detail);
  const kuasa = resolvePersonCtx(
    record.akta.nama_kuasa_bpn,
    settings.profil_kuasa,
    record.akta.kuasa_bpn_detail
  );

  const docDate = resolveNonAktaDocDate(record);
  const kotaKabupaten =
    (record.tanah.kabupaten || settings.nama_kabupaten || '')
      .replace(/^(Kabupaten|Kota)\s+/i, '')
      .trim() || '-';

  const sebutanPewaris = na?.sebutan_pewaris || 'Almarhum';
  const namaPewaris =
    na?.nama_pewaris?.trim() || sert.pemegangHak || '..............................';
  const parsedMeninggal = parseFlexibleDate(na?.tanggal_meninggal);
  const tglMeninggal = parsedMeninggal
    ? formatTanggalLengkapHuruf(parsedMeninggal.ddMmYyyy).tanggal_biasa
    : na?.tanggal_meninggal?.trim() || '..............................';
  const tempatMeninggal =
    na?.tempat_meninggal?.trim() || `Desa ${sert.desa}, Kecamatan ${sert.kecamatan}`;
  const namaPasangan = na?.nama_pasangan_pewaris?.trim();
  const hubPasangan = na?.hubungan_pasangan_pewaris || 'Istrinya';
  const tglLahirPasangan = na?.tgl_lahir_pasangan_pewaris?.trim()
    ? formatTanggalDDMMYYYY(na.tgl_lahir_pasangan_pewaris)
    : '-';
  const jumlahAnak = na?.jumlah_anak_waris ?? parties.length;
  const nopClean = record.pbb.nop?.trim() || '-';

  const ahliWarisListRows = parties.map((pt, idx) =>
    fp(`${idx + 1}. ${pt.nama} (${pt.ttl}) — NIK: ${pt.nik}`, {
      indentLeft: 420,
      spaceAfter: 20,
      line: 250,
    })
  );

  return [
    // HALAMAN 1: SURAT KETERANGAN WARISAN (SKW) — PENGESAHAN GANDA CAMAT & KADES ASAL SERTIPIKAT
    fp('SURAT KETERANGAN WARISAN', {
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 200,
      line: 260,
    }),
    fp(
      `Kami yang bertanda tangan dibawah ini, para ahli waris dari ${sebutanPewaris} ${namaPewaris}, menerangkan dengan sebenarnya dan sanggup diangkat sumpah bahwa ${sebutanPewaris} ${namaPewaris} telah meninggal dunia di ${tempatMeninggal} pada tanggal ${tglMeninggal}.`,
      { align: AlignmentType.JUSTIFIED, firstLine: 560, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    ...(namaPasangan
      ? [
          fp(
            `Semasa hidupnya almarhum/almarhumah menikah sah dengan ${hubPasangan} bernama ${namaPasangan} (lahir tanggal ${tglLahirPasangan}) dan dikaruniai ${jumlahAnak} orang anak / ahli waris yaitu :`,
            { align: AlignmentType.JUSTIFIED, firstLine: 560, spaceAfter: 60, line: F4_LINE_AIRY }
          ),
        ]
      : [
          fp(
            `Adapun para ahli waris yang sah menurut hukum dari ${sebutanPewaris} ${namaPewaris} berjumlah ${jumlahAnak} orang yaitu :`,
            { align: AlignmentType.JUSTIFIED, firstLine: 560, spaceAfter: 60, line: F4_LINE_AIRY }
          ),
        ]),
    ...ahliWarisListRows,
    ...(na?.catatan_skw?.trim()
      ? [
          fp(na.catatan_skw.trim(), {
            align: AlignmentType.JUSTIFIED,
            spaceBefore: 60,
            spaceAfter: 60,
            line: F4_LINE_AIRY,
          }),
        ]
      : []),
    fp(
      `Dengan demikian maka kami tersebut diatas adalah sebagai satu-satunya ahli waris yang sah dari ${sebutanPewaris} ${namaPewaris}, termasuk atas sebidang tanah ${sert.jenisHak} ${sert.ringkas} luas ${sert.luasStr} M² terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}.`,
      { align: AlignmentType.JUSTIFIED, spaceBefore: 60, spaceAfter: 100, line: F4_LINE_AIRY }
    ),
    fp('Demikian Surat Keterangan Warisan ini dibuat untuk dapat dipergunakan sebagaimana mestinya.', {
      align: AlignmentType.JUSTIFIED,
      spaceAfter: 140,
      line: F4_LINE_AIRY,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Saksi – saksi :', { spaceAfter: 100, line: 245 }),
                fp(`1. ${saksi1.nama} (...............)`, { spaceAfter: 240, line: 245 }),
                fp(`2. ${saksi2.nama} (...............)`, { spaceAfter: 20, line: 245 }),
              ],
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Para Ahli Waris :', { spaceAfter: 100, line: 245 }),
                ...parties.map((pt) =>
                  fp(`${pt.no}. ${pt.nama} (....................)`, {
                    spaceAfter: 160,
                    line: 245,
                  })
                ),
              ],
            }),
          ],
        }),
      ],
    }),
    fp('', { spaceAfter: 100, line: 160 }),
    buildDualPengesahanTable(ganda, `${sert.desa}, Tanggal ${docDate.tanggalText}`, {
      withNumbers: true,
      mode: 'skw',
    }),

    // HALAMAN 2-3: PERMOHONAN PENDAFTARAN PERALIHAN HAK (TURUN WARIS)
    ...buildPermohonanPage({
      record,
      parties,
      kuasa,
      item14Text: 'Alih Media',
      boldIndices: [5],
      lampiranItems: [
        'Sertipikat asli',
        'Pengecekan Sertipikat',
        'Surat Kuasa',
        'Fotokopi KTP Penerima Kuasa',
        'Surat Keterangan Warisan',
        'Fotokopi Akta Kematian',
        'Fotokopi KTP, KK Ahli Waris',
        'Fotokopi KTP para saksi',
        'SSPD BPHTB',
        'Fotokopi SPPT PBB Terakhir',
        'Surat Pernyataan fisik tanah dan tidak sengketa',
      ],
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
      tanggalDash: docDate.tanggalDash,
      pageBreakBefore: true,
    }),

    // HALAMAN 4-5: SURAT KUASA (PENGESAHAN TUNGGAL DEFAULT PPATS / PILIHAN)
    ...buildSuratKuasaNonAkta({
      record,
      parties,
      kuasa,
      pengesahan: pengesahanTunggal,
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
      objekPrefix: 'Sehubungan dengan Turun Waris atas Sertipikat Tanah',
      khususBullets: [
        `Untuk dan atas nama Pemberi Kuasa mendaftarkan permohonan Peralihan Hak karena Warisan (Turun Waris) atas sertipikat tanah tersebut ke Kantor Pertanahan ${sert.kabupaten}.`,
        'Guna keperluan tersebut Penerima Kuasa berhak menghadap kepada pejabat Kantor Pertanahan yang berwenang, membayar biaya pendaftaran, menandatangani surat-surat yang diperlukan, serta mengambil sertipikat tanah setelah selesai.',
      ],
    }),

    // HALAMAN 6-7: SURAT PERNYATAAN PENGUASAAN FISIK TANAH DAN TIDAK SENGKETA
    ...buildPenguasaanFisikNonAkta({
      record,
      parties,
      saksi1,
      saksi2,
      pengesahan: pengesahanTunggal,
      hariTanggal: docDate.hariTanggal,
      butir1Text: `Bahwa kami adalah ahli waris yang sah dan pemilik sebidang tanah seluas ${sert.luasStr} M2 (${sert.luasTerbilang}) yang terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten} yang sudah bersertipikat ${sert.jenisHak} ${sert.ringkas} Desa ${sert.desa}, tercatat atas nama ${sert.pemegangHak} dan merupakan produk Kantor Pertanahan ${sert.kabupaten}.`,
    }),

    // HALAMAN 8: SURAT PERNYATAAN SPPT / PBB WARIS
    fp('SURAT PERNYATAAN', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 180,
      line: 260,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: F4_LINE_AIRY }),
    ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
    fp(
      `Dengan ini menyatakan dengan sebenarnya bahwa Surat Pemberitahuan Pajak Terhutang Pajak Bumi dan Bangunan (SPPT/PBB) Nomor Obyek Pajak (NOP) : ${nopClean}`,
      { align: AlignmentType.JUSTIFIED, spaceBefore: 120, spaceAfter: 40, line: F4_LINE_AIRY }
    ),
    fp(`Atas tanah warisan bersertipikat ${sert.jenisHak} :`, { spaceAfter: 60, line: F4_LINE_AIRY }),
    kvTable(
      [
        [sert.labelNo, sert.nilaiNo],
        ['Desa', sert.desa],
        ['Kecamatan', sert.kecamatan],
        ['Kabupaten', sert.kabupaten],
        ['Atas nama', sert.pemegangHak],
      ],
      { indentTwips: 420, labelPct: 26, spaceAfterRow: 0, line: 255 }
    ),
    fp(
      'Masih atas nama almarhum/pemilik lama karena belum dibalik nama atas nama para ahli waris.',
      { align: AlignmentType.JUSTIFIED, spaceBefore: 80, spaceAfter: 40, line: F4_LINE_AIRY }
    ),
    fp('Demikian surat pernyataan ini dibuat agar dipergunakan sebagaimana mestinya.', {
      align: AlignmentType.JUSTIFIED,
      spaceAfter: 200,
      line: F4_LINE_AIRY,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: buildSinglePengesahanCentered(pengesahanTunggal, 40),
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp('Yang menyatakan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: parties.length > 1 ? 120 : 760,
                  line: 250,
                }),
                ...buildNumberedSignerLines(parties, 360),
              ],
            }),
          ],
        }),
      ],
    }),
    ...buildSaksiBlocks(saksi1.nama, saksi2.nama, 200),

    // HALAMAN 9: SURAT PERNYATAAN BPHTB WARISAN
    fp('SURAT PERNYATAAN', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 180,
      line: 260,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: F4_LINE_AIRY }),
    ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
    fp('Dengan ini menyatakan :', { spaceBefore: 100, spaceAfter: 60, line: F4_LINE_AIRY }),
    fp(
      `-  Bahwa kami telah menerima harta Warisan berupa sebidang tanah tersebut ${sert.ringkas} Desa ${sert.desa} dengan Luas ${sert.luasStr} m2 (${sert.luasTerbilang}) yang terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}.`,
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 240, spaceAfter: 80, line: F4_LINE_AIRY }
    ),
    fp(
      '-  Bahwa apabila dalam Proses Peralihan Hak (Turun Waris) atas tanah yang kami terima tersebut diatas ternyata terjadi kekurangan pembayaran BPHTB, maka kami sanggup untuk membayar kekurangan BPHTB tersebut sesuai ketentuan peraturan perundang-undangan yang berlaku.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 240, spaceAfter: 200, line: F4_LINE_AIRY }
    ),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: buildSinglePengesahanCentered(pengesahanTunggal, 40),
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp('Yang menyatakan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: parties.length > 1 ? 120 : 760,
                  line: 250,
                }),
                ...buildNumberedSignerLines(parties, 360),
              ],
            }),
          ],
        }),
      ],
    }),
    ...buildSaksiBlocks(saksi1.nama, saksi2.nama, 200),

    // HALAMAN 10-11: PEMROSESAN DATA PRIBADI
    ...buildDataPribadiPage({
      parties,
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
    }),

    // HALAMAN 12-13: PERTANGGUNGJAWABAN KEABSAHAN DAN KEBENARAN DOKUMEN (TANPA SAKSI)
    fp('SURAT PERTANGGUNGJAWABAN KEABSAHAN DAN', {
      pageBreakBefore: true,
      bold: true,
      align: AlignmentType.CENTER,
      spaceAfter: 0,
      line: 260,
    }),
    fp('KEBENARAN DOKUMEN', {
      bold: true,
      align: AlignmentType.CENTER,
      spaceAfter: 180,
      line: 260,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 50, line: 250 }),
    ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
    fp(
      `Merupakan para ahli waris / pemilik tanah dengan jenis tanah ${sert.jenisHak} ${sert.ringkas} terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}, Provinsi ${sert.provinsi}.`,
      { align: AlignmentType.JUSTIFIED, spaceBefore: 100, spaceAfter: 60, line: 255 }
    ),
    fp('Dengan ini menyatakan :', { spaceAfter: 40, line: 255 }),
    fp(
      '1. Sertipikat tersebut adalah asli dan dokumen Surat Keterangan Warisan beserta seluruh dokumen pendukung adalah benar dan sah menurut hukum.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 255 }
    ),
    fp(
      `2. Seluruh dokumen yang diunggah pada sistem aplikasi layanan pertanahan pada Kantor Pertanahan ${sert.kabupaten} adalah benar dan sesuai dengan aslinya.`,
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 255 }
    ),
    fp(
      '3. Beritikad baik dan bertanggungjawab sepenuhnya atas penggunaan data yang diakses serta bersedia menerima sanksi administratif maupun hukum apabila data dinyatakan tidak benar.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 160, line: 255 }
    ),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: buildSinglePengesahanCentered(pengesahanTunggal, 20),
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp('Yang menyatakan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: parties.length > 1 ? 120 : 760,
                  line: 250,
                }),
                ...buildNumberedSignerLines(parties, 360),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

function buildPemecahanUkurUlangBundle(
  record: BerkasRecord,
  settings: OfficeSettings
): Array<Paragraph | Table> {
  const parties = collectNonAktaParties(record);
  const sert = getSertipikatInfo(record);
  const na = record.non_akta;
  const pengesahanTunggal = resolvePengesahanTunggal(record, settings);
  const ganda = resolvePengesahanGanda(record, settings);
  const isPemecahan =
    record.layanan === 'Pemecahan' || record.layanan === 'Pemecahan/Ukur Ulang';

  const saksi1 = resolvePersonCtx(record.akta.saksi_1, settings.profil_saksi, record.akta.saksi_1_detail);
  const saksi2 = resolvePersonCtx(record.akta.saksi_2, settings.profil_saksi, record.akta.saksi_2_detail);
  const kuasa = resolvePersonCtx(
    record.akta.nama_kuasa_bpn,
    settings.profil_kuasa,
    record.akta.kuasa_bpn_detail
  );

  const docDate = resolveNonAktaDocDate(record);
  const kotaKabupaten =
    (record.tanah.kabupaten || settings.nama_kabupaten || '')
      .replace(/^(Kabupaten|Kota)\s+/i, '')
      .trim() || '-';

  const subPermohonan =
    na?.sub_permohonan_ukur || (isPemecahan ? 'Pemecahan' : 'Pengukuran');
  const includeValidasi = subPermohonan.includes('Validasi');
  const boldIndices = isPemecahan
    ? includeValidasi
      ? [8, 14]
      : [8]
    : includeValidasi
    ? [1, 14]
    : [1];

  const bidangPecahan =
    na?.bidang_pecahan && na.bidang_pecahan.length > 0
      ? na.bidang_pecahan
      : [
          {
            id: 'b-1',
            label: 'Bagian 1 (Induk / Sisa)',
            luas_m2: Math.max(0, Math.round(record.tanah.luas_m2 / 2)),
            atas_nama: parties[0]?.nama || '-',
            keterangan: 'Bidang Sisa',
          },
          {
            id: 'b-2',
            label: 'Bagian 2 (Pecahan)',
            luas_m2: Math.max(
              0,
              record.tanah.luas_m2 - Math.round(record.tanah.luas_m2 / 2)
            ),
            atas_nama: parties[0]?.nama || '-',
            keterangan: 'Bidang Pecahan',
          },
        ];

  const blocks: Array<Paragraph | Table> = [
    // 1. PERMOHONAN
    ...buildPermohonanPage({
      record,
      parties,
      kuasa,
      item14Text: 'Verifikasi / Validasi',
      boldIndices,
      lampiranItems: [
        'Permohonan',
        'Fotokopi/Sertipikat asli',
        'Surat Kuasa',
        'Fotokopi KTP dan KK Para Pihak',
        'Surat Penguasaan Fisik',
        'Surat Pemasangan Batas',
        ...(isPemecahan ? ['Gambar Sket Tanah'] : []),
        ...(na?.sertakan_pelepasan_jalan ? ['Surat Pelepasan Jalan'] : []),
        ...(na?.sertakan_beda_nama ? ['Surat Pernyataan Beda Nama'] : []),
        'Surat SPPT Terhutang',
        'Fotokopi SPPT PBB Terakhir',
        'Koordinat dan Poligon Bidang',
      ],
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
      tanggalDash: docDate.tanggalDash,
      pageBreakBefore: false,
    }),

    // 2. SURAT KUASA
    ...buildSuratKuasaNonAkta({
      record,
      parties,
      kuasa,
      pengesahan: pengesahanTunggal,
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
      objekPrefix: `Sehubungan dengan permohonan ${isPemecahan ? 'Pemecahan Bidang' : 'Pengukuran Ulang'} atas Sertipikat`,
      khususBullets: [
        `Untuk dan atas nama Pemberi Kuasa mendaftarkan permohonan ${subPermohonan} atas sertipikat tanah tersebut ke Kantor Pertanahan ${sert.kabupaten}.`,
        'Guna keperluan tersebut Penerima Kuasa berhak menghadap pejabat Kantor Pertanahan yang berwenang, mendampingi petugas ukur di lapangan, membayar PNBP, serta mengambil hasil sertipikat/surat ukur setelah selesai.',
      ],
    }),

    // 3. PENGUASAAN FISIK & TIDAK SENGKETA
    ...buildPenguasaanFisikNonAkta({
      record,
      parties,
      saksi1,
      saksi2,
      pengesahan: pengesahanTunggal,
      hariTanggal: docDate.hariTanggal,
      butir1Text: `Bahwa ${parties.length > 1 ? 'kami' : 'saya'} adalah benar-benar pemilik sebidang tanah seluas ${sert.luasStr} M2 (${sert.luasTerbilang}) yang terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten} yang sudah bersertipikat ${sert.jenisHak} ${sert.ringkas} Desa ${sert.desa}, tercatat atas nama ${sert.pemegangHak}.`,
    }),

    // 4. SURAT PERNYATAAN PERMOHONAN PEMECAHAN / UKUR ULANG
    fp('SURAT PERNYATAAN', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 180,
      line: 260,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: F4_LINE_AIRY }),
    ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
    fp(
      `Dengan ini menyatakan bahwa ${parties.length > 1 ? 'kami' : 'saya'} adalah pemegang hak atas sebidang tanah ${sert.jenisHak} ${sert.ringkas} seluas ${sert.luasStr} M² terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}.`,
      { align: AlignmentType.JUSTIFIED, spaceBefore: 120, spaceAfter: 60, line: F4_LINE_AIRY }
    ),
    fp(
      isPemecahan
        ? `Bahwa bidang tanah tersebut diatas benar-benar akan dilakukan PEMECAHAN menjadi ${bidangPecahan.length} bagian bidang tanah untuk keperluan penataan batas/pembagian aset dan tidak berada dalam sengketa dengan pihak manapun.`
        : 'Bahwa bidang tanah tersebut diatas diajukan PENGUKURAN ULANG / pengembalian batas untuk mencocokkan luas dan batas fisik di lapangan sesuai kondisi terkini.',
      { align: AlignmentType.JUSTIFIED, spaceAfter: 60, line: F4_LINE_AIRY }
    ),
    fp('Demikian surat pernyataan ini dibuat dengan sebenarnya untuk dapat dipergunakan sebagaimana mestinya.', {
      align: AlignmentType.JUSTIFIED,
      spaceAfter: 180,
      line: F4_LINE_AIRY,
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: buildSinglePengesahanCentered(pengesahanTunggal, 20),
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp('Yang menyatakan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: parties.length > 1 ? 120 : 760,
                  line: 250,
                }),
                ...buildNumberedSignerLines(parties, 360),
              ],
            }),
          ],
        }),
      ],
    }),
    ...buildSaksiBlocks(saksi1.nama, saksi2.nama, 180),

    // 5. SURAT PERNYATAAN PEMASANGAN TANDA BATAS & PERSETUJUAN PEMILIK BERBATASAN
    fp('SURAT PERNYATAAN PEMASANGAN TANDA BATAS DAN', {
      pageBreakBefore: true,
      bold: true,
      align: AlignmentType.CENTER,
      spaceAfter: 0,
      line: 255,
    }),
    fp('PERSETUJUAN PEMILIK YANG BERBATASAN', {
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 140,
      line: 255,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 40, line: 245 }),
    ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
    fp(
      `Selaku pemohon ${isPemecahan ? 'Pemecahan' : 'Pengukuran'} bidang tanah ${sert.jenisHak} ${sert.ringkas}, Luas ${sert.luasStr} M² di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}, dengan ini menyatakan :`,
      { align: AlignmentType.JUSTIFIED, spaceBefore: 80, spaceAfter: 50, line: 245 }
    ),
    fp(
      '1. Telah memasang tanda batas (patok) secara permanen pada setiap sudut bidang tanah sesuai dengan ketentuan peraturan pertanahan yang berlaku.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 300, hanging: 300, spaceAfter: 20, line: 245 }
    ),
    fp(
      '2. Tanda batas yang dipasang telah disepakati dan disetujui oleh pemilik bidang tanah yang berbatasan langsung.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 300, hanging: 300, spaceAfter: 20, line: 245 }
    ),
    fp(
      '3. Apabila terjadi selisih luas akibat pemasangan batas lapangan, pemohon bersedia menyesuaikan dengan hasil ukur resmi Kantor Pertanahan.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 300, hanging: 300, spaceAfter: 20, line: 245 }
    ),
    fp(
      '4. Apabila dikemudian hari terdapat keberatan dari pihak lain mengenai tanda batas, seluruh akibat hukum menjadi tanggung jawab pemohon.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 300, hanging: 300, spaceAfter: 90, line: 245 }
    ),
    fp('Menyetujui Pemilik Tanah Yang Berbatasan :', { bold: true, spaceAfter: 40, line: 245 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        ['Arah', 'Nama Pemilik / Batas Tanah', 'Tanda Tangan'],
        ['Sebelah Utara', record.tanah.batas_utara || '-', '1. ....................'],
        ['Sebelah Timur', record.tanah.batas_timur || '-', '2. ....................'],
        ['Sebelah Selatan', record.tanah.batas_selatan || '-', '3. ....................'],
        ['Sebelah Barat', record.tanah.batas_barat || '-', '4. ....................'],
      ].map((row, idx) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: 24, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [fp(row[0], { bold: idx === 0, spaceAfter: 20, line: 235 })],
            }),
            new TableCell({
              width: { size: 48, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [fp(row[1], { bold: idx === 0, spaceAfter: 20, line: 235 })],
            }),
            new TableCell({
              width: { size: 28, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [fp(row[2], { bold: idx === 0, spaceAfter: 20, line: 235 })],
            }),
          ],
        })
      ),
    }),
    fp('', { spaceAfter: 60, line: 160 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: buildSinglePengesahanCentered(pengesahanTunggal, 20),
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 245,
                }),
                fp('Pemohon / Yang Memasang Batas', {
                  align: AlignmentType.CENTER,
                  spaceAfter: parties.length > 1 ? 100 : 720,
                  line: 245,
                }),
                ...buildNumberedSignerLines(parties, 360),
              ],
            }),
          ],
        }),
      ],
    }),
    ...buildSaksiBlocks(saksi1.nama, saksi2.nama, 120),
  ];

  // 6. GAMBAR SKET TANAH (UNTUK PEMECAHAN)
  if (isPemecahan) {
    blocks.push(
      fp('GAMBAR SKET TANAH RENCANA PEMECAHAN', {
        pageBreakBefore: true,
        bold: true,
        underline: true,
        align: AlignmentType.CENTER,
        spaceAfter: 160,
        line: 260,
      }),
      kvTable(
        [
          ['Jenis & Nomor Sertipikat', `${sert.jenisHak} ${sert.ringkas}`],
          ['Luas Bidang Induk', `${sert.luasStr} M²`],
          ['Letak Tanah', `Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}`],
          ['Pemegang Hak', sert.pemegangHak],
        ],
        { labelPct: 30, spaceAfterRow: 0, line: 250 }
      ),
      fp('Rincian Rencana Pecahan Bidang Tanah :', {
        bold: true,
        spaceBefore: 120,
        spaceAfter: 60,
        line: 250,
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: ['No', 'Bagian Bidang', 'Luas (m²)', 'Keterangan / Penerima'].map(
              (h, cIdx) =>
                new TableCell({
                  width: {
                    size: cIdx === 0 ? 8 : cIdx === 1 ? 34 : cIdx === 2 ? 18 : 40,
                    type: WidthType.PERCENTAGE,
                  },
                  borders: THIN_BORDER,
                  children: [fp(h, { bold: true, align: AlignmentType.CENTER, spaceAfter: 20, line: 240 })],
                })
            ),
          }),
          ...bidangPecahan.map((b, i) =>
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 8, type: WidthType.PERCENTAGE },
                  borders: THIN_BORDER,
                  children: [fp(String(i + 1), { align: AlignmentType.CENTER, spaceAfter: 20, line: 240 })],
                }),
                new TableCell({
                  width: { size: 34, type: WidthType.PERCENTAGE },
                  borders: THIN_BORDER,
                  children: [fp(b.label, { spaceAfter: 20, line: 240 })],
                }),
                new TableCell({
                  width: { size: 18, type: WidthType.PERCENTAGE },
                  borders: THIN_BORDER,
                  children: [fp(`${formatNumberId(b.luas_m2)} m²`, { align: AlignmentType.RIGHT, spaceAfter: 20, line: 240 })],
                }),
                new TableCell({
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  borders: THIN_BORDER,
                  children: [
                    fp(
                      [b.atas_nama, b.keterangan].filter(Boolean).join(' · ') || '-',
                      { spaceAfter: 20, line: 240 }
                    ),
                  ],
                }),
              ],
            })
          ),
        ],
      }),
      fp('', { spaceAfter: 80, line: 160 }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: THIN_BORDER,
                children: [
                  fp('SKETSA BIDANG TANAH & BATAS-BATAS (UTARA DI ATAS)', {
                    bold: true,
                    align: AlignmentType.CENTER,
                    spaceBefore: 80,
                    spaceAfter: 1800,
                    line: 240,
                  }),
                  fp(
                    `Utara: ${record.tanah.batas_utara || '-'} | Timur: ${record.tanah.batas_timur || '-'} | Selatan: ${record.tanah.batas_selatan || '-'} | Barat: ${record.tanah.batas_barat || '-'}`,
                    { align: AlignmentType.CENTER, spaceAfter: 80, line: 235 }
                  ),
                ],
              }),
            ],
          }),
        ],
      }),
      fp('', { spaceAfter: 80, line: 160 }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 46, type: WidthType.PERCENTAGE },
                borders: NO_BORDER,
                children: buildSinglePengesahanCentered(pengesahanTunggal, 20),
              }),
              new TableCell({
                width: { size: 54, type: WidthType.PERCENTAGE },
                borders: NO_BORDER,
                children: [
                  fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                    align: AlignmentType.CENTER,
                    spaceAfter: 0,
                    line: 250,
                  }),
                  fp('Pemohon', {
                    align: AlignmentType.CENTER,
                    spaceAfter: parties.length > 1 ? 100 : 760,
                    line: 250,
                  }),
                  ...buildNumberedSignerLines(parties, 360),
                ],
              }),
            ],
          }),
        ],
      })
    );
  }

  // 7. OPSIONAL: PELEPASAN JALAN
  if (na?.sertakan_pelepasan_jalan) {
    const luasJalan = na.pelepasan_jalan_luas_m2 ?? 0;
    blocks.push(
      fp('SURAT PERNYATAAN PELEPASAN JALAN', {
        pageBreakBefore: true,
        bold: true,
        underline: true,
        align: AlignmentType.CENTER,
        spaceAfter: 180,
        line: 260,
      }),
      fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: F4_LINE_AIRY }),
      ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
      fp(
        `Dengan ini menyatakan dengan sebenarnya bahwa terhadap sebagian bidang tanah ${sert.jenisHak} ${sert.ringkas} seluas ${sert.luasStr} M² terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}, saya/kami melepas hak sebagian tanah seluas ± ${formatNumberId(luasJalan)} M² untuk dipergunakan sebagai Akses Jalan (${na.pelepasan_jalan_keterangan?.trim() || 'Jalan Umum/Akses Bidang Pecahan'}) tanpa menuntut ganti rugi apapun di kemudian hari.`,
        { align: AlignmentType.JUSTIFIED, spaceBefore: 120, spaceAfter: 100, line: F4_LINE_AIRY }
      ),
      fp('Demikian surat pernyataan pelepasan jalan ini dibuat untuk dapat dipergunakan sebagaimana mestinya.', {
        align: AlignmentType.JUSTIFIED,
        spaceAfter: 200,
        line: F4_LINE_AIRY,
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 46, type: WidthType.PERCENTAGE },
                borders: NO_BORDER,
                children: buildSinglePengesahanCentered(pengesahanTunggal, 20),
              }),
              new TableCell({
                width: { size: 54, type: WidthType.PERCENTAGE },
                borders: NO_BORDER,
                children: [
                  fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                    align: AlignmentType.CENTER,
                    spaceAfter: 0,
                    line: 250,
                  }),
                  fp('Yang menyatakan', {
                    align: AlignmentType.CENTER,
                    spaceAfter: parties.length > 1 ? 100 : 760,
                    line: 250,
                  }),
                  ...buildNumberedSignerLines(parties, 360),
                ],
              }),
            ],
          }),
        ],
      }),
      ...buildSaksiBlocks(saksi1.nama, saksi2.nama, 180)
    );
  }

  // 8. OPSIONAL: BEDA NAMA (DENGAN 2 PENGESAHAN CAMAT & KADES ASAL SERTIPIKAT)
  if (na?.sertakan_beda_nama) {
    blocks.push(
      ...buildBedaNamaPage({
        record,
        parties,
        saksi1,
        saksi2,
        ganda,
        kotaKabupaten,
        tanggalText: docDate.tanggalText,
      })
    );
  }

  // 9. SURAT PERNYATAAN PENUTUPAN PERMOHONAN DAN TIDAK MENUNTUT PENGEMBALIAN PNBP
  blocks.push(
    fp('SURAT PERNYATAAN PENUTUPAN PERMOHONAN DAN', {
      pageBreakBefore: true,
      bold: true,
      align: AlignmentType.CENTER,
      spaceAfter: 0,
      line: 255,
    }),
    fp('TIDAK MENUNTUT PENGEMBALIAN PNBP', {
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 180,
      line: 255,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 60, line: 250 }),
    ...buildNumberedPartiesKvBlocks(parties, { spaceAfterRow: 0 }),
    fp(
      `Selaku pemohon layanan ${subPermohonan} atas sertipikat ${sert.jenisHak} ${sert.ringkas} Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten}, dengan ini menyatakan :`,
      { align: AlignmentType.JUSTIFIED, spaceBefore: 100, spaceAfter: 60, line: 255 }
    ),
    fp(
      '1. Lokasi bidang tanah yang dimohonkan pengukuran/pemecahan adalah benar dan sesuai titik koordinat share location dengan toleransi radius maksimal 100 meter serta dapat ditunjukkan secara langsung kepada petugas ukur.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 255 }
    ),
    fp(
      '2. Tanda batas permanen telah terpasang dengan jelas dan pemilik batas telah menyetujui batas bidang tanah.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 40, line: 255 }
    ),
    fp(
      '3. Apabila syarat fisik lapangan/dokumen tidak terpenuhi hingga batas waktu sesuai SOP Kantor Pertanahan sehingga permohonan ditutup/dibatalkan, pemohon bersedia menerima penutupan berkas dan tidak menuntut pengembalian biaya PNBP yang telah disetorkan.',
      { align: AlignmentType.JUSTIFIED, indentLeft: 320, hanging: 320, spaceAfter: 180, line: 255 }
    ),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 46, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: buildSinglePengesahanCentered(pengesahanTunggal, 20),
            }),
            new TableCell({
              width: { size: 54, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 250,
                }),
                fp('Yang membuat pernyataan', {
                  align: AlignmentType.CENTER,
                  spaceAfter: parties.length > 1 ? 100 : 760,
                  line: 250,
                }),
                ...buildNumberedSignerLines(parties, 360),
              ],
            }),
          ],
        }),
      ],
    })
  );

  // 10. LAMPIRAN FOTO LOKASI, PATOK BATAS & POLIGON
  blocks.push(
    fp('LAMPIRAN FOTO LOKASI BIDANG TANAH, PATOK BATAS & POLIGON', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 140,
      line: 255,
    }),
    kvTable(
      [
        ['Sertipikat', `${sert.jenisHak} ${sert.ringkas}`],
        ['Letak Tanah', `Desa ${sert.desa}, Kec. ${sert.kecamatan}, ${sert.kabupaten}`],
        ['Koordinat GPS', na?.koordinat_gps?.trim() || '(Isi titik koordinat Lat, Lng)'],
        ['Tautan Google Maps', na?.link_maps?.trim() || '(Tempel tautan share location)'],
      ],
      { labelPct: 28, spaceAfterRow: 0, line: 245 }
    ),
    fp('', { spaceAfter: 80, line: 160 }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [
                fp('FOTO LOKASI TAMPAK DEPAN', {
                  bold: true,
                  align: AlignmentType.CENTER,
                  spaceBefore: 60,
                  spaceAfter: 1800,
                  line: 240,
                }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [
                fp('GAMBAR POLIGON / CITRA SATELIT', {
                  bold: true,
                  align: AlignmentType.CENTER,
                  spaceBefore: 60,
                  spaceAfter: 1800,
                  line: 240,
                }),
              ],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [
                fp('FOTO PATOK BATAS 1 & 2', {
                  bold: true,
                  align: AlignmentType.CENTER,
                  spaceBefore: 60,
                  spaceAfter: 1600,
                  line: 240,
                }),
              ],
            }),
            new TableCell({
              width: { size: 50, type: WidthType.PERCENTAGE },
              borders: THIN_BORDER,
              children: [
                fp('FOTO PATOK BATAS 3 & 4', {
                  bold: true,
                  align: AlignmentType.CENTER,
                  spaceBefore: 60,
                  spaceAfter: 1600,
                  line: 240,
                }),
              ],
            }),
          ],
        }),
      ],
    })
  );

  return blocks;
}

function buildRalatDataBundle(
  record: BerkasRecord,
  settings: OfficeSettings
): Array<Paragraph | Table> {
  const parties = collectNonAktaParties(record);
  const sert = getSertipikatInfo(record);
  const na = record.non_akta;
  const pengesahanTunggal = resolvePengesahanTunggal(record, settings);
  const ganda = resolvePengesahanGanda(record, settings);

  const saksi1 = resolvePersonCtx(record.akta.saksi_1, settings.profil_saksi, record.akta.saksi_1_detail);
  const saksi2 = resolvePersonCtx(record.akta.saksi_2, settings.profil_saksi, record.akta.saksi_2_detail);
  const kuasa = resolvePersonCtx(
    record.akta.nama_kuasa_bpn,
    settings.profil_kuasa,
    record.akta.kuasa_bpn_detail
  );

  const docDate = resolveNonAktaDocDate(record);
  const kotaKabupaten =
    (record.tanah.kabupaten || settings.nama_kabupaten || '')
      .replace(/^(Kabupaten|Kota)\s+/i, '')
      .trim() || '-';

  const namaSertipikat =
    na?.nama_di_sertipikat?.trim() || sert.pemegangHak || parties[0]?.nama || '-';
  const namaKtpKk = na?.nama_di_ktp_kk?.trim() || parties[0]?.nama || '-';
  const nomorKk = na?.nomor_kk?.trim() || '-';
  const dokLain = na?.jenis_dokumen_lain?.trim() || 'Kutipan Akta Nikah / Ijazah';
  const namaDokLain = na?.nama_di_dokumen_lain?.trim() || namaSertipikat;
  const tglLahirBaku = na?.tgl_lahir_baku?.trim() || parties[0]?.ttl || '-';
  const uraianKronologi =
    na?.uraian_kronologi?.trim() ||
    `Terjadi perbedaan penulisan nama pemegang hak pada sertipikat (${namaSertipikat}) dengan identitas kependudukan terbaru pada KTP dan Kartu Keluarga (${namaKtpKk}), dikarenakan kesalahan pencatatan administrasi/ejaan terdahulu.`;

  const pemohon = parties[0];

  return [
    // 1. PERMOHONAN (14. RALAT DATA)
    ...buildPermohonanPage({
      record,
      parties,
      kuasa,
      item14Text: 'RALAT DATA',
      boldIndices: [7, 14],
      lampiranItems: [
        'Fotokopi KTP dan KK',
        'Fotokopi/Sertipikat asli',
        'Surat Kuasa',
        'Surat Kronologi Ralat Data',
        'SK Penguasaan fisik',
        'Surat Pernyataan Beda nama',
      ],
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
      tanggalDash: docDate.tanggalDash,
      pageBreakBefore: false,
    }),

    // 2. SURAT KUASA RALAT DATA
    ...buildSuratKuasaNonAkta({
      record,
      parties,
      kuasa,
      pengesahan: pengesahanTunggal,
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
      objekPrefix: 'Sehubungan dengan Ralat Data pada Sertipikat Tanah',
      khususBullets: [
        `Untuk dan atas nama Pemberi Kuasa mendaftarkan permohonan Ralat Data Sertipikat (${sert.jenisHak} ${sert.ringkas}) ke Kantor Pertanahan ${sert.kabupaten}.`,
        'Guna keperluan tersebut Penerima Kuasa berhak menghadap kepada pejabat Kantor Pertanahan yang berwenang, menandatangani surat permohonan, membayar biaya pendaftaran, serta mengambil sertipikat setelah proses Ralat Data selesai.',
      ],
    }),

    // 3. PENGUASAAN FISIK & TIDAK SENGKETA (DENGAN CATATAN ALIAS)
    ...buildPenguasaanFisikNonAkta({
      record,
      parties,
      saksi1,
      saksi2,
      pengesahan: pengesahanTunggal,
      hariTanggal: docDate.hariTanggal,
      butir1Text: `Bahwa saya adalah benar-benar pemilik sebidang tanah seluas ${sert.luasStr} M2 (${sert.luasTerbilang}) yang terletak di Desa ${sert.desa}, Kecamatan ${sert.kecamatan}, ${sert.kabupaten} yang sudah bersertipikat ${sert.jenisHak} ${sert.ringkas} Desa ${sert.desa}, yang tercatat atas nama ${namaSertipikat} dan merupakan produk Kantor Pertanahan ${sert.kabupaten}.`,
      catatanAlias: `Nama ${namaSertipikat} di sertipikat alias ${namaKtpKk} sesuai KTP/KK`,
    }),

    // 4. SURAT PERNYATAAN BEDA NAMA (2 PENGESAHAN CAMAT & KADES)
    ...buildBedaNamaPage({
      record,
      parties,
      saksi1,
      saksi2,
      ganda,
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
    }),

    // 5. SURAT KRONOLOGI RALAT DATA (2 PENGESAHAN CAMAT & KADES)
    fp('SURAT KRONOLOGI RALAT DATA', {
      pageBreakBefore: true,
      bold: true,
      underline: true,
      align: AlignmentType.CENTER,
      spaceAfter: 180,
      line: 260,
    }),
    fp('Yang bertanda tangan dibawah ini :', { spaceAfter: 50, line: 250 }),
    kvTable(
      [
        ['Nama', pemohon.nama],
        ['Tempat/Tgl. Lahir', pemohon.ttl],
        ['Pekerjaan', pemohon.pekerjaan],
        ['Nomor KTP', pemohon.nik],
        ['Nomor KK', nomorKk],
        ['Alamat', pemohon.alamatLengkap],
      ],
      { indentTwips: 420, labelPct: 28, spaceAfterRow: 0, line: 250 }
    ),
    fp('Dengan ini menerangkan kronologi perbedaan data identitas sebagai berikut :', {
      spaceBefore: 120,
      spaceAfter: 60,
      line: 250,
    }),
    fp(`1. Pada Sertipikat ${sert.jenisHak} ${sert.ringkas} Desa ${sert.desa} tertulis nama : ${namaSertipikat}`, {
      indentLeft: 320,
      hanging: 320,
      spaceAfter: 30,
      line: 250,
    }),
    fp(`2. Pada KTP dan KK (No. ${nomorKk}) tertulis nama : ${namaKtpKk}`, {
      indentLeft: 320,
      hanging: 320,
      spaceAfter: 30,
      line: 250,
    }),
    fp(`3. Pada ${dokLain} tertulis nama : ${namaDokLain}`, {
      indentLeft: 320,
      hanging: 320,
      spaceAfter: 60,
      line: 250,
    }),
    fp(`Uraian Kronologi : ${uraianKronologi}`, {
      align: AlignmentType.JUSTIFIED,
      spaceAfter: 60,
      line: 255,
    }),
    fp(
      `Bahwa pemilik yang sah adalah satu orang yang sama, sehingga data identitas baku yang dimohonkan untuk diralat pada Sertipikat menjadi Nama : ${namaKtpKk}, Tempat/Tanggal Lahir : ${tglLahirBaku}. Pemohon bertanggungjawab penuh secara formil maupun materiil atas kebenaran data tersebut.`,
      { align: AlignmentType.JUSTIFIED, spaceAfter: 140, line: 255 }
    ),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 52, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp('Saksi – saksi :', { spaceAfter: 100, line: 245 }),
                fp(`1. ${saksi1.nama} (...............)`, { spaceAfter: 240, line: 245 }),
                fp(`2. ${saksi2.nama} (...............)`, { spaceAfter: 20, line: 245 }),
              ],
            }),
            new TableCell({
              width: { size: 48, type: WidthType.PERCENTAGE },
              borders: NO_BORDER,
              children: [
                fp(`${kotaKabupaten}, ${docDate.tanggalText}`, {
                  align: AlignmentType.CENTER,
                  spaceAfter: 0,
                  line: 245,
                }),
                fp('Yang membuat kronologi', {
                  align: AlignmentType.CENTER,
                  spaceAfter: 700,
                  line: 245,
                }),
                fp(pemohon.nama, {
                  bold: true,
                  underline: true,
                  align: AlignmentType.CENTER,
                  spaceAfter: 20,
                  line: 245,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    fp('', { spaceAfter: 80, line: 160 }),
    buildDualPengesahanTable(ganda, `${sert.desa}, Tanggal ${docDate.tanggalText}`, {
      withNumbers: true,
      mode: 'beda_nama',
    }),

    // 6. PERSETUJUAN PEMROSESAN DATA PRIBADI
    ...buildDataPribadiPage({
      parties,
      kotaKabupaten,
      tanggalText: docDate.tanggalText,
    }),
  ];
}

export async function buildNonAktaLampiranBuffer(
  record: BerkasRecord,
  settings: OfficeSettings
): Promise<Uint8Array> {
  let children: Array<Paragraph | Table>;
  if (record.layanan === 'Turun Waris') {
    children = buildTurunWarisBundle(record, settings);
  } else if (record.layanan === 'Ralat Data') {
    children = buildRalatDataBundle(record, settings);
  } else {
    children = buildPemecahanUkurUlangBundle(record, settings);
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: F4_FONT,
            size: F4_SIZE,
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: {
              line: F4_LINE,
              after: 20,
            },
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: PAPER_SIZE_F4_TWIPS,
            margin: LAMPIRAN_MARGINS,
          },
        },
        children,
      },
    ],
  });

  const arrayBuffer = await Packer.toArrayBuffer(doc);
  return new Uint8Array(arrayBuffer);
}
