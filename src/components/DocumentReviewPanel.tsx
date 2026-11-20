import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Eye,
  FileText,
  ZoomIn,
  ZoomOut,
  Edit3,
} from 'lucide-react';
import {
  BerkasRecord,
  CustomTemplateEntry,
  OfficeSettings,
  TemplateDocKey,
} from '../types/ppats';
import {
  DocxPreviewDocument,
  PreviewBlock,
  PreviewParagraphBlock,
  PreviewTableBlock,
  renderDocxPreviewPages,
} from '../utils/docxPreviewParser';
import { isAktaService } from '../utils/taxAndAudit';
import { resolvePengesahanGanda, resolvePengesahanTunggal } from '../utils/nonAktaTemplates';

interface DocumentReviewPanelProps {
  record: BerkasRecord;
  settings: OfficeSettings;
  customTemplates: Partial<Record<TemplateDocKey, CustomTemplateEntry>>;
  activeDocKey: '01_akta' | '02_lampiran_bundel';
  onSelectDocKey: (docKey: '01_akta' | '02_lampiran_bundel') => void;
  onEditStep: (step: 2 | 3) => void;
}

const RenderParagraph: React.FC<{ block: PreviewParagraphBlock; compact?: boolean }> = ({
  block,
  compact,
}) => {
  const hasText = block.runs.some((r) => r.text.trim().length > 0);
  const textAlign = block.align === 'justify' ? 'justify' : block.align;
  const textIndent = block.hangingPx
    ? -block.hangingPx
    : block.firstLinePx || 0;

  if (!hasText) {
    return <div className={compact ? 'h-1.5' : 'h-2.5'} aria-hidden="true" />;
  }

  return (
    <p
      className={`${compact ? 'text-[11px] leading-[1.35]' : 'text-[12.5px] leading-[1.55]'} text-slate-900 whitespace-pre-line ${
        block.borderBottom ? 'border-b-2 border-slate-900 pb-1.5 mb-2' : ''
      }`}
      style={{
        textAlign,
        paddingLeft: block.indentLeftPx > 0 ? `${block.indentLeftPx}px` : undefined,
        textIndent: textIndent !== 0 ? `${textIndent}px` : undefined,
        marginTop: block.spaceBeforePx > 0 ? `${block.spaceBeforePx}px` : undefined,
        marginBottom: block.spaceAfterPx > 0 ? `${block.spaceAfterPx}px` : undefined,
      }}
    >
      {block.runs.map((run, rIdx) => {
        const classes = [
          run.bold ? 'font-bold' : 'font-normal',
          run.italic ? 'italic' : '',
          run.underline ? 'underline underline-offset-2' : '',
        ]
          .filter(Boolean)
          .join(' ');
        const fontSizePx =
          !compact && run.sizePt && run.sizePt !== 12
            ? Math.round((run.sizePt / 12) * 12.5 * 10) / 10
            : undefined;
        return (
          <span
            key={rIdx}
            className={classes}
            style={fontSizePx ? { fontSize: `${fontSizePx}px`, lineHeight: 1.35 } : undefined}
          >
            {run.text}
          </span>
        );
      })}
    </p>
  );
};

const RenderBlock: React.FC<{ block: PreviewBlock; inTable?: boolean }> = ({
  block,
  inTable,
}) => {
  if (block.kind === 'paragraph') {
    return <RenderParagraph block={block} compact={inTable} />;
  }
  return <RenderTable block={block} />;
};

const RenderTable: React.FC<{ block: PreviewTableBlock }> = ({ block }) => {
  return (
    <div
      className="my-1.5"
      style={{
        paddingLeft: block.indentLeftPx > 0 ? `${block.indentLeftPx}px` : undefined,
      }}
    >
      <table
        className="border-collapse table-fixed"
        style={{ width: `${Math.min(100, Math.max(40, block.widthPct))}%` }}
      >
        <tbody>
          {block.rows.map((row, rIdx) => (
            <tr key={rIdx}>
              {row.cells.map((cell, cIdx) => {
                if (cell.hidden) return null;
                const borderStyle: React.CSSProperties = {
                  width: cell.widthPct ? `${cell.widthPct}%` : undefined,
                  borderTop: cell.borders.top ? '1px solid #334155' : 'none',
                  borderBottom: cell.borders.bottom ? '1px solid #334155' : 'none',
                  borderLeft: cell.borders.left ? '1px solid #334155' : 'none',
                  borderRight: cell.borders.right ? '1px solid #334155' : 'none',
                  verticalAlign: 'top',
                };
                const hasAnyBorder =
                  cell.borders.top ||
                  cell.borders.bottom ||
                  cell.borders.left ||
                  cell.borders.right;
                return (
                  <td
                    key={cIdx}
                    rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                    style={borderStyle}
                    className={hasAnyBorder ? 'px-2 py-1' : 'px-1 py-0.5'}
                  >
                    {cell.blocks.map((b, bIdx) => (
                      <RenderBlock key={bIdx} block={b} inTable />
                    ))}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export const DocumentReviewPanel: React.FC<DocumentReviewPanelProps> = ({
  record,
  settings,
  customTemplates,
  activeDocKey,
  onSelectDocKey,
  onEditStep,
}) => {
  const [preview, setPreview] = useState<DocxPreviewDocument | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedPage, setSelectedPage] = useState<number | 'all'>('all');
  const [zoomPct, setZoomPct] = useState<85 | 100 | 115>(100);

  useEffect(() => {
    setSelectedPage('all');
  }, [activeDocKey]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);

    renderDocxPreviewPages(record, settings, customTemplates, activeDocKey)
      .then((doc) => {
        if (!cancelled) {
          setPreview(doc);
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setError(
            err instanceof Error
              ? err.message
              : 'Gagal menyusun pratinjau dokumen .DOCX.'
          );
        }
      })
      .finally(() => {
        if (!cancelled) {
          setLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [record, settings, customTemplates, activeDocKey]);

  const isAkta = isAktaService(record.layanan);
  const pengesahanTunggal = !isAkta ? resolvePengesahanTunggal(record, settings) : null;
  const pengesahanGanda = !isAkta ? resolvePengesahanGanda(record, settings) : null;
  const nonAktaPemohonNames = !isAkta
    ? record.para_pihak
        .filter((p) => p.peran.startsWith('Pihak '))
        .map((p) => p.nama)
        .filter(Boolean)
        .join(', ') || '-'
    : '';
  const hasDualPengesahan =
    record.layanan === 'Turun Waris' ||
    record.layanan === 'Ralat Data' ||
    Boolean(record.non_akta?.sertakan_beda_nama);

  const visiblePages =
    preview && selectedPage !== 'all'
      ? preview.pages.filter((p) => p.pageNumber === selectedPage)
      : preview?.pages ?? [];

  const paperMaxWidthPx =
    zoomPct === 85 ? 660 : zoomPct === 115 ? 890 : 770;

  return (
    <section
      aria-label="Review Dokumen Sebelum Generate"
      className="border border-slate-200 rounded-xl bg-white overflow-hidden"
    >
      {/* Top Toolbar */}
      <div className="px-4 py-3.5 border-b border-slate-200 bg-slate-50/70 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
        <div className="flex items-start gap-2.5">
          <Eye className="w-4.5 h-4.5 text-slate-700 mt-0.5 shrink-0" />
          <div>
            <h3 className="font-display text-sm font-semibold text-slate-900">
              Review Dokumen Sebelum Generate
            </h3>
            <p className="text-xs text-slate-600 mt-0.5">
              {preview ? (
                <>
                  <span>{preview.fileName}</span>
                  <span aria-hidden="true"> · </span>
                  <span>
                    Kertas {preview.paperSize}{' '}
                    {preview.paperSize === 'A4' ? '(210 × 297 mm)' : '(215 × 330 mm)'}
                  </span>
                  <span aria-hidden="true"> · </span>
                  <span className="tabular-nums">{preview.pages.length} bagian halaman</span>
                </>
              ) : (
                <span>Memeriksa isi dokumen .DOCX aktual sesuai data berkas...</span>
              )}
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {/* Switcher Akta A4 vs Lampiran F4 */}
          <div
            role="tablist"
            aria-label="Pilih dokumen untuk direview"
            className="inline-flex items-center p-1 bg-slate-200/70 rounded-lg"
          >
            {isAkta && (
              <button
                type="button"
                role="tab"
                aria-selected={activeDocKey === '01_akta'}
                onClick={() => onSelectDocKey('01_akta')}
                className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                  activeDocKey === '01_akta'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                01. Blangko Akta (A4)
              </button>
            )}
            <button
              type="button"
              role="tab"
              aria-selected={!isAkta || activeDocKey === '02_lampiran_bundel'}
              onClick={() => onSelectDocKey('02_lampiran_bundel')}
              className={`px-3 py-1.5 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                !isAkta || activeDocKey === '02_lampiran_bundel'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {isAkta ? '02. Lampiran Bundel (F4)' : 'Bundel Lampiran Non-Akta (F4)'}
            </button>
          </div>

          {/* Zoom Controls */}
          <div className="inline-flex items-center border border-slate-200 rounded-lg bg-white overflow-hidden">
            <button
              type="button"
              onClick={() => setZoomPct(zoomPct === 115 ? 100 : 85)}
              disabled={zoomPct === 85}
              aria-label="Perkecil pratinjau"
              className="px-2 py-1.5 text-slate-600 hover:bg-slate-50 disabled:opacity-40"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="px-2 text-xs font-mono tabular-nums text-slate-700 border-x border-slate-200">
              {zoomPct}%
            </span>
            <button
              type="button"
              onClick={() => setZoomPct(zoomPct === 85 ? 100 : 115)}
              disabled={zoomPct === 115}
              aria-label="Perbesar pratinjau"
              className="px-2 py-1.5 text-slate-600 hover:bg-slate-50 disabled:opacity-40"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      </div>

      {/* Verification Strip (Resolved Roles & Leftover Tag Check) */}
      {preview && (
        <div className="px-4 py-2.5 border-b border-slate-200 bg-white flex flex-col lg:flex-row lg:items-center lg:justify-between gap-2 text-xs">
          <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-slate-700">
            <span
              className={`inline-flex items-center gap-1.5 font-medium ${
                preview.summary.leftoverTags.length === 0
                  ? 'text-emerald-700'
                  : 'text-rose-700'
              }`}
            >
              {preview.summary.leftoverTags.length === 0 ? (
                <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
              ) : (
                <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
              )}
              {preview.summary.leftoverTags.length === 0
                ? 'Seluruh variabel dokumen terisi'
                : `${preview.summary.leftoverTags.length} tag belum terisi (${preview.summary.leftoverTags.join(', ')})`}
            </span>
            <span aria-hidden="true" className="text-slate-300">|</span>
            {isAkta ? (
              <>
                <span>
                  <strong className="text-slate-900">{preview.summary.ketPihak1Label}:</strong>{' '}
                  {preview.summary.ketPihak1Names}
                </span>
                <span aria-hidden="true" className="text-slate-300">·</span>
                <span>
                  <strong className="text-slate-900">{preview.summary.ketPihak2Label}:</strong>{' '}
                  {preview.summary.ketPihak2Names}
                </span>
                <span aria-hidden="true" className="text-slate-300">·</span>
                <span>
                  <strong className="text-slate-900">TTD Kuasa/Permohonan/BPHTB/Fisik/Keabsahan:</strong>{' '}
                  {preview.summary.penerimaNama}
                </span>
                <span aria-hidden="true" className="text-slate-300">·</span>
                <span>
                  <strong className="text-slate-900">TTD SPPT/PBB:</strong>{' '}
                  {preview.summary.pemilikAwalNama}
                </span>
                <span aria-hidden="true" className="text-slate-300">·</span>
                <span className="tabular-nums">
                  <strong className="text-slate-900">TTD Data Pribadi:</strong>{' '}
                  {preview.summary.dataPribadiSignersCount} pihak (tanpa pasangan)
                </span>
              </>
            ) : (
              <>
                <span>
                  <strong className="text-slate-900">Berkas Non-Akta:</strong>{' '}
                  {record.layanan} (F4 Only · Non-Laporan Bulanan)
                </span>
                <span aria-hidden="true" className="text-slate-300">·</span>
                <span>
                  <strong className="text-slate-900">
                    {record.layanan === 'Turun Waris' ? 'Ahli Waris / Pemohon:' : 'Pemohon:'}
                  </strong>{' '}
                  {nonAktaPemohonNames}
                </span>
                {pengesahanTunggal && (
                  <>
                    <span aria-hidden="true" className="text-slate-300">·</span>
                    <span>
                      <strong className="text-slate-900">Pengesahan Lainnya:</strong>{' '}
                      {pengesahanTunggal.jabatanLines.join(' ')} ({pengesahanTunggal.nama})
                    </span>
                  </>
                )}
                {hasDualPengesahan && pengesahanGanda && (
                  <>
                    <span aria-hidden="true" className="text-slate-300">·</span>
                    <span>
                      <strong className="text-slate-900">Pengesahan SKW/Beda Nama:</strong>{' '}
                      {pengesahanGanda.camatJabatan} &amp; {pengesahanGanda.kadesJabatan}
                    </span>
                  </>
                )}
              </>
            )}
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => onEditStep(2)}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-slate-50 text-xs font-medium text-slate-700 whitespace-nowrap"
            >
              <Edit3 className="w-3 h-3" />
              <span>Revisi Pihak/Tanah</span>
            </button>
            <button
              type="button"
              onClick={() => onEditStep(3)}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-slate-50 text-xs font-medium text-slate-700 whitespace-nowrap"
            >
              <Edit3 className="w-3 h-3" />
              <span>{isAkta ? 'Ubah Parameter Akta' : 'Pengesahan & Parameter'}</span>
            </button>
          </div>
        </div>
      )}

      {/* Page Navigation Bar */}
      {preview && preview.pages.length > 1 && (
        <div className="px-4 py-2 border-b border-slate-200 bg-slate-50/50 flex items-center gap-1.5 overflow-x-auto">
          <button
            type="button"
            onClick={() => setSelectedPage('all')}
            className={`px-2.5 py-1 rounded-md text-xs font-medium whitespace-nowrap shrink-0 transition-colors ${
              selectedPage === 'all'
                ? 'bg-[#0f172a] text-white'
                : 'text-slate-600 hover:bg-slate-200/70'
            }`}
          >
            Semua Bagian ({preview.pages.length})
          </button>
          {preview.pages.map((p) => (
            <button
              key={p.pageNumber}
              type="button"
              onClick={() => setSelectedPage(p.pageNumber)}
              className={`px-2.5 py-1 rounded-md text-xs font-medium whitespace-nowrap shrink-0 transition-colors ${
                selectedPage === p.pageNumber
                  ? 'bg-[#0f172a] text-white'
                  : 'text-slate-600 hover:bg-slate-200/70'
              }`}
            >
              {p.title}
            </button>
          ))}
        </div>
      )}

      {/* Paper Viewport */}
      <div className="bg-slate-100/90 p-4 sm:p-6 max-h-[680px] overflow-y-auto overflow-x-auto space-y-6">
        {loading && (
          <div className="mx-auto bg-white border border-slate-200 rounded-md p-8 max-w-[770px] space-y-3 animate-pulse">
            <div className="h-5 w-1/2 bg-slate-200 rounded mx-auto" />
            <div className="h-3 w-full bg-slate-100 rounded" />
            <div className="h-3 w-11/12 bg-slate-100 rounded" />
            <div className="h-3 w-4/5 bg-slate-100 rounded" />
          </div>
        )}

        {error && !loading && (
          <div
            role="alert"
            className="mx-auto max-w-[770px] p-4 rounded-lg border border-rose-200 bg-rose-50 text-xs text-rose-800 flex items-start gap-2.5"
          >
            <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5 text-rose-600" />
            <div>
              <p className="font-semibold">Gagal Memuat Review Dokumen</p>
              <p className="mt-0.5">{error}</p>
            </div>
          </div>
        )}

        {!loading &&
          !error &&
          visiblePages.map((page) => (
            <div key={page.pageNumber} className="mx-auto" style={{ maxWidth: `${paperMaxWidthPx}px` }}>
              <div className="flex items-center justify-between text-[11px] text-slate-500 mb-1.5 px-1">
                <span className="inline-flex items-center gap-1.5 font-medium text-slate-700">
                  <FileText className="w-3.5 h-3.5 text-slate-500" />
                  <span>{page.title}</span>
                </span>
                <span className="font-mono tabular-nums">
                  {preview?.paperSize} · Bagian {page.pageNumber} / {preview?.pages.length}
                </span>
              </div>

              <div
                className="bg-white border border-slate-300 shadow-xs rounded-[2px] text-slate-900"
                style={{
                  fontFamily: '"Bookman Old Style", "URW Bookman", Georgia, "Times New Roman", serif',
                  padding:
                    preview?.paperSize === 'A4'
                      ? '38px 26px 36px 64px'
                      : '34px 44px 32px 44px',
                }}
              >
                <div className="space-y-0.5">
                  {page.blocks.map((block, bIdx) => (
                    <RenderBlock key={bIdx} block={block} />
                  ))}
                </div>

                {preview && preview.footerLines.length > 0 && (
                  <div className="mt-8 pt-2 border-t border-slate-800 text-[11px] text-slate-800 space-y-0.5">
                    {preview.footerLines.map((line, idx) => (
                      <p key={idx} className={idx === 1 ? 'font-bold italic' : ''}>
                        {line}
                      </p>
                    ))}
                  </div>
                )}
              </div>
            </div>
          ))}
      </div>
    </section>
  );
};
