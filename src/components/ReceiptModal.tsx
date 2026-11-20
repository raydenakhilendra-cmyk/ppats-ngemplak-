import React from 'react';
import { Printer, X, FileText } from 'lucide-react';
import { BerkasRecord, OfficeSettings } from '../types/ppats';
import {
  formatNumberId,
  formatRupiah,
  formatTanggalDDMMYYYY,
} from '../utils/indonesianFormatter';

interface ReceiptModalProps {
  record: BerkasRecord | null;
  settings: OfficeSettings;
  onClose: () => void;
}

export const ReceiptModal: React.FC<ReceiptModalProps> = ({
  record,
  settings,
  onClose,
}) => {
  if (!record) return null;

  const p1 = record.para_pihak
    .filter((p) => p.peran === 'Pihak I')
    .map((p) => p.nama)
    .join(', ');
  const p2 = record.para_pihak
    .filter((p) => p.peran === 'Pihak II')
    .map((p) => p.nama)
    .join(', ');
  const tglMasuk = formatTanggalDDMMYYYY(record.tanggal_masuk);
  const tglAkta = record.akta.tanggal_akta
    ? formatTanggalDDMMYYYY(record.akta.tanggal_akta)
    : 'Belum bertanggal';

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-950/60 flex items-center justify-center p-4 overflow-y-auto"
      role="dialog"
      aria-modal="true"
      aria-labelledby="receipt-modal-title"
    >
      <div className="bg-white border border-slate-200 rounded-xl w-full max-w-2xl shadow-xl overflow-hidden">
        <div className="no-print flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-slate-50">
          <div className="flex items-center gap-2">
            <FileText className="w-4.5 h-4.5 text-slate-700" />
            <h2
              id="receipt-modal-title"
              className="font-display font-semibold text-base text-slate-900"
            >
              Tanda Terima Pemohon — {record.kode_berkas}
            </h2>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => window.print()}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg bg-[#0f172a] hover:bg-slate-800 text-white text-xs font-medium transition-colors whitespace-nowrap"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Cetak / Simpan PDF</span>
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Tutup jendela tanda terima"
              className="p-2 rounded-lg text-slate-500 hover:text-slate-900 hover:bg-slate-200/70 transition-colors"
            >
              <X className="w-4.5 h-4.5" />
            </button>
          </div>
        </div>

        <div className="p-8 space-y-6 text-slate-900 text-sm">
          {/* Kop Kantor */}
          <div className="text-center border-b-2 border-slate-900 pb-4 space-y-0.5">
            <p className="text-xs font-semibold tracking-wide uppercase text-slate-600">
              Pemerintah {settings.nama_kabupaten}
            </p>
            <h3 className="font-display font-bold text-base uppercase tracking-tight">
              Pejabat Pembuat Akta Tanah Sementara (PPATS)
            </h3>
            <p className="font-semibold text-sm uppercase">{settings.nama_kecamatan}</p>
            <p className="text-xs text-slate-600">
              {settings.alamat_kantor} · Telp. {settings.telepon_kantor}
            </p>
          </div>

          <div className="text-center">
            <h4 className="font-display font-bold text-base underline underline-offset-4">
              TANDA TERIMA BERKAS PERMOHONAN AKTA
            </h4>
            <p className="font-mono text-xs text-slate-600 mt-1 tabular-nums">
              Kode Berkas: {record.kode_berkas} · Status: {record.status}
            </p>
          </div>

          <table className="w-full text-sm border-collapse">
            <tbody className="divide-y divide-slate-200">
              <tr>
                <td className="py-2 pr-4 text-slate-500 w-44">Tanggal Loket Masuk</td>
                <td className="py-2 font-medium">{tglMasuk}</td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Jenis Layanan Akta</td>
                <td className="py-2 font-semibold">{record.layanan}</td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Nomor & Tanggal Akta</td>
                <td className="py-2 font-mono tabular-nums">
                  {record.akta.nomor_akta || '-'} ({tglAkta})
                </td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Pihak I (Pelepas)</td>
                <td className="py-2 font-medium">{p1 || '-'}</td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Pihak II (Penerima)</td>
                <td className="py-2 font-medium">{p2 || '-'}</td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Objek Tanah</td>
                <td className="py-2">
                  {record.tanah.jenis_sertipikat === 'digital' || !record.tanah.nomor_hak_lama ? (
                    <>
                      {record.tanah.jenis_hak} · NIBel{' '}
                      <span className="font-mono">{record.tanah.nib || '-'}</span>
                    </>
                  ) : (
                    <>
                      {record.tanah.jenis_hak} No. {record.tanah.nomor_hak_lama} · NIB{' '}
                      <span className="font-mono">{record.tanah.nib || '-'}</span>
                    </>
                  )}{' '}
                  · Luas{' '}
                  <span className="font-mono tabular-nums">
                    {formatNumberId(record.tanah.luas_m2)} m²
                  </span>
                </td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Letak Tanah</td>
                <td className="py-2">
                  Desa/Kel. {record.tanah.desa}, Kec. {record.tanah.kecamatan},{' '}
                  {record.tanah.kabupaten}
                </td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Estimasi BPHTB / PPh</td>
                <td className="py-2 font-mono tabular-nums">
                  BPHTB {formatRupiah(record.pajak.bphtb)} · PPh {formatRupiah(record.pajak.pph)}
                </td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Kuasa Pengurusan BPN</td>
                <td className="py-2">{record.akta.nama_kuasa_bpn || '-'}</td>
              </tr>
              <tr>
                <td className="py-2 pr-4 text-slate-500">Catatan Loket</td>
                <td className="py-2 text-slate-700">{record.akta.catatan_loket || '-'}</td>
              </tr>
            </tbody>
          </table>

          <div className="grid grid-cols-2 gap-8 pt-6 text-center text-xs">
            <div>
              <p className="text-slate-600">Pemohon / Penyerah Berkas,</p>
              <div className="h-16" />
              <p className="font-semibold border-t border-slate-400 pt-1 inline-block px-6">
                {p2 || p1 || 'Pemohon'}
              </p>
            </div>
            <div>
              <p className="text-slate-600">
                {settings.nama_kecamatan}, {tglMasuk}
              </p>
              <p className="text-slate-600">Petugas Loket / PPATS,</p>
              <div className="h-12" />
              <p className="font-semibold underline">{settings.nama_ppats}</p>
              <p className="text-slate-500 font-mono">NIP. {settings.nip_ppats}</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
