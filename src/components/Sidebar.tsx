import React from 'react';
import {
  LayoutDashboard,
  FilePlus2,
  Calculator,
  BookOpenCheck,
  FileSpreadsheet,
  ClipboardList,
  Settings,
  Download,
  Upload,
  Menu,
  X,
  Scale,
} from 'lucide-react';
import { OfficeSettings } from '../types/ppats';

export type NavSection =
  | 'dashboard'
  | 'wizard'
  | 'tax'
  | 'register'
  | 'report'
  | 'non_akta_rekap'
  | 'settings';

interface SidebarProps {
  activeSection: NavSection;
  onSelectSection: (section: NavSection) => void;
  settings: OfficeSettings;
  onExportBackup: () => void;
  onRestoreBackup: (file: File) => void;
  isMobileOpen: boolean;
  onToggleMobile: () => void;
  editingBadge?: string | null;
}

const NAV_ITEMS: Array<{
  id: NavSection;
  label: string;
  icon: React.ComponentType<{ className?: string }>;
}> = [
  { id: 'dashboard', label: 'Dashboard & Tracking', icon: LayoutDashboard },
  { id: 'wizard', label: 'Buat Berkas Baru', icon: FilePlus2 },
  { id: 'tax', label: 'Perhitungan Pajak', icon: Calculator },
  { id: 'register', label: 'Buku Register Akta', icon: BookOpenCheck },
  { id: 'report', label: 'Laporan Bulanan PPATS', icon: FileSpreadsheet },
  { id: 'non_akta_rekap', label: 'Rekap Non-Akta', icon: ClipboardList },
  { id: 'settings', label: 'Pengaturan & Template', icon: Settings },
];

export const Sidebar: React.FC<SidebarProps> = ({
  activeSection,
  onSelectSection,
  settings,
  onExportBackup,
  onRestoreBackup,
  isMobileOpen,
  onToggleMobile,
  editingBadge,
}) => {
  const restoreInputRef = React.useRef<HTMLInputElement | null>(null);

  const handleRestoreChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onRestoreBackup(file);
      e.target.value = '';
    }
  };

  return (
    <>
      {/* Mobile Header Bar */}
      <div className="no-print lg:hidden flex items-center justify-between bg-[#0f172a] text-white px-4 py-3 border-b border-slate-800 sticky top-0 z-30">
        <div className="flex items-center gap-2.5">
          <Scale className="w-5 h-5 text-amber-400 shrink-0" />
          <span className="font-display font-bold text-base tracking-tight whitespace-nowrap">
            SIM-PPATS
          </span>
        </div>
        <button
          type="button"
          onClick={onToggleMobile}
          aria-label={isMobileOpen ? 'Tutup menu navigasi' : 'Buka menu navigasi'}
          className="p-2 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 transition-colors"
        >
          {isMobileOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
        </button>
      </div>

      {/* Backdrop for mobile drawer */}
      {isMobileOpen && (
        <div
          className="no-print fixed inset-0 bg-slate-950/50 z-30 lg:hidden"
          onClick={onToggleMobile}
          aria-hidden="true"
        />
      )}

      {/* Left Navigation Sidebar */}
      <aside
        aria-label="Navigasi Utama SIM-PPATS"
        className={`no-print fixed lg:static inset-y-0 left-0 z-40 w-64 lg:w-60 shrink-0 bg-[#0f172a] text-slate-100 flex flex-col border-r border-slate-800 transition-transform duration-100 ${
          isMobileOpen
            ? 'translate-x-0 visible'
            : '-translate-x-full invisible lg:visible lg:translate-x-0'
        }`}
      >
        {/* Office Brand Header */}
        <div className="px-4 py-5 border-b border-slate-800/80">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-lg bg-slate-800 border border-slate-700 flex items-center justify-center shrink-0">
              <Scale className="w-4.5 h-4.5 text-amber-400" />
            </div>
            <div className="min-w-0">
              <div className="font-display font-bold text-base tracking-tight text-white truncate">
                SIM-PPATS
              </div>
              <p className="text-xs text-slate-400 truncate">
                {settings.nama_kecamatan}
              </p>
            </div>
          </div>
        </div>

        {/* Navigation Links */}
        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
          {NAV_ITEMS.map((item) => {
            const Icon = item.icon;
            const isActive = activeSection === item.id;
            const label =
              item.id === 'wizard' && editingBadge
                ? 'Edit Berkas Aktif'
                : item.label;

            return (
              <button
                key={item.id}
                type="button"
                onClick={() => {
                  onSelectSection(item.id);
                  if (isMobileOpen) onToggleMobile();
                }}
                aria-current={isActive ? 'page' : undefined}
                className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors duration-75 text-left whitespace-nowrap ${
                  isActive
                    ? 'bg-sky-600 text-white shadow-sm'
                    : 'text-slate-300 hover:bg-slate-800/70 hover:text-white'
                }`}
              >
                <Icon className="w-4 h-4 shrink-0" />
                <span className="truncate">{label}</span>
              </button>
            );
          })}
        </nav>

        {/* Office Info & Backup Quick Actions */}
        <div className="p-4 border-t border-slate-800/80 space-y-3">
          <div className="text-xs text-slate-400 space-y-0.5">
            <p className="font-medium text-slate-200 truncate">
              {settings.nama_ppats}
            </p>
            <p className="truncate">{settings.nama_kabupaten}</p>
          </div>

          <div className="grid grid-cols-2 gap-2">
            <button
              type="button"
              onClick={onExportBackup}
              className="flex items-center justify-center gap-1.5 px-2.5 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 transition-colors whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5 shrink-0" />
              <span>Backup</span>
            </button>
            <button
              type="button"
              onClick={() => restoreInputRef.current?.click()}
              className="flex items-center justify-center gap-1.5 px-2.5 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-200 transition-colors whitespace-nowrap"
            >
              <Upload className="w-3.5 h-3.5 shrink-0" />
              <span>Restore</span>
            </button>
            <input
              ref={restoreInputRef}
              type="file"
              accept="application/json,.json"
              onChange={handleRestoreChange}
              className="hidden"
              aria-label="Pilih file backup JSON"
            />
          </div>
        </div>
      </aside>
    </>
  );
};
