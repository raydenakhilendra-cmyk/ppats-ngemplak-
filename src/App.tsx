/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useCallback, useEffect, useState } from 'react';
import { CheckCircle2, AlertCircle, X } from 'lucide-react';
import {
  BerkasRecord,
  BerkasStatus,
  CustomTemplateEntry,
  OfficeSettings,
  ServiceType,
  TaxDraftPrefill,
  TemplateDocKey,
} from './types/ppats';
import { DEFAULT_OFFICE_SETTINGS } from './utils/seedData';
import {
  deleteCustomTemplate,
  deleteRecord,
  exportDatabaseBackup,
  importDatabaseBackup,
  loadCustomTemplates,
  loadRecords,
  loadSettings,
  saveCustomTemplate,
  saveRecord,
  saveSettings,
} from './utils/storage';
import { NavSection, Sidebar } from './components/Sidebar';
import { DashboardModule } from './components/DashboardModule';
import { WizardModule } from './components/WizardModule';
import { TaxCalculatorModule } from './components/TaxCalculatorModule';
import { RegisterModule } from './components/RegisterModule';
import { ReportModule } from './components/ReportModule';
import { NonAktaRekapModule } from './components/NonAktaRekapModule';
import { SettingsModule } from './components/SettingsModule';
import { ReceiptModal } from './components/ReceiptModal';

export default function App() {
  const [activeSection, setActiveSection] = useState<NavSection>('dashboard');
  const [records, setRecords] = useState<BerkasRecord[]>([]);
  const [settings, setSettings] = useState<OfficeSettings>(DEFAULT_OFFICE_SETTINGS);
  const [customTemplates, setCustomTemplates] = useState<
    Partial<Record<TemplateDocKey, CustomTemplateEntry>>
  >({});
  const [loading, setLoading] = useState(true);
  const [editingRecord, setEditingRecord] = useState<BerkasRecord | null>(null);
  const [taxPrefill, setTaxPrefill] = useState<TaxDraftPrefill | null>(null);
  const [wizardDefaultLayanan, setWizardDefaultLayanan] = useState<ServiceType | undefined>(
    undefined
  );
  const [selectedTaxRecordId, setSelectedTaxRecordId] = useState<string | null>(null);
  const [wizardKey, setWizardKey] = useState(0);
  const [receiptRecord, setReceiptRecord] = useState<BerkasRecord | null>(null);
  const [isMobileSidebarOpen, setIsMobileSidebarOpen] = useState(false);
  const [toast, setToast] = useState<{
    message: string;
    type: 'success' | 'error';
  } | null>(null);

  const notify = useCallback(
    (message: string, type: 'success' | 'error' = 'success') => {
      setToast({ message, type });
    },
    []
  );

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 5000);
    return () => clearTimeout(t);
  }, [toast]);

  const reloadAll = useCallback(async () => {
    setLoading(true);
    try {
      const [recs, sets, tpls] = await Promise.all([
        loadRecords(),
        loadSettings(),
        loadCustomTemplates(),
      ]);
      setRecords(recs);
      setSettings(sets);
      setCustomTemplates(tpls);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    reloadAll();
  }, [reloadAll]);

  const handleCreateNew = () => {
    setEditingRecord(null);
    setTaxPrefill(null);
    setWizardDefaultLayanan(undefined);
    setWizardKey((k) => k + 1);
    setActiveSection('wizard');
  };

  const handleCreateNewNonAkta = () => {
    setEditingRecord(null);
    setTaxPrefill(null);
    setWizardDefaultLayanan('Turun Waris');
    setWizardKey((k) => k + 1);
    setActiveSection('wizard');
  };

  const handleEditRecord = (rec: BerkasRecord) => {
    setEditingRecord(rec);
    setTaxPrefill(null);
    setWizardDefaultLayanan(undefined);
    setWizardKey((k) => k + 1);
    setActiveSection('wizard');
  };

  const handleOpenTaxForRecord = (rec: BerkasRecord) => {
    setSelectedTaxRecordId(rec.id);
    setActiveSection('tax');
  };

  const handleCreateBerkasWithTax = (prefill: TaxDraftPrefill) => {
    setEditingRecord(null);
    setTaxPrefill(prefill);
    setWizardDefaultLayanan(undefined);
    setWizardKey((k) => k + 1);
    setActiveSection('wizard');
    notify(
      'Parameter luas, PBB, dan hasil perhitungan pajak dimuat ke formulir Buat Berkas Baru.',
      'success'
    );
  };

  const handleOpenTaxFromWizard = async (draftRecord: BerkasRecord) => {
    await handleSaveRecord(draftRecord);
    setSelectedTaxRecordId(draftRecord.id);
    setActiveSection('tax');
    notify(
      `Draft berkas ${draftRecord.kode_berkas} disimpan dan dibuka di jalur Perhitungan Pajak.`,
      'success'
    );
  };

  const handleChangeStatus = async (recordId: string, nextStatus: BerkasStatus) => {
    const target = records.find((r) => r.id === recordId);
    if (!target) return;
    const updated: BerkasRecord = {
      ...target,
      status: nextStatus,
      updated_at: new Date().toISOString(),
    };
    await saveRecord(updated);
    setRecords((prev) => prev.map((r) => (r.id === recordId ? updated : r)));
    notify(`Status berkas ${target.kode_berkas} diubah menjadi "${nextStatus}".`, 'success');
  };

  const handleDeleteRecord = async (recordId: string) => {
    const target = records.find((r) => r.id === recordId);
    await deleteRecord(recordId);
    setRecords((prev) => prev.filter((r) => r.id !== recordId));
    notify(`Berkas ${target?.kode_berkas || ''} dihapus dari register.`, 'success');
  };

  const handleSaveRecord = async (record: BerkasRecord) => {
    await saveRecord(record);
    setRecords((prev) => {
      const exists = prev.some((r) => r.id === record.id);
      if (exists) {
        return prev.map((r) => (r.id === record.id ? record : r));
      }
      return [...prev, record].sort((a, b) => a.no_urut - b.no_urut);
    });
    setEditingRecord(record);
  };

  const handleSaveSettings = async (next: OfficeSettings) => {
    await saveSettings(next);
    setSettings(next);
  };

  const handleUploadTemplate = async (entry: CustomTemplateEntry) => {
    await saveCustomTemplate(entry);
    setCustomTemplates((prev) => ({ ...prev, [entry.key]: entry }));
  };

  const handleDeleteTemplate = async (key: TemplateDocKey) => {
    await deleteCustomTemplate(key);
    setCustomTemplates((prev) => {
      const copy = { ...prev };
      delete copy[key];
      return copy;
    });
    notify('Template dikembalikan ke template bawaan sistem.', 'success');
  };

  const handleExportBackup = async () => {
    try {
      const payload = await exportDatabaseBackup();
      const blob = new Blob([JSON.stringify(payload, null, 2)], {
        type: 'application/json',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Backup_SIM_PPATS_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      notify(`Backup berhasil diekspor (${payload.records.length} berkas).`, 'success');
    } catch (err) {
      notify(
        err instanceof Error ? err.message : 'Gagal mengekspor backup JSON.',
        'error'
      );
    }
  };

  const handleRestoreBackup = async (file: File) => {
    try {
      const text = await file.text();
      const res = await importDatabaseBackup(text);
      await reloadAll();
      notify(
        `Restore berhasil: ${res.recordCount} berkas dan ${res.templateCount} template dimuat.`,
        'success'
      );
    } catch (err) {
      notify(
        err instanceof Error ? err.message : 'Gagal me-restore file backup JSON.',
        'error'
      );
    }
  };

  return (
    <div className="min-h-screen flex flex-col lg:flex-row bg-[#f8fafc] text-slate-900">
      <Sidebar
        activeSection={activeSection}
        onSelectSection={(sec) => {
          if (sec === 'wizard' && activeSection !== 'wizard' && !editingRecord) {
            setWizardDefaultLayanan(undefined);
            setWizardKey((k) => k + 1);
          }
          if (sec === 'tax') {
            setSelectedTaxRecordId(null);
          }
          setActiveSection(sec);
        }}
        settings={settings}
        onExportBackup={handleExportBackup}
        onRestoreBackup={handleRestoreBackup}
        isMobileOpen={isMobileSidebarOpen}
        onToggleMobile={() => setIsMobileSidebarOpen((v) => !v)}
        editingBadge={editingRecord?.kode_berkas ?? null}
      />

      <main className="flex-1 min-w-0 p-4 sm:p-6">
        {toast && (
          <div
            role="status"
            className={`no-print mb-5 px-4 py-3 rounded-xl border text-sm flex items-center justify-between gap-3 ${
              toast.type === 'success'
                ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                : 'bg-rose-50 border-rose-200 text-rose-900'
            }`}
          >
            <div className="flex items-center gap-2.5">
              {toast.type === 'success' ? (
                <CheckCircle2 className="w-4.5 h-4.5 text-emerald-600 shrink-0" />
              ) : (
                <AlertCircle className="w-4.5 h-4.5 text-rose-600 shrink-0" />
              )}
              <span className="font-medium">{toast.message}</span>
            </div>
            <button
              type="button"
              onClick={() => setToast(null)}
              aria-label="Tutup notifikasi"
              className="p-1 rounded hover:bg-black/5"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {loading ? (
          <div className="space-y-4">
            <div className="h-8 w-64 bg-slate-200 rounded animate-pulse" />
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              {[1, 2, 3, 4].map((i) => (
                <div key={i} className="h-28 bg-white border border-slate-200 rounded-xl animate-pulse" />
              ))}
            </div>
          </div>
        ) : (
          <>
            {activeSection === 'dashboard' && (
              <DashboardModule
                records={records}
                onCreateNew={handleCreateNew}
                onEditRecord={handleEditRecord}
                onCalculateTax={handleOpenTaxForRecord}
                onChangeStatus={handleChangeStatus}
                onPrintReceipt={(rec) => setReceiptRecord(rec)}
                onDeleteRecord={handleDeleteRecord}
              />
            )}

            {activeSection === 'wizard' && (
              <WizardModule
                key={wizardKey}
                existingRecords={records}
                editingRecord={editingRecord}
                taxPrefill={taxPrefill}
                initialLayanan={wizardDefaultLayanan}
                settings={settings}
                customTemplates={customTemplates}
                onSaveRecord={handleSaveRecord}
                onCancelEdit={() => {
                  setEditingRecord(null);
                  setTaxPrefill(null);
                  setWizardDefaultLayanan(undefined);
                  setWizardKey((k) => k + 1);
                }}
                onPrintReceipt={(rec) => setReceiptRecord(rec)}
                onOpenTaxCalculator={handleOpenTaxFromWizard}
                onNotify={notify}
              />
            )}

            {activeSection === 'tax' && (
              <TaxCalculatorModule
                records={records}
                settings={settings}
                initialRecordId={selectedTaxRecordId}
                onSaveRecord={handleSaveRecord}
                onCreateBerkasWithTax={handleCreateBerkasWithTax}
                onEditRecord={handleEditRecord}
                onNotify={notify}
              />
            )}

            {activeSection === 'register' && (
              <RegisterModule
                records={records}
                onEditRecord={handleEditRecord}
              />
            )}

            {activeSection === 'report' && (
              <ReportModule
                records={records}
                settings={settings}
                onOpenNonAktaRekap={() => setActiveSection('non_akta_rekap')}
                onNotify={notify}
              />
            )}

            {activeSection === 'non_akta_rekap' && (
              <NonAktaRekapModule
                records={records}
                settings={settings}
                onCreateNewNonAkta={handleCreateNewNonAkta}
                onEditRecord={handleEditRecord}
                onChangeStatus={handleChangeStatus}
                onPrintReceipt={(rec) => setReceiptRecord(rec)}
                onDeleteRecord={handleDeleteRecord}
                onNotify={notify}
              />
            )}

            {activeSection === 'settings' && (
              <SettingsModule
                settings={settings}
                customTemplates={customTemplates}
                onSaveSettings={handleSaveSettings}
                onUploadTemplate={handleUploadTemplate}
                onDeleteTemplate={handleDeleteTemplate}
                onExportBackup={handleExportBackup}
                onRestoreBackup={handleRestoreBackup}
                onNotify={notify}
              />
            )}
          </>
        )}
      </main>

      <ReceiptModal
        record={receiptRecord}
        settings={settings}
        onClose={() => setReceiptRecord(null)}
      />
    </div>
  );
}
