import { useEffect, useMemo, useState } from 'react';
import { CheckCircle2, AlertCircle, X } from 'lucide-react';
import type {
  ActiveTab,
  AttendanceConfig,
  Candidate,
  ManualCellEdit,
  PreviousAttendanceData,
  PythonFeatureFlags,
  SavedProfile,
} from './types/attendance';
import {
  DEFAULT_CONFIG,
  generateAttendanceBlocks,
  SAMPLE_CANDIDATES,
} from './utils/attendanceEngine';
import { exportAttendanceCSV, exportAttendanceExcel } from './utils/excelExporter';
import { DEFAULT_PYTHON_FLAGS } from './utils/pythonGenerator';
import { TopBar } from './components/TopBar';
import { BuilderView } from './components/BuilderView';
import { SheetPreviewView } from './components/SheetPreviewView';
import { AnalyticsView } from './components/AnalyticsView';
import { PythonCodeView } from './components/PythonCodeView';

const PROFILES_STORAGE_KEY = 'biometric_attendance_studio_profiles_v1';

function loadSavedProfiles(): SavedProfile[] {
  try {
    const raw = localStorage.getItem(PROFILES_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export default function App() {
  const [activeTab, setActiveTab] = useState<ActiveTab>('builder');
  const [config, setConfig] = useState<AttendanceConfig>(DEFAULT_CONFIG);
  const [candidates, setCandidates] = useState<Candidate[]>(SAMPLE_CANDIDATES);
  const [previousData, setPreviousData] = useState<PreviousAttendanceData | null>(null);
  const [manualEdits, setManualEdits] = useState<Record<string, ManualCellEdit>>({});
  const [profiles, setProfiles] = useState<SavedProfile[]>(loadSavedProfiles);
  const [pythonFlags, setPythonFlags] = useState<PythonFeatureFlags>(DEFAULT_PYTHON_FLAGS);
  const [isExporting, setIsExporting] = useState(false);
  const [toast, setToast] = useState<{ type: 'success' | 'error'; message: string } | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem(PROFILES_STORAGE_KEY, JSON.stringify(profiles));
    } catch {
      // Ignore storage quota errors
    }
  }, [profiles]);

  const showToast = (type: 'success' | 'error', message: string) => {
    setToast({ type, message });
    setTimeout(() => {
      setToast((cur) => (cur?.message === message ? null : cur));
    }, 4000);
  };

  const { blocks, error: generationError } = useMemo(
    () => generateAttendanceBlocks(candidates, config, manualEdits, previousData),
    [candidates, config, manualEdits, previousData]
  );

  const handleExportExcel = async () => {
    if (!blocks.length) {
      showToast('error', generationError || 'Nothing to export.');
      return;
    }
    setIsExporting(true);
    try {
      const filename = await exportAttendanceExcel(blocks, config);
      showToast('success', `Exported ${filename}`);
    } catch (err) {
      showToast('error', err instanceof Error ? err.message : 'Excel export failed.');
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportCSV = () => {
    if (!blocks.length) {
      showToast('error', generationError || 'Nothing to export.');
      return;
    }
    try {
      const filename = exportAttendanceCSV(blocks, config);
      showToast('success', `Exported ${filename}`);
    } catch (err) {
      showToast('error', err instanceof Error ? err.message : 'CSV export failed.');
    }
  };

  const handleSaveProfile = (name: string) => {
    const profile: SavedProfile = {
      id: `prof-${Date.now()}`,
      name,
      savedAt: new Date().toISOString(),
      config: JSON.parse(JSON.stringify(config)),
    };
    setProfiles((prev) => [profile, ...prev].slice(0, 12));
    showToast('success', `Saved profile "${name}"`);
  };

  const handleLoadProfile = (profile: SavedProfile) => {
    setConfig(JSON.parse(JSON.stringify(profile.config)));
    setManualEdits({});
    showToast('success', `Loaded profile "${profile.name}"`);
  };

  const handleDeleteProfile = (id: string) => {
    setProfiles((prev) => prev.filter((p) => p.id !== id));
  };

  const handleResetAll = () => {
    setConfig(DEFAULT_CONFIG);
    setCandidates(SAMPLE_CANDIDATES);
    setPreviousData(null);
    setManualEdits({});
    showToast('success', 'Reset configuration and demo roster to defaults.');
  };

  const handleRerollSeed = () => {
    setConfig((prev) => ({
      ...prev,
      seed: Math.floor(100000 + Math.random() * 900000),
    }));
    setManualEdits({});
    showToast(
      'success',
      previousData
        ? 'Re-rolled new dates (preserved previous Excel timestamps).'
        : 'Generated fresh biometric IN/OUT timestamps.'
    );
  };

  const handleEditCell = (key: string, edit: ManualCellEdit) => {
    setManualEdits((prev) => ({ ...prev, [key]: edit }));
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      <TopBar
        activeTab={activeTab}
        onSelectTab={setActiveTab}
        onExportExcel={handleExportExcel}
        isExporting={isExporting}
        canExport={blocks.length > 0}
      />

      <main className="flex-1">
        {activeTab === 'builder' && (
          <BuilderView
            config={config}
            candidates={candidates}
            previousData={previousData}
            profiles={profiles}
            onChangeConfig={setConfig}
            onChangeCandidates={setCandidates}
            onChangePreviousData={(data) => {
              setPreviousData(data);
              if (data) {
                showToast(
                  'success',
                  `Loaded ${data.datesCount} previous dates for ${data.candidateNames.length} trainees from ${data.fileName}`
                );
              }
            }}
            onSaveProfile={handleSaveProfile}
            onLoadProfile={handleLoadProfile}
            onDeleteProfile={handleDeleteProfile}
            onResetAll={handleResetAll}
            onOpenPreview={() => setActiveTab('preview')}
            onOpenPython={() => setActiveTab('python')}
            generationError={generationError}
          />
        )}

        {activeTab === 'preview' && (
          <SheetPreviewView
            blocks={blocks}
            config={config}
            previousData={previousData}
            manualEditCount={Object.keys(manualEdits).length}
            onRerollSeed={handleRerollSeed}
            onEditCell={handleEditCell}
            onClearManualEdits={() => setManualEdits({})}
            onExportExcel={handleExportExcel}
            onExportCSV={handleExportCSV}
            error={generationError}
          />
        )}

        {activeTab === 'analytics' && (
          <AnalyticsView blocks={blocks} config={config} error={generationError} />
        )}

        {activeTab === 'python' && (
          <PythonCodeView
            config={config}
            flags={pythonFlags}
            onChangeFlags={setPythonFlags}
          />
        )}
      </main>

      {/* Non-blocking feedback toast */}
      {toast && (
        <div
          role="status"
          className={`fixed bottom-5 right-5 z-50 flex items-center gap-2.5 px-4 py-3 text-sm font-medium rounded-lg shadow-lg border no-print ${
            toast.type === 'success'
              ? 'bg-slate-900 text-white border-slate-800'
              : 'bg-red-700 text-white border-red-800'
          }`}
        >
          {toast.type === 'success' ? (
            <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          ) : (
            <AlertCircle className="w-4 h-4 text-red-200 shrink-0" />
          )}
          <span>{toast.message}</span>
          <button
            type="button"
            onClick={() => setToast(null)}
            aria-label="Dismiss notification"
            className="ml-2 text-slate-400 hover:text-white cursor-pointer"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}
    </div>
  );
}
