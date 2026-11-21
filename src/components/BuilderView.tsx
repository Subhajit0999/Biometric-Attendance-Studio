import React, { useRef, useState } from 'react';
import {
  Calendar,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Plus,
  RefreshCw,
  RotateCcw,
  Save,
  Search,
  Settings2,
  Trash2,
  Upload,
  UserPlus,
  X,
} from 'lucide-react';
import type {
  AbsentPolicy,
  AttendanceConfig,
  Candidate,
  ExcelTheme,
  GenerationMode,
  PreviousAttendanceData,
  SavedProfile,
  TimeDistribution,
  TimeFormatOption,
  WeekendPolicy,
} from '../types/attendance';
import {
  extractCandidatesFromParsedFile,
  formatDisplayDate,
  formatISODate,
  INDIAN_HOLIDAY_PRESETS,
  MONTH_NAMES,
  parseISODate,
  parsePreviousAttendanceExcel,
  parseRosterFile,
  SAMPLE_CANDIDATES,
  SHIFT_PRESETS,
  type ParsedRosterFile,
} from '../utils/attendanceEngine';
import { downloadSampleRosterExcel } from '../utils/excelExporter';

interface BuilderViewProps {
  config: AttendanceConfig;
  candidates: Candidate[];
  previousData: PreviousAttendanceData | null;
  profiles: SavedProfile[];
  onChangeConfig: (updater: (prev: AttendanceConfig) => AttendanceConfig) => void;
  onChangeCandidates: (updater: (prev: Candidate[]) => Candidate[]) => void;
  onChangePreviousData: (data: PreviousAttendanceData | null) => void;
  onSaveProfile: (name: string) => void;
  onLoadProfile: (profile: SavedProfile) => void;
  onDeleteProfile: (id: string) => void;
  onResetAll: () => void;
  onOpenPreview: () => void;
  onOpenPython: () => void;
  generationError?: string;
}

const WEEKDAY_LABELS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

export function BuilderView({
  config,
  candidates,
  previousData,
  profiles,
  onChangeConfig,
  onChangeCandidates,
  onChangePreviousData,
  onSaveProfile,
  onLoadProfile,
  onDeleteProfile,
  onResetAll,
  onOpenPreview,
  onOpenPython,
  generationError,
}: BuilderViewProps) {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const folderInputRef = useRef<HTMLInputElement | null>(null);
  const prevExcelInputRef = useRef<HTMLInputElement | null>(null);
  const jsonInputRef = useRef<HTMLInputElement | null>(null);
  const [prevExcelError, setPrevExcelError] = useState('');

  const [rosterSearch, setRosterSearch] = useState('');
  const [newCandName, setNewCandName] = useState('');
  const [newCandRoll, setNewCandRoll] = useState('');
  const [showBulkPaste, setShowBulkPaste] = useState(false);
  const [bulkText, setBulkText] = useState('');
  const [parsedFile, setParsedFile] = useState<ParsedRosterFile | null>(null);
  const [selectedNameCol, setSelectedNameCol] = useState('');
  const [selectedRollCol, setSelectedRollCol] = useState('');
  const [deduplicateImport, setDeduplicateImport] = useState(true);
  const [importError, setImportError] = useState('');

  const [newHolDate, setNewHolDate] = useState(config.startDate);
  const [newHolLabel, setNewHolLabel] = useState('');
  const [profileNameInput, setProfileNameInput] = useState('');
  const [editingCandidate, setEditingCandidate] = useState<Candidate | null>(null);
  const [overrideAbsentDate, setOverrideAbsentDate] = useState('');

  const updateField = <K extends keyof AttendanceConfig>(key: K, value: AttendanceConfig[K]) => {
    onChangeConfig((prev) => ({ ...prev, [key]: value }));
  };

  const handleFileUpload = async (file: File) => {
    setImportError('');
    try {
      const parsed = await parseRosterFile(file);
      setParsedFile(parsed);
      setSelectedNameCol(parsed.detectedNameColumn);
      setSelectedRollCol(parsed.detectedRollColumn || '');
      const extracted = extractCandidatesFromParsedFile(
        parsed,
        parsed.detectedNameColumn,
        parsed.detectedRollColumn,
        deduplicateImport
      );
      if (!extracted.length) {
        setImportError('No candidate names were found in the detected column. Pick another column below.');
      } else {
        onChangeCandidates(() => extracted);
      }
    } catch (err) {
      setImportError(err instanceof Error ? err.message : 'Failed to parse uploaded file.');
    }
  };

  const handleFolderAutoFetch = async (fileList: FileList | null) => {
    if (!fileList || fileList.length === 0) return;
    setImportError('');
    const validExts = ['.xlsx', '.xls', '.xlsm', '.csv'];
    const candidatesFiles = Array.from(fileList).filter((f) => {
      const lower = f.name.toLowerCase();
      return (
        validExts.some((ext) => lower.endsWith(ext)) &&
        !lower.startsWith('~$') &&
        !lower.includes('_attendance_')
      );
    });

    if (!candidatesFiles.length) {
      setImportError('No Excel (.xlsx/.xls/.xlsm) or CSV roster files found in that folder.');
      return;
    }

    let bestParsed: ParsedRosterFile | null = null;
    let bestExtracted: Candidate[] = [];

    for (const file of candidatesFiles) {
      try {
        const parsed = await parseRosterFile(file);
        const extracted = extractCandidatesFromParsedFile(
          parsed,
          parsed.detectedNameColumn,
          parsed.detectedRollColumn,
          deduplicateImport
        );
        if (extracted.length > bestExtracted.length) {
          bestParsed = parsed;
          bestExtracted = extracted;
        }
      } catch {
        // Skip unreadable files in the folder
      }
    }

    if (!bestParsed || !bestExtracted.length) {
      setImportError('Found spreadsheet files in the folder, but none contained readable candidate names.');
      return;
    }

    setParsedFile(bestParsed);
    setSelectedNameCol(bestParsed.detectedNameColumn);
    setSelectedRollCol(bestParsed.detectedRollColumn || '');
    onChangeCandidates(() => bestExtracted);
  };

  const getMonthsBetween = (startIso: string, endIso: string): number[] => {
    const s = parseISODate(startIso);
    const e = parseISODate(endIso);
    const months = new Set<number>();
    let curY = s.getFullYear();
    let curM = s.getMonth() + 1;
    let guard = 0;
    while ((curY < e.getFullYear() || (curY === e.getFullYear() && curM <= e.getMonth() + 1)) && guard < 36) {
      months.add(curM);
      if (curM === 12) {
        curY++;
        curM = 1;
      } else {
        curM++;
      }
      guard++;
    }
    return Array.from(months).sort((a, b) => a - b);
  };

  const handleExtendMonths = (extraMonths: number) => {
    const anchorIso = previousData ? previousData.maxDate : config.endDate;
    const anchor = parseISODate(anchorIso);
    // Last day of (anchor month + extraMonths)
    const targetEnd = new Date(anchor.getFullYear(), anchor.getMonth() + extraMonths + 1, 0);
    const nextStart = previousData ? previousData.minDate : config.startDate;
    const nextEnd = formatISODate(targetEnd);
    onChangeConfig((prev) => ({
      ...prev,
      startDate: nextStart,
      endDate: nextEnd,
      selectedMonths: getMonthsBetween(nextStart, nextEnd),
    }));
  };

  const handleUploadPreviousExcel = async (file: File) => {
    setPrevExcelError('');
    try {
      const parsed = await parsePreviousAttendanceExcel(file);
      onChangePreviousData(parsed);

      // Load candidates from the previous attendance Excel preserving exact order
      const prevCandidates: Candidate[] = parsed.candidateNames.map((name, idx) => ({
        id: `prev-${Date.now()}-${idx}`,
        name,
      }));
      onChangeCandidates(() => prevCandidates);

      // Expand Start Date to previous minDate, and auto-extend End Date by 1 month if current endDate <= maxDate
      const prevMax = parseISODate(parsed.maxDate);
      const nextMonthEnd = formatISODate(new Date(prevMax.getFullYear(), prevMax.getMonth() + 2, 0));

      onChangeConfig((prev) => {
        const nextStart = parsed.minDate;
        const nextEnd = prev.endDate > parsed.maxDate ? prev.endDate : nextMonthEnd;
        const meta = parsed.detectedMeta || {};
        return {
          ...prev,
          centerName: prev.centerName.trim() || meta.centerName || '',
          district: prev.district.trim() || meta.district || '',
          batchNo: prev.batchNo.trim() || meta.batchNo || '',
          location: prev.location.trim() || meta.location || '',
          qualificationPack: prev.qualificationPack.trim() || meta.qualificationPack || '',
          startDate: nextStart,
          endDate: nextEnd,
          selectedMonths: getMonthsBetween(nextStart, nextEnd),
        };
      });
    } catch (err) {
      setPrevExcelError(
        err instanceof Error ? err.message : 'Failed to read previous attendance Excel file.'
      );
    }
  };

  const reApplyColumnSelection = (nameCol: string, rollCol: string, dedup: boolean) => {
    setSelectedNameCol(nameCol);
    setSelectedRollCol(rollCol);
    setDeduplicateImport(dedup);
    if (!parsedFile) return;
    const extracted = extractCandidatesFromParsedFile(
      parsedFile,
      nameCol,
      rollCol || null,
      dedup
    );
    onChangeCandidates(() => extracted);
  };

  const handleAddSingleCandidate = (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = newCandName.trim();
    if (!trimmed) return;
    onChangeCandidates((prev) => [
      ...prev,
      {
        id: `c-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        name: trimmed,
        rollNo: newCandRoll.trim() || undefined,
      },
    ]);
    setNewCandName('');
    setNewCandRoll('');
  };

  const handleBulkImport = () => {
    const lines = bulkText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    if (!lines.length) return;
    const newItems: Candidate[] = lines.map((line, i) => {
      const parts = line.split(/[,\t]+/).map((p) => p.trim());
      if (parts.length >= 2) {
        return {
          id: `bulk-${Date.now()}-${i}`,
          rollNo: parts[0],
          name: parts.slice(1).join(' '),
        };
      }
      return {
        id: `bulk-${Date.now()}-${i}`,
        name: parts[0],
      };
    });
    onChangeCandidates((prev) => [...prev, ...newItems]);
    setBulkText('');
    setShowBulkPaste(false);
  };

  const handleAddHoliday = () => {
    if (!newHolDate) return;
    if (config.holidays.some((h) => h.date === newHolDate)) return;
    updateField(
      'holidays',
      [...config.holidays, { date: newHolDate, label: newHolLabel.trim() || 'Public Holiday' }].sort(
        (a, b) => a.date.localeCompare(b.date)
      )
    );
    setNewHolLabel('');
  };

  const handleAddNationalHolidaysInRange = () => {
    const existing = new Set(config.holidays.map((h) => h.date));
    const inRange = INDIAN_HOLIDAY_PRESETS.filter(
      (h) => h.date >= config.startDate && h.date <= config.endDate && !existing.has(h.date)
    );
    if (!inRange.length) {
      // Fallback: add all for the start year
      const yr = config.startDate.slice(0, 4);
      const forYear = INDIAN_HOLIDAY_PRESETS.filter(
        (h) => h.date.startsWith(yr) && !existing.has(h.date)
      );
      if (forYear.length) {
        updateField(
          'holidays',
          [...config.holidays, ...forYear].sort((a, b) => a.date.localeCompare(b.date))
        );
      }
      return;
    }
    updateField(
      'holidays',
      [...config.holidays, ...inRange].sort((a, b) => a.date.localeCompare(b.date))
    );
  };

  const handleExportProfileJson = () => {
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${(config.batchNo || 'attendance').replace(/\s+/g, '_')}_profile.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImportProfileJson = async (file: File) => {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as Partial<AttendanceConfig>;
      onChangeConfig((prev) => ({ ...prev, ...parsed }));
    } catch {
      setImportError('Invalid JSON profile file.');
    }
  };

  const toggleMonth = (m: number) => {
    const has = config.selectedMonths.includes(m);
    const next = has
      ? config.selectedMonths.filter((x) => x !== m)
      : [...config.selectedMonths, m].sort((a, b) => a - b);
    updateField('selectedMonths', next);
  };

  const toggleCustomWeekendDay = (dow: number) => {
    const has = config.customWeekendDays.includes(dow);
    const next = has
      ? config.customWeekendDays.filter((x) => x !== dow)
      : [...config.customWeekendDays, dow].sort((a, b) => a - b);
    updateField('customWeekendDays', next);
  };

  const filteredCandidates = candidates.filter((c) => {
    const q = rosterSearch.trim().toLowerCase();
    if (!q) return true;
    return (
      c.name.toLowerCase().includes(q) ||
      (c.rollNo && c.rollNo.toLowerCase().includes(q))
    );
  });

  const saveCandidateOverrides = () => {
    if (!editingCandidate) return;
    onChangeCandidates((prev) =>
      prev.map((c) => (c.id === editingCandidate.id ? editingCandidate : c))
    );
    setEditingCandidate(null);
  };

  return (
    <div className="max-w-[1440px] mx-auto px-6 py-6 space-y-6">
      {/* Contextual Sub-header & Action Strip */}
      <div className="flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-slate-200">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Attendance Configuration & Roster Builder
          </h1>
          <p className="mt-1 text-sm text-slate-600">
            Configure training batch metadata, biometric timing curves, compliance guardrails, and holidays — then inspect the live sheet or download the customized Python application.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={onOpenPreview}
            className="px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            Inspect Live Sheet
          </button>
          <button
            type="button"
            onClick={onOpenPython}
            className="px-3.5 py-2 text-sm font-medium text-blue-700 bg-blue-50 border border-blue-200 rounded-lg hover:bg-blue-100 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            Get Python Code
          </button>
          <button
            type="button"
            onClick={onResetAll}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-slate-600 hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <RotateCcw className="w-4 h-4" />
            <span>Reset</span>
          </button>
        </div>
      </div>

      {generationError && (
        <div
          role="alert"
          className="px-4 py-3 text-sm text-red-800 bg-red-50 border border-red-200 rounded-lg"
        >
          {generationError}
        </div>
      )}

      {/* Main 12-column Workspace Split */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* LEFT 7 COLUMNS: Structured Configuration Sections */}
        <div className="lg:col-span-7 bg-white border border-slate-200 rounded-xl divide-y divide-slate-200">
          {/* 1. Saved Profiles & JSON Import/Export */}
          <section className="p-6 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div>
                <h2 className="text-base font-semibold text-slate-900">
                  01. Center & Batch Metadata
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Appears in the 6-row Excel banner and baked into your Python script defaults.
                </p>
              </div>
              <div className="flex items-center gap-2">
                <input
                  ref={jsonInputRef}
                  type="file"
                  accept=".json"
                  className="hidden"
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) handleImportProfileJson(f);
                    e.target.value = '';
                  }}
                />
                <button
                  type="button"
                  onClick={() => jsonInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-50 border border-slate-200 rounded-md hover:bg-slate-100 cursor-pointer whitespace-nowrap"
                >
                  <FolderOpen className="w-3.5 h-3.5" />
                  <span>Import JSON</span>
                </button>
                <button
                  type="button"
                  onClick={handleExportProfileJson}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-50 border border-slate-200 rounded-md hover:bg-slate-100 cursor-pointer whitespace-nowrap"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Export JSON</span>
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Training Center Name
                </label>
                <input
                  type="text"
                  value={config.centerName}
                  onChange={(e) => updateField('centerName', e.target.value)}
                  placeholder="Enter center name"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  District
                </label>
                <input
                  type="text"
                  value={config.district}
                  onChange={(e) => updateField('district', e.target.value)}
                  placeholder="Enter district"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Location / Town
                </label>
                <input
                  type="text"
                  value={config.location}
                  onChange={(e) => updateField('location', e.target.value)}
                  placeholder="Enter location"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Batch Number / ID
                </label>
                <input
                  type="text"
                  value={config.batchNo}
                  onChange={(e) => updateField('batchNo', e.target.value)}
                  placeholder="Enter batch no"
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Qualification Pack / Job Role
                </label>
                <input
                  type="text"
                  value={config.qualificationPack}
                  onChange={(e) => updateField('qualificationPack', e.target.value)}
                  placeholder="Enter qualification pack"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Scheme / Organization
                </label>
                <input
                  type="text"
                  value={config.schemeName}
                  onChange={(e) => updateField('schemeName', e.target.value)}
                  placeholder="Enter scheme or organization"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Trainer Name (Signature Block)
                </label>
                <input
                  type="text"
                  value={config.trainerName}
                  onChange={(e) => updateField('trainerName', e.target.value)}
                  placeholder="Enter trainer name"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Center Manager (Signature Block)
                </label>
                <input
                  type="text"
                  value={config.centerManager}
                  onChange={(e) => updateField('centerManager', e.target.value)}
                  placeholder="Enter center manager name"
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
            </div>

            {/* Browser Saved Profiles Bar */}
            <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center gap-2">
              <input
                type="text"
                value={profileNameInput}
                onChange={(e) => setProfileNameInput(e.target.value)}
                placeholder="Save current setup as profile..."
                className="flex-1 min-w-[180px] px-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <button
                type="button"
                onClick={() => {
                  const name = profileNameInput.trim() || `${config.centerName} (${config.batchNo})`;
                  onSaveProfile(name);
                  setProfileNameInput('');
                }}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors cursor-pointer whitespace-nowrap"
              >
                <Save className="w-3.5 h-3.5" />
                <span>Save Profile</span>
              </button>
              {profiles.length > 0 && (
                <div className="w-full flex flex-wrap items-center gap-2 pt-1">
                  <span className="text-xs text-slate-500">Saved:</span>
                  {profiles.map((p) => (
                    <span
                      key={p.id}
                      className="inline-flex items-center gap-1.5 text-xs text-slate-700 bg-slate-100 rounded-md px-2 py-1"
                    >
                      <button
                        type="button"
                        onClick={() => onLoadProfile(p)}
                        className="hover:text-blue-700 font-medium cursor-pointer"
                      >
                        {p.name}
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteProfile(p.id)}
                        aria-label={`Delete ${p.name}`}
                        className="text-slate-400 hover:text-red-600 cursor-pointer"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          </section>

          {/* 2. Date Span, Continuation & Weekend Policy */}
          <section className="p-6 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-base font-semibold text-slate-900">
                02. Schedule, Continue Previous Excel & Weekends
              </h2>
            </div>

            {/* Continue from Previous Attendance Excel */}
            <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-3">
              <input
                ref={prevExcelInputRef}
                type="file"
                accept=".xlsx,.xls,.xlsm,.csv"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleUploadPreviousExcel(f);
                  e.target.value = '';
                }}
              />
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <p className="text-xs font-semibold text-slate-900">
                    Continue from Previous Attendance Excel
                  </p>
                  <p className="text-[11px] text-slate-600 mt-0.5">
                    Upload an earlier attendance sheet (e.g. April) to keep all original IN/OUT times intact while extending into May, June, or any new date range.
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => prevExcelInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer whitespace-nowrap"
                  >
                    <Upload className="w-3.5 h-3.5" />
                    <span>{previousData ? 'Replace Previous Excel' : 'Upload Previous Excel'}</span>
                  </button>
                  {previousData && (
                    <button
                      type="button"
                      onClick={() => {
                        onChangePreviousData(null);
                        setPrevExcelError('');
                      }}
                      className="px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:text-red-700 cursor-pointer whitespace-nowrap"
                    >
                      Clear
                    </button>
                  )}
                </div>
              </div>

              {prevExcelError && (
                <p className="text-xs text-red-700 bg-red-50 border border-red-200 rounded-md px-3 py-2">
                  {prevExcelError}
                </p>
              )}

              {previousData && (
                <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs">
                  <div className="text-slate-700">
                    <span className="font-medium text-slate-900">{previousData.fileName}</span>
                    <span className="mx-1.5 text-slate-400">·</span>
                    <span className="font-mono tabular-nums">
                      {formatDisplayDate(previousData.minDate)} to {formatDisplayDate(previousData.maxDate)}
                    </span>
                    <span className="mx-1.5 text-slate-400">·</span>
                    <span className="tabular-nums">{previousData.datesCount} dates locked</span>
                    <span className="mx-1.5 text-slate-400">·</span>
                    <span className="tabular-nums">{previousData.candidateNames.length} trainees</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => handleExtendMonths(1)}
                      className="px-2.5 py-1 text-xs font-medium text-blue-700 bg-white border border-blue-200 hover:bg-blue-50 rounded-md cursor-pointer whitespace-nowrap"
                    >
                      +1 Month
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExtendMonths(2)}
                      className="px-2.5 py-1 text-xs font-medium text-blue-700 bg-white border border-blue-200 hover:bg-blue-50 rounded-md cursor-pointer whitespace-nowrap"
                    >
                      +2 Months
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExtendMonths(3)}
                      className="px-2.5 py-1 text-xs font-medium text-blue-700 bg-white border border-blue-200 hover:bg-blue-50 rounded-md cursor-pointer whitespace-nowrap"
                    >
                      +3 Months
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Segmented Mode Control */}
            <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg w-fit">
              {(
                [
                  { id: 'single', label: 'Single Range' },
                  { id: 'months_stacked', label: 'By Months (Stacked)' },
                  { id: 'months_separate', label: 'By Months (Separate Sheets)' },
                ] as { id: GenerationMode; label: string }[]
              ).map((m) => (
                <button
                  key={m.id}
                  type="button"
                  onClick={() => updateField('mode', m.id)}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    config.mode === m.id
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {m.label}
                </button>
              ))}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Start Date
                </label>
                <input
                  type="date"
                  value={config.startDate}
                  onChange={(e) => updateField('startDate', e.target.value)}
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  End Date
                </label>
                <input
                  type="date"
                  value={config.endDate}
                  onChange={(e) => updateField('endDate', e.target.value)}
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Weekend Policy
                </label>
                <select
                  value={config.weekendPolicy}
                  onChange={(e) => updateField('weekendPolicy', e.target.value as WeekendPolicy)}
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  <option value="sunday">Sundays Only</option>
                  <option value="sat_sun">Saturday & Sunday</option>
                  <option value="second_fourth_sat">2nd & 4th Sat + Sunday</option>
                  <option value="custom">Custom Weekdays</option>
                </select>
              </div>
            </div>

            {config.weekendPolicy === 'custom' && (
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs text-slate-600">Off days:</span>
                {WEEKDAY_LABELS.map((day, dow) => {
                  const active = config.customWeekendDays.includes(dow);
                  return (
                    <button
                      key={day}
                      type="button"
                      onClick={() => toggleCustomWeekendDay(dow)}
                      className={`px-2.5 py-1 text-xs font-medium rounded-md border transition-colors cursor-pointer ${
                        active
                          ? 'bg-slate-900 text-white border-slate-900'
                          : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'
                      }`}
                    >
                      {day}
                    </button>
                  );
                })}
              </div>
            )}

            {config.mode !== 'single' && (
              <div className="space-y-2 pt-1">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-medium text-slate-700">
                    Select Months to Include
                  </span>
                  <div className="flex items-center gap-3 text-xs">
                    <button
                      type="button"
                      onClick={() => updateField('selectedMonths', [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])}
                      className="text-blue-700 hover:underline cursor-pointer"
                    >
                      Select All
                    </button>
                    <button
                      type="button"
                      onClick={() => updateField('selectedMonths', [])}
                      className="text-slate-500 hover:underline cursor-pointer"
                    >
                      Clear
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-4 sm:grid-cols-6 gap-1.5">
                  {MONTH_NAMES.map((name, idx) => {
                    const m = idx + 1;
                    const selected = config.selectedMonths.includes(m);
                    return (
                      <button
                        key={name}
                        type="button"
                        onClick={() => toggleMonth(m)}
                        className={`px-2 py-1.5 text-xs font-medium rounded-md border transition-colors cursor-pointer ${
                          selected
                            ? 'bg-blue-700 text-white border-blue-700'
                            : 'bg-white text-slate-600 border-slate-200 hover:border-slate-300'
                        }`}
                      >
                        {name.slice(0, 3)}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </section>

          {/* 3. Biometric Timing Engine & Shift Presets */}
          <section className="p-6 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-base font-semibold text-slate-900">
                03. Biometric Punch Timing & Realism
              </h2>
              <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer">
                <input
                  type="checkbox"
                  checked={config.autoFillTimes}
                  onChange={(e) => updateField('autoFillTimes', e.target.checked)}
                  className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                />
                <span>Auto-fill IN/OUT timestamps</span>
              </label>
            </div>

            {/* Shift Quick Presets */}
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs text-slate-500">Shift Presets:</span>
              {SHIFT_PRESETS.map((preset) => (
                <button
                  key={preset.id}
                  type="button"
                  onClick={() =>
                    onChangeConfig((prev) => ({
                      ...prev,
                      inMin: preset.inMin,
                      inMax: preset.inMax,
                      outMin: preset.outMin,
                      outMax: preset.outMax,
                      halfDayOutTime: preset.halfDayOut,
                    }))
                  }
                  className="px-2.5 py-1 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors cursor-pointer whitespace-nowrap"
                >
                  {preset.label}
                </button>
              ))}
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  IN Min
                </label>
                <input
                  type="time"
                  value={config.inMin}
                  onChange={(e) => updateField('inMin', e.target.value)}
                  disabled={!config.autoFillTimes}
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg disabled:bg-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  IN Max
                </label>
                <input
                  type="time"
                  value={config.inMax}
                  onChange={(e) => updateField('inMax', e.target.value)}
                  disabled={!config.autoFillTimes}
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg disabled:bg-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  OUT Min
                </label>
                <input
                  type="time"
                  value={config.outMin}
                  onChange={(e) => updateField('outMin', e.target.value)}
                  disabled={!config.autoFillTimes}
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg disabled:bg-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  OUT Max
                </label>
                <input
                  type="time"
                  value={config.outMax}
                  onChange={(e) => updateField('outMax', e.target.value)}
                  disabled={!config.autoFillTimes}
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg disabled:bg-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Display Time Format
                </label>
                <select
                  value={config.timeFormat}
                  onChange={(e) => updateField('timeFormat', e.target.value as TimeFormatOption)}
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  <option value="12h_short">12-Hour Short (10:14)</option>
                  <option value="12h_ampm">12-Hour AM/PM (10:14 AM)</option>
                  <option value="24h">24-Hour Military (16:14)</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Sampling Distribution
                </label>
                <select
                  value={config.timeDistribution}
                  onChange={(e) =>
                    updateField('timeDistribution', e.target.value as TimeDistribution)
                  }
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  <option value="gaussian">Realistic Bell Curve (Gaussian)</option>
                  <option value="uniform">Uniform Random</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Deterministic Seed
                </label>
                <div className="flex gap-1.5">
                  <input
                    type="number"
                    value={config.seed}
                    onChange={(e) => updateField('seed', parseInt(e.target.value, 10) || 1)}
                    className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                  <button
                    type="button"
                    onClick={() => updateField('seed', Math.floor(100000 + Math.random() * 900000))}
                    title="Generate new random seed"
                    className="px-2.5 py-2 text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors cursor-pointer"
                  >
                    <RefreshCw className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </div>

            {/* Late & Half-Day Realism Toggles */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2 border-t border-slate-100">
              <div className="space-y-2">
                <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-800 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={config.lateSimulationEnabled}
                    onChange={(e) => updateField('lateSimulationEnabled', e.target.checked)}
                    className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                  />
                  <span>Simulate Occasional Late Arrivals</span>
                </label>
                {config.lateSimulationEnabled && (
                  <div className="flex items-center gap-3 pl-5">
                    <div className="flex-1">
                      <span className="block text-[11px] text-slate-500">Chance (%)</span>
                      <input
                        type="number"
                        min={1}
                        max={40}
                        value={config.lateChancePercent}
                        onChange={(e) =>
                          updateField('lateChancePercent', Number(e.target.value) || 0)
                        }
                        className="w-full px-2 py-1 text-xs font-mono tabular-nums border border-slate-300 rounded-md"
                      />
                    </div>
                    <div className="flex-1">
                      <span className="block text-[11px] text-slate-500">Max +Mins</span>
                      <input
                        type="number"
                        min={5}
                        max={90}
                        value={config.lateMaxExtraMinutes}
                        onChange={(e) =>
                          updateField('lateMaxExtraMinutes', Number(e.target.value) || 15)
                        }
                        className="w-full px-2 py-1 text-xs font-mono tabular-nums border border-slate-300 rounded-md"
                      />
                    </div>
                  </div>
                )}
              </div>

              <div className="space-y-2">
                <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-800 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={config.halfDaySimulationEnabled}
                    onChange={(e) => updateField('halfDaySimulationEnabled', e.target.checked)}
                    className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                  />
                  <span>Simulate Half-Day Early Out</span>
                </label>
                {config.halfDaySimulationEnabled && (
                  <div className="flex items-center gap-3 pl-5">
                    <div className="flex-1">
                      <span className="block text-[11px] text-slate-500">Chance (%)</span>
                      <input
                        type="number"
                        min={1}
                        max={30}
                        value={config.halfDayChancePercent}
                        onChange={(e) =>
                          updateField('halfDayChancePercent', Number(e.target.value) || 0)
                        }
                        className="w-full px-2 py-1 text-xs font-mono tabular-nums border border-slate-300 rounded-md"
                      />
                    </div>
                    <div className="flex-1">
                      <span className="block text-[11px] text-slate-500">Early Out</span>
                      <input
                        type="time"
                        value={config.halfDayOutTime}
                        onChange={(e) => updateField('halfDayOutTime', e.target.value)}
                        className="w-full px-2 py-1 text-xs font-mono tabular-nums border border-slate-300 rounded-md"
                      />
                    </div>
                  </div>
                )}
              </div>
            </div>
          </section>

          {/* 4. Absent Policy & Compliance Guardrails */}
          <section className="p-6 space-y-4">
            <h2 className="text-base font-semibold text-slate-900">
              04. Absent Allocation & Audit Compliance
            </h2>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Absent Strategy
                </label>
                <select
                  value={config.absentPolicy}
                  onChange={(e) => updateField('absentPolicy', e.target.value as AbsentPolicy)}
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  <option value="none">No Auto Absents (100%)</option>
                  <option value="fixed">Fixed Absents / Month</option>
                  <option value="range">Random Min–Max / Month</option>
                  <option value="target_percent">Target Attendance % Band</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Audit Pass Threshold (%)
                </label>
                <input
                  type="number"
                  min={0}
                  max={100}
                  value={config.complianceThresholdPct}
                  onChange={(e) =>
                    updateField('complianceThresholdPct', Number(e.target.value) || 75)
                  }
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Absent / Off Marker
                </label>
                <input
                  type="text"
                  value={config.absentMarker}
                  onChange={(e) => updateField('absentMarker', e.target.value)}
                  placeholder="0.00"
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
            </div>

            {config.absentPolicy === 'fixed' && (
              <div className="max-w-xs">
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Exact Absents per Candidate per Month
                </label>
                <input
                  type="number"
                  min={0}
                  max={25}
                  value={config.fixedAbsentsPerMonth}
                  onChange={(e) =>
                    updateField('fixedAbsentsPerMonth', Math.max(0, Number(e.target.value) || 0))
                  }
                  className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg"
                />
              </div>
            )}

            {config.absentPolicy === 'range' && (
              <div className="grid grid-cols-2 gap-4 max-w-sm">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Min Absents / Month
                  </label>
                  <input
                    type="number"
                    min={0}
                    max={25}
                    value={config.minAbsentsPerMonth}
                    onChange={(e) =>
                      updateField('minAbsentsPerMonth', Math.max(0, Number(e.target.value) || 0))
                    }
                    className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Max Absents / Month
                  </label>
                  <input
                    type="number"
                    min={0}
                    max={25}
                    value={config.maxAbsentsPerMonth}
                    onChange={(e) =>
                      updateField('maxAbsentsPerMonth', Math.max(0, Number(e.target.value) || 0))
                    }
                    className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            )}

            {config.absentPolicy === 'target_percent' && (
              <div className="grid grid-cols-2 gap-4 max-w-sm">
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Min Target Attendance (%)
                  </label>
                  <input
                    type="number"
                    min={50}
                    max={100}
                    value={config.targetMinAttendancePct}
                    onChange={(e) =>
                      updateField('targetMinAttendancePct', Number(e.target.value) || 80)
                    }
                    className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-700 mb-1">
                    Max Target Attendance (%)
                  </label>
                  <input
                    type="number"
                    min={50}
                    max={100}
                    value={config.targetMaxAttendancePct}
                    onChange={(e) =>
                      updateField('targetMaxAttendancePct', Number(e.target.value) || 96)
                    }
                    className="w-full px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            )}

            <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                checked={config.avoidConsecutiveAbsents}
                onChange={(e) => updateField('avoidConsecutiveAbsents', e.target.checked)}
                className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
              />
              <span>Spread absents naturally (avoid back-to-back consecutive absent days)</span>
            </label>
          </section>

          {/* 5. Holiday Calendar Manager */}
          <section className="p-6 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-base font-semibold text-slate-900">
                05. Holiday Calendar
              </h2>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleAddNationalHolidaysInRange}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md transition-colors cursor-pointer whitespace-nowrap"
                >
                  <Calendar className="w-3.5 h-3.5" />
                  <span>Load National Presets</span>
                </button>
                {config.holidays.length > 0 && (
                  <button
                    type="button"
                    onClick={() => updateField('holidays', [])}
                    className="text-xs text-slate-500 hover:text-red-600 cursor-pointer"
                  >
                    Clear All
                  </button>
                )}
              </div>
            </div>

            <div className="flex flex-wrap gap-2">
              <input
                type="date"
                value={newHolDate}
                onChange={(e) => setNewHolDate(e.target.value)}
                className="px-3 py-2 text-sm font-mono tabular-nums bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <input
                type="text"
                value={newHolLabel}
                onChange={(e) => setNewHolLabel(e.target.value)}
                placeholder="Occasion (e.g. Diwali, Independence Day)"
                className="flex-1 min-w-[180px] px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <button
                type="button"
                onClick={handleAddHoliday}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer whitespace-nowrap"
              >
                <Plus className="w-4 h-4" />
                <span>Add Holiday</span>
              </button>
            </div>

            {config.holidays.length === 0 ? (
              <p className="text-xs text-slate-500 py-2">
                No holidays configured. Add dates above or load national holiday presets.
              </p>
            ) : (
              <div className="divide-y divide-slate-100 border border-slate-200 rounded-lg max-h-44 overflow-y-auto">
                {config.holidays.map((h) => (
                  <div
                    key={h.date}
                    className="px-3 py-2 flex items-center justify-between text-xs hover:bg-slate-50"
                  >
                    <div className="flex items-center gap-3">
                      <span className="font-mono tabular-nums font-medium text-slate-900">
                        {h.date}
                      </span>
                      <span className="text-slate-600">{h.label}</span>
                    </div>
                    <button
                      type="button"
                      onClick={() =>
                        updateField(
                          'holidays',
                          config.holidays.filter((x) => x.date !== h.date)
                        )
                      }
                      aria-label={`Remove ${h.label}`}
                      className="text-slate-400 hover:text-red-600 p-1 cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* 6. Excel Styling & Layout Options */}
          <section className="p-6 space-y-4">
            <h2 className="text-base font-semibold text-slate-900">
              06. Excel Workbook Styling & Print Layout
            </h2>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Workbook Visual Theme
                </label>
                <select
                  value={config.excelTheme}
                  onChange={(e) => updateField('excelTheme', e.target.value as ExcelTheme)}
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  <option value="classic_navy">Classic Navy (Corporate)</option>
                  <option value="slate_minimal">Slate Minimal</option>
                  <option value="emerald_gov">Emerald Institutional</option>
                  <option value="monochrome">Monochrome Print-Only</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Report Subtitle
                </label>
                <input
                  type="text"
                  value={config.reportTitle}
                  onChange={(e) => updateField('reportTitle', e.target.value)}
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5 pt-1">
              {(
                [
                  ['colorCodeCells', 'Color-code Weekend, Holiday, Absent & Late cells'],
                  ['includeSummarySheet', 'Add separate Executive_Summary worksheet'],
                  ['freezePanes', 'Freeze candidate names & date headers (D9)'],
                  ['landscapePrintSetup', 'Configure A4 Landscape fit-to-1-page-wide print'],
                  ['includeSignatureBlock', 'Append Trainer & Center Manager signature block'],
                ] as [keyof AttendanceConfig, string][]
              ).map(([key, label]) => (
                <label
                  key={key}
                  className="inline-flex items-start gap-2 text-xs text-slate-700 cursor-pointer"
                >
                  <input
                    type="checkbox"
                    checked={Boolean(config[key])}
                    onChange={(e) => updateField(key, e.target.checked as never)}
                    className="mt-0.5 rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                  />
                  <span>{label}</span>
                </label>
              ))}
            </div>
          </section>
        </div>

        {/* RIGHT 5 COLUMNS: Candidate Roster & File Importer */}
        <div className="lg:col-span-5 bg-white border border-slate-200 rounded-xl divide-y divide-slate-200 lg:sticky lg:top-20">
          <div className="p-5 space-y-3">
            <div className="flex items-center justify-between gap-2">
              <div>
                <h2 className="text-base font-semibold text-slate-900">
                  Candidate Roster
                </h2>
                <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
                  {candidates.length} trainees active
                  {parsedFile ? ` · Imported from ${parsedFile.fileName}` : ''}
                </p>
              </div>
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => downloadSampleRosterExcel(SAMPLE_CANDIDATES)}
                  className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-50 border border-slate-200 rounded-md hover:bg-slate-100 cursor-pointer whitespace-nowrap"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  <span>Sample .xlsx</span>
                </button>
                <button
                  type="button"
                  onClick={() => onChangeCandidates(() => [...SAMPLE_CANDIDATES])}
                  className="px-2.5 py-1.5 text-xs font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md cursor-pointer whitespace-nowrap"
                >
                  Load Demo
                </button>
              </div>
            </div>

            {/* File Dropzone */}
            <div
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const file = e.dataTransfer.files?.[0];
                if (file) handleFileUpload(file);
              }}
              className="border border-dashed border-slate-300 rounded-lg p-4 text-center bg-slate-50/70 hover:bg-slate-50 transition-colors"
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.xlsm,.csv"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleFileUpload(f);
                  e.target.value = '';
                }}
              />
              <input
                ref={(el) => {
                  folderInputRef.current = el;
                  if (el) {
                    el.setAttribute('webkitdirectory', '');
                    el.setAttribute('directory', '');
                  }
                }}
                type="file"
                multiple
                className="hidden"
                onChange={(e) => {
                  handleFolderAutoFetch(e.target.files);
                  e.target.value = '';
                }}
              />
              <Upload className="w-5 h-5 text-slate-400 mx-auto mb-1.5" />
              <p className="text-xs text-slate-700 font-medium">
                Drop an Excel (.xlsx / .xls) or CSV roster file here
              </p>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Auto-detects Candidate Name & Roll No columns just like the Python engine
              </p>
              <div className="mt-2.5 flex flex-wrap items-center justify-center gap-2">
                <button
                  type="button"
                  onClick={() => folderInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer whitespace-nowrap"
                >
                  <FolderOpen className="w-3.5 h-3.5" />
                  <span>Auto-Fetch Folder</span>
                </button>
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs font-medium text-white bg-slate-900 hover:bg-slate-800 rounded-md cursor-pointer whitespace-nowrap"
                >
                  Browse File
                </button>
                <button
                  type="button"
                  onClick={() => setShowBulkPaste((v) => !v)}
                  className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md cursor-pointer whitespace-nowrap"
                >
                  {showBulkPaste ? 'Hide Bulk Paste' : 'Bulk Paste'}
                </button>
              </div>
            </div>

            {importError && (
              <p className="text-xs text-red-700 bg-red-50 border border-red-200 rounded-md px-3 py-2">
                {importError}
              </p>
            )}

            {/* Dynamic Column Mapping after Upload */}
            {parsedFile && (
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-medium text-slate-800 truncate">
                    Column Mapping · {parsedFile.fileName}
                  </span>
                  <label className="inline-flex items-center gap-1.5 text-[11px] text-slate-600 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={deduplicateImport}
                      onChange={(e) =>
                        reApplyColumnSelection(selectedNameCol, selectedRollCol, e.target.checked)
                      }
                      className="rounded border-slate-300"
                    />
                    <span>Deduplicate</span>
                  </label>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <span className="block text-[11px] text-slate-500 mb-0.5">Name Column</span>
                    <select
                      value={selectedNameCol}
                      onChange={(e) =>
                        reApplyColumnSelection(e.target.value, selectedRollCol, deduplicateImport)
                      }
                      className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded-md"
                    >
                      {parsedFile.headers.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <span className="block text-[11px] text-slate-500 mb-0.5">Roll / ID Column</span>
                    <select
                      value={selectedRollCol}
                      onChange={(e) =>
                        reApplyColumnSelection(selectedNameCol, e.target.value, deduplicateImport)
                      }
                      className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded-md"
                    >
                      <option value="">None</option>
                      {parsedFile.headers.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            )}

            {/* Bulk Paste Drawer */}
            {showBulkPaste && (
              <div className="space-y-2 pt-1">
                <textarea
                  rows={4}
                  value={bulkText}
                  onChange={(e) => setBulkText(e.target.value)}
                  placeholder={'Paste one name per line, or RollNo, Name:\nBT-101, Rahul Verma\nBT-102, Sneha Roy'}
                  className="w-full px-3 py-2 text-xs font-mono bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
                <div className="flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowBulkPaste(false)}
                    className="px-2.5 py-1 text-xs text-slate-600 hover:text-slate-900 cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="button"
                    onClick={handleBulkImport}
                    className="px-3 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
                  >
                    Append Names
                  </button>
                </div>
              </div>
            )}

            {/* Single Candidate Quick Add */}
            <form onSubmit={handleAddSingleCandidate} className="flex gap-2">
              <input
                type="text"
                value={newCandName}
                onChange={(e) => setNewCandName(e.target.value)}
                placeholder="Add candidate full name..."
                className="flex-1 px-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <input
                type="text"
                value={newCandRoll}
                onChange={(e) => setNewCandRoll(e.target.value)}
                placeholder="Roll ID"
                className="w-24 px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <button
                type="submit"
                aria-label="Add candidate"
                className="inline-flex items-center justify-center px-2.5 py-1.5 text-white bg-blue-700 hover:bg-blue-800 rounded-lg cursor-pointer"
              >
                <UserPlus className="w-4 h-4" />
              </button>
            </form>
          </div>

          {/* Search & Roster Table Tools */}
          <div className="px-5 py-2.5 bg-slate-50 flex items-center justify-between gap-2">
            <div className="relative flex-1">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={rosterSearch}
                onChange={(e) => setRosterSearch(e.target.value)}
                placeholder="Filter trainees..."
                className="w-full pl-8 pr-3 py-1 text-xs bg-white border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
            </div>
            <div className="flex items-center gap-2 text-xs">
              <button
                type="button"
                onClick={() =>
                  onChangeCandidates((prev) =>
                    [...prev].sort((a, b) => a.name.localeCompare(b.name))
                  )
                }
                className="text-slate-600 hover:text-slate-900 cursor-pointer whitespace-nowrap"
              >
                Sort A–Z
              </button>
              <span className="text-slate-300">·</span>
              <button
                type="button"
                onClick={() => onChangeCandidates(() => [])}
                className="text-slate-500 hover:text-red-600 cursor-pointer whitespace-nowrap"
              >
                Clear
              </button>
            </div>
          </div>

          {/* Scrollable Candidate List */}
          <div className="max-h-[460px] overflow-y-auto divide-y divide-slate-100">
            {filteredCandidates.length === 0 ? (
              <div className="p-8 text-center">
                <p className="text-sm font-medium text-slate-700">No candidates in roster</p>
                <p className="text-xs text-slate-500 mt-1">
                  Upload an Excel/CSV file, paste names, or load the demo batch to get started.
                </p>
              </div>
            ) : (
              filteredCandidates.map((cand, idx) => {
                const hasOverrides = Boolean(
                  cand.joinDate || cand.dropDate || (cand.manualAbsents && cand.manualAbsents.length)
                );
                return (
                  <div
                    key={cand.id}
                    className="px-5 py-2.5 flex items-center justify-between gap-2 hover:bg-slate-50/80 text-sm"
                  >
                    <div className="min-w-0 flex items-center gap-3">
                      <span className="w-6 text-xs font-mono tabular-nums text-slate-400 shrink-0">
                        {String(idx + 1).padStart(2, '0')}
                      </span>
                      <div className="min-w-0">
                        <p className="text-sm font-medium text-slate-900 truncate">
                          {cand.name}
                        </p>
                        <p className="text-xs text-slate-500 truncate">
                          {cand.rollNo || 'No ID'}
                          {hasOverrides ? ' · Custom dates set' : ''}
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          setEditingCandidate({
                            ...cand,
                            manualAbsents: [...(cand.manualAbsents || [])],
                          });
                          setOverrideAbsentDate(config.startDate);
                        }}
                        title="Per-candidate join/drop & leave overrides"
                        className={`p-1.5 rounded-md transition-colors cursor-pointer ${
                          hasOverrides
                            ? 'text-blue-700 bg-blue-50 hover:bg-blue-100'
                            : 'text-slate-400 hover:text-slate-700 hover:bg-slate-100'
                        }`}
                      >
                        <Settings2 className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          onChangeCandidates((prev) => prev.filter((x) => x.id !== cand.id))
                        }
                        aria-label={`Remove ${cand.name}`}
                        className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-md transition-colors cursor-pointer"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>
      </div>

      {/* Per-Candidate Override Modal */}
      {editingCandidate && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-xl max-w-md w-full p-6 space-y-4 shadow-xl">
            <div className="flex items-start justify-between">
              <div>
                <h3 className="text-base font-semibold text-slate-900">
                  Candidate Overrides
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">{editingCandidate.name}</p>
              </div>
              <button
                type="button"
                onClick={() => setEditingCandidate(null)}
                className="text-slate-400 hover:text-slate-700 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Join Date (Optional)
                </label>
                <input
                  type="date"
                  value={editingCandidate.joinDate || ''}
                  onChange={(e) =>
                    setEditingCandidate({
                      ...editingCandidate,
                      joinDate: e.target.value || undefined,
                    })
                  }
                  className="w-full px-2.5 py-1.5 text-xs font-mono border border-slate-300 rounded-md"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Drop / Exit Date
                </label>
                <input
                  type="date"
                  value={editingCandidate.dropDate || ''}
                  onChange={(e) =>
                    setEditingCandidate({
                      ...editingCandidate,
                      dropDate: e.target.value || undefined,
                    })
                  }
                  className="w-full px-2.5 py-1.5 text-xs font-mono border border-slate-300 rounded-md"
                />
              </div>
            </div>

            <div className="space-y-2">
              <label className="block text-xs font-medium text-slate-700">
                Specific Planned Leave / Absent Dates
              </label>
              <div className="flex gap-2">
                <input
                  type="date"
                  value={overrideAbsentDate}
                  onChange={(e) => setOverrideAbsentDate(e.target.value)}
                  className="flex-1 px-2.5 py-1.5 text-xs font-mono border border-slate-300 rounded-md"
                />
                <button
                  type="button"
                  onClick={() => {
                    if (!overrideAbsentDate) return;
                    const cur = editingCandidate.manualAbsents || [];
                    if (!cur.includes(overrideAbsentDate)) {
                      setEditingCandidate({
                        ...editingCandidate,
                        manualAbsents: [...cur, overrideAbsentDate].sort(),
                      });
                    }
                  }}
                  className="px-3 py-1.5 text-xs font-medium text-white bg-slate-900 rounded-md hover:bg-slate-800 cursor-pointer"
                >
                  Add Date
                </button>
              </div>
              {(editingCandidate.manualAbsents || []).length > 0 && (
                <div className="flex flex-wrap gap-1.5 pt-1">
                  {(editingCandidate.manualAbsents || []).map((ds) => (
                    <span
                      key={ds}
                      className="inline-flex items-center gap-1 px-2 py-0.5 text-xs font-mono bg-red-50 text-red-800 border border-red-200 rounded"
                    >
                      {ds}
                      <button
                        type="button"
                        onClick={() =>
                          setEditingCandidate({
                            ...editingCandidate,
                            manualAbsents: (editingCandidate.manualAbsents || []).filter(
                              (x) => x !== ds
                            ),
                          })
                        }
                        className="hover:text-red-950 cursor-pointer"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  ))}
                </div>
              )}
            </div>

            <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setEditingCandidate(null)}
                className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={saveCandidateOverrides}
                className="px-4 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
              >
                Save Overrides
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
