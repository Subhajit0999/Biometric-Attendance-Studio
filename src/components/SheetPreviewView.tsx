import React, { useState } from 'react';
import {
  Download,
  FileSpreadsheet,
  Printer,
  RefreshCw,
  RotateCcw,
  Search,
  X,
} from 'lucide-react';
import type {
  AttendanceBlock,
  AttendanceConfig,
  CellStatus,
  DayCellData,
  ManualCellEdit,
  PreviousAttendanceData,
} from '../types/attendance';
import { formatDisplayDate, formatMinutes, parseHHMMToMinutes } from '../utils/attendanceEngine';

interface SheetPreviewViewProps {
  blocks: AttendanceBlock[];
  config: AttendanceConfig;
  previousData: PreviousAttendanceData | null;
  manualEditCount: number;
  onRerollSeed: () => void;
  onEditCell: (key: string, edit: ManualCellEdit) => void;
  onClearManualEdits: () => void;
  onExportExcel: () => void;
  onExportCSV: () => void;
  error?: string;
}

type StatusFilter = 'all' | 'low' | 'has_absents' | 'has_late';

function getCellClass(status: CellStatus, colorCode: boolean): string {
  if (!colorCode) return 'text-slate-800 bg-white';
  switch (status) {
    case 'sunday':
      return 'bg-slate-200/80 text-slate-500';
    case 'holiday':
      return 'bg-blue-100/80 text-blue-900 font-medium';
    case 'absent':
      return 'bg-red-100/90 text-red-800 font-semibold';
    case 'late':
      return 'bg-amber-100/80 text-amber-900';
    case 'half_day':
      return 'bg-purple-100/80 text-purple-900';
    case 'inactive':
      return 'bg-slate-100 text-slate-400';
    default:
      return 'bg-white text-slate-800';
  }
}

export function SheetPreviewView({
  blocks,
  config,
  previousData,
  manualEditCount,
  onRerollSeed,
  onEditCell,
  onClearManualEdits,
  onExportExcel,
  onExportCSV,
  error,
}: SheetPreviewViewProps) {
  const [activeBlockIdx, setActiveBlockIdx] = useState(0);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');

  const [editingCell, setEditingCell] = useState<{
    blockId: string;
    candidateId: string;
    candidateName: string;
    day: DayCellData;
    status: CellStatus;
    inTime: string;
    outTime: string;
  } | null>(null);

  if (error || blocks.length === 0) {
    return (
      <div className="max-w-[1440px] mx-auto px-6 py-16 text-center">
        <p className="text-base font-semibold text-slate-900">No Attendance Sheet Generated</p>
        <p className="mt-1 text-sm text-slate-600">
          {error || 'Add candidates and select a valid date span in the Builder view.'}
        </p>
      </div>
    );
  }

  const safeIdx = Math.min(activeBlockIdx, blocks.length - 1);
  const block = blocks[safeIdx];

  const filteredRows = block.rows.filter((r) => {
    const q = search.trim().toLowerCase();
    if (
      q &&
      !r.candidate.name.toLowerCase().includes(q) &&
      !(r.candidate.rollNo && r.candidate.rollNo.toLowerCase().includes(q))
    ) {
      return false;
    }
    if (statusFilter === 'low') return !r.compliant;
    if (statusFilter === 'has_absents') return r.absentDays > 0;
    if (statusFilter === 'has_late') return r.lateDays > 0;
    return true;
  });

  const openEditor = (candidateId: string, candidateName: string, day: DayCellData) => {
    setEditingCell({
      blockId: block.id,
      candidateId,
      candidateName,
      day,
      status: day.status,
      inTime: day.inTime,
      outTime: day.outTime,
    });
  };

  const applyQuickStatus = (nextStatus: CellStatus) => {
    if (!editingCell) return;
    if (nextStatus === 'absent' || nextStatus === 'holiday' || nextStatus === 'sunday') {
      setEditingCell({
        ...editingCell,
        status: nextStatus,
        inTime: config.absentMarker,
        outTime: config.absentMarker,
      });
    } else if (nextStatus === 'present') {
      const inM = parseHHMMToMinutes(config.inMin) + 8;
      const outM = parseHHMMToMinutes(config.outMin) + 10;
      setEditingCell({
        ...editingCell,
        status: 'present',
        inTime: formatMinutes(inM, config.timeFormat),
        outTime: formatMinutes(outM, config.timeFormat),
      });
    } else if (nextStatus === 'late') {
      const inM = parseHHMMToMinutes(config.inMax) + 15;
      const outM = parseHHMMToMinutes(config.outMin) + 10;
      setEditingCell({
        ...editingCell,
        status: 'late',
        inTime: formatMinutes(inM, config.timeFormat),
        outTime: formatMinutes(outM, config.timeFormat),
      });
    } else if (nextStatus === 'half_day') {
      const inM = parseHHMMToMinutes(config.inMin) + 5;
      const outM = parseHHMMToMinutes(config.halfDayOutTime);
      setEditingCell({
        ...editingCell,
        status: 'half_day',
        inTime: formatMinutes(inM, config.timeFormat),
        outTime: formatMinutes(outM, config.timeFormat),
      });
    }
  };

  const handleSaveCellEdit = () => {
    if (!editingCell) return;
    const key = `${editingCell.blockId}__${editingCell.candidateId}__${editingCell.day.dateStr}`;
    onEditCell(key, {
      status: editingCell.status,
      inTime: editingCell.inTime,
      outTime: editingCell.outTime,
    });
    setEditingCell(null);
  };

  return (
    <div className="max-w-[1440px] mx-auto px-6 py-6 space-y-4">
      {/* Toolbar */}
      <div className="flex flex-wrap items-center justify-between gap-4 no-print">
        <div className="flex flex-wrap items-center gap-3">
          {blocks.length > 1 && (
            <div className="flex items-center gap-1 p-1 bg-slate-200/70 rounded-lg">
              {blocks.map((b, i) => (
                <button
                  key={b.id}
                  type="button"
                  onClick={() => setActiveBlockIdx(i)}
                  className={`px-3 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    safeIdx === i
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {b.monthName}
                </button>
              ))}
            </div>
          )}

          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search trainee..."
              className="pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg w-48 focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>

          <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
            {(
              [
                ['all', 'All'],
                ['low', `Below ${config.complianceThresholdPct}%`],
                ['has_absents', 'With Absents'],
                ['has_late', 'With Late'],
              ] as [StatusFilter, string][]
            ).map(([id, label]) => (
              <button
                key={id}
                type="button"
                onClick={() => setStatusFilter(id)}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  statusFilter === id
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {manualEditCount > 0 && (
            <button
              type="button"
              onClick={onClearManualEdits}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-amber-800 bg-amber-50 border border-amber-200 rounded-lg hover:bg-amber-100 cursor-pointer whitespace-nowrap"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Clear {manualEditCount} Manual Edit{manualEditCount > 1 ? 's' : ''}</span>
            </button>
          )}
          <button
            type="button"
            onClick={onRerollSeed}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 cursor-pointer whitespace-nowrap"
          >
            <RefreshCw className="w-3.5 h-3.5" />
            <span>Re-roll Times</span>
          </button>
          <button
            type="button"
            onClick={onExportCSV}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 cursor-pointer whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" />
            <span>CSV</span>
          </button>
          <button
            type="button"
            onClick={() => window.print()}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 cursor-pointer whitespace-nowrap"
          >
            <Printer className="w-3.5 h-3.5" />
            <span>Print / PDF</span>
          </button>
          <button
            type="button"
            onClick={onExportExcel}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-slate-900 hover:bg-slate-800 rounded-lg cursor-pointer whitespace-nowrap"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>Download .xlsx</span>
          </button>
        </div>
      </div>

      {/* Quiet Legend & Edit Instruction Line */}
      <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600 no-print">
        <div className="flex flex-wrap items-center gap-x-4 gap-y-1">
          <span className="font-medium text-slate-800">Click any day cell to edit IN/OUT or status</span>
          {previousData && (
            <>
              <span aria-hidden="true">·</span>
              <span className="font-medium text-blue-700">
                Preserving {previousData.datesCount} previous dates ({formatDisplayDate(previousData.minDate)} – {formatDisplayDate(previousData.maxDate)}) from {previousData.fileName}
              </span>
            </>
          )}
          <span aria-hidden="true">·</span>
          <span>Working: <strong className="font-mono tabular-nums text-slate-900">{block.workingDaysCount}d</strong></span>
          <span aria-hidden="true">·</span>
          <span>Weekends: <strong className="font-mono tabular-nums text-slate-900">{block.weekendCount}d</strong></span>
          <span aria-hidden="true">·</span>
          <span>Holidays: <strong className="font-mono tabular-nums text-slate-900">{block.holidayCount}d</strong></span>
          <span aria-hidden="true">·</span>
          <span>Seed: <strong className="font-mono tabular-nums text-slate-900">{config.seed}</strong></span>
        </div>
        <div className="flex flex-wrap items-center gap-3 text-[11px]">
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 bg-red-100 border border-red-300 inline-block" /> Absent
          </span>
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 bg-amber-100 border border-amber-300 inline-block" /> Late
          </span>
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 bg-blue-100 border border-blue-300 inline-block" /> Holiday
          </span>
          <span className="inline-flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 bg-slate-200 border border-slate-300 inline-block" /> Weekend
          </span>
        </div>
      </div>

      {/* Sheet Container */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        {/* 6-row Excel Header Preview */}
        <div className="text-center border-b border-slate-200">
          <div className="bg-slate-900 text-white py-2.5 px-4 text-sm font-bold tracking-wide">
            BIOMETRIC ATTENDANCE SHEET
          </div>
          <div className="bg-slate-50 py-2 px-4 space-y-0.5 text-xs text-slate-700">
            {(config.centerName.trim() || config.district.trim() || config.batchNo.trim()) && (
              <p className="font-semibold text-slate-900 text-sm">
                {[
                  config.centerName.trim(),
                  config.district.trim() ? `(Dist: ${config.district.trim()})` : '',
                  config.batchNo.trim() ? `(Batch: ${config.batchNo.trim()})` : '',
                ]
                  .filter(Boolean)
                  .join(' ')}
              </p>
            )}
            <p>
              {config.reportTitle}
              {config.schemeName.trim() ? ` — ${config.schemeName.trim()}` : ''} · {block.monthName} · {block.reportRangeLabel}
            </p>
            {(config.qualificationPack.trim() || config.location.trim()) && (
              <p className="font-medium text-slate-800">
                {[
                  config.qualificationPack.trim()
                    ? `Qualification Pack: ${config.qualificationPack.trim()}`
                    : '',
                  config.location.trim() ? `Location: ${config.location.trim()}` : '',
                ]
                  .filter(Boolean)
                  .join(' · ')}
              </p>
            )}
          </div>
        </div>

        {/* High-Density Scrollable Biometric Grid */}
        <div className="overflow-x-auto max-h-[65vh]">
          <table className="w-full border-collapse text-xs font-mono tabular-nums">
            <thead className="sticky top-0 z-20">
              <tr className="bg-slate-800 text-white">
                <th className="sticky left-0 z-30 bg-slate-800 border border-slate-700 px-2 py-2 text-center w-11">
                  Sl
                </th>
                <th className="sticky left-11 z-30 bg-slate-800 border border-slate-700 px-3 py-2 text-left min-w-[175px] font-sans">
                  Candidate Name
                </th>
                <th className="sticky left-[219px] z-30 bg-slate-800 border border-slate-700 px-2 py-2 text-center w-12 font-sans">
                  Type
                </th>
                {block.dates.map((dt) => (
                  <th
                    key={dt.dateStr}
                    title={dt.holidayLabel ? `Holiday: ${dt.holidayLabel}` : dt.displayDate}
                    className="border border-slate-700 px-2 py-1.5 text-center whitespace-nowrap min-w-[66px]"
                  >
                    <div>{dt.displayDate.slice(0, 5)}</div>
                    <div
                      className={`text-[10px] font-normal ${
                        dt.isHoliday
                          ? 'text-sky-300 font-semibold'
                          : dt.isWeekend
                          ? 'text-rose-300'
                          : 'text-slate-300'
                      }`}
                    >
                      {dt.weekdayShort}
                      {dt.isHoliday ? '*' : ''}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredRows.map((row) => (
                <React.Fragment key={row.candidate.id}>
                  {/* IN ROW */}
                  <tr className="group hover:bg-blue-50/30">
                    <td
                      rowSpan={2}
                      className="sticky left-0 z-10 bg-white group-hover:bg-slate-50 border border-slate-200 px-2 py-1 text-center font-semibold text-slate-600 align-middle"
                    >
                      {row.slNo}
                    </td>
                    <td
                      rowSpan={2}
                      className="sticky left-11 z-10 bg-white group-hover:bg-slate-50 border border-slate-200 px-3 py-1 text-left font-sans font-medium text-slate-900 align-middle"
                    >
                      <div className="truncate max-w-[165px]" title={row.candidate.name}>
                        {row.candidate.name}
                      </div>
                      {row.candidate.rollNo && (
                        <div className="text-[10px] font-mono text-slate-400">
                          {row.candidate.rollNo}
                        </div>
                      )}
                    </td>
                    <td className="sticky left-[219px] z-10 bg-slate-50 border border-slate-200 px-2 py-1 text-center font-sans font-semibold text-sky-700 text-[11px]">
                      In
                    </td>
                    {block.dates.map((dt) => {
                      const cell = row.days[dt.dateStr];
                      return (
                        <td
                          key={dt.dateStr}
                          onClick={() => openEditor(row.candidate.id, row.candidate.name, cell)}
                          title={`Click to edit ${row.candidate.name} on ${dt.displayDate}${
                            cell.holidayLabel ? ` (${cell.holidayLabel})` : ''
                          }${cell.isPreservedFromPrevious ? ' — Preserved from previous Excel' : ''}`}
                          className={`border border-slate-200 px-1.5 py-1 text-center whitespace-nowrap cursor-pointer hover:ring-1 hover:ring-inset hover:ring-blue-600 ${getCellClass(
                            cell.status,
                            config.colorCodeCells
                          )} ${cell.isManuallyEdited ? 'underline decoration-blue-600' : ''}`}
                        >
                          {cell.inTime || '—'}
                        </td>
                      );
                    })}
                  </tr>

                  {/* OUT ROW */}
                  <tr className="group hover:bg-blue-50/30">
                    <td className="sticky left-[219px] z-10 bg-slate-50 border border-slate-200 px-2 py-1 text-center font-sans font-medium text-slate-500 text-[11px]">
                      Out
                    </td>
                    {block.dates.map((dt) => {
                      const cell = row.days[dt.dateStr];
                      return (
                        <td
                          key={dt.dateStr}
                          onClick={() => openEditor(row.candidate.id, row.candidate.name, cell)}
                          className={`border border-slate-200 px-1.5 py-1 text-center whitespace-nowrap cursor-pointer hover:ring-1 hover:ring-inset hover:ring-blue-600 ${getCellClass(
                            cell.status,
                            config.colorCodeCells
                          )}`}
                        >
                          {cell.outTime || '—'}
                        </td>
                      );
                    })}
                  </tr>
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>

        {/* Signature Footer Preview */}
        {config.includeSignatureBlock && (
          <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs text-slate-700">
            <div>
              <span className="text-slate-500">Trainer Signature: </span>
              <span className="font-medium text-slate-900">
                {config.trainerName || '____________________'}
              </span>
            </div>
            <div className="sm:text-center">
              <span className="text-slate-500">Center Manager: </span>
              <span className="font-medium text-slate-900">
                {config.centerManager || '____________________'}
              </span>
            </div>
            <div className="sm:text-right">
              <span className="text-slate-500">Verified By / Seal: </span>
              <span className="font-medium text-slate-900">____________________</span>
            </div>
          </div>
        )}
      </div>

      {/* Cell Edit Modal */}
      {editingCell && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4 no-print">
          <div className="bg-white border border-slate-200 rounded-xl max-w-sm w-full p-5 space-y-4 shadow-xl">
            <div className="flex items-start justify-between">
              <div>
                <h3 className="text-sm font-semibold text-slate-900">Edit Attendance Cell</h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  {editingCell.candidateName} · {editingCell.day.displayDate} ({editingCell.day.weekdayShort})
                </p>
              </div>
              <button
                type="button"
                onClick={() => setEditingCell(null)}
                className="text-slate-400 hover:text-slate-700 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-1.5">
              <span className="block text-xs font-medium text-slate-700">Quick Status</span>
              <div className="grid grid-cols-3 gap-1.5">
                {(
                  [
                    ['present', 'Present'],
                    ['absent', 'Absent'],
                    ['late', 'Late'],
                    ['half_day', 'Half Day'],
                    ['holiday', 'Holiday'],
                    ['sunday', 'Weekend'],
                  ] as [CellStatus, string][]
                ).map(([st, lbl]) => (
                  <button
                    key={st}
                    type="button"
                    onClick={() => applyQuickStatus(st)}
                    className={`px-2 py-1.5 text-xs font-medium rounded-md border transition-colors cursor-pointer ${
                      editingCell.status === st
                        ? 'bg-blue-700 text-white border-blue-700'
                        : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    {lbl}
                  </button>
                ))}
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">IN Time</label>
                <input
                  type="text"
                  value={editingCell.inTime}
                  onChange={(e) => setEditingCell({ ...editingCell, inTime: e.target.value })}
                  className="w-full px-2.5 py-1.5 text-xs font-mono border border-slate-300 rounded-md"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">OUT Time</label>
                <input
                  type="text"
                  value={editingCell.outTime}
                  onChange={(e) => setEditingCell({ ...editingCell, outTime: e.target.value })}
                  className="w-full px-2.5 py-1.5 text-xs font-mono border border-slate-300 rounded-md"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setEditingCell(null)}
                className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSaveCellEdit}
                className="px-4 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
              >
                Apply Change
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
