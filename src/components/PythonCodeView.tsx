import { useMemo, useRef, useState } from 'react';
import { Check, Copy, Download, FileCode2, FileText } from 'lucide-react';
import type { AttendanceConfig, PythonFeatureFlags } from '../types/attendance';
import {
  generateDynamicPythonCode,
  generateRequirementsTxt,
} from '../utils/pythonGenerator';

interface PythonCodeViewProps {
  config: AttendanceConfig;
  flags: PythonFeatureFlags;
  onChangeFlags: (updater: (prev: PythonFeatureFlags) => PythonFeatureFlags) => void;
}

const FEATURE_CONTROLS: {
  key: keyof PythonFeatureFlags;
  title: string;
  desc: string;
}[] = [
  {
    key: 'bakeCurrentDefaults',
    title: 'Bake Current Builder Defaults',
    desc: 'Pre-populates Center, District, Batch, Qualification, shift times & rules into AttendanceConfig.',
  },
  {
    key: 'bakeCurrentHolidays',
    title: 'Bake Configured Holidays',
    desc: 'Embeds your current holiday calendar directly into default Python startup state.',
  },
  {
    key: 'includeGaussianTiming',
    title: 'Gaussian Bell-Curve Punch Engine',
    desc: 'Realistic clustered arrivals/departures with minimum shift-duration enforcement.',
  },
  {
    key: 'includeLateAndHalfDay',
    title: 'Late Arrival & Half-Day Simulation',
    desc: 'Adds configurable late punch percentage, max extra minutes, and half-day early-out.',
  },
  {
    key: 'includeComplianceGuard',
    title: 'Audit Compliance Floor Guard',
    desc: 'Target attendance % band with automatic compliance floor cap (e.g., >=75%).',
  },
  {
    key: 'includeFlexibleWeekends',
    title: 'Flexible Weekend Policies',
    desc: 'Supports Sunday only, Sat+Sun, 2nd & 4th Saturday + Sunday, or custom weekday sets.',
  },
  {
    key: 'includeStyledExcelAndSummary',
    title: 'Styled Excel & Theme Formatting',
    desc: 'Theme palettes, conditional cell colors, optional Executive Summary sheet, freeze panes & landscape print.',
  },
  {
    key: 'includeCsvAndHtmlExport',
    title: 'CSV & Printable HTML Report Exporters',
    desc: 'Adds RFC-4180 CSV and browser-printable A4 landscape HTML/PDF exporters.',
  },
  {
    key: 'includeJsonProfiles',
    title: 'JSON Configuration Profiles',
    desc: 'Save and reload training center / batch profiles from JSON files.',
  },
  {
    key: 'includeTabbedScrollableGui',
    title: 'Modern 5-Tab Tkinter Desktop GUI',
    desc: 'Replaces cramped single-frame Tkinter form with Notebook tabs, live preview, holiday delete & background thread.',
  },
  {
    key: 'includeCliMode',
    title: 'Headless CLI Mode (argparse)',
    desc: 'Run automated batch generation from terminal or cron jobs via --input, --config, --output.',
  },
];

const UPGRADE_MATRIX = [
  {
    area: 'Architecture & Error Handling',
    original: 'Global HOLIDAYS = [] list; multiple bare except: blocks swallowing failures silently.',
    upgraded: 'Typed dataclasses (AttendanceConfig, CandidateRecord, HolidayEntry), structured logging, zero bare excepts.',
  },
  {
    area: 'Candidate & Continuation Loading',
    original: 'Folder scan only; auto-picks largest file with no continuation or deduplication.',
    upgraded: 'One-click "Auto-Fetch Excel (Script Folder)", "Continue Previous Excel" (preserves earlier IN/OUT times when extending months/dates), direct file picker, deduplication & in-GUI editor.',
  },
  {
    area: 'Time & Shift Realism',
    original: 'Uniform random 12h %I:%M only; no duration validation, no late or half-day modeling.',
    upgraded: 'Shift presets, Gaussian bell-curve clustering, 12h/AM-PM/24h formats, min-shift guard, late & half-day simulation, deterministic seed.',
  },
  {
    area: 'Absents & Weekends',
    original: 'Hardcoded Sunday-only weekends and hardcoded 2 random absents/month toggle.',
    upgraded: '4 weekend policies (Sun, Sat+Sun, 2nd/4th Sat, Custom), 4 absent strategies (None/Fixed/Range/Target %), compliance floor & non-consecutive spread.',
  },
  {
    area: 'Holiday Manager',
    original: 'Add-only holiday listbox with no delete button and no holiday occasion names.',
    upgraded: 'Labeled holidays, Treeview with Add/Remove, and one-click Indian National Holiday presets.',
  },
  {
    area: 'Excel Workbook Output',
    original: 'Plain black-and-white table; stacked sheet only; no themes, print setup, or signatures.',
    upgraded: '4 color themes, conditional fills, Separate Monthly Sheets mode, clean biometric-only date columns, Freeze Panes, Landscape A4 print, Signature block.',
  },
  {
    area: 'Desktop UI & Automation',
    original: 'Single tall Tkinter frame (clips on laptops), no live preview, no profile saving, no CLI.',
    upgraded: '5-tab ttk.Notebook GUI, Live Summary Preview tab, background worker thread + progress bar, JSON profiles, CSV/HTML export, and full argparse CLI.',
  },
];

export function PythonCodeView({
  config,
  flags,
  onChangeFlags,
}: PythonCodeViewProps) {
  const [copied, setCopied] = useState(false);
  const codeContainerRef = useRef<HTMLDivElement | null>(null);

  const pythonCode = useMemo(
    () => generateDynamicPythonCode(config, flags),
    [config, flags]
  );

  const requirementsTxt = useMemo(() => generateRequirementsTxt(flags), [flags]);

  const lines = useMemo(() => pythonCode.split('\n'), [pythonCode]);

  const sectionAnchors = useMemo(() => {
    const anchors: { label: string; lineIdx: number }[] = [];
    lines.forEach((line, idx) => {
      const m = line.match(/^# SECTION (\d+) — (.+)$/);
      if (m) {
        anchors.push({ label: `${m[1]}. ${m[2]}`, lineIdx: idx });
      }
    });
    return anchors;
  }, [lines]);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(pythonCode);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch {
      // Fallback textarea copy
      const ta = document.createElement('textarea');
      ta.value = pythonCode;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    }
  };

  const downloadTextFile = (content: string, filename: string) => {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const jumpToLine = (lineIdx: number) => {
    const el = codeContainerRef.current;
    if (!el) return;
    // Approximate 20px line height
    el.scrollTo({ top: Math.max(0, lineIdx * 20 - 24), behavior: 'smooth' });
  };

  const toggleFlag = (key: keyof PythonFeatureFlags) => {
    onChangeFlags((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  return (
    <div className="max-w-[1440px] mx-auto px-6 py-6 space-y-6">
      {/* Header */}
      <div className="flex flex-wrap items-end justify-between gap-4 pb-5 border-b border-slate-200">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Dynamic Python Code Builder
          </h1>
          <p className="mt-1 text-sm text-slate-600">
            Production-grade Python 3 rewrite of your attendance script ·{' '}
            <span className="font-mono tabular-nums text-slate-900">{lines.length} lines</span> ·
            Synced dynamically with your current center, timing & holiday setup
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={handleCopy}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors cursor-pointer whitespace-nowrap"
          >
            {copied ? <Check className="w-4 h-4 text-emerald-600" /> : <Copy className="w-4 h-4" />}
            <span>{copied ? 'Copied' : 'Copy Code'}</span>
          </button>
          <button
            type="button"
            onClick={() => downloadTextFile(requirementsTxt, 'requirements.txt')}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors cursor-pointer whitespace-nowrap"
          >
            <FileText className="w-4 h-4" />
            <span>requirements.txt</span>
          </button>
          <button
            type="button"
            onClick={() => downloadTextFile(pythonCode, 'attendance_maker_pro.py')}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer whitespace-nowrap"
          >
            <Download className="w-4 h-4" />
            <span>Download attendance_maker_pro.py</span>
          </button>
        </div>
      </div>

      {/* Main Split: Feature Toggles + Code Inspector */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        {/* Left 4 columns: Dynamic Module Toggles + Section Jump */}
        <div className="lg:col-span-4 space-y-6">
          <div className="bg-white border border-slate-200 rounded-xl divide-y divide-slate-200">
            <div className="p-4">
              <h2 className="text-sm font-semibold text-slate-900">
                Customize Python Modules
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Toggle capabilities to regenerate a lean or full-featured script instantly.
              </p>
            </div>
            <div className="divide-y divide-slate-100 max-h-[460px] overflow-y-auto">
              {FEATURE_CONTROLS.map((ctrl) => (
                <label
                  key={ctrl.key}
                  className="flex items-start gap-3 px-4 py-3 hover:bg-slate-50 cursor-pointer"
                >
                  <input
                    type="checkbox"
                    checked={flags[ctrl.key]}
                    onChange={() => toggleFlag(ctrl.key)}
                    className="mt-0.5 rounded border-slate-300 text-blue-700 focus:ring-blue-600 shrink-0"
                  />
                  <div className="min-w-0">
                    <p className="text-xs font-semibold text-slate-900">{ctrl.title}</p>
                    <p className="text-[11px] text-slate-500 mt-0.5 leading-relaxed">
                      {ctrl.desc}
                    </p>
                  </div>
                </label>
              ))}
            </div>
          </div>

          {/* Quick Run Instructions */}
          <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
            <h3 className="text-sm font-semibold text-slate-900">
              How to Run Locally
            </h3>
            <pre className="p-3 bg-slate-900 text-slate-100 rounded-lg text-[11px] font-mono overflow-x-auto leading-relaxed">
{`# 1. Install dependencies
pip install -r requirements.txt

# 2. Launch modern desktop GUI
python attendance_maker_pro.py

# 3. Continue from a previous Excel (e.g. April -> April+May+June)
python attendance_maker_pro.py \\
  --continue-from April_Attendance.xlsx \\
  --end 2026-06-30 \\
  --output April_May_June.xlsx

# 4. Or auto-fetch roster Excel from script folder via CLI
python attendance_maker_pro.py \\
  --auto-fetch \\
  --start ${config.startDate} \\
  --end ${config.endDate}`}
            </pre>
          </div>
        </div>

        {/* Right 8 columns: Section Navigator + Source Code Viewer */}
        <div className="lg:col-span-8 bg-slate-950 border border-slate-800 rounded-xl overflow-hidden flex flex-col">
          {/* Code Bar & Section Jump Tabs */}
          <div className="px-4 py-3 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs text-slate-200 font-mono">
              <FileCode2 className="w-4 h-4 text-blue-400 shrink-0" />
              <span>attendance_maker_pro.py</span>
              <span className="text-slate-500">·</span>
              <span className="text-slate-400 tabular-nums">{lines.length} lines</span>
            </div>
            <div className="flex flex-wrap items-center gap-1.5">
              {sectionAnchors.map((sec) => (
                <button
                  key={sec.label}
                  type="button"
                  onClick={() => jumpToLine(sec.lineIdx)}
                  className="px-2 py-1 text-[11px] font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 hover:text-white rounded transition-colors cursor-pointer whitespace-nowrap"
                >
                  {sec.label.split(' — ')[0]}
                </button>
              ))}
            </div>
          </div>

          {/* Monospace Numbered Code Viewport */}
          <div
            ref={codeContainerRef}
            className="overflow-auto max-h-[640px] p-4 font-mono text-xs leading-5 text-slate-100"
          >
            <table className="w-full border-collapse">
              <tbody>
                {lines.map((line, i) => {
                  const isSectionBanner = line.startsWith('# ===') || line.startsWith('# SECTION ');
                  const isComment = line.trim().startsWith('#');
                  return (
                    <tr key={i} className="hover:bg-slate-900/70">
                      <td className="select-none pr-4 text-right text-slate-600 tabular-nums w-12 align-top">
                        {i + 1}
                      </td>
                      <td
                        className={`whitespace-pre align-top ${
                          isSectionBanner
                            ? 'text-sky-400 font-semibold'
                            : isComment
                            ? 'text-slate-400'
                            : 'text-slate-100'
                        }`}
                      >
                        {line || ' '}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Original Script vs Upgraded Pro Comparison Matrix */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200">
          <h2 className="text-base font-semibold text-slate-900">
            Upgrade Summary: Original Script vs. Attendance Maker Pro
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Key architectural improvements, bug fixes, and new capabilities built on top of your uploaded script
          </p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold">
                <th className="px-6 py-3 w-48">Capability Area</th>
                <th className="px-6 py-3">Original Script</th>
                <th className="px-6 py-3">Upgraded Dynamic Pro Edition</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {UPGRADE_MATRIX.map((item) => (
                <tr key={item.area} className="hover:bg-slate-50/70">
                  <td className="px-6 py-3 font-semibold text-slate-900 align-top">
                    {item.area}
                  </td>
                  <td className="px-6 py-3 text-slate-600 align-top leading-relaxed">
                    {item.original}
                  </td>
                  <td className="px-6 py-3 text-slate-900 align-top leading-relaxed">
                    {item.upgraded}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
