import { useMemo, useState } from 'react';
import { ArrowUpDown } from 'lucide-react';
import type { AttendanceBlock, AttendanceConfig } from '../types/attendance';

interface AnalyticsViewProps {
  blocks: AttendanceBlock[];
  config: AttendanceConfig;
  error?: string;
}

type SortKey = 'name' | 'attendancePct' | 'absentDays' | 'lateDays' | 'totalHours';

export function AnalyticsView({ blocks, config, error }: AnalyticsViewProps) {
  const [sortKey, setSortKey] = useState<SortKey>('attendancePct');
  const [sortAsc, setSortAsc] = useState(true);

  const analytics = useMemo(() => {
    if (!blocks.length) return null;

    const candMap = new Map<
      string,
      {
        id: string;
        name: string;
        rollNo: string;
        workingDays: number;
        presentDays: number;
        absentDays: number;
        lateDays: number;
        halfDays: number;
        totalHours: number;
      }
    >();

    const dailyStats: {
      dateStr: string;
      displayDate: string;
      weekdayShort: string;
      present: number;
      late: number;
      absent: number;
      isOff: boolean;
      offType: string;
    }[] = [];

    const inMinutesBucket = new Map<string, number>();

    let totalWorkingDays = 0;
    let totalWeekends = 0;
    let totalHolidays = 0;

    for (const blk of blocks) {
      totalWorkingDays += blk.workingDaysCount;
      totalWeekends += blk.weekendCount;
      totalHolidays += blk.holidayCount;

      for (const dt of blk.dates) {
        let p = 0;
        let l = 0;
        let a = 0;
        for (const r of blk.rows) {
          const cell = r.days[dt.dateStr];
          if (!cell) continue;
          if (cell.status === 'present' || cell.status === 'half_day') p++;
          else if (cell.status === 'late') l++;
          else if (cell.status === 'absent') a++;

          if (cell.inMinutes !== null) {
            const bucket5 = Math.floor(cell.inMinutes / 5) * 5;
            const h = String(Math.floor(bucket5 / 60)).padStart(2, '0');
            const m = String(bucket5 % 60).padStart(2, '0');
            const label = `${h}:${m}`;
            inMinutesBucket.set(label, (inMinutesBucket.get(label) || 0) + 1);
          }
        }
        dailyStats.push({
          dateStr: dt.dateStr,
          displayDate: dt.displayDate,
          weekdayShort: dt.weekdayShort,
          present: p,
          late: l,
          absent: a,
          isOff: dt.isWeekend || dt.isHoliday,
          offType: dt.isHoliday ? dt.holidayLabel || 'Holiday' : dt.isWeekend ? 'Weekend' : '',
        });
      }

      for (const r of blk.rows) {
        const cur = candMap.get(r.candidate.id) || {
          id: r.candidate.id,
          name: r.candidate.name,
          rollNo: r.candidate.rollNo || '—',
          workingDays: 0,
          presentDays: 0,
          absentDays: 0,
          lateDays: 0,
          halfDays: 0,
          totalHours: 0,
        };
        cur.workingDays += r.workingDays;
        cur.presentDays += r.presentDays;
        cur.absentDays += r.absentDays;
        cur.lateDays += r.lateDays;
        cur.halfDays += r.halfDays;
        cur.totalHours += r.totalHours;
        candMap.set(r.candidate.id, cur);
      }
    }

    const candidates = Array.from(candMap.values()).map((c) => {
      const attendancePct =
        c.workingDays > 0 ? Number(((c.presentDays / c.workingDays) * 100).toFixed(1)) : 100;
      const compliant = attendancePct >= config.complianceThresholdPct;
      const neededForCompliance = Math.max(
        0,
        Math.ceil((config.complianceThresholdPct / 100) * c.workingDays - c.presentDays)
      );
      return {
        ...c,
        totalHours: Number(c.totalHours.toFixed(1)),
        attendancePct,
        compliant,
        neededForCompliance,
      };
    });

    const avgAttendance =
      candidates.length > 0
        ? Number(
            (
              candidates.reduce((sum, c) => sum + c.attendancePct, 0) / candidates.length
            ).toFixed(1)
          )
        : 0;
    const compliantCount = candidates.filter((c) => c.compliant).length;
    const complianceRate =
      candidates.length > 0
        ? Number(((compliantCount / candidates.length) * 100).toFixed(1))
        : 0;
    const totalLatePunches = candidates.reduce((sum, c) => sum + c.lateDays, 0);

    const arrivalHistogram = Array.from(inMinutesBucket.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([bucket, count]) => ({ bucket, count }));

    return {
      candidates,
      dailyStats,
      arrivalHistogram,
      totalWorkingDays,
      totalWeekends,
      totalHolidays,
      avgAttendance,
      compliantCount,
      complianceRate,
      totalLatePunches,
    };
  }, [blocks, config.complianceThresholdPct]);

  if (error || !analytics) {
    return (
      <div className="max-w-[1440px] mx-auto px-6 py-16 text-center">
        <p className="text-base font-semibold text-slate-900">No Analytics Data Available</p>
        <p className="mt-1 text-sm text-slate-600">
          {error || 'Configure a roster and date span in the Builder to view compliance analytics.'}
        </p>
      </div>
    );
  }

  const sortedCandidates = [...analytics.candidates].sort((a, b) => {
    const dir = sortAsc ? 1 : -1;
    if (sortKey === 'name') return a.name.localeCompare(b.name) * dir;
    return (a[sortKey] - b[sortKey]) * dir;
  });

  const handleSort = (key: SortKey) => {
    if (sortKey === key) setSortAsc((v) => !v);
    else {
      setSortKey(key);
      setSortAsc(key === 'name' || key === 'attendancePct');
    }
  };

  const maxDailyTotal = Math.max(1, analytics.candidates.length);
  const maxHistCount = Math.max(1, ...analytics.arrivalHistogram.map((h) => h.count));

  return (
    <div className="max-w-[1440px] mx-auto px-6 py-6 space-y-6">
      <div className="pb-5 border-b border-slate-200">
        <h1 className="text-2xl font-bold tracking-tight text-slate-900">
          Compliance & Biometric Analytics
        </h1>
        <p className="mt-1 text-sm text-slate-600">
          {[
            config.centerName.trim() ? config.centerName.trim() : 'Consolidated attendance metrics',
            config.batchNo.trim() ? `Batch ${config.batchNo.trim()}` : '',
            `Threshold ${config.complianceThresholdPct}%`,
          ]
            .filter(Boolean)
            .join(' · ')}
        </p>
      </div>

      {/* 4-column Single-Elevation KPI Strip */}
      <div className="grid grid-cols-2 lg:grid-cols-4 bg-white border border-slate-200 rounded-xl divide-y lg:divide-y-0 lg:divide-x divide-slate-200">
        <div className="p-5">
          <p className="text-xs font-medium text-slate-500">Batch Average Attendance</p>
          <p className="mt-2 text-3xl font-bold font-mono tabular-nums text-slate-900">
            {analytics.avgAttendance.toFixed(1)}%
          </p>
          <p className="mt-1 text-xs text-slate-500">
            Across {analytics.candidates.length} trainees
          </p>
        </div>
        <div className="p-5">
          <p className="text-xs font-medium text-slate-500">
            Audit Compliance (&ge;{config.complianceThresholdPct}%)
          </p>
          <p className="mt-2 text-3xl font-bold font-mono tabular-nums text-emerald-700">
            {analytics.compliantCount}/{analytics.candidates.length}
          </p>
          <p className="mt-1 text-xs text-slate-500 tabular-nums">
            {analytics.complianceRate.toFixed(1)}% pass rate
          </p>
        </div>
        <div className="p-5">
          <p className="text-xs font-medium text-slate-500">Schedule Calendar Split</p>
          <p className="mt-2 text-3xl font-bold font-mono tabular-nums text-slate-900">
            {analytics.totalWorkingDays}d
          </p>
          <p className="mt-1 text-xs text-slate-500 tabular-nums">
            {analytics.totalWeekends} weekends · {analytics.totalHolidays} holidays
          </p>
        </div>
        <div className="p-5">
          <p className="text-xs font-medium text-slate-500">Late Arrivals Recorded</p>
          <p className="mt-2 text-3xl font-bold font-mono tabular-nums text-amber-700">
            {analytics.totalLatePunches}
          </p>
          <p className="mt-1 text-xs text-slate-500">
            Curve: {config.timeDistribution === 'gaussian' ? 'Gaussian Bell' : 'Uniform'}
          </p>
        </div>
      </div>

      {/* Charts Row */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Daily Attendance Trend */}
        <div className="lg:col-span-8 bg-white border border-slate-200 rounded-xl p-6 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div>
              <h2 className="text-base font-semibold text-slate-900">
                Daily Attendance Headcount
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Working day turnout (on-time, late, absent) vs weekends and holidays
              </p>
            </div>
            <div className="flex items-center gap-3 text-xs text-slate-600">
              <span className="inline-flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 bg-blue-600 inline-block" /> On-Time
              </span>
              <span className="inline-flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 bg-amber-500 inline-block" /> Late
              </span>
              <span className="inline-flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 bg-red-300 inline-block" /> Absent
              </span>
            </div>
          </div>

          <div className="overflow-x-auto pb-2">
            <div className="flex items-end gap-1.5 h-48 min-w-[540px] pt-4 border-b border-slate-200">
              {analytics.dailyStats.slice(0, 62).map((d) => {
                const pH = (d.present / maxDailyTotal) * 100;
                const lH = (d.late / maxDailyTotal) * 100;
                const aH = (d.absent / maxDailyTotal) * 100;
                return (
                  <div
                    key={d.dateStr}
                    title={
                      d.isOff
                        ? `${d.displayDate} (${d.offType})`
                        : `${d.displayDate}: ${d.present} on-time, ${d.late} late, ${d.absent} absent`
                    }
                    className="flex-1 min-w-[14px] h-full flex flex-col justify-end items-center group"
                  >
                    {d.isOff ? (
                      <div className="w-full h-full bg-slate-100/90 flex items-center justify-center rounded-t-xs">
                        <span className="text-[9px] font-mono text-slate-400 [writing-mode:vertical-rl] rotate-180">
                          {d.offType === 'Weekend' ? 'OFF' : 'HOL'}
                        </span>
                      </div>
                    ) : (
                      <div className="w-full h-full flex flex-col justify-end rounded-t-xs overflow-hidden">
                        {aH > 0 && (
                          <div style={{ height: `${aH}%` }} className="w-full bg-red-300" />
                        )}
                        {lH > 0 && (
                          <div style={{ height: `${lH}%` }} className="w-full bg-amber-500" />
                        )}
                        {pH > 0 && (
                          <div style={{ height: `${pH}%` }} className="w-full bg-blue-600" />
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
            <div className="flex items-center justify-between text-[10px] font-mono text-slate-500 mt-1.5">
              <span>{analytics.dailyStats[0]?.displayDate}</span>
              <span>
                {analytics.dailyStats.length > 62
                  ? 'Showing first 62 days'
                  : `${analytics.dailyStats.length} calendar days`}
              </span>
              <span>
                {analytics.dailyStats[Math.min(61, analytics.dailyStats.length - 1)]?.displayDate}
              </span>
            </div>
          </div>
        </div>

        {/* Arrival Punch Distribution */}
        <div className="lg:col-span-4 bg-white border border-slate-200 rounded-xl p-6 space-y-4">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              IN Punch Time Distribution
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              5-minute arrival buckets verifying realistic biometric spread
            </p>
          </div>

          {analytics.arrivalHistogram.length === 0 ? (
            <p className="text-xs text-slate-500 py-8 text-center">
              Enable auto-fill timestamps to inspect punch distribution.
            </p>
          ) : (
            <div className="space-y-2 max-h-52 overflow-y-auto pr-1">
              {analytics.arrivalHistogram.map((h) => {
                const widthPct = Math.max(4, Math.round((h.count / maxHistCount) * 100));
                return (
                  <div key={h.bucket} className="flex items-center gap-2 text-xs font-mono tabular-nums">
                    <span className="w-12 text-slate-600 shrink-0">{h.bucket}</span>
                    <div className="flex-1 bg-slate-100 h-4 rounded-xs overflow-hidden">
                      <div
                        style={{ width: `${widthPct}%` }}
                        className="h-full bg-slate-800"
                      />
                    </div>
                    <span className="w-8 text-right text-slate-700 shrink-0">{h.count}</span>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {/* Candidate Compliance Audit Table */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Candidate Audit Ledger
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Click any column heading to sort trainees by compliance risk, absents, or hours
            </p>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                <th className="px-6 py-3 w-12">#</th>
                <th className="px-4 py-3">
                  <button
                    type="button"
                    onClick={() => handleSort('name')}
                    className="inline-flex items-center gap-1 hover:text-slate-900 cursor-pointer"
                  >
                    <span>Candidate</span>
                    <ArrowUpDown className="w-3 h-3" />
                  </button>
                </th>
                <th className="px-4 py-3 text-right">Working</th>
                <th className="px-4 py-3 text-right">Present</th>
                <th className="px-4 py-3 text-right">
                  <button
                    type="button"
                    onClick={() => handleSort('absentDays')}
                    className="inline-flex items-center gap-1 hover:text-slate-900 cursor-pointer"
                  >
                    <span>Absent</span>
                    <ArrowUpDown className="w-3 h-3" />
                  </button>
                </th>
                <th className="px-4 py-3 text-right">
                  <button
                    type="button"
                    onClick={() => handleSort('lateDays')}
                    className="inline-flex items-center gap-1 hover:text-slate-900 cursor-pointer"
                  >
                    <span>Late</span>
                    <ArrowUpDown className="w-3 h-3" />
                  </button>
                </th>
                <th className="px-4 py-3 text-right">
                  <button
                    type="button"
                    onClick={() => handleSort('totalHours')}
                    className="inline-flex items-center gap-1 hover:text-slate-900 cursor-pointer"
                  >
                    <span>Total Hrs</span>
                    <ArrowUpDown className="w-3 h-3" />
                  </button>
                </th>
                <th className="px-4 py-3 text-right">
                  <button
                    type="button"
                    onClick={() => handleSort('attendancePct')}
                    className="inline-flex items-center gap-1 hover:text-slate-900 cursor-pointer"
                  >
                    <span>Attendance %</span>
                    <ArrowUpDown className="w-3 h-3" />
                  </button>
                </th>
                <th className="px-6 py-3 text-right">Audit Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {sortedCandidates.map((c, idx) => (
                <tr key={c.id} className="hover:bg-slate-50/80">
                  <td className="px-6 py-2.5 text-xs font-mono tabular-nums text-slate-400">
                    {String(idx + 1).padStart(2, '0')}
                  </td>
                  <td className="px-4 py-2.5">
                    <span className="font-medium text-slate-900">{c.name}</span>
                    <span className="ml-2 text-xs font-mono text-slate-400">{c.rollNo}</span>
                  </td>
                  <td className="px-4 py-2.5 text-right font-mono tabular-nums text-slate-600">
                    {c.workingDays}
                  </td>
                  <td className="px-4 py-2.5 text-right font-mono tabular-nums text-slate-900 font-medium">
                    {c.presentDays}
                  </td>
                  <td className="px-4 py-2.5 text-right font-mono tabular-nums text-red-700">
                    {c.absentDays}
                  </td>
                  <td className="px-4 py-2.5 text-right font-mono tabular-nums text-amber-700">
                    {c.lateDays}
                  </td>
                  <td className="px-4 py-2.5 text-right font-mono tabular-nums text-slate-700">
                    {c.totalHours.toFixed(1)}h
                  </td>
                  <td className="px-4 py-2.5 text-right font-mono tabular-nums font-semibold text-slate-900">
                    {c.attendancePct.toFixed(1)}%
                  </td>
                  <td className="px-6 py-2.5 text-right text-xs font-semibold">
                    {c.compliant ? (
                      <span className="text-emerald-700">Compliant</span>
                    ) : (
                      <span className="text-red-700">
                        Shortfall (-{c.neededForCompliance}d)
                      </span>
                    )}
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
