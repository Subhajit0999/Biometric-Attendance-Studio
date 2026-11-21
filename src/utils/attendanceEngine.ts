import * as XLSX from 'xlsx';
import type {
  AttendanceBlock,
  AttendanceConfig,
  Candidate,
  CandidateBlockRow,
  CellStatus,
  DayCellData,
  HolidayItem,
  ManualCellEdit,
  PreviousAttendanceData,
  PreviousPunchPair,
  ShiftPreset,
} from '../types/attendance';

export const CANDIDATE_HEADER_KEYWORDS = [
  'candidate',
  'name',
  'student',
  'trainee',
  'participant',
  'candidate name',
  'student name',
  'full name',
];

export const SHIFT_PRESETS: ShiftPreset[] = [
  {
    id: 'standard',
    label: 'Standard (10:00 – 16:30)',
    inMin: '10:00',
    inMax: '10:30',
    outMin: '16:00',
    outMax: '16:30',
    halfDayOut: '13:15',
  },
  {
    id: 'morning',
    label: 'Morning (09:00 – 13:30)',
    inMin: '09:00',
    inMax: '09:20',
    outMin: '13:15',
    outMax: '13:40',
    halfDayOut: '11:15',
  },
  {
    id: 'afternoon',
    label: 'Afternoon (13:30 – 18:00)',
    inMin: '13:20',
    inMax: '13:45',
    outMin: '17:50',
    outMax: '18:15',
    halfDayOut: '15:45',
  },
  {
    id: 'full_day',
    label: 'Full Day (09:30 – 17:30)',
    inMin: '09:15',
    inMax: '09:45',
    outMin: '17:20',
    outMax: '17:50',
    halfDayOut: '13:30',
  },
];

export const INDIAN_HOLIDAY_PRESETS: HolidayItem[] = [
  // 2025
  { date: '2025-01-26', label: 'Republic Day' },
  { date: '2025-02-26', label: 'Maha Shivaratri' },
  { date: '2025-03-14', label: 'Holi' },
  { date: '2025-03-31', label: 'Id-ul-Fitr' },
  { date: '2025-04-14', label: 'Ambedkar Jayanti' },
  { date: '2025-04-18', label: 'Good Friday' },
  { date: '2025-05-01', label: 'Labour Day' },
  { date: '2025-08-15', label: 'Independence Day' },
  { date: '2025-10-02', label: 'Gandhi Jayanti / Dussehra' },
  { date: '2025-10-20', label: 'Diwali' },
  { date: '2025-11-05', label: 'Guru Nanak Jayanti' },
  { date: '2025-12-25', label: 'Christmas Day' },
  // 2026
  { date: '2026-01-26', label: 'Republic Day' },
  { date: '2026-02-15', label: 'Maha Shivaratri' },
  { date: '2026-03-04', label: 'Holi' },
  { date: '2026-03-21', label: 'Id-ul-Fitr' },
  { date: '2026-04-03', label: 'Good Friday' },
  { date: '2026-04-14', label: 'Ambedkar Jayanti' },
  { date: '2026-05-01', label: 'Labour Day' },
  { date: '2026-08-15', label: 'Independence Day' },
  { date: '2026-10-02', label: 'Gandhi Jayanti' },
  { date: '2026-10-20', label: 'Dussehra' },
  { date: '2026-11-08', label: 'Diwali' },
  { date: '2026-11-24', label: 'Guru Nanak Jayanti' },
  { date: '2026-12-25', label: 'Christmas Day' },
  // 2027
  { date: '2027-01-26', label: 'Republic Day' },
  { date: '2027-03-22', label: 'Holi' },
  { date: '2027-04-14', label: 'Ambedkar Jayanti' },
  { date: '2027-05-01', label: 'Labour Day' },
  { date: '2027-08-15', label: 'Independence Day' },
  { date: '2027-10-02', label: 'Gandhi Jayanti' },
  { date: '2027-10-29', label: 'Diwali' },
  { date: '2027-12-25', label: 'Christmas Day' },
];

export const SAMPLE_CANDIDATES: Candidate[] = [
  { id: 'c-1', name: 'Aarav Chatterjee', rollNo: 'BT-2026-001' },
  { id: 'c-2', name: 'Ananya Mukherjee', rollNo: 'BT-2026-002' },
  { id: 'c-3', name: 'Rohan Dasgupta', rollNo: 'BT-2026-003' },
  { id: 'c-4', name: 'Priya Sen', rollNo: 'BT-2026-004' },
  { id: 'c-5', name: 'Subhajit Saha', rollNo: 'BT-2026-005' },
  { id: 'c-6', name: 'Debashis Mondal', rollNo: 'BT-2026-006' },
  { id: 'c-7', name: 'Srijita Banerjee', rollNo: 'BT-2026-007' },
  { id: 'c-8', name: 'Koushik Ghosh', rollNo: 'BT-2026-008' },
  { id: 'c-9', name: 'Moumita Roy', rollNo: 'BT-2026-009' },
  { id: 'c-10', name: 'Sourav Paul', rollNo: 'BT-2026-010' },
  { id: 'c-11', name: 'Ritika Bose', rollNo: 'BT-2026-011' },
  { id: 'c-12', name: 'Arjun Halder', rollNo: 'BT-2026-012' },
  { id: 'c-13', name: 'Poulami Dutta', rollNo: 'BT-2026-013' },
  { id: 'c-14', name: 'Tanmoy Sarkar', rollNo: 'BT-2026-014' },
  { id: 'c-15', name: 'Swagata Karmakar', rollNo: 'BT-2026-015' },
  { id: 'c-16', name: 'Dipankar Biswas', rollNo: 'BT-2026-016' },
  { id: 'c-17', name: 'Nandini Dey', rollNo: 'BT-2026-017' },
  { id: 'c-18', name: 'Prasenjit Naskar', rollNo: 'BT-2026-018' },
];

export const DEFAULT_CONFIG: AttendanceConfig = {
  centerName: '',
  district: '',
  location: '',
  batchNo: '',
  qualificationPack: '',
  schemeName: '',
  trainerName: '',
  centerManager: '',
  reportTitle: 'Monthly Attendance Report',

  mode: 'single',
  startDate: '2026-11-01',
  endDate: '2026-11-30',
  selectedMonths: [11],

  weekendPolicy: 'sunday',
  customWeekendDays: [0],

  autoFillTimes: true,
  inMin: '10:00',
  inMax: '10:30',
  outMin: '16:00',
  outMax: '16:30',
  timeFormat: '12h_short',
  timeDistribution: 'gaussian',
  minShiftHours: 4,

  lateSimulationEnabled: true,
  lateChancePercent: 8,
  lateMaxExtraMinutes: 25,
  halfDaySimulationEnabled: false,
  halfDayChancePercent: 4,
  halfDayOutTime: '13:15',

  absentPolicy: 'range',
  fixedAbsentsPerMonth: 2,
  minAbsentsPerMonth: 1,
  maxAbsentsPerMonth: 3,
  targetMinAttendancePct: 80,
  targetMaxAttendancePct: 96,
  complianceThresholdPct: 75,
  avoidConsecutiveAbsents: true,
  absentMarker: '0.00',

  holidays: [
    { date: '2026-11-08', label: 'Diwali' },
    { date: '2026-11-24', label: 'Guru Nanak Jayanti' },
  ],

  seed: 202611,

  excelTheme: 'classic_navy',
  colorCodeCells: true,
  includeSummaryColumns: false,
  includeSummarySheet: false,
  freezePanes: true,
  landscapePrintSetup: true,
  includeSignatureBlock: true,
};

export const MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

const WEEKDAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

/** Deterministic seeded PRNG (Mulberry32) */
export function createSeededRng(seed: number): () => number {
  let t = seed >>> 0;
  return () => {
    t += 0x6d2b79f5;
    let r = Math.imul(t ^ (t >>> 15), 1 | t);
    r ^= r + Math.imul(r ^ (r >>> 7), 61 | r);
    return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
  };
}

/** Parse YYYY-MM-DD cleanly into local UTC-neutral Date */
export function parseISODate(iso: string): Date {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
}

export function formatISODate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

export function formatDisplayDate(iso: string): string {
  const [y, m, d] = iso.split('-');
  return `${d}-${m}-${y}`;
}

export function parseHHMMToMinutes(hhmm: string): number {
  const parts = hhmm.trim().split(':');
  const h = Math.max(0, Math.min(23, parseInt(parts[0] || '0', 10) || 0));
  const m = Math.max(0, Math.min(59, parseInt(parts[1] || '0', 10) || 0));
  return h * 60 + m;
}

export function formatMinutes(totalMins: number, format: AttendanceConfig['timeFormat']): string {
  const clamped = ((Math.round(totalMins) % 1440) + 1440) % 1440;
  const h24 = Math.floor(clamped / 60);
  const m = clamped % 60;
  const mm = String(m).padStart(2, '0');

  if (format === '24h') {
    return `${String(h24).padStart(2, '0')}:${mm}`;
  }

  const ampm = h24 >= 12 ? 'PM' : 'AM';
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  const hh = String(h12).padStart(2, '0');

  if (format === '12h_ampm') {
    return `${hh}:${mm} ${ampm}`;
  }
  return `${hh}:${mm}`;
}

/** Sample minutes in [minM, maxM] using either Gaussian (bell curve) or uniform */
function sampleMinutes(
  minM: number,
  maxM: number,
  distribution: AttendanceConfig['timeDistribution'],
  rng: () => number
): number {
  const lo = Math.min(minM, maxM);
  const hi = Math.max(minM, maxM);
  if (hi <= lo) return lo;

  if (distribution === 'uniform') {
    return lo + Math.floor(rng() * (hi - lo + 1));
  }

  // Box-Muller Gaussian centered slightly before midpoint (realistic biometric clustering)
  const u1 = Math.max(1e-7, rng());
  const u2 = rng();
  const z = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
  const mean = lo + (hi - lo) * 0.45;
  const std = (hi - lo) / 4.2;
  const val = Math.round(mean + z * std);
  return Math.max(lo, Math.min(hi, val));
}

export function isWeekendDate(d: Date, config: AttendanceConfig): boolean {
  const dow = d.getDay(); // 0=Sun..6=Sat
  switch (config.weekendPolicy) {
    case 'sunday':
      return dow === 0;
    case 'sat_sun':
      return dow === 0 || dow === 6;
    case 'second_fourth_sat': {
      if (dow === 0) return true;
      if (dow === 6) {
        const nthSat = Math.ceil(d.getDate() / 7);
        return nthSat === 2 || nthSat === 4;
      }
      return false;
    }
    case 'custom':
      return config.customWeekendDays.includes(dow);
    default:
      return dow === 0;
  }
}

function getDateSpanList(startIso: string, endIso: string): Date[] {
  const start = parseISODate(startIso);
  const end = parseISODate(endIso);
  if (end < start) return [];

  const list: Date[] = [];
  const cur = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  let guard = 0;
  while (cur <= end && guard < 400) {
    list.push(new Date(cur.getFullYear(), cur.getMonth(), cur.getDate()));
    cur.setDate(cur.getDate() + 1);
    guard++;
  }
  return list;
}

/** Pick N absent dates from workingDates, optionally avoiding consecutive days */
function pickAbsentDates(
  workingDates: string[],
  count: number,
  avoidConsecutive: boolean,
  rng: () => number
): Set<string> {
  const target = Math.max(0, Math.min(count, workingDates.length));
  if (target === 0) return new Set();

  // Fisher-Yates shuffle indices
  const indices = workingDates.map((_, idx) => idx);
  for (let i = indices.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [indices[i], indices[j]] = [indices[j], indices[i]];
  }

  if (!avoidConsecutive) {
    return new Set(indices.slice(0, target).map((idx) => workingDates[idx]));
  }

  const pickedIndices: number[] = [];
  for (const idx of indices) {
    if (pickedIndices.length >= target) break;
    const isAdj = pickedIndices.some((p) => Math.abs(p - idx) === 1);
    if (!isAdj) {
      pickedIndices.push(idx);
    }
  }
  // If non-consecutive couldn't reach count, top up from remaining
  if (pickedIndices.length < target) {
    for (const idx of indices) {
      if (pickedIndices.length >= target) break;
      if (!pickedIndices.includes(idx)) {
        pickedIndices.push(idx);
      }
    }
  }
  return new Set(pickedIndices.map((idx) => workingDates[idx]));
}

function buildBlockForSpan(
  blockId: string,
  sheetName: string,
  monthName: string,
  spanStartIso: string,
  spanEndIso: string,
  candidates: Candidate[],
  config: AttendanceConfig,
  manualEdits: Record<string, ManualCellEdit>,
  rng: () => number,
  previousData?: PreviousAttendanceData | null
): AttendanceBlock {
  const holidayMap = new Map<string, string>();
  for (const h of config.holidays) {
    holidayMap.set(h.date, h.label || 'Holiday');
  }

  const dateObjs = getDateSpanList(spanStartIso, spanEndIso);
  let weekendCount = 0;
  let holidayCount = 0;
  let workingDaysCount = 0;

  const dates = dateObjs.map((d) => {
    const dateStr = formatISODate(d);
    const isWeekend = isWeekendDate(d, config);
    const isHoliday = !isWeekend && holidayMap.has(dateStr);
    if (isWeekend) weekendCount++;
    else if (isHoliday) holidayCount++;
    else workingDaysCount++;

    return {
      dateStr,
      displayDate: formatDisplayDate(dateStr),
      weekdayShort: WEEKDAY_SHORT[d.getDay()],
      isWeekend,
      isHoliday,
      holidayLabel: holidayMap.get(dateStr),
    };
  });

  // Group working dates by YYYY-MM so monthly absent quotas apply cleanly even in multi-month Single Range
  const workingByMonth = new Map<string, string[]>();
  for (const dt of dates) {
    if (!dt.isWeekend && !dt.isHoliday) {
      const ym = dt.dateStr.slice(0, 7);
      const arr = workingByMonth.get(ym) || [];
      arr.push(dt.dateStr);
      workingByMonth.set(ym, arr);
    }
  }

  const inMinM = parseHHMMToMinutes(config.inMin);
  const inMaxM = parseHHMMToMinutes(config.inMax);
  const outMinM = parseHHMMToMinutes(config.outMin);
  const outMaxM = parseHHMMToMinutes(config.outMax);
  const halfDayOutM = parseHHMMToMinutes(config.halfDayOutTime);
  const minShiftMins = Math.max(0, Math.round(config.minShiftHours * 60));

  const rows: CandidateBlockRow[] = candidates.map((cand, index) => {
    const forcedAbsents = new Set(cand.manualAbsents || []);
    const autoAbsents = new Set<string>();
    const prevCandRecords = previousData?.records[cand.name.trim().toLowerCase()] || {};

    // Determine working dates eligible for this candidate (respecting joinDate/dropDate & previous locked dates)
    for (const [, monthWorkDates] of workingByMonth.entries()) {
      const activeWorkDates = monthWorkDates.filter((ds) => {
        if (cand.joinDate && ds < cand.joinDate) return false;
        if (cand.dropDate && ds > cand.dropDate) return false;
        return true;
      });

      const preservedAbsentsInMonth = activeWorkDates.filter((ds) => {
        const p = prevCandRecords[ds];
        if (!p) return false;
        const v = p.inTime.trim();
        return v === '0.00' || v === config.absentMarker || v.toUpperCase() === 'A';
      }).length;

      const alreadyForced = activeWorkDates.filter(
        (ds) => !prevCandRecords[ds] && forcedAbsents.has(ds)
      ).length;
      const availableForAuto = activeWorkDates.filter(
        (ds) => !prevCandRecords[ds] && !forcedAbsents.has(ds)
      );

      let desiredTotalAbsents = 0;
      if (config.absentPolicy === 'fixed') {
        desiredTotalAbsents = config.fixedAbsentsPerMonth;
      } else if (config.absentPolicy === 'range') {
        const lo = Math.min(config.minAbsentsPerMonth, config.maxAbsentsPerMonth);
        const hi = Math.max(config.minAbsentsPerMonth, config.maxAbsentsPerMonth);
        desiredTotalAbsents = lo + Math.floor(rng() * (hi - lo + 1));
      } else if (config.absentPolicy === 'target_percent') {
        const loPct = Math.min(config.targetMinAttendancePct, config.targetMaxAttendancePct);
        const hiPct = Math.max(config.targetMinAttendancePct, config.targetMaxAttendancePct);
        const chosenPct = loPct + rng() * (hiPct - loPct);
        desiredTotalAbsents = Math.round(activeWorkDates.length * (1 - chosenPct / 100));
      }

      // Enforce compliance threshold cap if target_percent is selected
      if (config.absentPolicy === 'target_percent' && activeWorkDates.length > 0) {
        const maxAllowedAbsents = Math.floor(
          activeWorkDates.length * (1 - config.complianceThresholdPct / 100)
        );
        desiredTotalAbsents = Math.min(desiredTotalAbsents, Math.max(0, maxAllowedAbsents));
      }

      const additionalNeeded = Math.max(
        0,
        desiredTotalAbsents - alreadyForced - preservedAbsentsInMonth
      );
      const picked = pickAbsentDates(
        availableForAuto,
        additionalNeeded,
        config.avoidConsecutiveAbsents,
        rng
      );
      picked.forEach((ds) => autoAbsents.add(ds));
    }

    const days: Record<string, DayCellData> = {};
    let presentDays = 0;
    let lateDays = 0;
    let halfDays = 0;
    let absentDays = 0;
    let candWorkingDays = 0;
    let candNonWorkingDays = 0;
    let totalMinutes = 0;

    for (const dt of dates) {
      const editKey = `${blockId}__${cand.id}__${dt.dateStr}`;
      const manual = manualEdits[editKey];
      const prevPair = prevCandRecords[dt.dateStr];

      let status: CellStatus = 'present';
      let inTime = '';
      let outTime = '';
      let inMins: number | null = null;
      let outMins: number | null = null;
      let durationMins = 0;
      let isPreserved = false;

      const isBeforeJoin = Boolean(cand.joinDate && dt.dateStr < cand.joinDate);
      const isAfterDrop = Boolean(cand.dropDate && dt.dateStr > cand.dropDate);

      if (manual) {
        status = manual.status;
        inTime = manual.inTime;
        outTime = manual.outTime;
        if (status === 'present' || status === 'late' || status === 'half_day') {
          const parsedIn = tryParseDisplayTime(inTime);
          const parsedOut = tryParseDisplayTime(outTime);
          inMins = parsedIn;
          outMins = parsedOut;
          if (parsedIn !== null && parsedOut !== null && parsedOut > parsedIn) {
            durationMins = parsedOut - parsedIn;
          }
        }
      } else if (prevPair) {
        isPreserved = true;
        inTime = prevPair.inTime;
        outTime = prevPair.outTime;
        const normIn = inTime.trim();
        if (dt.isWeekend) {
          status = 'sunday';
        } else if (dt.isHoliday) {
          status = 'holiday';
        } else if (normIn === '' || normIn === '—') {
          status = 'blank';
        } else if (normIn === '-') {
          status = 'inactive';
        } else if (normIn === '0.00' || normIn === config.absentMarker || normIn.toUpperCase() === 'A') {
          status = 'absent';
        } else {
          const parsedIn = tryParseDisplayTime(inTime);
          const parsedOut = tryParseDisplayTime(outTime);
          inMins = parsedIn;
          outMins = parsedOut;
          if (parsedIn !== null && parsedOut !== null) {
            let effOut = parsedOut;
            // Handle 12h short format without AM/PM (e.g. 04:15 after 10:15 -> 16:15)
            if (effOut <= parsedIn && effOut + 720 > parsedIn && effOut < 720) {
              effOut += 720;
              outMins = effOut;
            }
            if (effOut > parsedIn) {
              durationMins = effOut - parsedIn;
            }
          }
          if (parsedIn !== null && parsedIn > Math.max(inMinM, inMaxM) + 2) {
            status = 'late';
          } else {
            status = 'present';
          }
        }
      } else if (dt.isWeekend) {
        status = 'sunday';
        inTime = config.absentMarker;
        outTime = config.absentMarker;
      } else if (dt.isHoliday) {
        status = 'holiday';
        inTime = config.absentMarker;
        outTime = config.absentMarker;
      } else if (isBeforeJoin || isAfterDrop) {
        status = 'inactive';
        inTime = '-';
        outTime = '-';
      } else if (forcedAbsents.has(dt.dateStr) || autoAbsents.has(dt.dateStr)) {
        status = 'absent';
        inTime = config.absentMarker;
        outTime = config.absentMarker;
      } else if (!config.autoFillTimes) {
        status = 'blank';
        inTime = '';
        outTime = '';
      } else {
        // Generate biometric IN/OUT times
        const isLate =
          config.lateSimulationEnabled && rng() * 100 < config.lateChancePercent;
        const isHalf =
          !isLate &&
          config.halfDaySimulationEnabled &&
          rng() * 100 < config.halfDayChancePercent;

        if (isLate) {
          status = 'late';
          const extra = 3 + Math.floor(rng() * Math.max(1, config.lateMaxExtraMinutes));
          inMins = Math.min(1430, Math.max(inMinM, inMaxM) + extra);
          outMins = sampleMinutes(outMinM, outMaxM, config.timeDistribution, rng);
        } else if (isHalf) {
          status = 'half_day';
          inMins = sampleMinutes(inMinM, inMaxM, config.timeDistribution, rng);
          const jitter = Math.floor((rng() - 0.5) * 20);
          outMins = Math.max(inMins + 90, halfDayOutM + jitter);
        } else {
          status = 'present';
          inMins = sampleMinutes(inMinM, inMaxM, config.timeDistribution, rng);
          outMins = sampleMinutes(outMinM, outMaxM, config.timeDistribution, rng);
          if (outMins - inMins < minShiftMins && minShiftMins > 0) {
            outMins = Math.min(1435, Math.max(outMins, inMins + minShiftMins));
          }
        }

        if (outMins <= inMins) {
          outMins = Math.min(1435, inMins + 60);
        }

        inTime = formatMinutes(inMins, config.timeFormat);
        outTime = formatMinutes(outMins, config.timeFormat);
        durationMins = outMins - inMins;
      }

      // Tally metrics
      if (status === 'sunday' || status === 'holiday' || status === 'inactive') {
        candNonWorkingDays++;
      } else {
        candWorkingDays++;
        if (status === 'present') {
          presentDays++;
          totalMinutes += durationMins;
        } else if (status === 'late') {
          presentDays++;
          lateDays++;
          totalMinutes += durationMins;
        } else if (status === 'half_day') {
          presentDays += 0.5;
          halfDays++;
          totalMinutes += durationMins;
        } else if (status === 'absent') {
          absentDays++;
        }
      }

      days[dt.dateStr] = {
        dateStr: dt.dateStr,
        displayDate: dt.displayDate,
        weekdayShort: dt.weekdayShort,
        status,
        inTime,
        outTime,
        inMinutes: inMins,
        outMinutes: outMins,
        durationMinutes: durationMins,
        holidayLabel: dt.holidayLabel,
        isManuallyEdited: Boolean(manual),
        isPreservedFromPrevious: isPreserved,
      };
    }

    const attendancePct =
      candWorkingDays > 0 ? Number(((presentDays / candWorkingDays) * 100).toFixed(1)) : 100;
    const totalHours = Number((totalMinutes / 60).toFixed(1));
    const effectiveAttendedDays = presentDays > 0 ? presentDays : 1;
    const avgDailyHours =
      presentDays > 0 ? Number((totalMinutes / 60 / effectiveAttendedDays).toFixed(2)) : 0;

    return {
      slNo: index + 1,
      candidate: cand,
      days,
      presentDays,
      lateDays,
      halfDays,
      absentDays,
      workingDays: candWorkingDays,
      nonWorkingDays: candNonWorkingDays,
      attendancePct,
      totalHours,
      avgDailyHours,
      compliant: attendancePct >= config.complianceThresholdPct,
    };
  });

  return {
    id: blockId,
    sheetName,
    monthName,
    startDate: spanStartIso,
    endDate: spanEndIso,
    reportRangeLabel: `${formatDisplayDate(spanStartIso)} - ${formatDisplayDate(spanEndIso)}`,
    dates,
    rows,
    workingDaysCount,
    weekendCount,
    holidayCount,
  };
}

function tryParseDisplayTime(val: string): number | null {
  const s = val.trim().toUpperCase();
  const match = s.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?$/);
  if (!match) return null;
  let h = parseInt(match[1], 10);
  const m = parseInt(match[2], 10);
  const ampm = match[3];
  if (ampm === 'PM' && h !== 12) h += 12;
  if (ampm === 'AM' && h === 12) h = 0;
  return h * 60 + m;
}

/** Generate all attendance blocks according to mode (single, stacked months, separate monthly sheets) */
export function generateAttendanceBlocks(
  candidates: Candidate[],
  config: AttendanceConfig,
  manualEdits: Record<string, ManualCellEdit> = {},
  previousData?: PreviousAttendanceData | null
): { blocks: AttendanceBlock[]; error?: string } {
  if (!candidates.length) {
    return { blocks: [], error: 'Add at least one candidate to generate an attendance sheet.' };
  }

  const start = parseISODate(config.startDate);
  const end = parseISODate(config.endDate);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    return { blocks: [], error: 'Invalid start or end date.' };
  }
  if (end < start) {
    return { blocks: [], error: 'End date must be on or after the start date.' };
  }

  const rng = createSeededRng(config.seed);

  if (config.mode === 'single') {
    const startMonth = MONTH_NAMES[start.getMonth()];
    const endMonth = MONTH_NAMES[end.getMonth()];
    const monthLabel =
      start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()
        ? `${startMonth} ${start.getFullYear()}`
        : `${startMonth} ${start.getFullYear()} – ${endMonth} ${end.getFullYear()}`;

    const block = buildBlockForSpan(
      'single',
      'Attendance_ALL',
      monthLabel,
      config.startDate,
      config.endDate,
      candidates,
      config,
      manualEdits,
      rng,
      previousData
    );
    return { blocks: [block] };
  }

  // Multi-month mode (stacked or separate)
  if (!config.selectedMonths.length) {
    return {
      blocks: [],
      error: 'Select at least one month for By-Months mode.',
    };
  }

  const ymList: { year: number; month: number }[] = [];
  let curY = start.getFullYear();
  let curM = start.getMonth() + 1;
  const endY = end.getFullYear();
  const endM = end.getMonth() + 1;

  let guard = 0;
  while ((curY < endY || (curY === endY && curM <= endM)) && guard < 36) {
    ymList.push({ year: curY, month: curM });
    if (curM === 12) {
      curY++;
      curM = 1;
    } else {
      curM++;
    }
    guard++;
  }

  const blocks: AttendanceBlock[] = [];
  for (const { year, month } of ymList) {
    if (!config.selectedMonths.includes(month)) continue;
    const firstOfMonth = new Date(year, month - 1, 1);
    const lastOfMonth = new Date(year, month, 0);
    const spanStart = firstOfMonth > start ? firstOfMonth : start;
    const spanEnd = lastOfMonth < end ? lastOfMonth : end;
    if (spanEnd < spanStart) continue;

    const mName = MONTH_NAMES[month - 1];
    const sheetName =
      config.mode === 'months_separate'
        ? `${mName.slice(0, 3)}_${year}`
        : 'Attendance_ALL';

    blocks.push(
      buildBlockForSpan(
        `${year}-${String(month).padStart(2, '0')}`,
        sheetName,
        `${mName} ${year}`,
        formatISODate(spanStart),
        formatISODate(spanEnd),
        candidates,
        config,
        manualEdits,
        rng,
        previousData
      )
    );
  }

  if (!blocks.length) {
    return {
      blocks: [],
      error:
        'Selected month(s) do not overlap with your Start Date and End Date range. Adjust either the date span or selected months.',
    };
  }

  return { blocks };
}

/** Smart candidate column sniffer mirroring & improving original Python script */
export function sniffCandidateColumn(headers: string[]): string {
  if (!headers.length) return '';
  for (const col of headers) {
    const norm = col.trim().toLowerCase();
    for (const kw of CANDIDATE_HEADER_KEYWORDS) {
      if (norm.includes(kw)) return col;
    }
  }
  const first = headers[0].trim().toLowerCase();
  if (['sl', 's.no', 's/no', 'sl no', 'slno', 'sr', '#'].some((x) => first.includes(x))) {
    if (headers.length >= 2) return headers[1];
  }
  return headers[0];
}

export interface ParsedRosterFile {
  fileName: string;
  headers: string[];
  rows: Record<string, string>[];
  detectedNameColumn: string;
  detectedRollColumn: string | null;
}

export async function parseRosterFile(file: File): Promise<ParsedRosterFile> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) {
    throw new Error('The uploaded file contains no worksheets.');
  }
  const sheet = workbook.Sheets[firstSheet];
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: '',
    raw: false,
  });

  if (!rawRows.length) {
    throw new Error('The uploaded sheet has no rows.');
  }

  const headers = Object.keys(rawRows[0]);
  const rows: Record<string, string>[] = rawRows.map((r) => {
    const out: Record<string, string> = {};
    for (const h of headers) {
      out[h] = String(r[h] ?? '').trim();
    }
    return out;
  });

  const detectedNameColumn = sniffCandidateColumn(headers);
  const detectedRollColumn =
    headers.find((h) => {
      const n = h.trim().toLowerCase();
      return (
        h !== detectedNameColumn &&
        (n.includes('roll') || n.includes('id') || n.includes('reg') || n.includes('code'))
      );
    }) || null;

  return {
    fileName: file.name,
    headers,
    rows,
    detectedNameColumn,
    detectedRollColumn,
  };
}

export function extractCandidatesFromParsedFile(
  parsed: ParsedRosterFile,
  nameColumn: string,
  rollColumn: string | null,
  deduplicate: boolean
): Candidate[] {
  const seen = new Set<string>();
  const result: Candidate[] = [];

  for (let i = 0; i < parsed.rows.length; i++) {
    const name = (parsed.rows[i][nameColumn] || '').trim();
    if (!name || name.toLowerCase() === 'nan' || name.toLowerCase() === 'none') continue;
    if (name.toLowerCase() === nameColumn.trim().toLowerCase()) continue;

    const key = name.toLowerCase();
    if (deduplicate && seen.has(key)) continue;
    seen.add(key);

    const rollNo = rollColumn ? (parsed.rows[i][rollColumn] || '').trim() : undefined;
    result.push({
      id: `imp-${Date.now()}-${i}`,
      name,
      rollNo: rollNo || undefined,
    });
  }

  return result;
}

/** Try to parse a date header cell (DD-MM-YYYY, YYYY-MM-DD, DD/MM/YYYY, or Excel Date) -> YYYY-MM-DD */
function parseAttendanceHeaderDate(raw: unknown): string | null {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return formatISODate(raw);
  }
  const s = String(raw ?? '').trim();
  if (!s) return null;

  // DD-MM-YYYY or DD/MM/YYYY (optionally followed by weekday text)
  const dmy = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]);
    const year = Number(dmy[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // YYYY-MM-DD or YYYY/MM/DD
  const ymd = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (ymd) {
    const year = Number(ymd[1]);
    const month = Number(ymd[2]);
    const day = Number(ymd[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  return null;
}

/**
 * Parse a previously generated Biometric Attendance Excel (.xlsx/.xls/.csv)
 * so the user can continue generating additional months or date ranges
 * while keeping all earlier IN/OUT timestamps intact.
 */
export async function parsePreviousAttendanceExcel(
  file: File
): Promise<PreviousAttendanceData> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });

  if (!workbook.SheetNames.length) {
    throw new Error('The uploaded workbook contains no worksheets.');
  }

  const candidateOrder: string[] = [];
  const candidateSeen = new Set<string>();
  const records: Record<string, Record<string, PreviousPunchPair>> = {};
  const allDates = new Set<string>();
  const detectedMeta: PreviousAttendanceData['detectedMeta'] = {};

  for (const sheetName of workbook.SheetNames) {
    if (sheetName.toLowerCase().includes('summary')) continue;
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: '',
      raw: false,
    });

    let activeDateCols: Map<number, string> | null = null;
    let nameColIdx = 1;
    let typeColIdx = 2;
    let lastCandidateName = '';

    for (let r = 0; r < matrix.length; r++) {
      const row = matrix[r];
      if (!Array.isArray(row) || row.length === 0) continue;

      const firstCell = String(row[0] ?? '').trim();

      // Opportunistically extract banner metadata if available
      if (firstCell && !activeDateCols) {
        const distMatch = firstCell.match(/^(.*?)\s*\(Dist:\s*([^)]+)\)/i);
        if (distMatch) {
          if (!detectedMeta.centerName && distMatch[1].trim()) {
            detectedMeta.centerName = distMatch[1].trim();
          }
          if (!detectedMeta.district && distMatch[2].trim()) {
            detectedMeta.district = distMatch[2].trim();
          }
        }
        const batchMatch = firstCell.match(/\(Batch[:-]?\s*([^)]+)\)/i);
        if (batchMatch && !detectedMeta.batchNo) {
          detectedMeta.batchNo = batchMatch[1].trim();
        }
        const locMatch = firstCell.match(/Location:\s*(.+)$/i);
        if (locMatch && !detectedMeta.location) {
          detectedMeta.location = locMatch[1].trim();
        }
        const qpMatch = firstCell.match(/Qualification\s*Pack:\s*(.+)$/i);
        if (qpMatch && !detectedMeta.qualificationPack) {
          detectedMeta.qualificationPack = qpMatch[1].trim();
        }
      }

      // Check if this row is a Date Header row (has at least 1 parseable date column)
      const dateCols = new Map<number, string>();
      for (let c = 0; c < row.length; c++) {
        const iso = parseAttendanceHeaderDate(row[c]);
        if (iso) {
          dateCols.set(c, iso);
        }
      }

      if (dateCols.size > 0) {
        activeDateCols = dateCols;
        // Locate Candidate Name & Type columns in this header row
        nameColIdx = 1;
        typeColIdx = 2;
        for (let c = 0; c < Math.min(row.length, 5); c++) {
          const hdr = String(row[c] ?? '').trim().toLowerCase();
          if (hdr.includes('name') || hdr.includes('candidate')) nameColIdx = c;
          if (hdr === 'type') typeColIdx = c;
        }
        lastCandidateName = '';
        continue;
      }

      // Reset block if another banner or signature line starts
      if (
        firstCell.toUpperCase().includes('BIOMETRIC ATTENDANCE') ||
        String(row[1] ?? '').toLowerCase().startsWith('trainer')
      ) {
        activeDateCols = null;
        lastCandidateName = '';
        continue;
      }

      if (!activeDateCols) continue;

      const rowType = String(row[typeColIdx] ?? '').trim().toLowerCase();
      if (rowType !== 'in' && rowType !== 'out') continue;

      const rawName = String(row[nameColIdx] ?? '').trim();
      const candName = rawName || lastCandidateName;
      if (!candName) continue;

      if (rowType === 'in') {
        lastCandidateName = candName;
      }

      const normKey = candName.toLowerCase();
      if (!candidateSeen.has(normKey)) {
        candidateSeen.add(normKey);
        candidateOrder.push(candName);
      }
      if (!records[normKey]) {
        records[normKey] = {};
      }

      for (const [colIdx, isoDate] of activeDateCols.entries()) {
        const cellVal = String(row[colIdx] ?? '').trim();
        allDates.add(isoDate);
        if (!records[normKey][isoDate]) {
          records[normKey][isoDate] = { inTime: '', outTime: '' };
        }
        if (rowType === 'in') {
          records[normKey][isoDate].inTime = cellVal;
        } else {
          records[normKey][isoDate].outTime = cellVal;
        }
      }
    }
  }

  const sortedDates = Array.from(allDates).sort();
  if (!sortedDates.length || !candidateOrder.length) {
    throw new Error(
      'Could not find previous attendance dates and In/Out rows in this Excel file. Make sure you upload an attendance sheet generated by this tool or your original Python script.'
    );
  }

  return {
    fileName: file.name,
    minDate: sortedDates[0],
    maxDate: sortedDates[sortedDates.length - 1],
    datesCount: sortedDates.length,
    candidateNames: candidateOrder,
    records,
    detectedMeta,
  };
}

