export type GenerationMode = 'single' | 'months_stacked' | 'months_separate';

export type WeekendPolicy = 'sunday' | 'sat_sun' | 'second_fourth_sat' | 'custom';

export type TimeFormatOption = '12h_short' | '12h_ampm' | '24h';

export type TimeDistribution = 'gaussian' | 'uniform';

export type AbsentPolicy = 'none' | 'fixed' | 'range' | 'target_percent';

export type ExcelTheme = 'classic_navy' | 'slate_minimal' | 'emerald_gov' | 'monochrome';

export type CellStatus =
  | 'present'
  | 'late'
  | 'half_day'
  | 'absent'
  | 'sunday'
  | 'holiday'
  | 'inactive'
  | 'blank';

export interface HolidayItem {
  date: string; // YYYY-MM-DD
  label: string;
}

export interface Candidate {
  id: string;
  name: string;
  rollNo?: string;
  joinDate?: string; // YYYY-MM-DD optional override
  dropDate?: string; // YYYY-MM-DD optional override
  manualAbsents?: string[]; // YYYY-MM-DD forced absents
}

export interface ManualCellEdit {
  status: CellStatus;
  inTime: string;
  outTime: string;
}

export interface PreviousPunchPair {
  inTime: string;
  outTime: string;
}

export interface PreviousAttendanceData {
  fileName: string;
  minDate: string; // YYYY-MM-DD
  maxDate: string; // YYYY-MM-DD
  datesCount: number;
  candidateNames: string[];
  /** key: normalized lowercase candidate name -> Record<YYYY-MM-DD, PreviousPunchPair> */
  records: Record<string, Record<string, PreviousPunchPair>>;
  detectedMeta?: {
    centerName?: string;
    district?: string;
    batchNo?: string;
    location?: string;
    qualificationPack?: string;
  };
}

export interface AttendanceConfig {
  // Center & Batch Metadata
  centerName: string;
  district: string;
  location: string;
  batchNo: string;
  qualificationPack: string;
  schemeName: string;
  trainerName: string;
  centerManager: string;
  reportTitle: string;

  // Date & Mode
  mode: GenerationMode;
  startDate: string; // YYYY-MM-DD
  endDate: string; // YYYY-MM-DD
  selectedMonths: number[]; // 1..12

  // Weekend Policy
  weekendPolicy: WeekendPolicy;
  customWeekendDays: number[]; // 0=Sun..6=Sat

  // Timing Rules
  autoFillTimes: boolean;
  inMin: string; // HH:MM (24h internally)
  inMax: string;
  outMin: string;
  outMax: string;
  timeFormat: TimeFormatOption;
  timeDistribution: TimeDistribution;
  minShiftHours: number;

  // Late & Half-Day Simulation
  lateSimulationEnabled: boolean;
  lateChancePercent: number;
  lateMaxExtraMinutes: number;
  halfDaySimulationEnabled: boolean;
  halfDayChancePercent: number;
  halfDayOutTime: string; // HH:MM 24h

  // Absent & Compliance Rules
  absentPolicy: AbsentPolicy;
  fixedAbsentsPerMonth: number;
  minAbsentsPerMonth: number;
  maxAbsentsPerMonth: number;
  targetMinAttendancePct: number;
  targetMaxAttendancePct: number;
  complianceThresholdPct: number;
  avoidConsecutiveAbsents: boolean;
  absentMarker: string; // default "0.00"

  // Holidays
  holidays: HolidayItem[];

  // Seed
  seed: number;

  // Excel Styling & Layout
  excelTheme: ExcelTheme;
  colorCodeCells: boolean;
  includeSummaryColumns: boolean;
  includeSummarySheet: boolean;
  freezePanes: boolean;
  landscapePrintSetup: boolean;
  includeSignatureBlock: boolean;
}

export interface DayCellData {
  dateStr: string; // YYYY-MM-DD
  displayDate: string; // DD-MM-YYYY
  weekdayShort: string; // Mon, Tue...
  status: CellStatus;
  inTime: string;
  outTime: string;
  inMinutes: number | null;
  outMinutes: number | null;
  durationMinutes: number;
  holidayLabel?: string;
  isManuallyEdited?: boolean;
  isPreservedFromPrevious?: boolean;
}

export interface CandidateBlockRow {
  slNo: number;
  candidate: Candidate;
  days: Record<string, DayCellData>; // key: YYYY-MM-DD
  presentDays: number;
  lateDays: number;
  halfDays: number;
  absentDays: number;
  workingDays: number;
  nonWorkingDays: number;
  attendancePct: number;
  totalHours: number;
  avgDailyHours: number;
  compliant: boolean;
}

export interface AttendanceBlock {
  id: string;
  sheetName: string;
  monthName: string;
  startDate: string; // YYYY-MM-DD
  endDate: string; // YYYY-MM-DD
  reportRangeLabel: string;
  dates: {
    dateStr: string;
    displayDate: string;
    weekdayShort: string;
    isWeekend: boolean;
    isHoliday: boolean;
    holidayLabel?: string;
  }[];
  rows: CandidateBlockRow[];
  workingDaysCount: number;
  weekendCount: number;
  holidayCount: number;
}

export interface SavedProfile {
  id: string;
  name: string;
  savedAt: string;
  config: AttendanceConfig;
}

export interface PythonFeatureFlags {
  bakeCurrentDefaults: boolean;
  bakeCurrentHolidays: boolean;
  includeGaussianTiming: boolean;
  includeLateAndHalfDay: boolean;
  includeComplianceGuard: boolean;
  includeFlexibleWeekends: boolean;
  includeStyledExcelAndSummary: boolean;
  includeCsvAndHtmlExport: boolean;
  includeTabbedScrollableGui: boolean;
  includeCliMode: boolean;
  includeJsonProfiles: boolean;
}

export interface ShiftPreset {
  id: string;
  label: string;
  inMin: string;
  inMax: string;
  outMin: string;
  outMax: string;
  halfDayOut: string;
}

export type ActiveTab = 'builder' | 'preview' | 'analytics' | 'python';
