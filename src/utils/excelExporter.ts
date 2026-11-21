import ExcelJS from 'exceljs';
import type {
  AttendanceBlock,
  AttendanceConfig,
  Candidate,
  CellStatus,
  ExcelTheme,
} from '../types/attendance';
import { formatDisplayDate } from './attendanceEngine';

interface ThemePalette {
  bannerBg: string;
  bannerText: string;
  subBannerBg: string;
  subBannerText: string;
  tableHeaderBg: string;
  tableHeaderText: string;
  weekdayBg: string;
  weekdayText: string;
  borderHex: string;
  zebraBg: string;
}

const THEME_PALETTES: Record<ExcelTheme, ThemePalette> = {
  classic_navy: {
    bannerBg: '1E3A5F',
    bannerText: 'FFFFFF',
    subBannerBg: 'EFF6FF',
    subBannerText: '0F172A',
    tableHeaderBg: '1E293B',
    tableHeaderText: 'FFFFFF',
    weekdayBg: 'E2E8F0',
    weekdayText: '334155',
    borderHex: 'CBD5E1',
    zebraBg: 'F8FAFC',
  },
  slate_minimal: {
    bannerBg: '0F172A',
    bannerText: 'FFFFFF',
    subBannerBg: 'F1F5F9',
    subBannerText: '0F172A',
    tableHeaderBg: '334155',
    tableHeaderText: 'FFFFFF',
    weekdayBg: 'E2E8F0',
    weekdayText: '334155',
    borderHex: 'CBD5E1',
    zebraBg: 'F8FAFC',
  },
  emerald_gov: {
    bannerBg: '065F46',
    bannerText: 'FFFFFF',
    subBannerBg: 'ECFDF5',
    subBannerText: '064E3B',
    tableHeaderBg: '047857',
    tableHeaderText: 'FFFFFF',
    weekdayBg: 'D1FAE5',
    weekdayText: '065F46',
    borderHex: 'A7F3D0',
    zebraBg: 'F0FDF4',
  },
  monochrome: {
    bannerBg: 'FFFFFF',
    bannerText: '000000',
    subBannerBg: 'FFFFFF',
    subBannerText: '000000',
    tableHeaderBg: 'F1F5F9',
    tableHeaderText: '000000',
    weekdayBg: 'F8FAFC',
    weekdayText: '334155',
    borderHex: '000000',
    zebraBg: 'FFFFFF',
  },
};

function getCellStatusColors(
  status: CellStatus,
  colorCode: boolean
): { bg?: string; fg?: string; bold?: boolean } {
  if (!colorCode) return {};
  switch (status) {
    case 'sunday':
      return { bg: 'E2E8F0', fg: '475569' };
    case 'holiday':
      return { bg: 'DBEAFE', fg: '1E40AF', bold: true };
    case 'absent':
      return { bg: 'FEE2E2', fg: '991B1B', bold: true };
    case 'late':
      return { bg: 'FEF3C7', fg: '92400E' };
    case 'half_day':
      return { bg: 'F3E8FF', fg: '6B21A8' };
    case 'inactive':
      return { bg: 'F1F5F9', fg: '94A3B8' };
    default:
      return {};
  }
}

function colLetter(colIdx1Based: number): string {
  let n = colIdx1Based;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function buildOutputFilename(config: AttendanceConfig, ext: 'xlsx' | 'csv'): string {
  const safe = (v: string) =>
    v.trim().replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '');
  const prefixParts = [safe(config.district), safe(config.location), safe(config.batchNo)].filter(Boolean);
  const s = formatDisplayDate(config.startDate).replace(/-/g, '');
  const e = formatDisplayDate(config.endDate).replace(/-/g, '');
  const prefix = prefixParts.length ? `${prefixParts.join('_')}_` : '';
  return `${prefix}Attendance_${s}_${e}.${ext}`;
}

function triggerBrowserDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function exportAttendanceExcel(
  blocks: AttendanceBlock[],
  config: AttendanceConfig
): Promise<string> {
  if (!blocks.length) {
    throw new Error('No attendance blocks available to export.');
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Biometric Attendance Studio';
  wb.created = new Date();

  const palette = THEME_PALETTES[config.excelTheme];
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: `FF${palette.borderHex}` } },
    bottom: { style: 'thin', color: { argb: `FF${palette.borderHex}` } },
    left: { style: 'thin', color: { argb: `FF${palette.borderHex}` } },
    right: { style: 'thin', color: { argb: `FF${palette.borderHex}` } },
  };

  // Track current start row per worksheet (supports stacked multi-month or separate sheets)
  const sheetRowCursor = new Map<string, number>();

  for (const block of blocks) {
    let ws = wb.getWorksheet(block.sheetName);
    if (!ws) {
      ws = wb.addWorksheet(block.sheetName);
      sheetRowCursor.set(block.sheetName, 1);

      if (config.landscapePrintSetup) {
        ws.pageSetup = {
          orientation: 'landscape',
          paperSize: 9, // A4
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          margins: {
            left: 0.3,
            right: 0.3,
            top: 0.4,
            bottom: 0.4,
            header: 0.2,
            footer: 0.2,
          },
        };
      }
    }

    const startRow = sheetRowCursor.get(block.sheetName) || 1;
    const totalCols = 3 + block.dates.length;
    const lastLetter = colLetter(totalCols);

    // 6-row merged header block (matching & enhancing original format)
    const centerMetaParts = [
      config.centerName.trim(),
      config.district.trim() ? `(Dist: ${config.district.trim()})` : '',
      config.batchNo.trim() ? `(Batch: ${config.batchNo.trim()})` : '',
    ].filter(Boolean);

    const periodLine = config.location.trim()
      ? `Period: ${block.reportRangeLabel} | Location: ${config.location.trim()}`
      : `Period: ${block.reportRangeLabel}`;

    const bannerLines = [
      'BIOMETRIC ATTENDANCE SHEET',
      centerMetaParts.join(' '),
      `${config.reportTitle || 'Monthly Attendance Report'}${config.schemeName.trim() ? ` — ${config.schemeName.trim()}` : ''}`,
      block.monthName || '',
      periodLine,
      config.qualificationPack.trim() ? `Qualification Pack: ${config.qualificationPack.trim()}` : '',
    ];

    bannerLines.forEach((line, idx) => {
      const r = startRow + idx;
      ws!.mergeCells(`A${r}:${lastLetter}${r}`);
      const cell = ws!.getCell(`A${r}`);
      cell.value = line;
      const isPrimary = idx === 0;
      cell.font = {
        name: 'Calibri',
        size: isPrimary ? 14 : idx === 1 ? 12 : 10,
        bold: idx <= 2 || idx === 5,
        color: { argb: `FF${isPrimary ? palette.bannerText : palette.subBannerText}` },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${isPrimary ? palette.bannerBg : palette.subBannerBg}` },
      };
      ws!.getRow(r).height = isPrimary ? 26 : 19;

      for (let c = 1; c <= totalCols; c++) {
        ws!.getCell(r, c).border = thinBorder;
      }
    });

    const rDate = startRow + 6;
    const rWeekday = startRow + 7;
    const rDataStart = startRow + 8;

    // Date header row + weekday row
    const fixedHeaders = ['Sl No', 'Candidate Name', 'Type'];
    fixedHeaders.forEach((h, i) => {
      const c = i + 1;
      const cell = ws!.getCell(rDate, c);
      cell.value = h;
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: `FF${palette.tableHeaderText}` },
      };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${palette.tableHeaderBg}` },
      };
      cell.border = thinBorder;

      const sub = ws!.getCell(rWeekday, c);
      sub.value = '';
      sub.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${palette.weekdayBg}` },
      };
      sub.border = thinBorder;
    });

    block.dates.forEach((dt, idx) => {
      const c = 4 + idx;
      const dCell = ws!.getCell(rDate, c);
      dCell.value = dt.displayDate;
      dCell.font = {
        name: 'Calibri',
        size: 9,
        bold: true,
        color: { argb: `FF${palette.tableHeaderText}` },
      };
      dCell.alignment = { horizontal: 'center', vertical: 'middle' };
      dCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${palette.tableHeaderBg}` },
      };
      dCell.border = thinBorder;

      const wCell = ws!.getCell(rWeekday, c);
      wCell.value = dt.isHoliday ? `${dt.weekdayShort}*` : dt.weekdayShort;
      wCell.font = {
        name: 'Calibri',
        size: 9,
        bold: dt.isWeekend || dt.isHoliday,
        color: {
          argb: dt.isHoliday
            ? 'FF1E40AF'
            : dt.isWeekend
            ? 'FF991B1B'
            : `FF${palette.weekdayText}`,
        },
      };
      wCell.alignment = { horizontal: 'center', vertical: 'middle' };
      wCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: dt.isHoliday ? 'FFDBEAFE' : `FF${palette.weekdayBg}`,
        },
      };
      wCell.border = thinBorder;
    });

    ws.getRow(rDate).height = 22;
    ws.getRow(rWeekday).height = 18;

    // Candidate rows (2 rows per candidate: In + Out)
    let curRow = rDataStart;
    block.rows.forEach((candRow, rowIndex) => {
      const inRowIdx = curRow;
      const outRowIdx = curRow + 1;
      const isZebra = rowIndex % 2 === 1;

      const slCell = ws!.getCell(inRowIdx, 1);
      slCell.value = String(candRow.slNo);
      const nameCell = ws!.getCell(inRowIdx, 2);
      nameCell.value = candRow.candidate.name;

      // Merge Sl No and Candidate Name vertically across In & Out rows for clean readability
      ws!.mergeCells(inRowIdx, 1, outRowIdx, 1);
      ws!.mergeCells(inRowIdx, 2, outRowIdx, 2);

      [
        { r: inRowIdx, c: 1, align: 'center' as const, bold: true },
        { r: inRowIdx, c: 2, align: 'left' as const, bold: true },
      ].forEach(({ r, c, align, bold }) => {
        const cell = ws!.getCell(r, c);
        cell.font = { name: 'Calibri', size: 10, bold, color: { argb: 'FF0F172A' } };
        cell.alignment = { horizontal: align, vertical: 'middle', wrapText: true };
        if (isZebra) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${palette.zebraBg}` } };
        }
        ws!.getCell(inRowIdx, c).border = thinBorder;
        ws!.getCell(outRowIdx, c).border = thinBorder;
      });

      const inTypeCell = ws!.getCell(inRowIdx, 3);
      inTypeCell.value = 'In';
      inTypeCell.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF0369A1' } };
      inTypeCell.alignment = { horizontal: 'center', vertical: 'middle' };
      inTypeCell.border = thinBorder;

      const outTypeCell = ws!.getCell(outRowIdx, 3);
      outTypeCell.value = 'Out';
      outTypeCell.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF475569' } };
      outTypeCell.alignment = { horizontal: 'center', vertical: 'middle' };
      outTypeCell.border = thinBorder;

      if (isZebra) {
        inTypeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${palette.zebraBg}` } };
        outTypeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${palette.zebraBg}` } };
      }

      block.dates.forEach((dt, dIdx) => {
        const c = 4 + dIdx;
        const day = candRow.days[dt.dateStr];
        const statusStyle = getCellStatusColors(day.status, config.colorCodeCells);

        const inC = ws!.getCell(inRowIdx, c);
        const outC = ws!.getCell(outRowIdx, c);
        inC.value = day.inTime;
        outC.value = day.outTime;

        [inC, outC].forEach((cell) => {
          cell.font = {
            name: 'Calibri',
            size: 9,
            bold: Boolean(statusStyle.bold),
            color: { argb: `FF${statusStyle.fg || '0F172A'}` },
          };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = thinBorder;
          if (statusStyle.bg) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: `FF${statusStyle.bg}` },
            };
          } else if (isZebra) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: `FF${palette.zebraBg}` },
            };
          }
        });
      });

      ws!.getRow(inRowIdx).height = 17;
      ws!.getRow(outRowIdx).height = 17;
      curRow += 2;
    });

    // Optional Signature Footer block
    if (config.includeSignatureBlock) {
      const sigRow = curRow + 2;
      ws.getCell(sigRow, 2).value = `Trainer Signature: ${config.trainerName || '____________________'}`;
      ws.getCell(sigRow, 2).font = { name: 'Calibri', size: 10, bold: true };

      const midCol = Math.max(5, Math.floor(totalCols / 2));
      ws.getCell(sigRow, midCol).value = `Center Manager: ${config.centerManager || '____________________'}`;
      ws.getCell(sigRow, midCol).font = { name: 'Calibri', size: 10, bold: true };

      const rightCol = Math.max(midCol + 4, totalCols - 4);
      ws.getCell(sigRow, rightCol).value = 'Verified By / Seal: ____________________';
      ws.getCell(sigRow, rightCol).font = { name: 'Calibri', size: 10, bold: true };
      curRow = sigRow + 2;
    } else {
      curRow += 2;
    }

    // Freeze panes on first block of a sheet
    if (config.freezePanes && startRow === 1) {
      ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 8 }];
    }

    // Column widths
    ws.getColumn(1).width = 7;
    ws.getColumn(2).width = 26;
    ws.getColumn(3).width = 6.5;
    for (let i = 0; i < block.dates.length; i++) {
      ws.getColumn(4 + i).width = config.timeFormat === '12h_ampm' ? 11.5 : 10.5;
    }

    sheetRowCursor.set(block.sheetName, curRow + 1);
  }

  // Optional Executive Summary Dashboard Sheet
  if (config.includeSummarySheet) {
    const sumWs = wb.addWorksheet('Executive_Summary');
    sumWs.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

    sumWs.mergeCells('A1:I1');
    const titleCell = sumWs.getCell('A1');
    titleCell.value = config.centerName.trim()
      ? `${config.centerName.trim()} — Consolidated Attendance Summary`
      : 'Consolidated Attendance Summary';
    titleCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: `FF${palette.bannerText}` } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${palette.bannerBg}` } };
    sumWs.getRow(1).height = 28;

    sumWs.mergeCells('A2:I2');
    const metaCell = sumWs.getCell('A2');
    const summaryMeta = [
      config.batchNo.trim() ? `Batch: ${config.batchNo.trim()}` : '',
      config.district.trim() ? `District: ${config.district.trim()}` : '',
      config.qualificationPack.trim() ? `Qualification: ${config.qualificationPack.trim()}` : '',
      `Period: ${formatDisplayDate(config.startDate)} to ${formatDisplayDate(config.endDate)}`,
    ].filter(Boolean);
    metaCell.value = summaryMeta.join(' | ');
    metaCell.font = { name: 'Calibri', size: 10, color: { argb: `FF${palette.subBannerText}` } };
    metaCell.alignment = { horizontal: 'center', vertical: 'middle' };
    metaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${palette.subBannerBg}` } };
    sumWs.getRow(2).height = 20;

    // Aggregate across all blocks per candidate
    const aggMap = new Map<
      string,
      {
        name: string;
        rollNo: string;
        working: number;
        present: number;
        absent: number;
        late: number;
        half: number;
        totalHrs: number;
      }
    >();

    for (const blk of blocks) {
      for (const r of blk.rows) {
        const existing = aggMap.get(r.candidate.id) || {
          name: r.candidate.name,
          rollNo: r.candidate.rollNo || '-',
          working: 0,
          present: 0,
          absent: 0,
          late: 0,
          half: 0,
          totalHrs: 0,
        };
        existing.working += r.workingDays;
        existing.present += r.presentDays;
        existing.absent += r.absentDays;
        existing.late += r.lateDays;
        existing.half += r.halfDays;
        existing.totalHrs += r.totalHours;
        aggMap.set(r.candidate.id, existing);
      }
    }

    const headers = [
      'Sl No',
      'Roll / ID',
      'Candidate Name',
      'Working Days',
      'Present Days',
      'Absent Days',
      'Late Punches',
      'Attendance %',
      `Compliance (>=${config.complianceThresholdPct}%)`,
    ];

    headers.forEach((h, idx) => {
      const c = sumWs.getCell(4, idx + 1);
      c.value = h;
      c.font = { name: 'Calibri', size: 10, bold: true, color: { argb: `FF${palette.tableHeaderText}` } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${palette.tableHeaderBg}` } };
      c.border = thinBorder;
    });
    sumWs.getRow(4).height = 22;

    let idx = 1;
    for (const item of aggMap.values()) {
      const r = 4 + idx;
      const pct = item.working > 0 ? Number(((item.present / item.working) * 100).toFixed(1)) : 100;
      const compliant = pct >= config.complianceThresholdPct;
      const vals = [
        idx,
        item.rollNo,
        item.name,
        item.working,
        item.present,
        item.absent,
        item.late,
        `${pct.toFixed(1)}%`,
        compliant ? 'COMPLIANT' : 'SHORTFALL',
      ];
      vals.forEach((v, cIdx) => {
        const cell = sumWs.getCell(r, cIdx + 1);
        cell.value = v;
        cell.font = {
          name: 'Calibri',
          size: 10,
          bold: cIdx === 2 || cIdx === 8,
          color: {
            argb: cIdx === 8 ? (compliant ? 'FF166534' : 'FF991B1B') : 'FF0F172A',
          },
        };
        cell.alignment = { horizontal: cIdx === 2 ? 'left' : 'center', vertical: 'middle' };
        cell.border = thinBorder;
        if (cIdx === 8 && config.colorCodeCells) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: compliant ? 'FFDCFCE7' : 'FFFEE2E2' },
          };
        }
      });
      sumWs.getRow(r).height = 19;
      idx++;
    }

    const colWidths = [8, 15, 28, 14, 14, 14, 14, 15, 20];
    colWidths.forEach((w, i) => {
      sumWs.getColumn(i + 1).width = w;
    });
  }

  const buffer = await wb.xlsx.writeBuffer();
  const filename = buildOutputFilename(config, 'xlsx');
  triggerBrowserDownload(
    new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    filename
  );
  return filename;
}

export function exportAttendanceCSV(blocks: AttendanceBlock[], config: AttendanceConfig): string {
  const lines: string[] = [];
  const esc = (v: string | number) => {
    const s = String(v);
    if (s.includes(',') || s.includes('"') || s.includes('\n')) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  };

  for (const blk of blocks) {
    lines.push(esc(`BIOMETRIC ATTENDANCE SHEET${config.centerName.trim() ? ` — ${config.centerName.trim()}` : ''} (${blk.monthName})`));
    const header = [
      'Sl No',
      'Candidate Name',
      'Type',
      ...blk.dates.map((d) => `${d.displayDate} (${d.weekdayShort})`),
    ];
    lines.push(header.map(esc).join(','));

    for (const row of blk.rows) {
      const inCols = [
        row.slNo,
        row.candidate.name,
        'In',
        ...blk.dates.map((d) => row.days[d.dateStr].inTime),
      ];
      const outCols = [
        '',
        '',
        'Out',
        ...blk.dates.map((d) => row.days[d.dateStr].outTime),
      ];
      lines.push(inCols.map(esc).join(','));
      lines.push(outCols.map(esc).join(','));
    }
    lines.push('');
  }

  const filename = buildOutputFilename(config, 'csv');
  triggerBrowserDownload(new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' }), filename);
  return filename;
}

export async function downloadSampleRosterExcel(candidates: Candidate[]) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Candidates');
  ws.addRow(['Sl No', 'Roll No', 'Candidate Name']);
  ws.getRow(1).font = { bold: true };
  candidates.forEach((c, idx) => {
    ws.addRow([idx + 1, c.rollNo || `BT-${String(idx + 1).padStart(3, '0')}`, c.name]);
  });
  ws.getColumn(1).width = 8;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 28;
  const buf = await wb.xlsx.writeBuffer();
  triggerBrowserDownload(
    new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    'Sample_Candidate_Roster.xlsx'
  );
}
