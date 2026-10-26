import ExcelJS from 'exceljs';
import type { Project, ProjectMember, Snapshot, TaskRegion } from './db';

type Row = Pick<
  TaskRegion,
  'taskId' | 'annotationId' | 'completedBy' | 'regionCount'
> & {
  updatedBy?: number | null;
  updatedAt?: string | null;
};
type ExportInput = {
  project: Project;
  snapshot: Pick<
    Snapshot,
    | 'recordDate'
    | 'fileName'
    | 'totalTasks'
    | 'totalRegions'
    | 'zeroRegionTasks'
  >;
  rows: Row[];
  previousRows?: Row[];
  previousTotal?: number;
  members: ProjectMember[];
  history?: Snapshot[];
};

const C = {
  blue: '28558B',
  deep: '193F6A',
  lightBlue: 'CADDEC',
  labelGray: 'F2F5F9',
  lightGray: 'F4F7FB',
  white: 'FFFFFF',
  yellow: 'FFF0C7',
  pink: 'F8DFE3',
  grid: 'E7ECF1',
  ink: '23354C',
  greenBg: 'E2F2E9',
  green: '008A4B',
  redBg: 'FBE6E7',
  red: 'C62842',
  muted: '63758B',
  blueText: '244C7D',
} as const;
const fontName = '맑은 고딕';
const argb = (hex: string) => `FF${hex}`;
const fill = (hex: string): ExcelJS.Fill => ({
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: argb(hex) },
});
const safeName = (s: string) => s.replace(/[\\/:*?"<>|]/g, '_').slice(0, 65);
const memberLabel = (id: number | null, members: ProjectMember[]) => {
  if (id === null) return '미상';
  const member = members.find((m) => m.labelStudioUserId === id);
  return member?.name || `작업자 #${id}`;
};
function kst(value?: string | null) {
  if (!value) return '';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Seoul',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  }).formatToParts(parsed);
  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? '';
  return `${get('year')}-${get('month')}-${get('day')} ${get('hour')}:${get('minute')}:${get('second')}`;
}
function styleHeader(
  row: ExcelJS.Row,
  columns: number,
  color: string = C.blue,
  light = false,
) {
  row.height = 34;
  for (let i = 1; i <= columns; i++) {
    const cell = row.getCell(i);
    cell.fill = fill(color);
    cell.font = {
      name: fontName,
      size: 12,
      bold: true,
      color: { argb: argb(light ? C.deep : C.white) },
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    cell.border = { bottom: { style: 'thin', color: { argb: argb(C.grid) } } };
  }
}
function styleBody(
  row: ExcelJS.Row,
  columns: number,
  background: string = C.white,
  horizontalBordersOnly = false,
) {
  row.height = 30;
  for (let i = 1; i <= columns; i++) {
    const cell = row.getCell(i);
    cell.fill = fill(background);
    cell.font = { name: fontName, size: 12, color: { argb: argb(C.ink) } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'hair', color: { argb: argb(C.grid) } },
      bottom: { style: 'hair', color: { argb: argb(C.grid) } },
      ...(horizontalBordersOnly
        ? {}
        : {
            left: { style: 'hair' as const, color: { argb: argb(C.grid) } },
            right: { style: 'hair' as const, color: { argb: argb(C.grid) } },
          }),
    };
  }
}
function sectionTitle(ws: ExcelJS.Worksheet, text: string, columns: number) {
  const index = ws.rowCount + 1;
  ws.mergeCells(index, 1, index, columns);
  const row = ws.getRow(index);
  row.height = 37;
  const cell = row.getCell(1);
  cell.value = text;
  cell.fill = fill(C.deep);
  cell.font = {
    name: fontName,
    size: 15,
    bold: true,
    color: { argb: argb(C.white) },
  };
  cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
}
function spacer(ws: ExcelJS.Worksheet, height = 18) {
  const row = ws.addRow([]);
  row.height = height;
}
async function saveBook(workbook: ExcelJS.Workbook, name: string) {
  const bytes = await workbook.xlsx.writeBuffer();
  const blob = new Blob([bytes as BlobPart], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export async function downloadDraftExcel(
  project: Project,
  date: string,
  taskIds: number[],
) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Draft 목록', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.columns = [{ width: 22 }];
  styleHeader(ws.addRow(['Task ID']), 1);
  for (const id of [...new Set(taskIds)].sort((a, b) => a - b))
    styleBody(ws.addRow([id]), 1, C.yellow);
  ws.autoFilter = { from: 'A1', to: `A${Math.max(1, ws.rowCount)}` };
  await saveBook(wb, `${safeName(project.name)}_${date}_Draft목록.xlsx`);
}
export async function downloadDetailExcel({
  project,
  snapshot,
  rows,
  previousTotal,
  members,
  history,
}: ExportInput) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Audio Labeling Tracker';
  const summary = wb.addWorksheet('요약', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  summary.columns = [
    { width: 28 },
    { width: 42 },
    { width: 23 },
    { width: 23 },
    { width: 25 },
  ];
  summary.mergeCells('A1:E1');
  const title = summary.getCell('A1');
  title.value = `${project.name} · 작업 현황`;
  title.fill = fill(C.blue);
  title.font = {
    name: fontName,
    bold: true,
    size: 20,
    color: { argb: argb(C.white) },
  };
  title.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  summary.getRow(1).height = 48;

  spacer(summary, 12);
  sectionTitle(summary, '프로젝트 정보', 5);
  const info: (string | number)[][] = [
    ['프로젝트명', project.name],
    ['프로젝트 ID', project.labelStudioProjectId],
    ['기록 날짜', snapshot.recordDate],
    ['파일명', snapshot.fileName],
    ['전체 Task 수', snapshot.totalTasks],
    ['누적 Region 수', snapshot.totalRegions],
    ['Region 0개 Task 수', snapshot.zeroRegionTasks],
    [
      '직전 기록 대비 순증감',
      previousTotal === undefined
        ? '기준값'
        : snapshot.totalRegions - previousTotal,
    ],
  ];
  for (const entry of info) {
    const i = info.indexOf(entry);
    const row = summary.addRow(entry);
    styleBody(row, 5, C.white, true);
    row.getCell(1).fill = fill(C.labelGray);
    row.getCell(1).font = {
      name: fontName,
      size: 12,
      bold: true,
      color: { argb: argb(C.blueText) },
    };
    row.getCell(2).alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
      indent: 1,
    };
    if (i === 3) row.height = 36;
    if (i >= 4 && i <= 6)
      row.getCell(2).font = {
        name: fontName,
        size: 13,
        bold: true,
        color: { argb: argb(C.ink) },
      };
    if (i === 7 && typeof entry[1] === 'number')
      styleDelta(row.getCell(2), entry[1]);
  }

  spacer(summary);
  sectionTitle(summary, '작업자별 누적 실적', 5);
  styleHeader(
    summary.addRow(['작업자 ID', '작업자 이름', '누적 Region', 'Task 수', '']),
    5,
    C.lightBlue,
    true,
  );
  const totals = new Map<
    string,
    { id: number | null; regions: number; tasks: number }
  >();
  for (const r of rows) {
    const key = r.completedBy == null ? 'unknown' : String(r.completedBy);
    const item = totals.get(key) ?? {
      id: r.completedBy ?? null,
      regions: 0,
      tasks: 0,
    };
    item.regions += r.regionCount;
    item.tasks++;
    totals.set(key, item);
  }
  for (const item of [...totals.values()].sort(
    (a, b) => b.regions - a.regions,
  )) {
    const row = summary.addRow([
      item.id ?? '미상',
      memberLabel(item.id, members),
      item.regions,
      item.tasks,
    ]);
    styleBody(row, 5, C.white, true);
    row.getCell(2).font = {
      name: fontName,
      size: 12,
      bold: true,
      color: { argb: argb(C.blueText) },
    };
    row.getCell(3).font = {
      name: fontName,
      size: 13,
      bold: true,
      color: { argb: argb(C.ink) },
    };
  }

  spacer(summary, 22);
  sectionTitle(summary, '날짜별 작업 현황', 5);
  const dailyHeader = summary.addRow([
    '기록 날짜',
    '전체 Task',
    '누적 Region',
    '순증감',
    'Region 0개 Task',
  ]);
  styleHeader(dailyHeader, 5, C.lightBlue, true);
  const daily = new Map<string, Snapshot>();
  for (const s of [...(history ?? [])].sort(
    (a, b) =>
      a.recordDate.localeCompare(b.recordDate) ||
      a.uploadedAt.localeCompare(b.uploadedAt) ||
      (a.id ?? 0) - (b.id ?? 0),
  ))
    daily.set(s.recordDate, s);
  const timeline = [...daily.values()].filter(
    (s) => s.recordDate <= snapshot.recordDate,
  );
  if (!timeline.some((s) => s.recordDate === snapshot.recordDate)) {
    timeline.push({
      ...snapshot,
      projectId: project.id!,
      uploadedAt: '',
      id: undefined,
    });
    timeline.sort((a, b) => a.recordDate.localeCompare(b.recordDate));
  }
  timeline.forEach((s, i) => {
    const prev = timeline[i - 1];
    const delta = prev ? s.totalRegions - prev.totalRegions : '기준값';
    const row = summary.addRow([
      s.recordDate,
      s.totalTasks,
      s.totalRegions,
      delta,
      s.zeroRegionTasks,
    ]);
    styleBody(row, 5, i % 2 === 0 ? C.white : C.lightGray, true);
    row.height = 34;
    row.getCell(3).font = {
      name: fontName,
      size: 13,
      bold: true,
      color: { argb: argb(C.ink) },
    };
    styleDelta(row.getCell(4), delta);
  });
  spacer(summary, 16);
  const note = summary.addRow([
    '참고',
    '순증감에는 Region 수정 및 Annotation 집계 대상 변경이 반영될 수 있습니다.',
  ]);
  styleBody(note, 5, C.lightGray);
  note.getCell(1).font = {
    name: fontName,
    size: 11,
    bold: true,
    color: { argb: argb(C.muted) },
  };
  note.getCell(2).alignment = {
    wrapText: true,
    vertical: 'middle',
    horizontal: 'left',
    indent: 1,
  };
  note.height = 42;

  const detail = wb.addWorksheet('Task 상세', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  detail.columns = [
    { width: 18 },
    { width: 18 },
    { width: 27 },
    { width: 18 },
    { width: 23 },
    { width: 58 },
  ];
  styleHeader(
    detail.addRow([
      '작업자 ID',
      'Task ID',
      '작업 날짜(KST)',
      'Region 수',
      'Region 0개 여부',
      '비고',
    ]),
    6,
  );
  for (const r of [...rows].sort((a, b) => a.taskId - b.taskId)) {
    const row = detail.addRow([
      r.completedBy ?? '',
      r.taskId,
      kst(r.updatedAt),
      r.regionCount,
      r.regionCount === 0 ? 'Y' : 'N',
      '',
    ]);
    styleBody(row, 6, r.regionCount === 0 ? C.pink : C.yellow);
    // Task 상세: 배경색과 어울리는 자연스러운 셀 구분선
    const borderColor = r.regionCount === 0 ? 'E8C9CF' : 'E8DDB8';

    for (let col = 1; col <= 6; col++) {
      row.getCell(col).border = {
        top: {
          style: 'hair',
          color: { argb: `FF${borderColor}` },
        },
        bottom: {
          style: 'hair',
          color: { argb: `FF${borderColor}` },
        },
        left: {
          style: 'hair',
          color: { argb: `FF${borderColor}` },
        },
        right: {
          style: 'hair',
          color: { argb: `FF${borderColor}` },
        },
      };
    }
    row.getCell(6).alignment = {
      horizontal: 'left',
      vertical: 'middle',
      wrapText: true,
    };
    if (r.regionCount === 0) {
      row.getCell(5).font = {
        name: fontName,
        size: 12,
        bold: true,
        color: { argb: argb(C.red) },
      };
    }
  }
  detail.autoFilter = { from: 'A1', to: `F${Math.max(1, detail.rowCount)}` };
  await saveBook(
    wb,
    `${safeName(project.name)}_${snapshot.recordDate}_작업현황.xlsx`,
  );
}
function styleDelta(cell: ExcelJS.Cell, value: string | number) {
  const positive = typeof value === 'number' && value > 0;
  const negative = typeof value === 'number' && value < 0;
  const foreground = positive ? C.green : negative ? C.red : C.muted;
  cell.font = {
    name: fontName,
    size: 13,
    bold: true,
    color: { argb: argb(foreground) },
  };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
  if (typeof value === 'number') cell.numFmt = '+0;-0;0';
}
