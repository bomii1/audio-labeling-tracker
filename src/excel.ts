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
const BLUE = '24548C';
const YELLOW = 'FFF2CB';
const PINK = 'F6D9DD';
const GRID = 'D9D3C7';
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
function header(row: ExcelJS.Row) {
  row.height = 25;
  row.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${BLUE}` },
    };
    cell.font = {
      name: '맑은 고딕',
      size: 11,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = { bottom: { style: 'thin', color: { argb: `FF${GRID}` } } };
  });
}
function body(row: ExcelJS.Row, background = 'FFFFFF') {
  row.height = 22;
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${background}` },
    };
    cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF17243A' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'hair', color: { argb: `FF${GRID}` } },
      bottom: { style: 'hair', color: { argb: `FF${GRID}` } },
      left: { style: 'hair', color: { argb: `FF${GRID}` } },
      right: { style: 'hair', color: { argb: `FF${GRID}` } },
    };
  });
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
  ws.columns = [{ width: 19 }];
  header(ws.addRow(['Task ID']));
  for (const id of [...new Set(taskIds)].sort((a, b) => a - b))
    body(ws.addRow([id]), YELLOW);
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
  const summary = wb.addWorksheet('요약');
  summary.columns = [
    { width: 28 },
    { width: 38 },
    { width: 22 },
    { width: 22 },
    { width: 22 },
  ];
  summary.mergeCells('A1:E1');
  const title = summary.getCell('A1');
  title.value = `${project.name} · 작업 현황`;
  title.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: `FF${BLUE}` },
  };
  title.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  title.alignment = { vertical: 'middle' };
  summary.getRow(1).height = 35;
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
    const row = summary.addRow(entry);
    body(row, row.number % 2 ? 'F3F6FB' : 'FFFFFF');
    row.getCell(1).font = { bold: true, color: { argb: `FF${BLUE}` } };
    row.getCell(2).alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
  }
  summary.addRow([]);
  header(
    summary.addRow(['작업자 ID', '작업자 이름', '누적 Region', 'Task 수']),
  );
  const totals = new Map<
    string,
    { id: number | null; regions: number; tasks: number }
  >();
  for (const r of rows) {
    const key = r.completedBy === null ? 'unknown' : String(r.completedBy);
    const item = totals.get(key) ?? { id: r.completedBy, regions: 0, tasks: 0 };
    item.regions += r.regionCount;
    item.tasks++;
    totals.set(key, item);
  }
  for (const item of [...totals.values()].sort(
    (a, b) => b.regions - a.regions,
  )) {
    body(
      summary.addRow([
        item.id ?? '미상',
        memberLabel(item.id, members),
        item.regions,
        item.tasks,
      ]),
      YELLOW,
    );
  }
  summary.addRow([]);
  header(
    summary.addRow([
      '기록 날짜',
      '전체 Task',
      '누적 Region',
      '순증감',
      'Region 0개 Task',
    ]),
  );
  const daily = new Map<string, Snapshot>();
  for (const s of [...(history ?? [])].sort(
    (a, b) =>
      a.recordDate.localeCompare(b.recordDate) ||
      a.uploadedAt.localeCompare(b.uploadedAt),
  ))
    daily.set(s.recordDate, s);
  // 기록 조회에서 해당 날짜까지의 저장 이력, 업로드 미리보기에서는 임시 분석 결과를 사용한다.
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
    body(
      summary.addRow([
        s.recordDate,
        s.totalTasks,
        s.totalRegions,
        prev ? s.totalRegions - prev.totalRegions : '기준값',
        s.zeroRegionTasks,
      ]),
      i % 2 ? 'F3F6FB' : 'FFFFFF',
    );
  });
  summary.addRow([]);
  const note = summary.addRow([
    '참고',
    '순증감에는 Region 수정 및 Annotation 집계 대상 변경이 반영될 수 있습니다.',
  ]);
  note.getCell(2).alignment = { wrapText: true, vertical: 'middle' };
  summary.getRow(note.number).height = 30;

  const detail = wb.addWorksheet('Task 상세', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  detail.columns = [
    { width: 17 },
    { width: 17 },
    { width: 25 },
    { width: 16 },
    { width: 20 },
    { width: 56 },
  ];
  header(
    detail.addRow([
      '작업자 ID',
      'Task ID',
      '작업 날짜(KST)',
      'Region 수',
      'Region 0개 여부',
      '비고',
    ]),
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
    body(row, r.regionCount === 0 ? PINK : YELLOW);
    row.getCell(6).alignment = {
      horizontal: 'left',
      vertical: 'middle',
      wrapText: true,
    };
  }
  detail.autoFilter = { from: 'A1', to: `F${Math.max(1, detail.rowCount)}` };
  await saveBook(
    wb,
    `${safeName(project.name)}_${snapshot.recordDate}_작업현황.xlsx`,
  );
}
