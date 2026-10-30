import Dexie, { type EntityTable } from 'dexie';

export type Project = {
  id?: number;
  labelStudioProjectId: number;
  name: string;
  gameName: string;
  description: string;
  createdAt: string;
};
export type Snapshot = {
  id?: number;
  projectId: number;
  recordDate: string;
  uploadedAt: string;
  fileName: string;
  totalTasks: number;
  totalRegions: number;
  zeroRegionTasks: number;
};
export type TaskRegion = {
  id?: number;
  snapshotId: number;
  taskId: number;
  annotationId: number | null;
  completedBy: number | null;
  updatedBy?: number | null;
  updatedAt?: string | null;
  regionCount: number;
};
export type AnnotationChoice = {
  id?: number;
  projectId: number;
  taskId: number;
  annotationId: number;
  knownAnnotationIds: number[];
};
export type ProjectMember = {
  id?: number;
  projectId: number;
  labelStudioUserId: number;
  name: string;
  role: 'worker' | 'pl' | 'both';
};
export const db = new Dexie('audioLabelingTracker') as Dexie & {
  projects: EntityTable<Project, 'id'>;
  snapshots: EntityTable<Snapshot, 'id'>;
  taskRegions: EntityTable<TaskRegion, 'id'>;
  annotationChoices: EntityTable<AnnotationChoice, 'id'>;
  projectMembers: EntityTable<ProjectMember, 'id'>;
};

db.version(1).stores({
  projects: '++id,&labelStudioProjectId,name',
  snapshots: '++id,projectId,recordDate,uploadedAt',
  taskRegions: '++id,snapshotId,[snapshotId+taskId]',
  annotationChoices: '++id,&[projectId+taskId]',
});
db.version(2).stores({
  projects: '++id,&labelStudioProjectId,name',
  snapshots: '++id,projectId,recordDate,uploadedAt',
  taskRegions: '++id,snapshotId,[snapshotId+taskId]',
  annotationChoices: '++id,&[projectId+taskId]',
  projectMembers:
    '++id,&[projectId+labelStudioUserId],projectId,labelStudioUserId',
});

// 기존 같은 날짜의 중복 기록은 최신 업로드 하나만 남긴다.
// 고유 인덱스를 만들기 전에 별도 버전에서 정리해야 마이그레이션이 안전하다.
db.version(3)
  .stores({
    projects: '++id,&labelStudioProjectId,name',
    snapshots: '++id,projectId,recordDate,uploadedAt,[projectId+recordDate]',
    taskRegions: '++id,snapshotId,[snapshotId+taskId]',
    annotationChoices: '++id,&[projectId+taskId]',
    projectMembers:
      '++id,&[projectId+labelStudioUserId],projectId,labelStudioUserId',
  })
  .upgrade(async (tx) => {
    const snapshots = (await tx.table('snapshots').toArray()) as Snapshot[];
    const latestByDate = new Map<string, Snapshot>();
    const sorted = snapshots.sort(
      (a, b) =>
        a.uploadedAt.localeCompare(b.uploadedAt) || (a.id ?? 0) - (b.id ?? 0),
    );
    for (const snapshot of sorted) {
      latestByDate.set(
        JSON.stringify([snapshot.projectId, snapshot.recordDate]),
        snapshot,
      );
    }
    const keepIds = new Set([...latestByDate.values()].map((s) => s.id));
    for (const snapshot of snapshots) {
      if (snapshot.id === undefined || keepIds.has(snapshot.id)) continue;
      await tx
        .table('taskRegions')
        .where('snapshotId')
        .equals(snapshot.id)
        .delete();
      await tx.table('snapshots').delete(snapshot.id);
    }
  });

// 프로젝트별 날짜 하나만 허용하는 데이터베이스 제약.
db.version(4).stores({
  projects: '++id,&labelStudioProjectId,name',
  snapshots: '++id,projectId,recordDate,uploadedAt,&[projectId+recordDate]',
  taskRegions: '++id,snapshotId,[snapshotId+taskId]',
  annotationChoices: '++id,&[projectId+taskId]',
  projectMembers:
    '++id,&[projectId+labelStudioUserId],projectId,labelStudioUserId',
});
