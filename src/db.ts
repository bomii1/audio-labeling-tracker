import Dexie, { type EntityTable } from 'dexie';
export type Project = { id?: number; labelStudioProjectId: number; name: string; gameName: string; description: string; createdAt: string };
export type Snapshot = { id?: number; projectId: number; recordDate: string; uploadedAt: string; fileName: string; totalTasks: number; totalRegions: number; zeroRegionTasks: number };
export type TaskRegion = { id?: number; snapshotId: number; taskId: number; annotationId: number | null; completedBy: number | null; regionCount: number };
export type AnnotationChoice = { id?: number; projectId: number; taskId: number; annotationId: number; knownAnnotationIds: number[] };
export type ProjectMember = { id?: number; projectId: number; labelStudioUserId: number; name: string; role: 'worker' | 'pl' | 'both' };
export const db = new Dexie('audioLabelingTracker') as Dexie & {
  projects: EntityTable<Project,'id'>;
  snapshots: EntityTable<Snapshot,'id'>;
  taskRegions: EntityTable<TaskRegion,'id'>;
  annotationChoices: EntityTable<AnnotationChoice,'id'>;
  projectMembers: EntityTable<ProjectMember,'id'>;
};
// v1 데이터를 보존하면서 작업자 관리 테이블만 추가합니다.
db.version(1).stores({ projects:'++id,&labelStudioProjectId,name', snapshots:'++id,projectId,recordDate,uploadedAt', taskRegions:'++id,snapshotId,[snapshotId+taskId]', annotationChoices:'++id,&[projectId+taskId]' });
db.version(2).stores({ projects:'++id,&labelStudioProjectId,name', snapshots:'++id,projectId,recordDate,uploadedAt', taskRegions:'++id,snapshotId,[snapshotId+taskId]', annotationChoices:'++id,&[projectId+taskId]', projectMembers:'++id,&[projectId+labelStudioUserId],projectId,labelStudioUserId' });
