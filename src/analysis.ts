export type Annotation = {
  id: number;
  completed_by?: number | null;
  updated_by?: number | null;
  result?: unknown[];
  was_cancelled?: boolean;
  created_at?: string;
  updated_at?: string;
};
export type Task = {
  id: number;
  project: number;
  annotations?: Annotation[];
  drafts?: unknown[];
};
export type InspectedTask = { taskId: number; candidates: Annotation[] };
export function inspectExport(input: unknown, expectedProjectId: number) {
  if (!Array.isArray(input))
    throw new Error('JSON 최상위 구조는 Task 배열이어야 합니다.');
  const ids = new Set<number>();
  const tasks: Task[] = input.map((item: unknown) => {
    if (!item || typeof item !== 'object')
      throw new Error('올바르지 않은 Task가 있습니다.');
    const task = item as Task;
    if (
      !Number.isInteger(task.id) ||
      !Number.isInteger(task.project) ||
      !Array.isArray(task.annotations) ||
      !Array.isArray(task.drafts)
    )
      throw new Error(
        `Task 필수 필드가 올바르지 않습니다. (ID: ${String(task.id)})`,
      );
    if (task.project !== expectedProjectId)
      throw new Error(
        `다른 프로젝트의 JSON입니다. Task ${task.id}, Project ${task.project}`,
      );
    if (ids.has(task.id)) throw new Error(`중복 Task ID: ${task.id}`);
    ids.add(task.id);
    const annotationIds = new Set<number>();

    for (const a of task.annotations) {
      if (!Number.isInteger(a.id) || !Array.isArray(a.result)) {
        throw new Error(
          `Task ${task.id}의 Annotation 형식이 올바르지 않습니다.`,
        );
      }

      if (annotationIds.has(a.id)) {
        throw new Error(
          `Task ${task.id}에 중복된 Annotation ID가 있습니다: ${a.id}`,
        );
      }

      annotationIds.add(a.id);
    }
    return task;
  });
  const draftTaskIds = tasks
    .filter((t) => t.drafts!.length > 0)
    .map((t) => t.id)
    .sort((a, b) => a - b);
  const multiple: InspectedTask[] = tasks
    .map((t) => ({
      taskId: t.id,
      candidates: t.annotations!.filter((a) => !a.was_cancelled),
    }))
    .filter((t) => t.candidates.length > 1);
  return { tasks, draftTaskIds, multiple };
}
