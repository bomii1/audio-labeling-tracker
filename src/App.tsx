import { useEffect, useState } from 'react';
import { Link, NavLink, Route, Routes, useParams } from 'react-router-dom';
import { db, type Project, type Snapshot, type ProjectMember } from './db';
import { inspectExport, type Task } from './analysis';
import { downloadDraftExcel, downloadDetailExcel } from './excel';
import {
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  FileJson,
  Folder,
  History,
  LayoutDashboard,
  Plus,
  UploadCloud,
  X,
  Pencil,
  Trash2,
  Users,
  AlertTriangle,
} from 'lucide-react';
function useProjects() {
  const [items, setItems] = useState<Project[]>([]);
  const refresh = () => db.projects.toArray().then(setItems);
  useEffect(() => {
    void refresh();
  }, []);
  return { items, refresh };
}
function ProjectHome() {
  const { items, refresh } = useProjects();
  const [editing, setEditing] = useState<Project | null | undefined>(undefined);
  const [deleting, setDeleting] = useState<Project | null>(null);
  const [name, setName] = useState('');
  const [game, setGame] = useState('');
  const [lsId, setLsId] = useState('');
  const [desc, setDesc] = useState('');
  const [error, setError] = useState('');
  const [confirmation, setConfirmation] = useState('');
  const [busy, setBusy] = useState(false);
  const [stats, setStats] = useState<Record<number, Snapshot>>({});
  useEffect(() => {
    void db.snapshots.toArray().then((rows) => {
      const next: Record<number, Snapshot> = {};
      for (const row of rows.sort(
        (a, b) =>
          a.recordDate.localeCompare(b.recordDate) ||
          a.uploadedAt.localeCompare(b.uploadedAt),
      ))
        next[row.projectId] = row;
      setStats(next);
    });
  }, [items]);
  function openEditor(project: Project | null) {
    setEditing(project);
    setName(project?.name ?? '');
    setGame(project?.gameName ?? '');
    setLsId(String(project?.labelStudioProjectId ?? ''));
    setDesc(project?.description ?? '');
    setError('');
  }
  async function submitProject() {
    const id = Number(lsId);
    if (!name.trim() || !Number.isInteger(id) || id <= 0) {
      setError('프로젝트명과 올바른 Label Studio 프로젝트 ID를 입력해 주세요.');
      return;
    }
    setBusy(true);
    setError('');
    try {
      if (editing?.id) {
        if (id !== editing.labelStudioProjectId) {
          setError(
            '기존 기록 보호를 위해 Label Studio 프로젝트 ID는 변경할 수 없습니다.',
          );
          return;
        }
        await db.projects.update(editing.id, {
          name: name.trim(),
          gameName: game.trim(),
          description: desc.trim(),
        });
      } else {
        await db.projects.add({
          name: name.trim(),
          labelStudioProjectId: id,
          gameName: game.trim(),
          description: desc.trim(),
          createdAt: new Date().toISOString(),
        });
      }
      setEditing(undefined);
      refresh();
    } catch (e) {
      setError(
        e instanceof Error && e.name === 'ConstraintError'
          ? '이미 등록된 Label Studio 프로젝트 ID입니다.'
          : e instanceof Error
            ? e.message
            : '저장에 실패했습니다.',
      );
    } finally {
      setBusy(false);
    }
  }
  async function deleteProject() {
    if (!deleting?.id || confirmation !== `delete ${deleting.name}`) return;
    setBusy(true);
    setError('');
    try {
      const id = deleting.id;
      await db.transaction(
        'rw',
        db.projects,
        db.snapshots,
        db.taskRegions,
        db.annotationChoices,
        db.projectMembers,
        async () => {
          const snaps = await db.snapshots
            .where('projectId')
            .equals(id)
            .primaryKeys();
          for (const snapshotId of snaps) {
            if (snapshotId === undefined) {
              continue;
            }
            await db.taskRegions
              .where('snapshotId')
              .equals(snapshotId)
              .delete();
          }
          await db.snapshots.where('projectId').equals(id).delete();
          await db.annotationChoices.where('projectId').equals(id).delete();
          await db.projectMembers.where('projectId').equals(id).delete();
          await db.projects.delete(id);
        },
      );
      setDeleting(null);
      setConfirmation('');
      refresh();
    } catch (e) {
      setError(
        e instanceof Error ? e.message : '프로젝트 삭제에 실패했습니다.',
      );
    } finally {
      setBusy(false);
    }
  }
  return (
    <div className="content">
      <div className="heading">
        <div>
          <span className="eyebrow">WORKSPACE</span>
          <h1>내 프로젝트</h1>
          <p>담당하는 라벨링 프로젝트를 선택해 작업 현황을 확인하세요.</p>
        </div>
        <button className="primary" onClick={() => openEditor(null)}>
          <Plus size={17} /> 프로젝트 추가
        </button>
      </div>
      {items.length === 0 ? (
        <div className="empty">
          <Folder size={40} />
          <h2>아직 등록된 프로젝트가 없어요</h2>
          <p>Label Studio 프로젝트 ID로 첫 프로젝트를 등록해 보세요.</p>
          <button className="primary" onClick={() => openEditor(null)}>
            <Plus size={16} /> 프로젝트 추가
          </button>
        </div>
      ) : (
        <div className="project-grid">
          {items.map((project) => (
            <div className="project-card" key={project.id}>
              <div className="project-card-top">
                <div className="folder">
                  <Folder size={22} />
                </div>
                <div className="project-controls">
                  <button
                    className="icon-button"
                    title="프로젝트 수정"
                    aria-label={`${project.name} 수정`}
                    onClick={() => openEditor(project)}
                  >
                    <Pencil size={17} />
                  </button>
                  <button
                    className="icon-button danger-icon"
                    title="프로젝트 삭제"
                    aria-label={`${project.name} 삭제`}
                    onClick={() => {
                      setDeleting(project);
                      setConfirmation('');
                      setError('');
                    }}
                  >
                    <Trash2 size={17} />
                  </button>
                </div>
              </div>
              <Link
                className="project-card-link"
                to={`/projects/${project.id}`}
              >
                <h2>{project.name}</h2>
                <span className="muted">
                  Project ID {project.labelStudioProjectId}
                  {project.gameName && ` · ${project.gameName}`}
                </span>
                <div className="card-footer">
                  <div>
                    <small>최근 누적 Region</small>
                    <strong>
                      {project.id && stats[project.id]
                        ? stats[project.id].totalRegions.toLocaleString()
                        : '—'}
                    </strong>
                  </div>
                  <span className="muted">
                    {project.id && stats[project.id]
                      ? stats[project.id].recordDate
                      : '기록 없음'}
                  </span>
                  <ArrowRight size={18} />
                </div>
              </Link>
            </div>
          ))}
        </div>
      )}
      {editing !== undefined && (
        <div
          className="overlay"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setEditing(undefined);
          }}
        >
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-label={editing ? '프로젝트 수정' : '프로젝트 추가'}
          >
            <div className="modal-title">
              <h2>{editing ? '프로젝트 수정' : '프로젝트 추가'}</h2>
              <button
                className="icon-button"
                onClick={() => setEditing(undefined)}
              >
                <X size={18} />
              </button>
            </div>
            <label>
              프로젝트명 *
              <input
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="예: RAFT 오디오 라벨링"
              />
            </label>
            <label>
              Label Studio 프로젝트 ID *
              <input
                type="number"
                min="1"
                value={lsId}
                disabled={!!editing}
                onChange={(e) => setLsId(e.target.value)}
                placeholder="예: 54"
              />
            </label>
            <label>
              게임명 (선택)
              <input
                value={game}
                onChange={(e) => setGame(e.target.value)}
                placeholder="예: RAFT"
              />
            </label>
            <label>
              설명 (선택)
              <textarea
                value={desc}
                onChange={(e) => setDesc(e.target.value)}
                placeholder="프로젝트 메모"
              />
            </label>
            {editing && (
              <p className="muted">
                기존 기록 보호를 위해 Label Studio 프로젝트 ID는 수정할 수
                없습니다.
              </p>
            )}
            {error && <p className="error">{error}</p>}
            <div className="actions">
              <button
                className="secondary"
                onClick={() => setEditing(undefined)}
              >
                취소
              </button>
              <button
                className="primary"
                disabled={busy}
                onClick={() => void submitProject()}
              >
                {busy ? '저장 중...' : editing ? '변경 사항 저장' : '추가하기'}
              </button>
            </div>
          </div>
        </div>
      )}
      {deleting && (
        <div
          className="overlay"
          onMouseDown={(e) => {
            if (e.target === e.currentTarget) setDeleting(null);
          }}
        >
          <div
            className="modal"
            role="alertdialog"
            aria-modal="true"
            aria-label="프로젝트 영구 삭제 확인"
          >
            <div className="modal-title">
              <h2>프로젝트 영구 삭제</h2>
              <button className="icon-button" onClick={() => setDeleting(null)}>
                <X size={18} />
              </button>
            </div>
            <p className="delete-warning">
              <strong>{deleting.name}</strong> 프로젝트와 모든 날짜별 분석 기록,
              Task별 집계, Annotation 선택 이력이 영구 삭제됩니다. 되돌릴 수
              없습니다.
            </p>
            <label>
              삭제하려면 아래 문구를 정확히 입력하세요.
              <div className="confirmation-text">delete {deleting.name}</div>
              <input
                autoComplete="off"
                spellCheck={false}
                value={confirmation}
                onChange={(e) => setConfirmation(e.target.value)}
                placeholder={`delete ${deleting.name}`}
              />
            </label>
            {error && <p className="error">{error}</p>}
            <div className="actions">
              <button className="secondary" onClick={() => setDeleting(null)}>
                취소
              </button>
              <button
                className="delete-button"
                disabled={busy || confirmation !== `delete ${deleting.name}`}
                onClick={() => void deleteProject()}
              >
                {busy ? '삭제 중...' : '프로젝트 영구 삭제'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function ProjectShell() {
  const { projectId } = useParams();
  const [project, setProject] = useState<Project>();
  useEffect(() => {
    void db.projects.get(Number(projectId)).then(setProject);
  }, [projectId]);
  if (!project)
    return (
      <div className="content">
        <p>프로젝트를 불러오는 중이거나 존재하지 않습니다.</p>
        <Link to="/">프로젝트 목록으로</Link>
      </div>
    );
  return <ProjectPage project={project} />;
}
function ProjectPage({ project }: { project: Project }) {
  const [tab, setTab] = useState<
    'dashboard' | 'analyze' | 'history' | 'members'
  >('dashboard');
  const [snapshots, setSnapshots] = useState<Snapshot[]>([]);
  const refresh = () =>
    db.snapshots
      .where('projectId')
      .equals(project.id!)
      .toArray()
      .then(setSnapshots);
  useEffect(() => {
    void refresh();
  }, [project.id]);
  const ordered = [...snapshots].sort(
    (a, b) =>
      a.recordDate.localeCompare(b.recordDate) ||
      a.uploadedAt.localeCompare(b.uploadedAt) ||
      (a.id ?? 0) - (b.id ?? 0),
  );
  const perDay = new Map<string, Snapshot>();
  ordered.forEach((s) => perDay.set(s.recordDate, s));
  const days = [...perDay.values()];
  const latest = days.at(-1);
  const previous = days.at(-2);
  return (
    <div className="content">
      <Link to="/" className="back">
        <ArrowLeft size={16} /> 모든 프로젝트
      </Link>
      <div className="heading">
        <div>
          <span className="eyebrow">
            PROJECT {project.labelStudioProjectId}
          </span>
          <h1>{project.name}</h1>
          <p>
            {project.description || project.gameName || '프로젝트 작업 현황'}
          </p>
        </div>
        <button className="primary" onClick={() => setTab('analyze')}>
          <UploadCloud size={17} /> JSON 업로드
        </button>
      </div>
      <div className="tabs">
        <button
          className={tab === 'dashboard' ? 'active' : ''}
          onClick={() => setTab('dashboard')}
        >
          <LayoutDashboard size={17} /> 대시보드
        </button>
        <button
          className={tab === 'analyze' ? 'active' : ''}
          onClick={() => setTab('analyze')}
        >
          <FileJson size={17} /> JSON 분석
        </button>
        <button
          className={tab === 'history' ? 'active' : ''}
          onClick={() => setTab('history')}
        >
          <History size={17} /> 기록 조회
        </button>
        <button
          className={tab === 'members' ? 'active' : ''}
          onClick={() => setTab('members')}
        >
          <Users size={17} /> 작업자 관리
        </button>
      </div>
      {tab === 'dashboard' && (
        <>
          <div className="metrics">
            <Metric
              label="누적 Region"
              value={latest?.totalRegions.toLocaleString() ?? '—'}
            />
            <Metric
              label="최근 순증감"
              value={
                latest
                  ? previous
                    ? `${latest.totalRegions - previous.totalRegions >= 0 ? '+' : ''}${(latest.totalRegions - previous.totalRegions).toLocaleString()}`
                    : '기준값'
                  : '—'
              }
              tone={
                latest && previous
                  ? latest.totalRegions > previous.totalRegions
                    ? 'positive'
                    : latest.totalRegions < previous.totalRegions
                      ? 'negative'
                      : 'neutral'
                  : 'neutral'
              }
            />
            <Metric
              label="전체 Task"
              value={latest?.totalTasks.toLocaleString() ?? '—'}
            />
            <Metric
              label="Region 0개 Task"
              value={latest?.zeroRegionTasks.toLocaleString() ?? '—'}
            />
          </div>
          <section className="panel">
            <h2>최근 기록</h2>
            {latest ? (
              <p>
                {latest.recordDate} 기준 · {latest.fileName}
              </p>
            ) : (
              <div className="empty-small">
                <p>아직 분석 기록이 없어요.</p>
                <button className="primary" onClick={() => setTab('analyze')}>
                  첫 JSON 업로드
                </button>
              </div>
            )}
          </section>
        </>
      )}
      {tab === 'analyze' && (
        <Analyze project={project} refresh={refresh} latest={latest} />
      )}
      {tab === 'members' && <ProjectMembers project={project} />}{' '}
      {tab === 'history' && (
        <HistoryPanel project={project} days={days} allSnapshots={ordered} />
      )}
    </div>
  );
}
function HistoryPanel({
  project,
  days,
  allSnapshots,
}: {
  project: Project;
  days: Snapshot[];
  allSnapshots: Snapshot[];
}) {
  const [error, setError] = useState('');
  const [busy, setBusy] = useState<number | null>(null);
  async function download(snapshot: Snapshot, index: number) {
    if (!snapshot.id) return;
    setBusy(snapshot.id);
    setError('');
    try {
      const historyToDate = allSnapshots.filter(
        (s) =>
          s.recordDate < snapshot.recordDate ||
          (s.recordDate === snapshot.recordDate &&
            (s.uploadedAt < snapshot.uploadedAt ||
              (s.uploadedAt === snapshot.uploadedAt &&
                (s.id ?? 0) <= (snapshot.id ?? 0)))),
      );
      const previous = days
        .filter((s) => s.recordDate < snapshot.recordDate)
        .at(-1);
      const [rows, previousRows, members] = await Promise.all([
        db.taskRegions.where('snapshotId').equals(snapshot.id).toArray(),
        previous?.id
          ? db.taskRegions.where('snapshotId').equals(previous.id).toArray()
          : Promise.resolve(undefined),
        db.projectMembers.where('projectId').equals(project.id!).toArray(),
      ]);
      await downloadDetailExcel({
        project,
        snapshot,
        rows,
        previousRows,
        previousTotal: previous?.totalRegions,
        members,
        history: historyToDate,
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : '다운로드에 실패했습니다.');
    } finally {
      setBusy(null);
    }
  }
  return (
    <section className="panel">
      <h2>기록 조회</h2>
      {error && <p className="error">{error}</p>}
      {allSnapshots.length === 0 ? (
        <p>저장된 기록이 없습니다.</p>
      ) : (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>기록 날짜</th>
                <th>업로드 시간</th>
                <th>전체 Task</th>
                <th>누적 Region</th>
                <th>순증감</th>
                <th>Region 0개</th>
                <th>Download</th>
              </tr>
            </thead>
            <tbody>
              {allSnapshots
                .map((s, i) => ({ s, i }))
                .reverse()
                .map(({ s, i }) => {
                  const prior = days
                    .filter((d) => d.recordDate < s.recordDate)
                    .at(-1);
                  return (
                    <tr key={s.id}>
                      <td>{s.recordDate}</td>
                      <td>{new Date(s.uploadedAt).toLocaleString('ko-KR')}</td>
                      <td>{s.totalTasks.toLocaleString()}</td>
                      <td>{s.totalRegions.toLocaleString()}</td>
                      <td>
                        {prior
                          ? (
                              s.totalRegions - prior.totalRegions
                            ).toLocaleString()
                          : '기준값'}
                      </td>
                      <td>{s.zeroRegionTasks}</td>
                      <td>
                        <button
                          className="secondary"
                          disabled={busy !== null}
                          onClick={() => void download(s, i)}
                        >
                          {busy === s.id ? '생성 중...' : 'Excel 다운로드'}
                        </button>
                      </td>
                    </tr>
                  );
                })}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}
function Metric({
  label,
  value,
  tone,
}: {
  label: string;
  value: string;
  tone?: 'positive' | 'negative' | 'neutral';
}) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong className={tone ? `delta-${tone}` : undefined}>{value}</strong>
    </div>
  );
}
function ProjectMembers({ project }: { project: Project }) {
  const [members, setMembers] = useState<ProjectMember[]>([]);
  const [userId, setUserId] = useState('');
  const [name, setName] = useState('');
  const [role, setRole] = useState<ProjectMember['role']>('worker');
  const [editing, setEditing] = useState<number | null>(null);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const refresh = () =>
    db.projectMembers
      .where('projectId')
      .equals(project.id!)
      .sortBy('labelStudioUserId')
      .then(setMembers);
  useEffect(() => {
    void refresh();
  }, [project.id]);
  const reset = () => {
    setEditing(null);
    setUserId('');
    setName('');
    setRole('worker');
  };
  const edit = (member: ProjectMember) => {
    setEditing(member.id!);
    setUserId(String(member.labelStudioUserId));
    setName(member.name);
    setRole(member.role);
    setError('');
    setNotice('');
  };
  async function saveMember() {
    setError('');
    setNotice('');
    const id = Number(userId);
    if (!Number.isSafeInteger(id) || id <= 0) {
      setError('올바른 Label Studio 사용자 ID를 입력해 주세요.');
      return;
    }
    try {
      const existing = await db.projectMembers
        .where('[projectId+labelStudioUserId]')
        .equals([project.id!, id])
        .first();
      if (existing && existing.id !== editing) {
        setError('이 프로젝트에 이미 등록된 사용자 ID입니다.');
        return;
      }
      const member = {
        projectId: project.id!,
        labelStudioUserId: id,
        name: name.trim(),
        role,
      };
      if (editing !== null) await db.projectMembers.update(editing, member);
      else await db.projectMembers.add(member);
      reset();
      await refresh();
      setNotice('작업자 정보가 저장되었습니다.');
    } catch (e) {
      setError(e instanceof Error ? e.message : '저장에 실패했습니다.');
    }
  }
  async function removeMember(member: ProjectMember) {
    if (
      !window.confirm(
        `작업자 #${member.labelStudioUserId}의 이름과 역할 정보를 삭제할까요?\n기존 분석 기록은 삭제되지 않습니다.`,
      )
    )
      return;
    try {
      await db.projectMembers.delete(member.id!);
      await refresh();
      if (editing === member.id) reset();
      setNotice('등록 정보가 삭제되었습니다.');
    } catch (e) {
      setError(e instanceof Error ? e.message : '삭제 실패');
    }
  }
  const roleText = (r: ProjectMember['role']) =>
    r === 'pl' ? 'PL' : r === 'both' ? '작업자 · PL' : '작업자';
  return (
    <section className="panel">
      <h2>작업자 관리</h2>
      <p>
        JSON에서 발견된 사용자 ID를 자동 등록해요. 이름과 역할은 이 프로젝트에만
        저장됩니다.
      </p>
      <div className="member-summary">
        등록된 사용자 <strong>{members.length}명</strong> · 이름 미등록{' '}
        <strong>{members.filter((m) => !m.name).length}명</strong>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Label Studio ID</th>
              <th>표시 이름</th>
              <th>역할</th>
              <th>관리</th>
            </tr>
          </thead>
          <tbody>
            {members.length ? (
              members.map((member) => (
                <tr key={member.id}>
                  <td>#{member.labelStudioUserId}</td>
                  <td>
                    {member.name || <span className="muted">미등록</span>}
                  </td>
                  <td>{roleText(member.role)}</td>
                  <td>
                    <div className="member-actions">
                      <button
                        className="secondary"
                        onClick={() => edit(member)}
                      >
                        <Pencil size={14} /> 수정
                      </button>
                      <button
                        className="secondary"
                        onClick={() => void removeMember(member)}
                      >
                        <Trash2 size={14} /> 삭제
                      </button>
                    </div>
                  </td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={4} className="muted">
                  아직 등록된 사용자가 없어요. JSON을 업로드하거나 직접 등록해
                  주세요.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <div className="member-form">
        <h3>{editing !== null ? '사용자 수정' : '사용자 직접 등록'}</h3>
        <div className="member-fields">
          <label>
            Label Studio ID
            <input
              type="number"
              min="1"
              value={userId}
              onChange={(e) => setUserId(e.target.value)}
              placeholder="예: 30"
            />
          </label>
          <label>
            표시 이름
            <input
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="예: 홍길동 (선택)"
            />
          </label>
          <label>
            역할
            <select
              value={role}
              onChange={(e) => setRole(e.target.value as ProjectMember['role'])}
            >
              <option value="worker">작업자</option>
              <option value="pl">PL</option>
              <option value="both">작업자 · PL</option>
            </select>
          </label>
        </div>
        <div className="actions">
          {editing !== null && (
            <button className="secondary" onClick={reset}>
              수정 취소
            </button>
          )}
          <button className="primary" onClick={() => void saveMember()}>
            <Plus size={16} />
            {editing !== null ? '변경 사항 저장' : '사용자 등록'}
          </button>
        </div>
        {error && <p className="error">{error}</p>}
        {notice && <p className="success">{notice}</p>}
      </div>
      <p className="member-hint">
        참고: 역할은 표시용 정보이며 Region 실적은 항상 Annotation의 최초
        제출자(completed_by)에게 귀속돼요.
      </p>
    </section>
  );
}
function Analyze({
  project,
  refresh,
  latest,
}: {
  project: Project;
  refresh: () => void;
  latest?: Snapshot;
}) {
  const [date, setDate] = useState(() =>
    new Date().toLocaleDateString('sv-SE'),
  );
  const [fileName, setFileName] = useState('');
  const [tasks, setTasks] = useState<Task[]>([]);
  const [drafts, setDrafts] = useState<number[]>([]);
  const [showAllDrafts, setShowAllDrafts] = useState(false);
  const [error, setError] = useState('');
  const [status, setStatus] = useState('');
  const [saving, setSaving] = useState(false);
  const [members, setMembers] = useState<ProjectMember[]>([]);
  async function upload(file?: File) {
    if (!file) return;
    setError('');
    setStatus('');
    setTasks([]);
    setShowAllDrafts(false);
    setDrafts([]);
    setFileName(file.name);
    try {
      const raw = JSON.parse(await file.text()) as unknown;
      const result = inspectExport(raw, project.labelStudioProjectId);
      setTasks(result.tasks);
      setDrafts(result.draftTaskIds);
      if (result.draftTaskIds.length) {
        setStatus(
          'Draft가 발견되었습니다. Draft는 집계에서 제외하고 완료된 Annotation만 분석합니다.',
        );
      }
      // Export에서 확인된 제출자/수정자 ID를 프로젝트에 자동 등록한다. 기존 이름과 역할은 보존한다.
      const foundIds = new Set<number>();
      for (const task of result.tasks)
        for (const a of task.annotations ?? []) {
          if (typeof a.completed_by === 'number') foundIds.add(a.completed_by);
          if (typeof a.updated_by === 'number') foundIds.add(a.updated_by);
        }
      await db.transaction('rw', db.projectMembers, async () => {
        for (const id of foundIds) {
          const found = await db.projectMembers
            .where('[projectId+labelStudioUserId]')
            .equals([project.id!, id])
            .first();
          if (!found)
            await db.projectMembers.add({
              projectId: project.id!,
              labelStudioUserId: id,
              name: '',
              role: 'worker',
            });
        }
      });
      const currentMembers = await db.projectMembers
        .where('projectId')
        .equals(project.id!)
        .toArray();
      setMembers(currentMembers);
    } catch (e) {
      setError(e instanceof Error ? e.message : '파일을 분석하지 못했습니다.');
    }
  }
  const multi = tasks
    .map((task) => ({
      task,
      candidates: (task.annotations ?? []).filter((a) => !a.was_cancelled),
    }))
    .filter((item) => item.candidates.length > 1);
  const selectedRows = tasks.flatMap((task) =>
    (task.annotations ?? [])
      .filter((annotation) => !annotation.was_cancelled)
      .map((annotation) => ({
        taskId: task.id,
        annotationId: annotation.id,
        completedBy: annotation.completed_by ?? null,
        updatedBy: annotation.updated_by ?? null,
        updatedAt: annotation.updated_at ?? null,
        regionCount: annotation.result?.length ?? 0,
      })),
  );
  const totalTasks = new Set(selectedRows.map((row) => row.taskId)).size;

  const zeroRegionTasks = [
    ...new Set(selectedRows.map((row) => row.taskId)),
  ].filter((taskId) =>
    selectedRows
      .filter((row) => row.taskId === taskId)
      .every((row) => row.regionCount === 0),
  ).length;

  const totalAnnotations = selectedRows.length;
  const total = selectedRows.reduce((sum, row) => sum + row.regionCount, 0);
  const [previousForDate, setPreviousForDate] = useState<Snapshot>();

  useEffect(() => {
    let cancelled = false;

    async function loadPrevious() {
      const snapshots = await db.snapshots
        .where('projectId')
        .equals(project.id!)
        .toArray();

      const previous = snapshots
        .filter((snapshot) => snapshot.recordDate < date)
        .sort(
          (a, b) =>
            b.recordDate.localeCompare(a.recordDate) ||
            b.uploadedAt.localeCompare(a.uploadedAt) ||
            (b.id ?? 0) - (a.id ?? 0),
        )[0];

      if (!cancelled) setPreviousForDate(previous);
    }

    void loadPrevious();

    return () => {
      cancelled = true;
    };
  }, [project.id, date, latest]);
  async function downloadPreview() {
    if (!tasks.length) return;
    setError('');
    try {
      const earlier = await db.snapshots
        .where('projectId')
        .equals(project.id!)
        .toArray();
      const previousForDate = earlier
        .filter((s) => s.recordDate < date)
        .sort(
          (a, b) =>
            b.recordDate.localeCompare(a.recordDate) ||
            b.uploadedAt.localeCompare(a.uploadedAt) ||
            (b.id ?? 0) - (a.id ?? 0),
        )[0];
      const previousRows = previousForDate?.id
        ? await db.taskRegions
            .where('snapshotId')
            .equals(previousForDate.id)
            .toArray()
        : undefined;
      const history = await db.snapshots
        .where('projectId')
        .equals(project.id!)
        .toArray();
      const perDay = new Map<string, Snapshot>();
      history
        .sort(
          (a, b) =>
            a.recordDate.localeCompare(b.recordDate) ||
            a.uploadedAt.localeCompare(b.uploadedAt) ||
            (a.id ?? 0) - (b.id ?? 0),
        )
        .forEach((s) => perDay.set(s.recordDate, s));
      const snapshot = {
        recordDate: date,
        fileName,
        totalTasks,
        totalRegions: total,
        zeroRegionTasks,
      };
      await downloadDetailExcel({
        project,
        snapshot,
        rows: selectedRows,
        previousRows,
        previousTotal: previousForDate?.totalRegions,
        members,
        history: [...perDay.values()]
          .filter((s) => s.recordDate < date)
          .concat([
            {
              ...snapshot,
              projectId: project.id!,
              uploadedAt: new Date().toISOString(),
            },
          ]),
      });
    } catch (e) {
      setError(
        e instanceof Error ? e.message : 'Excel 다운로드에 실패했습니다.',
      );
    }
  }
  async function save() {
    if (!tasks.length || !date) return;
    setSaving(true);
    setError('');
    try {
      await db.transaction('rw', db.snapshots, db.taskRegions, async () => {
        const snapshotId = await db.snapshots.add({
          projectId: project.id!,
          recordDate: date,
          uploadedAt: new Date().toISOString(),
          fileName,
          totalTasks,
          totalRegions: total,
          zeroRegionTasks,
        });
        if (snapshotId === undefined)
          throw new Error('분석 기록 ID를 생성하지 못했습니다.');
        await db.taskRegions.bulkAdd(
          selectedRows.map((row) => ({ ...row, snapshotId })),
        );
      });
      setStatus('분석 결과를 저장했습니다.');
      setTasks([]);
      setDrafts([]);
      setShowAllDrafts(false);
      refresh();
    } catch (e) {
      setError(e instanceof Error ? e.message : '저장에 실패했습니다.');
    } finally {
      setSaving(false);
    }
  }
  return (
    <section className="panel">
      <h2>JSON 업로드 및 분석</h2>
      <p>Label Studio에서 Export한 프로젝트 전체 JSON을 업로드하세요.</p>
      <div className="form-row">
        <label>
          기록 날짜
          <input
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
          />
        </label>
        <label className="file-picker">
          <UploadCloud size={19} /> JSON 파일 선택
          <input
            type="file"
            accept=".json,application/json"
            onChange={(e) => void upload(e.target.files?.[0])}
          />
        </label>
      </div>
      {fileName && (
        <p className="filename">
          <FileJson size={16} />
          {fileName}
        </p>
      )}
      {error && <p className="error">{error}</p>}
      {status && !(drafts.length > 0 && tasks.length > 0) && (
        <p
          className={
            tasks.length > 0 && drafts.length > 0 ? 'warning' : 'success'
          }
        >
          {status}
        </p>
      )}
      {tasks.length > 0 && (drafts.length > 0 || multi.length > 0) && (
        <div
          className="analysis-notices"
          style={{
            border: '1px solid #e6d4b5',
            background: '#fffbf3',
            borderRadius: 14,
            padding: '20px 24px',
            margin: '24px 0',
          }}
        >
          <AlertTriangle
            size={20}
            color="#86580f"
            aria-label="분석 시 확인 사항"
            style={{ marginBottom: 16 }}
          />
          {drafts.length > 0 && (
            <div>
              <h4
                style={{ margin: '0 0 10px', fontSize: 14, color: '#86580f' }}
              >
                Draft가 있는 Task · {drafts.length}개
              </h4>
              <div
                className="multi-annotation-tasks"
                style={{ marginBottom: 14 }}
              >
                {(showAllDrafts ? drafts : drafts.slice(0, 5)).map((taskId) => (
                  <span className="multi-annotation-tag" key={taskId}>
                    #{taskId}
                  </span>
                ))}
                {drafts.length > 5 && (
                  <button
                    type="button"
                    onClick={() => setShowAllDrafts((current) => !current)}
                    style={{
                      border: '1px solid #ead3ac',
                      borderRadius: 8,
                      padding: '4px 10px',
                      background: '#fff9ed',
                      color: '#86580f',
                      fontSize: 12,
                      cursor: 'pointer',
                    }}
                  >
                    {showAllDrafts ? '접기' : `${drafts.length - 5}개 더 보기`}
                  </button>
                )}
              </div>
              <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                <button
                  className="secondary"
                  onClick={() =>
                    void navigator.clipboard.writeText(drafts.join('\\n'))
                  }
                >
                  Task ID 복사
                </button>
                <button
                  className="secondary"
                  onClick={() => void downloadDraftExcel(project, date, drafts)}
                >
                  Draft 목록 다운로드
                </button>
              </div>
            </div>
          )}
          {drafts.length > 0 && multi.length > 0 && (
            <hr
              style={{
                border: 0,
                borderTop: '1px solid #eadfce',
                margin: '20px 0',
              }}
            />
          )}
          {multi.length > 0 && (
            <div>
              <h4
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  margin: '0 0 10px',
                  fontSize: 14,
                  color: '#86580f',
                }}
              >
                다중 Annotation Task · {multi.length}개
              </h4>
              <div className="multi-annotation-tasks">
                {multi.map(({ task }) => (
                  <span className="multi-annotation-tag" key={task.id}>
                    #{task.id}
                  </span>
                ))}
              </div>
            </div>
          )}
          <p style={{ margin: '16px 0 0', fontSize: 12, color: '#786a58' }}>
            Draft는 제외하고 모든 유효 Annotation을 집계했습니다.
          </p>
        </div>
      )}
      {tasks.length > 0 && (
        <>
          <div className="metrics">
            <Metric label="전체 Task" value={totalTasks.toLocaleString()} />
            <Metric
              label="전체 Annotation"
              value={totalAnnotations.toLocaleString()}
            />
            <Metric label="누적 Region" value={total.toLocaleString()} />
            <Metric
              label="Region 0개 Task"
              value={zeroRegionTasks.toLocaleString()}
            />
            <Metric
              label="이전 기록 대비"
              value={
                previousForDate
                  ? `${total - previousForDate.totalRegions >= 0 ? '+' : ''}${(
                      total - previousForDate.totalRegions
                    ).toLocaleString()}`
                  : '기준값'
              }
              tone={
                previousForDate
                  ? total > previousForDate.totalRegions
                    ? 'positive'
                    : total < previousForDate.totalRegions
                      ? 'negative'
                      : 'neutral'
                  : 'neutral'
              }
            />
          </div>
          <div className="actions">
            <span className="muted">
              {multi.length > 0 ? '전체 집계 완료' : '모든 검사 완료'}
            </span>
            <button
              className="secondary"
              disabled={saving}
              onClick={() => void downloadPreview()}
            >
              상세 분석 Excel 다운로드
            </button>
            <button
              className="primary"
              disabled={saving}
              onClick={() => void save()}
            >
              <CheckCircle2 size={17} />
              {saving ? '저장 중...' : '분석 결과 저장'}
            </button>
          </div>
        </>
      )}
    </section>
  );
}
export default function App() {
  return (
    <div className="app">
      <aside className="sidebar">
        <Link className="brand" to="/">
          <div className="brand-icon">
            <FileJson size={20} />
          </div>
          <span>
            Audio Labeling
            <br />
            Tracker
          </span>
        </Link>
        <nav>
          <NavLink
            to="/"
            className={({ isActive }) => (isActive ? 'nav active' : 'nav')}
          >
            <Folder size={17} /> 프로젝트
          </NavLink>
        </nav>
        <div className="version">MVP · v0.4</div>
      </aside>
      <main>
        <Routes>
          <Route path="/" element={<ProjectHome />} />
          <Route path="/projects/:projectId/*" element={<ProjectShell />} />
          <Route path="*" element={<ProjectHome />} />
        </Routes>
      </main>
    </div>
  );
}
