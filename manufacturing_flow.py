"""加工工程と納品後の実績。既存procurements.statusとは独立し、加工工程は明示登録のみ。"""
from datetime import date
import sqlite3

MILESTONES = ('納品', '検査合格', '入庫', '出庫準備', '出庫')

class FlowError(ValueError):
    pass

def create_tables(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS manufacturing_steps (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        part_id INTEGER NOT NULL REFERENCES parts(id) ON DELETE CASCADE,
        process_order INTEGER NOT NULL,
        process_name TEXT NOT NULL,
        progress_percent INTEGER NOT NULL DEFAULT 0 CHECK(progress_percent BETWEEN 0 AND 100),
        completed_on TEXT,
        UNIQUE(part_id, process_order),
        CHECK ((progress_percent = 100 AND completed_on IS NOT NULL) OR
               (progress_percent < 100 AND completed_on IS NULL)))''')
    conn.execute('''CREATE TABLE IF NOT EXISTS part_flow_milestones (
        part_id INTEGER NOT NULL REFERENCES parts(id) ON DELETE CASCADE,
        milestone TEXT NOT NULL CHECK (milestone IN ('納品','検査合格','入庫','出庫準備','出庫')),
        completed_on TEXT NOT NULL,
        PRIMARY KEY (part_id, milestone))''')
    conn.commit()

def _day(value):
    try:
        if date.fromisoformat(value).isoformat() != value: raise ValueError
    except (ValueError, TypeError):
        raise FlowError('実績日はYYYY-MM-DD形式で指定してください') from None

def _steps(conn, part_id):
    return conn.execute('SELECT id,process_order,progress_percent,completed_on FROM manufacturing_steps WHERE part_id=? ORDER BY process_order', (part_id,)).fetchall()

def add_step(conn, part_id, process_order, name):
    if not conn.execute('SELECT 1 FROM parts WHERE id=?', (part_id,)).fetchone(): raise FlowError('部品がありません')
    if type(process_order) is not int or process_order < 0 or not name.strip(): raise FlowError('工程情報が不正です')
    if conn.execute('SELECT 1 FROM part_flow_milestones WHERE part_id=?', (part_id,)).fetchone():
        raise FlowError('納品後の工程は追加できません')
    if any(step[2] > 0 for step in _steps(conn, part_id)):
        raise FlowError('加工実績がある部品には工程を追加できません')
    with conn:
        return conn.execute('INSERT INTO manufacturing_steps(part_id,process_order,process_name) VALUES (?,?,?)', (part_id, process_order, name.strip())).lastrowid

def can_start(conn, step_id):
    row = conn.execute('SELECT part_id FROM manufacturing_steps WHERE id=?', (step_id,)).fetchone()
    if not row: raise FlowError('工程がありません')
    steps = _steps(conn,row[0]); index = next(i for i,s in enumerate(steps) if s[0]==step_id)
    return index==0 or steps[index-1][2]==100

def set_step_progress(conn, step_id, percent, completed_on=None):
    if type(percent) is not int or not 0<=percent<=100: raise FlowError('進捗率は0～100の整数です')
    if (percent==100) != (completed_on is not None): raise FlowError('完了日と進捗率が不一致です')
    if completed_on is not None: _day(completed_on)
    row = conn.execute('SELECT part_id FROM manufacturing_steps WHERE id=?', (step_id,)).fetchone()
    if not row: raise FlowError('工程がありません')
    part_id=row[0]; steps=_steps(conn,part_id); idx=next(i for i,s in enumerate(steps) if s[0]==step_id)
    if percent>0 and not can_start(conn,step_id): raise FlowError('先行工程が未完了です')
    if idx>0 and completed_on is not None and completed_on<steps[idx-1][3]:
        raise FlowError('先行工程より完了日が早いです')
    if completed_on is not None and idx+1<len(steps) and steps[idx+1][3] and completed_on>steps[idx+1][3]:
        raise FlowError('後続工程より完了日が遅いです')
    if percent<100 and any(s[2]>0 for s in steps[idx+1:]): raise FlowError('後続工程に実績があります')
    milestone = conn.execute("SELECT completed_on FROM part_flow_milestones WHERE part_id=? AND milestone='納品'", (part_id,)).fetchone()
    if idx==len(steps)-1 and milestone:
        if percent<100: raise FlowError('納品実績があるので未完了に戻せません')
        if completed_on>milestone[0]: raise FlowError('加工完了日が納品日より後です')
    with conn:
        conn.execute('UPDATE manufacturing_steps SET progress_percent=?,completed_on=? WHERE id=?',(percent,completed_on,step_id))

def next_milestone(conn,part_id):
    if not conn.execute('SELECT 1 FROM parts WHERE id=?',(part_id,)).fetchone(): raise FlowError('部品がありません')
    steps=_steps(conn,part_id)
    if not steps or any(s[2]<100 for s in steps): return None
    done={r[0] for r in conn.execute('SELECT milestone FROM part_flow_milestones WHERE part_id=?',(part_id,))}
    return next((m for m in MILESTONES if m not in done),None)

def record_milestone(conn,part_id,milestone,completed_on):
    if milestone not in MILESTONES: raise FlowError('未定義の段階です')
    _day(completed_on)
    if next_milestone(conn,part_id)!=milestone: raise FlowError('前段階が未完了です')
    idx=MILESTONES.index(milestone)
    prev=(_steps(conn,part_id)[-1][3] if idx==0 else conn.execute(
        'SELECT completed_on FROM part_flow_milestones WHERE part_id=? AND milestone=?',
        (part_id,MILESTONES[idx-1])).fetchone()[0])
    if completed_on<prev: raise FlowError('実績日が前段階より早いです')
    with conn:
        conn.execute('INSERT INTO part_flow_milestones(part_id,milestone,completed_on) VALUES(?,?,?)',
                     (part_id,milestone,completed_on))
