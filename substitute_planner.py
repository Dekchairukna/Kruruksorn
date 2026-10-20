"""ระบบจัดสอนแทน (Substitute Planner)

ขั้นตอน
1. ผู้จัดสอนแทน (แอดมิน หรือครูที่ได้รับสิทธิ์) เลือกวันที่
2. ติ๊กครูที่ไม่มา (ลา/ไปราชการ/ป่วย/อื่น ๆ) ทั้งวันหรือบางคาบ
3. ระบบดึงคาบสอนของครูที่ไม่มาในวันนั้นจากตารางสอนภาคเรียนปัจจุบัน
   และแสดงครูที่ "ว่าง" ในคาบนั้น (ไม่มีสอน ไม่ได้ลา ไม่ได้รับมอบหมายสอนแทนซ้อน)
   เรียงตามจำนวนครั้งที่สอนแทนในเดือนนี้ (น้อยก่อน) เพื่อความเป็นธรรม
4. เลือกครูสอนแทนทีละคาบ หรือกด "จัดให้อัตโนมัติ"
5. พิมพ์ "แบบมอบหมายการปฏิบัติการสอนแทน" หรือบันทึกเป็นรูปภาพไปส่งในไลน์
6. ครูผู้สอนแทนเห็นคาบที่ต้องสอนแทนในหน้า "สอนแทนของฉัน" และกดเช็กชื่อแทนได้ทันที

ใช้งาน: ``substitute_planner.register(app, globals())`` ใน app.py ก่อน init_db()
"""
from datetime import datetime, date, timedelta
from collections import defaultdict

from flask import render_template, request, redirect, url_for, flash, abort
from flask_login import login_required, current_user

DAY_NAMES = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
THAI_MONTHS = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
               'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
REASONS = ['ลากิจ', 'ลาป่วย', 'ไปราชการ', 'อบรม/ประชุม', 'พาเด็กไปกิจกรรม', 'อื่น ๆ']
NO_SUB = 'ไม่ได้จัดสอนแทน'


def thai_full_date(d):
    return f'วัน{DAY_NAMES[d.weekday()]} ที่ {d.day} {THAI_MONTHS[d.month]} {d.year + 543}'


def register(app, g):
    db = g['db']
    User = g['User']
    TeachingSchedule = g['TeachingSchedule']
    role_required = g['role_required']
    active_schedule_query = g['active_schedule_query']
    get_active_semester = g['get_active_semester']
    local_today = g['local_today']
    get_school_setting = g.get('get_school_setting')
    is_school_blocked_day = g.get('is_school_blocked_day')
    _current_school_id = g.get('current_school_id')

    def sid():
        try:
            return _current_school_id() if _current_school_id else getattr(current_user, 'school_id', None)
        except Exception:
            return getattr(current_user, 'school_id', None)

    # ------------------------------------------------------------------
    # Models
    # ------------------------------------------------------------------
    class SubSetting(db.Model):
        """สิทธิ์ต่อผู้ใช้: can_manage = จัดสอนแทนได้, can_substitute = ถูกจัดให้สอนแทนได้"""
        __tablename__ = 'sub_setting'
        id = db.Column(db.Integer, primary_key=True)
        school_id = db.Column(db.Integer, nullable=True, index=True)
        user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False, index=True)
        can_manage = db.Column(db.Boolean, default=False)
        can_substitute = db.Column(db.Boolean, default=True)

    class SubDay(db.Model):
        __tablename__ = 'sub_day'
        id = db.Column(db.Integer, primary_key=True)
        school_id = db.Column(db.Integer, nullable=True, index=True)
        sub_date = db.Column(db.Date, nullable=False, index=True)
        note = db.Column(db.Text, default='')
        created_by_id = db.Column(db.Integer, nullable=True)
        updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    class SubAbsence(db.Model):
        __tablename__ = 'sub_absence'
        id = db.Column(db.Integer, primary_key=True)
        day_id = db.Column(db.Integer, db.ForeignKey('sub_day.id'), nullable=False, index=True)
        teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
        reason = db.Column(db.String(80), default='ลากิจ')
        periods = db.Column(db.String(60), default='')   # ว่าง = ทั้งวัน, หรือ "1,2,5"
        teacher = db.relationship('User')

        @property
        def period_set(self):
            return {int(x) for x in (self.periods or '').split(',') if x.strip().isdigit()}

    class SubAssign(db.Model):
        __tablename__ = 'sub_assign'
        id = db.Column(db.Integer, primary_key=True)
        day_id = db.Column(db.Integer, db.ForeignKey('sub_day.id'), nullable=False, index=True)
        schedule_id = db.Column(db.Integer, nullable=True)
        absent_teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
        period_no = db.Column(db.Integer, nullable=False)
        subject_label = db.Column(db.String(255), default='')
        classroom_label = db.Column(db.String(120), default='')
        subject_id = db.Column(db.Integer, nullable=True)
        classroom_id = db.Column(db.Integer, nullable=True)
        substitute_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
        note = db.Column(db.String(200), default='')
        absent_teacher = db.relationship('User', foreign_keys=[absent_teacher_id])
        substitute = db.relationship('User', foreign_keys=[substitute_id])

    for _c in (SubSetting, SubDay, SubAbsence, SubAssign):
        g[_c.__name__] = _c

    # ------------------------------------------------------------------
    # helpers
    # ------------------------------------------------------------------
    def teachers_q():
        q = User.query.filter(User.role == 'teacher')
        q = q.filter(db.or_(User.is_active == True, User.is_active.is_(None)))
        if hasattr(User, 'school_id'):
            q = q.filter(User.school_id == sid())
        return q.order_by(User.full_name)

    def settings_map():
        return {s.user_id: s for s in SubSetting.query.filter_by(school_id=sid()).all()}

    def can_manage(user=None):
        user = user or current_user
        if not user.is_authenticated:
            return False
        if user.role == 'admin':
            return True
        s = SubSetting.query.filter_by(school_id=sid(), user_id=user.id).first()
        return bool(s and s.can_manage)

    def require_manager():
        if not can_manage():
            flash('บัญชีนี้ยังไม่ได้รับสิทธิ์จัดสอนแทน ให้แอดมินเปิดสิทธิ์ที่หน้า ตั้งค่าสอนแทน', 'danger')
            return redirect(url_for('subplan_mine'))
        return None

    def parse_date(text):
        try:
            return datetime.strptime((text or '').strip(), '%Y-%m-%d').date()
        except ValueError:
            return local_today()

    def get_day(d, create=False):
        day = SubDay.query.filter_by(school_id=sid(), sub_date=d).first()
        if not day and create:
            day = SubDay(school_id=sid(), sub_date=d, created_by_id=current_user.id)
            db.session.add(day)
            db.session.flush()
        return day

    def schedules_on(d):
        """คาบสอนทั้งหมดของวันนั้น (ภาคเรียนปัจจุบัน) -> list"""
        return active_schedule_query().filter(TeachingSchedule.weekday == d.weekday()).all()

    def month_counts(d):
        start = d.replace(day=1)
        end = (start + timedelta(days=32)).replace(day=1)
        rows = (db.session.query(SubAssign.substitute_id, db.func.count(SubAssign.id))
                .join(SubDay, SubAssign.day_id == SubDay.id)
                .filter(SubDay.school_id == sid(), SubDay.sub_date >= start, SubDay.sub_date < end,
                        SubAssign.substitute_id.isnot(None))
                .group_by(SubAssign.substitute_id).all())
        return dict(rows)

    def sync_assignments(day):
        """สร้าง/ลบแถว SubAssign ให้ตรงกับครูที่ไม่มาและคาบสอนของวันนั้น (คงครูแทนที่เลือกไว้)"""
        absences = {a.teacher_id: a for a in SubAbsence.query.filter_by(day_id=day.id).all()}
        existing = {(a.absent_teacher_id, a.schedule_id): a for a in SubAssign.query.filter_by(day_id=day.id).all()}
        wanted = set()
        for s in schedules_on(day.sub_date):
            ab = absences.get(s.teacher_id)
            if not ab:
                continue
            if ab.period_set and s.period_no not in ab.period_set:
                continue
            key = (s.teacher_id, s.id)
            wanted.add(key)
            if key not in existing:
                db.session.add(SubAssign(day_id=day.id, schedule_id=s.id, absent_teacher_id=s.teacher_id,
                                         period_no=s.period_no, subject_id=s.subject_id, classroom_id=s.classroom_id,
                                         subject_label=(s.subject.name if s.subject else ''),
                                         classroom_label=(s.classroom.name if s.classroom else '')))
        for key, a in existing.items():
            if key not in wanted:
                db.session.delete(a)

    def busy_map(day, scheds):
        """(teacher_id, period) ที่ไม่ว่าง: มีสอนเอง / ลาคาบนั้น / ได้รับมอบหมายสอนแทนแล้ว"""
        busy = defaultdict(set)   # (tid, period) -> reasons
        for s in scheds:
            busy[(s.teacher_id, s.period_no)].add('สอน ' + (s.classroom.name if s.classroom else ''))
        for a in SubAbsence.query.filter_by(day_id=day.id).all():
            ps = a.period_set or set(range(1, 13))
            for p in ps:
                busy[(a.teacher_id, p)].add('ไม่อยู่')
        for x in SubAssign.query.filter_by(day_id=day.id).filter(SubAssign.substitute_id.isnot(None)).all():
            busy[(x.substitute_id, x.period_no)].add(f'สอนแทน {x.classroom_label}')
        return busy

    def candidates_for(assign, teachers, busy, counts, loads, smap):
        out = []
        for t in teachers:
            st = smap.get(t.id)
            if st is not None and st.can_substitute is False:
                continue
            if t.id == assign.absent_teacher_id:
                continue
            reasons = busy.get((t.id, assign.period_no), set())
            mine = assign.substitute_id == t.id
            if reasons and not mine:
                continue
            out.append({'t': t, 'count': counts.get(t.id, 0), 'load': loads.get(t.id, 0)})
        out.sort(key=lambda c: (c['count'], c['load'], c['t'].full_name))
        return out

    @app.context_processor
    def _subplan_ctx():
        try:
            return {'subplan_enabled': True, 'subplan_can_manage': can_manage()}
        except Exception:
            return {'subplan_enabled': True, 'subplan_can_manage': False}

    # ------------------------------------------------------------------
    # หน้าจัดสอนแทนรายวัน
    # ------------------------------------------------------------------
    @app.route('/substitute-plan', methods=['GET', 'POST'])
    @login_required
    @role_required('teacher', 'admin')
    def subplan_day():
        deny = require_manager()
        if deny:
            return deny
        d = parse_date(request.values.get('date'))
        teachers = teachers_q().all()
        tids = {t.id for t in teachers}
        if request.method == 'POST':
            action = request.form.get('action')
            day = get_day(d, create=True)
            if action == 'absent':
                keep = set()
                for tid in request.form.getlist('absent'):
                    if not tid.isdigit() or int(tid) not in tids:
                        continue
                    tid = int(tid)
                    keep.add(tid)
                    a = SubAbsence.query.filter_by(day_id=day.id, teacher_id=tid).first()
                    if not a:
                        a = SubAbsence(day_id=day.id, teacher_id=tid)
                        db.session.add(a)
                    a.reason = (request.form.get(f'reason_{tid}') or 'ลากิจ')[:80]
                    a.periods = ','.join(sorted({p for p in request.form.getlist(f'p_{tid}') if p.isdigit()}, key=int))
                for a in SubAbsence.query.filter_by(day_id=day.id).all():
                    if a.teacher_id not in keep:
                        db.session.delete(a)
                db.session.flush()
                sync_assignments(day)
                db.session.commit()
                flash(f'บันทึกครูที่ไม่มา {len(keep)} คนแล้ว — เลือกครูสอนแทนในตารางด้านล่าง', 'success')
            elif action == 'assign':
                for a in SubAssign.query.filter_by(day_id=day.id).all():
                    v = request.form.get(f'sub_{a.id}')
                    if v is None:
                        continue
                    a.substitute_id = int(v) if v.isdigit() and int(v) in tids else None
                    a.note = (request.form.get(f'note_{a.id}') or '')[:200]
                day.note = (request.form.get('day_note') or '').strip()
                # ตรวจไม่ให้ครูคนเดียวสอนแทนสองที่ในคาบเดียวกัน
                seen = {}
                clash = []
                for a in SubAssign.query.filter_by(day_id=day.id).filter(SubAssign.substitute_id.isnot(None)).all():
                    k = (a.substitute_id, a.period_no)
                    if k in seen:
                        clash.append(a)
                        a.substitute_id = None
                    seen[k] = a
                db.session.commit()
                if clash:
                    flash(f'มี {len(clash)} คาบที่เลือกครูซ้ำในคาบเดียวกัน ระบบล้างออกให้แล้ว กรุณาเลือกใหม่', 'warning')
                else:
                    flash('บันทึกการจัดสอนแทนแล้ว', 'success')
            elif action == 'auto':
                sync_assignments(day)
                db.session.flush()
                scheds = schedules_on(d)
                counts = month_counts(d)
                loads = defaultdict(int)
                for s in scheds:
                    loads[s.teacher_id] += 1
                smap = settings_map()
                n = 0
                for a in sorted(SubAssign.query.filter_by(day_id=day.id).all(), key=lambda a: a.period_no):
                    if a.substitute_id:
                        continue
                    busy = busy_map(day, scheds)
                    cands = candidates_for(a, teachers, busy, counts, loads, smap)
                    if cands:
                        a.substitute_id = cands[0]['t'].id
                        counts[a.substitute_id] = counts.get(a.substitute_id, 0) + 1
                        loads[a.substitute_id] += 1
                        db.session.flush()
                        n += 1
                db.session.commit()
                flash(f'จัดครูสอนแทนให้อัตโนมัติ {n} คาบ (เลือกคนที่ว่างและสอนแทนน้อยที่สุดในเดือนนี้) — ตรวจและแก้ได้ก่อนพิมพ์', 'success')
            elif action == 'clear':
                SubAssign.query.filter_by(day_id=day.id).update({'substitute_id': None})
                db.session.commit()
                flash('ล้างครูสอนแทนของวันนี้แล้ว', 'success')
            return redirect(url_for('subplan_day', date=d.isoformat()))

        day = get_day(d)
        if day:
            sync_assignments(day)
            db.session.commit()
        scheds = schedules_on(d)
        loads = defaultdict(int)
        my_periods = defaultdict(list)
        for s in scheds:
            loads[s.teacher_id] += 1
            my_periods[s.teacher_id].append(s.period_no)
        absences = {a.teacher_id: a for a in SubAbsence.query.filter_by(day_id=day.id).all()} if day else {}
        assigns = SubAssign.query.filter_by(day_id=day.id).order_by(SubAssign.absent_teacher_id, SubAssign.period_no).all() if day else []
        counts = month_counts(d)
        smap = settings_map()
        busy = busy_map(day, scheds) if day else defaultdict(set)
        cands = {a.id: candidates_for(a, teachers, busy, counts, loads, smap) for a in assigns}
        # ครูว่างรายคาบ (ภาพรวม)
        periods = sorted({s.period_no for s in scheds} | set(range(1, 9)))
        free_by_period = {}
        for p in periods:
            free_by_period[p] = [t for t in teachers if not busy.get((t.id, p))
                                 and not (smap.get(t.id) is not None and smap[t.id].can_substitute is False)]
        by_teacher = defaultdict(list)
        for a in assigns:
            by_teacher[a.absent_teacher_id].append(a)
        sem = get_active_semester()
        return render_template('subplan/day.html', d=d, day=day, teachers=teachers, absences=absences,
                               by_teacher=by_teacher, cands=cands, loads=loads, my_periods=my_periods, counts=counts,
                               periods=periods, free_by_period=free_by_period, reasons=REASONS, sem=sem,
                               date_label=thai_full_date(d), no_schedule=not scheds,
                               blocked=(is_school_blocked_day(d) if is_school_blocked_day else False),
                               prev_d=(d - timedelta(days=1)).isoformat(), next_d=(d + timedelta(days=1)).isoformat(),
                               n_unassigned=sum(1 for a in assigns if not a.substitute_id))

    @app.route('/substitute-plan/print')
    @login_required
    @role_required('teacher', 'admin')
    def subplan_print():
        d = parse_date(request.args.get('date'))
        day = get_day(d)
        if not day:
            flash('ยังไม่มีการจัดสอนแทนของวันนี้', 'warning')
            return redirect(url_for('subplan_day', date=d.isoformat()) if can_manage() else url_for('subplan_mine'))
        absences = SubAbsence.query.filter_by(day_id=day.id).all()
        assigns = SubAssign.query.filter_by(day_id=day.id).order_by(SubAssign.period_no).all()
        by_teacher = defaultdict(list)
        for a in assigns:
            by_teacher[a.absent_teacher_id].append(a)
        sem = get_active_semester()
        setting = get_school_setting() if get_school_setting else None
        sections = []
        for ab in sorted(absences, key=lambda a: (a.teacher.full_name if a.teacher else '')):
            sections.append({'ab': ab, 'rows': by_teacher.get(ab.teacher_id, [])})
        return render_template('subplan/print.html', d=d, day=day, sections=sections, sem=sem, setting=setting,
                               date_label=thai_full_date(d), no_sub=NO_SUB)

    # ------------------------------------------------------------------
    # สอนแทนของฉัน (ครูทุกคน)
    # ------------------------------------------------------------------
    @app.route('/substitute-plan/mine')
    @login_required
    @role_required('teacher', 'admin')
    def subplan_mine():
        today = local_today()
        rows = (SubAssign.query.join(SubDay, SubAssign.day_id == SubDay.id)
                .filter(SubDay.school_id == sid(), SubAssign.substitute_id == current_user.id,
                        SubDay.sub_date >= today - timedelta(days=30))
                .order_by(SubDay.sub_date.desc(), SubAssign.period_no).all())
        days = {x.id: x for x in SubDay.query.filter(SubDay.id.in_({r.day_id for r in rows} or {-1})).all()}
        upcoming = [r for r in rows if days[r.day_id].sub_date >= today]
        past = [r for r in rows if days[r.day_id].sub_date < today]
        recent_days = SubDay.query.filter(SubDay.school_id == sid(), SubDay.sub_date >= today - timedelta(days=7)).order_by(SubDay.sub_date.desc()).all()
        return render_template('subplan/mine.html', upcoming=sorted(upcoming, key=lambda r: (days[r.day_id].sub_date, r.period_no)),
                               past=past, days=days, today=today, recent_days=recent_days, thai_full_date=thai_full_date)

    # ------------------------------------------------------------------
    # สถิติ / ประวัติ
    # ------------------------------------------------------------------
    @app.route('/substitute-plan/history')
    @login_required
    @role_required('teacher', 'admin')
    def subplan_history():
        deny = require_manager()
        if deny:
            return deny
        today = local_today()
        sem = get_active_semester()
        start = sem.start_date if sem else today.replace(month=1, day=1)
        end = sem.end_date if sem else today
        if request.args.get('month'):
            try:
                y, m = [int(x) for x in request.args['month'].split('-')]
                start = date(y, m, 1)
                end = (start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
            except ValueError:
                pass
        days = SubDay.query.filter(SubDay.school_id == sid(), SubDay.sub_date >= start, SubDay.sub_date <= end).order_by(SubDay.sub_date.desc()).all()
        ids = [x.id for x in days]
        assigns = SubAssign.query.filter(SubAssign.day_id.in_(ids or [-1])).all()
        absences = SubAbsence.query.filter(SubAbsence.day_id.in_(ids or [-1])).all()
        sub_count = defaultdict(int)
        absent_count = defaultdict(int)
        uncovered = 0
        for a in assigns:
            if a.substitute_id:
                sub_count[a.substitute_id] += 1
            else:
                uncovered += 1
            absent_count[a.absent_teacher_id] += 1
        teachers = teachers_q().all()
        per_day = defaultdict(lambda: {'absent': 0, 'periods': 0, 'covered': 0})
        for ab in absences:
            per_day[ab.day_id]['absent'] += 1
        for a in assigns:
            per_day[a.day_id]['periods'] += 1
            per_day[a.day_id]['covered'] += 1 if a.substitute_id else 0
        return render_template('subplan/history.html', days=days, per_day=per_day, teachers=teachers, sub_count=sub_count,
                               absent_count=absent_count, uncovered=uncovered, start=start, end=end,
                               thai_full_date=thai_full_date, total=len(assigns))

    # ------------------------------------------------------------------
    # ตั้งค่าสิทธิ์
    # ------------------------------------------------------------------
    @app.route('/substitute-plan/settings', methods=['GET', 'POST'])
    @login_required
    @role_required('admin')
    def subplan_settings():
        teachers = teachers_q().all()
        smap = settings_map()
        if request.method == 'POST':
            managers = {int(x) for x in request.form.getlist('manage') if x.isdigit()}
            subs = {int(x) for x in request.form.getlist('substitute') if x.isdigit()}
            for t in teachers:
                s = smap.get(t.id)
                if not s:
                    s = SubSetting(school_id=sid(), user_id=t.id)
                    db.session.add(s)
                s.can_manage = t.id in managers
                s.can_substitute = t.id in subs
            db.session.commit()
            flash('บันทึกสิทธิ์การจัดสอนแทนแล้ว', 'success')
            return redirect(url_for('subplan_settings'))
        return render_template('subplan/settings.html', teachers=teachers, smap=smap)

    return g
