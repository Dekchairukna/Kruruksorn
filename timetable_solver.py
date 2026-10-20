"""ตัวจัดตารางสอนอัตโนมัติ (ไม่ขึ้นกับ Flask ทดสอบแยกได้)

แนวคิด
- แต่ละรายวิชา-ห้อง (lesson) มีจำนวนคาบ/สัปดาห์ และ "คาบติดกัน" (block) เช่น 2 = เรียนทีละ 2 คาบติดกัน
- แตกเป็นก้อน (chunk) แล้ววางลงช่อง (วัน, คาบเริ่ม) โดยก้อนไม่ข้ามพักกลางวัน
- เงื่อนไขบังคับ (hard): ครูไม่สอนซ้อน, ห้อง/ชั้นไม่เรียนซ้อน, ห้ามวางในช่องที่ห้ามเด็ดขาด
- เงื่อนไขพยายาม (soft): เกลี่ยคาบของครูให้เท่ากันทุกวัน, ไม่เกินคาบสูงสุด/วัน, วิชาเดียวกันไม่ซ้ำวัน,
  ไม่สอนติดกันยาวเกินไป, หลีกเลี่ยงช่วงที่ครูไม่สะดวก, เกลี่ยคาบของห้องเรียนในแต่ละวัน
- ใช้ greedy วางเริ่มต้น แล้วปรับด้วย simulated annealing ภายในเวลาที่กำหนด
"""
import math
import random
import time
from collections import defaultdict

HARD = 1000          # ชนกัน/ผิดข้อห้ามเด็ดขาด ต่อคาบ
UNPLACED = 3000      # วางไม่ได้ ต่อคาบ
SAME_DAY = 40        # วิชาเดียวกัน (ห้องเดียวกัน) มากกว่า 1 ก้อนในวันเดียว
OVER_MAX_DAY = 60    # ครูสอนเกินคาบสูงสุด/วัน ต่อคาบที่เกิน
TEACHER_BALANCE = 3  # ถ่วงน้ำหนัก sum(load^2) ของครูรายวัน (ยิ่งเท่ากันยิ่งน้อย)
GROUP_BALANCE = 1    # ถ่วงน้ำหนักการกระจายคาบของห้องเรียนรายวัน
LONG_RUN = 20        # ครูสอนติดกันเกินที่กำหนด ต่อคาบที่เกิน
MOVED = 25           # โหมดปรับจากตารางเดิม: ย้ายคาบออกจากตำแหน่งเดิม ต่อคาบ


def split_chunks(n, block):
    block = max(1, int(block or 1))
    out = []
    while n > 0:
        size = min(block, n)
        out.append(size)
        n -= size
    return out


class Solver:
    def __init__(self, problem, seed=None):
        self.P = problem
        self.rng = random.Random(seed)
        self.days = int(problem.get('days', 5))
        self.periods = int(problem.get('periods', 8))
        self.lunch_after = int(problem.get('lunch_after', 4))
        self.max_run = int(problem.get('max_consecutive', 4))
        self.teachers = problem.get('teachers', {})
        self.lessons = {l['id']: l for l in problem['lessons']}

        # ข้อห้าม: cost ต่อ (scope,ref,d,p)
        self.t_block = defaultdict(int)
        self.g_block = defaultdict(int)
        self.a_block = defaultdict(int)
        for b in problem.get('blocks', []):
            w = HARD if b.get('hard') else int(b.get('weight') or 30)
            d, p = int(b['day']), int(b['period'])
            if b['scope'] == 'teacher':
                self.t_block[(b['ref'], d, p)] = max(self.t_block[(b['ref'], d, p)], w)
            elif b['scope'] == 'group':
                self.g_block[(b['ref'], d, p)] = max(self.g_block[(b['ref'], d, p)], w)
            else:
                self.a_block[(d, p)] = max(self.a_block[(d, p)], w)

        # ล็อกไว้
        locked_by_lesson = defaultdict(list)
        for s in problem.get('locked', []):
            if s['lesson_id'] in self.lessons:
                locked_by_lesson[s['lesson_id']].append((int(s['day']), int(s['period'])))

        # chunks: dict id -> {lesson, size, fixed, pos}
        self.chunks = []
        for lid, l in self.lessons.items():
            fixed = locked_by_lesson.get(lid, [])
            for (d, p) in fixed:
                self.chunks.append({'lesson': lid, 'size': 1, 'fixed': True, 'pos': (d, p)})
            remain = max(0, int(l['periods']) - len(fixed))
            for size in split_chunks(remain, l.get('block', 1)):
                self.chunks.append({'lesson': lid, 'size': size, 'fixed': False, 'pos': None})

        self.t_occ = defaultdict(int)   # (t,d,p) -> count
        self.g_occ = defaultdict(int)   # (g,d,p) -> count
        self.l_day = defaultdict(int)   # (lesson,d) -> chunks count
        for c in self.chunks:
            if c['fixed']:
                self._apply(c, c['pos'], +1)

        self._pos_cache = {}
        self.prefer = {}  # lesson -> set((d,p)) ตำแหน่งเดิมที่อยากคงไว้

    # ---------------- geometry ----------------
    def starts_for(self, size):
        key = size
        if key in self._pos_cache:
            return self._pos_cache[key]
        halves = [(1, self.lunch_after), (self.lunch_after + 1, self.periods)]
        out = []
        for d in range(1, self.days + 1):
            fit = False
            for lo, hi in halves:
                for s in range(lo, hi - size + 2):
                    out.append((d, s))
                    fit = True
            if not fit:  # ก้อนใหญ่เกินครึ่งวัน: ยอมให้ข้ามพักกลางวัน
                for s in range(1, self.periods - size + 2):
                    out.append((d, s))
        self._pos_cache[key] = out
        return out

    def cells(self, c, pos):
        d, s = pos
        return [(d, s + i) for i in range(c['size'])]

    # ---------------- bookkeeping ----------------
    def _apply(self, c, pos, sign):
        l = self.lessons[c['lesson']]
        t = l['teacher_id']
        for (d, p) in self.cells(c, pos):
            self.t_occ[(t, d, p)] += sign
            for g in l['groups']:
                self.g_occ[(g, d, p)] += sign
        self.l_day[(c['lesson'], pos[0])] += sign

    # ---------------- costs ----------------
    def pos_cost(self, c, pos):
        if pos is None:
            return UNPLACED * c['size']
        l = self.lessons[c['lesson']]
        t = l['teacher_id']
        cost = 0
        pref = self.prefer.get(c['lesson'])
        for (d, p) in self.cells(c, pos):
            cost += self.t_block.get((t, d, p), 0) + self.a_block.get((d, p), 0)
            if pref is not None and (d, p) not in pref:
                cost += MOVED
            for g in l['groups']:
                cost += self.g_block.get((g, d, p), 0)
        return cost

    def teacher_cost(self, t):
        maxpd = (self.teachers.get(t) or {}).get('max_per_day') or 0
        cost = 0
        for d in range(1, self.days + 1):
            load, run = 0, 0
            for p in range(1, self.periods + 1):
                n = self.t_occ.get((t, d, p), 0)
                if n > 1:
                    cost += HARD * (n - 1)
                if n:
                    load += 1
                    run = 0 if p == self.lunch_after + 1 else run  # พักกลางวันตัดช่วงติดกัน
                    run += 1
                    if run > self.max_run:
                        cost += LONG_RUN
                else:
                    run = 0
            cost += TEACHER_BALANCE * load * load
            if maxpd and load > maxpd:
                cost += OVER_MAX_DAY * (load - maxpd)
        return cost

    def group_cost(self, g):
        cost = 0
        for d in range(1, self.days + 1):
            load = 0
            for p in range(1, self.periods + 1):
                n = self.g_occ.get((g, d, p), 0)
                if n > 1:
                    cost += HARD * (n - 1)
                load += 1 if n else 0
            cost += GROUP_BALANCE * load * load
        return cost

    def lesson_cost(self, lid):
        cost = 0
        for d in range(1, self.days + 1):
            n = self.l_day.get((lid, d), 0)
            if n > 1:
                cost += SAME_DAY * (n - 1)
        return cost

    def local_cost(self, c):
        l = self.lessons[c['lesson']]
        return (self.teacher_cost(l['teacher_id']) + sum(self.group_cost(g) for g in l['groups'])
                + self.lesson_cost(c['lesson']))

    def move_delta(self, c, new_pos):
        old_pos = c['pos']
        before = self.local_cost(c) + self.pos_cost(c, old_pos)
        if old_pos is not None:
            self._apply(c, old_pos, -1)
        if new_pos is not None:
            self._apply(c, new_pos, +1)
        after = self.local_cost(c) + self.pos_cost(c, new_pos)
        # คืนค่าเดิม
        if new_pos is not None:
            self._apply(c, new_pos, -1)
        if old_pos is not None:
            self._apply(c, old_pos, +1)
        return after - before

    def do_move(self, c, new_pos):
        if c['pos'] is not None:
            self._apply(c, c['pos'], -1)
        if new_pos is not None:
            self._apply(c, new_pos, +1)
        c['pos'] = new_pos

    def total_cost(self):
        tids = {l['teacher_id'] for l in self.lessons.values()}
        gids = {g for l in self.lessons.values() for g in l['groups']}
        return (sum(self.pos_cost(c, c['pos']) for c in self.chunks)
                + sum(self.teacher_cost(t) for t in tids)
                + sum(self.group_cost(g) for g in gids)
                + sum(self.lesson_cost(lid) for lid in self.lessons))

    # ---------------- search ----------------
    def seed_from(self, initial):
        """วางตามตารางเดิม (ถ้ามี) เพื่อเริ่มจากของที่ใช้อยู่"""
        by_lesson = defaultdict(set)
        for s in initial or []:
            by_lesson[s['lesson_id']].add((int(s['day']), int(s['period'])))
        for c in sorted(self.chunks, key=lambda c: -c['size']):
            if c['fixed'] or c['pos'] is not None:
                continue
            cells = by_lesson.get(c['lesson'])
            if not cells:
                continue
            for pos in self.starts_for(c['size']):
                want = self.cells(c, pos)
                if all(x in cells for x in want):
                    self.do_move(c, pos)
                    for x in want:
                        cells.discard(x)
                    break

    def greedy(self):
        todo = [c for c in self.chunks if not c['fixed'] and c['pos'] is None]
        tload = defaultdict(int)
        for l in self.lessons.values():
            tload[l['teacher_id']] += int(l['periods'])
        self.rng.shuffle(todo)
        todo.sort(key=lambda c: (-c['size'], -len(self.lessons[c['lesson']]['groups']),
                                 -tload[self.lessons[c['lesson']]['teacher_id']]))
        for c in todo:
            best, best_d = None, None
            for pos in self.starts_for(c['size']):
                dlt = self.move_delta(c, pos) + self.rng.random()
                if best_d is None or dlt < best_d:
                    best, best_d = pos, dlt
            self.do_move(c, best)

    def anneal(self, time_limit=15.0, t0=60.0, t1=0.3):
        movable = [c for c in self.chunks if not c['fixed']]
        if not movable:
            return
        cur = self.total_cost()
        best = cur
        best_state = [c['pos'] for c in self.chunks]
        start = time.time()
        it = 0
        while True:
            el = time.time() - start
            if el >= time_limit:
                break
            temp = t0 * (t1 / t0) ** (el / time_limit)
            it += 1
            c = self.rng.choice(movable)
            r = self.rng.random()
            if r < 0.15 and c['pos'] is not None:
                # สลับกับก้อนอื่นของครูคนเดียวกันที่ขนาดเท่ากัน
                t = self.lessons[c['lesson']]['teacher_id']
                cand = [o for o in movable if o is not c and o['size'] == c['size'] and o['pos'] is not None
                        and self.lessons[o['lesson']]['teacher_id'] == t]
                if cand:
                    o = self.rng.choice(cand)
                    a, b = c['pos'], o['pos']
                    dswap = self._swap_delta(c, o)
                    if dswap <= 0 or self.rng.random() < math.exp(-dswap / max(temp, 1e-6)):
                        self.do_move(c, None); self.do_move(o, None)
                        self.do_move(c, b); self.do_move(o, a)
                        cur += dswap
                        if cur < best:
                            best = cur
                            best_state = [x['pos'] for x in self.chunks]
                    continue
            positions = self.starts_for(c['size'])
            if r < 0.6:
                pos = self.rng.choice(positions)
            else:
                sample = self.rng.sample(positions, min(8, len(positions)))
                pos = min(sample, key=lambda p: self.move_delta(c, p))
            if pos == c['pos']:
                continue
            dlt = self.move_delta(c, pos)
            if dlt <= 0 or self.rng.random() < math.exp(-dlt / max(temp, 1e-6)):
                self.do_move(c, pos)
                cur += dlt
                if cur < best:
                    best = cur
                    best_state = [x['pos'] for x in self.chunks]
        for c, pos in zip(self.chunks, best_state):
            if not c['fixed'] and c['pos'] != pos:
                self.do_move(c, pos)
        self.iterations = it

    def _swap_delta(self, c, o):
        a, b = c['pos'], o['pos']
        ents_before = self._entities_cost([c, o])
        self.do_move(c, None); self.do_move(o, None)
        self.do_move(c, b); self.do_move(o, a)
        ents_after = self._entities_cost([c, o])
        pos_after = self.pos_cost(c, b) + self.pos_cost(o, a)
        self.do_move(c, None); self.do_move(o, None)
        self.do_move(c, a); self.do_move(o, b)
        return (ents_after + pos_after) - (ents_before + self.pos_cost(c, a) + self.pos_cost(o, b))

    def _entities_cost(self, chunks):
        ts, gs, ls = set(), set(), set()
        for c in chunks:
            l = self.lessons[c['lesson']]
            ts.add(l['teacher_id']); gs.update(l['groups']); ls.add(c['lesson'])
        return (sum(self.teacher_cost(t) for t in ts) + sum(self.group_cost(g) for g in gs)
                + sum(self.lesson_cost(x) for x in ls))

    # ---------------- result ----------------
    def result(self):
        placements, unplaced = [], []
        for c in self.chunks:
            if c['pos'] is None:
                unplaced.append({'lesson_id': c['lesson'], 'size': c['size']})
                continue
            for (d, p) in self.cells(c, c['pos']):
                placements.append({'lesson_id': c['lesson'], 'day': d, 'period': p, 'locked': c['fixed']})
        return {'placements': placements, 'unplaced': unplaced, 'cost': self.total_cost(),
                'iterations': getattr(self, 'iterations', 0)}


def solve(problem, time_limit=15.0, seed=None, initial=None, keep_initial=False):
    """จัดตาราง คืน dict(placements, unplaced, cost, iterations)
    initial = ตารางเดิมสำหรับเริ่มต้น, keep_initial=True จะเริ่มจากตารางเดิมแล้วปรับเฉพาะส่วนที่มีปัญหา"""
    s = Solver(problem, seed=seed)
    if initial:
        if keep_initial:
            for x in initial:
                s.prefer.setdefault(x['lesson_id'], set()).add((int(x['day']), int(x['period'])))
        s.seed_from(initial)
    s.greedy()
    s.anneal(time_limit=time_limit, t0=(8.0 if keep_initial and initial else 60.0))
    return s.result()


def evaluate(problem, placements):
    """ตรวจตารางที่มีอยู่: คืนรายการปัญหา (ชนกัน/ผิดข้อห้าม/เกินคาบต่อวัน) สำหรับแสดงบนหน้าเว็บ"""
    lessons = {l['id']: l for l in problem['lessons']}
    teachers = problem.get('teachers', {})
    t_cells, g_cells = defaultdict(list), defaultdict(list)
    for s in placements:
        l = lessons.get(s['lesson_id'])
        if not l:
            continue
        t_cells[(l['teacher_id'], s['day'], s['period'])].append(s['lesson_id'])
        for g in l['groups']:
            g_cells[(g, s['day'], s['period'])].append(s['lesson_id'])
    issues = []
    for (t, d, p), ls in t_cells.items():
        if len(ls) > 1:
            issues.append({'type': 'teacher_clash', 'teacher_id': t, 'day': d, 'period': p, 'lessons': ls})
    for (g, d, p), ls in g_cells.items():
        if len(set(ls)) > 1:
            issues.append({'type': 'group_clash', 'group': g, 'day': d, 'period': p, 'lessons': sorted(set(ls))})
    hard_t = {(b['ref'], int(b['day']), int(b['period'])): b for b in problem.get('blocks', []) if b['scope'] == 'teacher'}
    hard_g = {(b['ref'], int(b['day']), int(b['period'])): b for b in problem.get('blocks', []) if b['scope'] == 'group'}
    hard_a = {(int(b['day']), int(b['period'])): b for b in problem.get('blocks', []) if b['scope'] == 'all'}
    for s in placements:
        l = lessons.get(s['lesson_id'])
        if not l:
            continue
        key = (s['day'], s['period'])
        for b in [hard_t.get((l['teacher_id'],) + key), hard_a.get(key)] + [hard_g.get((g,) + key) for g in l['groups']]:
            if b:
                issues.append({'type': 'blocked_hard' if b.get('hard') else 'blocked_soft', 'lesson_id': s['lesson_id'],
                               'day': s['day'], 'period': s['period'], 'reason': b.get('reason', '')})
    loads = defaultdict(int)
    for (t, d, p), ls in t_cells.items():
        loads[(t, d)] += 1
    for (t, d), n in loads.items():
        mx = (teachers.get(t) or {}).get('max_per_day') or 0
        if mx and n > mx:
            issues.append({'type': 'over_max', 'teacher_id': t, 'day': d, 'load': n, 'max': mx})
    return issues
