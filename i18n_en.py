"""English version of KRURUKSORN (TH/EN switch).

How it works: every HTML/JSON response is translated just before it is sent,
when the viewer chose English. Templates and app.py stay in Thai, so other
people's (Codex's) work never conflicts with this module.

Dictionary: i18n/en.json  {"ข้อความไทย": "English"}
  - "{}" in a key matches any text (Jinja values / f-string values) and is
    put back into the English text at "{}" (or "{0}", "{1}" to reorder).
  - user data (names, subject names, notes) is not touched unless the whole
    text is a dictionary entry.
"""
import html
import json
import os
import re
from collections import OrderedDict

from flask import request, session, redirect, url_for

TH = re.compile('[฀-๿]')
TH_CHAR = re.compile('[฀-๿]')
WS = re.compile(r'\s+')
LANG_COOKIE = 'krs_lang'
# short words translated only when they are the whole text, never inside other text
FALLBACK_STOP = {'ที่', 'จาก', 'ถึง', 'และ', 'หรือ', 'ของ', 'ใน', 'ได้', 'ไม่', 'มา', 'ลา', 'สาย', 'ขาด', 'ดี',
                 'ห้อง', 'ปี', 'วัน', 'ครั้ง', 'คน', 'เต็ม', 'รับ', 'ส่ง', 'ออก', 'สอน', 'งาน', 'ผล', 'ดู', 'ชั้น', 'ยาก',
                 'ง่าย', 'สูง', 'ปิด', 'เปิด', 'ถูก', 'ผิด', 'หยุด', 'กิจ', 'ห้าม', 'เลี่ยง', 'หลีก', 'แล้ว', 'ทุก', 'เช้า',
                 'บ่าย', 'พัก', 'โดด', 'ตัด', 'ย้าย', 'ลบ', 'เพิ่ม', 'แสดง', 'ช่วง', 'ระยะ', 'เวลา', 'เดือน', 'ชนิด', 'ค่า',
                 'นาย', 'นาง', 'นางสาว', 'เด็กชาย', 'เด็กหญิง', 'ครู', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัส', 'พฤหัสบดี', 'ศุกร์',
                 'เสาร์', 'อาทิตย์', 'จ.', 'อ.', 'พ.', 'ศ.', 'ส.', 'อา.', 'พฤ.', 'ด.ช.', 'ด.ญ.', 'ปกติ', 'ชาย', 'หญิง', 'ใช้', 'สื่อ', 'วิชา', 'หน่วย', 'เรื่อง'}

THAI_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม',
               'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']
THAI_MONTHS_ABBR = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.']
EN_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
             'September', 'October', 'November', 'December']
EN_MONTHS_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
THAI_DAYS = {'วันจันทร์': 'Monday', 'วันอังคาร': 'Tuesday', 'วันพุธ': 'Wednesday', 'วันพฤหัสบดี': 'Thursday',
             'วันพฤหัส': 'Thursday', 'วันศุกร์': 'Friday', 'วันเสาร์': 'Saturday', 'วันอาทิตย์': 'Sunday',
             'จันทร์': 'Monday', 'อังคาร': 'Tuesday', 'พุธ': 'Wednesday', 'พฤหัสบดี': 'Thursday',
             'พฤหัส': 'Thursday', 'ศุกร์': 'Friday', 'เสาร์': 'Saturday', 'อาทิตย์': 'Sunday'}

SHORT_MONTHS = ['มกร', 'กุม', 'มีน', 'เมย', 'พฤษ', 'มิถ', 'กรก', 'สิง', 'กัน', 'ตุล', 'พฤศ', 'ธัน']
_month_alt = '|'.join(re.escape(m) for m in THAI_MONTHS + THAI_MONTHS_ABBR)
DATE_RE = re.compile(r'(?:(\d{1,2})\s*)?(' + _month_alt + r')(?:\s*(?:พ\.ศ\.\s*)?(\d{4}|\d{2})\b)?')
DAY_PREFIX_RE = re.compile(r'(วัน(?:จันทร์|อังคาร|พุธ|พฤหัสบดี|พฤหัส|ศุกร์|เสาร์|อาทิตย์))(?:\s*ที่)?')
SHORT_MONTH_RE = re.compile(r'(?<![\u0e00-\u0e7f])(มกร|กุม|มีน|เมย|พฤษ|มิถ|กรก|สิง|กัน|ตุล|พฤศ|ธัน)\s+(25\d\d)\b')
SYMBOL_PREFIX = re.compile('^[^\\w\u0e00-\u0e7f(\\[{"\u201c\'\u2018<]+')
WEEKDAY_DATE_RE = re.compile(r'(?<![\u0e00-\u0e7f])(จันทร์|อังคาร|พุธ|พฤหัสบดี|พฤหัส|ศุกร์|เสาร์|อาทิตย์|จ\.|อ\.|พ\.|พฤ\.|ศ\.|ส\.|อา\.)\s*(?=\d)|(?<![฀-๿])(พฤ|อา|จ|อ|พ|ศ)\s+(?=\d{1,2}\b)')
SHORT_DAYS = {'จันทร์': 'Mon', 'อังคาร': 'Tue', 'พุธ': 'Wed', 'พฤหัสบดี': 'Thu', 'พฤหัส': 'Thu', 'ศุกร์': 'Fri',
              'เสาร์': 'Sat', 'อาทิตย์': 'Sun', 'จ.': 'Mon', 'อ.': 'Tue', 'พ.': 'Wed', 'พฤ.': 'Thu', 'ศ.': 'Fri',
              'ส.': 'Sat', 'อา.': 'Sun', 'จ': 'Mon', 'อ': 'Tue', 'พ': 'Wed', 'พฤ': 'Thu', 'ศ': 'Fri', 'อา': 'Sun'}
CLASS_RE = re.compile(r'(?<![\u0e00-\u0e7f.])ม\.(?=\d)')
NAME_RE = re.compile(r'(?:นางสาว|นาง|นาย|เด็กชาย|เด็กหญิง|ด\.ช\.|ด\.ญ\.|ว่าที่ร้อยตรี|ว่าที่ ร\.ต\.)\s*[\u0e00-\u0e7fA-Za-z]+(?:\s+[\u0e00-\u0e7f]+)?')


def _fill(tr, vals):
    """Put captured values into the English text: "{}" in order, or "{0}" "{1}" by position."""
    if re.search(r'\{\d+\}', tr):
        return re.sub(r'\{(\d+)\}', lambda m: vals[int(m.group(1))] if int(m.group(1)) < len(vals) else '', tr)
    it = iter(vals)
    return re.sub(r'\{\}', lambda m: next(it, ''), tr)


def _norm(s):
    return WS.sub(' ', html.unescape(s)).strip()


class Translator:
    def __init__(self, path):
        self.path = path
        self.mtime = None
        self.exact = {}
        self.buckets = {}
        self.free = []
        self.fallback_re = None
        self.fallback = {}
        self.cache = OrderedDict()
        self.load()

    def _files(self):
        folder = os.path.dirname(self.path)
        names = ['en.json'] + sorted(n for n in os.listdir(folder) if n.startswith('en_') and n.endswith('.json')) \
            if os.path.isdir(folder) else []
        return [os.path.join(folder, n) for n in names if os.path.exists(os.path.join(folder, n))]

    def load(self):
        files = self._files()
        try:
            mtime = tuple(os.path.getmtime(f) for f in files)
        except OSError:
            return
        if not files or mtime == self.mtime:
            return
        raw = {}
        for fp in files:   # later files (en_*.json, school data) override en.json
            try:
                with open(fp, encoding='utf-8') as f:
                    raw.update(json.load(f))
            except (OSError, ValueError):
                continue
        exact, buckets, free, fallback = {}, {}, [], {}
        for k, v in raw.items():
            if not v or k.startswith('_'):
                continue
            key = _norm(k)
            if not TH.search(key):
                continue
            if '{}' in key:
                parts = key.split('{}')
                rx = '(.+?)'.join(re.escape(p) for p in parts)
                try:
                    pat = re.compile(rx + r'\Z', re.S)
                except re.error:
                    continue
                item = (pat, v, len(key) - 2 * key.count('{}'))
                head = parts[0][:2]
                if len(parts[0]) >= 2:
                    buckets.setdefault(head, []).append(item)
                else:
                    free.append(item)
            else:
                exact[key] = v
                if len(key) >= 3 and key not in FALLBACK_STOP:
                    fallback[key] = v
        for lst in list(buckets.values()) + [free]:
            lst.sort(key=lambda x: -x[2])
        keys = sorted(fallback, key=len, reverse=True)
        self.exact, self.buckets, self.free, self.fallback = exact, buckets, free, fallback
        self.fallback_re = re.compile('|'.join(re.escape(k) for k in keys)) if keys else None
        self.cache.clear()
        self.mtime = mtime

    # ---- core ---------------------------------------------------------
    def text(self, s, depth=0):
        """Translate one plain (unescaped) text. Returns the same string if nothing known."""
        if not s or not TH.search(s):
            return s
        hit = self.cache.get(s)
        if hit is not None:
            return hit
        lead = s[:len(s) - len(s.lstrip())]
        trail = s[len(s.rstrip()):]
        core = WS.sub(' ', s).strip()
        out = self._core(core, depth)
        res = lead + out + trail if out is not None else s
        if len(self.cache) > 60000:
            self.cache.popitem(last=False)
        self.cache[s] = res
        return res

    def _core(self, core, depth):
        v = self.exact.get(core)
        if v is not None:
            return v
        # "🗓️ ภาคเรียนที่ 2/2569", "• รายการ" : translate after an emoji/symbol prefix
        m = SYMBOL_PREFIX.match(core)
        if m and m.end() < len(core):
            rest = self.exact.get(core[m.end():])
            if rest is None:
                rest = self._patterns(core[m.end():], depth)
            if rest is not None:
                return core[:m.end()] + rest
        v = self._patterns(core, depth)
        if v is not None:
            return v
        return self._fallback(core)

    def _patterns(self, core, depth):
        cands = self.buckets.get(core[:2], [])
        for pat, tr, _ in cands + self.free:
            m = pat.match(core)
            if m:
                vals = [self.text(g, depth + 1) if depth < 3 else g for g in m.groups()]
                return _fill(tr, vals)
        return None

    def _dates(self, s):
        def rep(m):
            day, mon, year = m.group(1), m.group(2), m.group(3)
            if mon in THAI_MONTHS:
                name = EN_MONTHS[THAI_MONTHS.index(mon)]
            else:
                name = EN_MONTHS_ABBR[THAI_MONTHS_ABBR.index(mon)]
            out = (day + ' ' if day else '') + name
            if year:
                y = int(year)
                if len(year) == 2:
                    y += 2500
                out += ' ' + str(y - 543 if y > 2400 else y)
            return ' ' + out + ' '
        s = DATE_RE.sub(rep, s)
        s = SHORT_MONTH_RE.sub(lambda m: EN_MONTHS_ABBR[SHORT_MONTHS.index(m.group(1))] + ' ' + str(int(m.group(2)) - 543), s)
        s = DAY_PREFIX_RE.sub(lambda m: ' ' + THAI_DAYS[m.group(1)] + ', ', s)
        s = WEEKDAY_DATE_RE.sub(lambda m: SHORT_DAYS[m.group(1) or m.group(2)] + ' ', s)
        s = s.replace('พ.ศ.', 'B.E.')
        return s

    def _fallback(self, core):
        s = self._dates(core)
        s = CLASS_RE.sub('M.', s)
        if self.fallback_re is not None and TH.search(s):
            protected = [m.span() for m in NAME_RE.finditer(s)]
            matches = [m for m in self.fallback_re.finditer(s)
                       if not any(a < m.end() and m.start() < b for a, b in protected)]
            ok = []
            for i, m in enumerate(matches):
                a, b = m.start(), m.end()
                left_ok = a == 0 or not TH_CHAR.match(s[a - 1]) or (i > 0 and matches[i - 1].end() == a)
                right_ok = b == len(s) or not TH_CHAR.match(s[b]) or (i + 1 < len(matches) and matches[i + 1].start() == b)
                if left_ok and right_ok:
                    ok.append(m)
            if ok:
                parts, pos = [], 0
                for m in ok:
                    parts.append(s[pos:m.start()])
                    parts.append('\x00' + self.fallback[m.group(0)] + '\x00')
                    pos = m.end()
                parts.append(s[pos:])
                s = ''.join(parts)
                # put spaces between joined words
                s = re.sub(r'(?<=[\w\)\.,:;!?%])\x00(?=[\w\(])', ' ', s)
                s = re.sub(r'(?<=[\w\)])\x00(?=[\w\(])', ' ', s)
                s = s.replace('\x00\x00', ' ').replace('\x00', '')
        s = WS.sub(' ', s).strip()
        s = re.sub(r'\s+([,.:;)])', r'\1', s)
        return s if s != core else None

    # ---- html ---------------------------------------------------------
    def html_text(self, raw):
        if not TH.search(raw):
            return raw
        plain = html.unescape(raw)
        out = self.text(plain)
        if out == plain:
            return raw
        return html.escape(out, quote=False)

    def js_literals(self, code):
        def rep(m):
            q, body = m.group(1), m.group(2)
            if not TH.search(body):
                return m.group(0)
            v = self.exact.get(_norm(body))
            if v is None:
                return m.group(0)
            v = v.replace(q, '\\' + q).replace('\n', '\\n')
            return q + v + q
        return JS_STR.sub(rep, code)

    def tag(self, t):
        if not TH.search(t):
            return t
        def rep_attr(m):
            name, q, val = m.group(1), m.group(2), m.group(3)
            low = name.lower()
            if low == 'value':
                if not re.search(r'type\s*=\s*["\']?(submit|button|reset)', t, re.I):
                    return m.group(0)
            if low.startswith('on'):
                nv = self.js_literals(html.unescape(val))
                return f'{name}={q}{html.escape(nv, quote=True)}{q}' if nv != html.unescape(val) else m.group(0)
            plain = html.unescape(val)
            nv = self.text(plain)
            if nv == plain:
                return m.group(0)
            return f'{name}={q}{html.escape(nv, quote=True)}{q}'
        return ATTR_RE.sub(rep_attr, t)

    def page(self, body):
        if not TH.search(body):
            return body
        pieces = TOKEN_RE.split(body)
        out = []
        skip = None          # 'style' | 'textarea' | 'script'
        pending_option = None  # index in out of <option> tag without value
        for p in pieces:
            if not p:
                continue
            if p[0] == '<' and TAG_RE.match(p):
                low = p[:12].lower()
                if skip:
                    if low.startswith('</' + skip):
                        skip = None
                    out.append(p)
                    continue
                if low.startswith('<script'):
                    skip = 'script'
                elif low.startswith('<style'):
                    skip = 'style'
                elif low.startswith('<textarea'):
                    skip = 'textarea'
                if low.startswith('<option') and not re.search(r'\svalue\s*=', p, re.I):
                    pending_option = len(out)
                elif not low.startswith('</option'):
                    pass
                if low.startswith('</option') or low.startswith('</select'):
                    pending_option = None
                if low.startswith('<html'):
                    p = re.sub(r'lang\s*=\s*"th"', 'lang="en"', p)
                out.append(self.tag(p) if p[1] != '!' and p[1] != '/' else p)
                if low.startswith('<option') and pending_option is not None and pending_option != len(out) - 1:
                    pending_option = None
                continue
            if skip == 'script':
                out.append(self.js_literals(p))
                continue
            if skip:
                out.append(p)
                continue
            t = self.html_text(p)
            if t != p and pending_option is not None:
                orig = html.unescape(p).strip()
                tagp = out[pending_option]
                out[pending_option] = tagp[:-1].rstrip('/') + f' value="{html.escape(orig, quote=True)}">'
            pending_option = None
            out.append(t)
        return ''.join(out)


TOKEN_RE = re.compile(r'(<!--.*?-->|<[A-Za-z/!][^<>]*>)', re.S)
TAG_RE = re.compile(r'<[A-Za-z/!]')
ATTR_RE = re.compile(r'''\b(placeholder|title|data-confirm|alt|aria-label|value|label|data-title|on[a-z]+)\s*=\s*(["'])(.*?)\2''', re.I | re.S)
JS_STR = re.compile(r'''(['"])((?:(?!\1)[^\\\n]|\\.)*)\1''')

TOGGLE_CSS = ('<style>.krs-lang{display:inline-flex;gap:2px;background:rgba(15,23,42,.08);border-radius:999px;padding:3px;'
              'font:700 12px/1 system-ui,sans-serif;z-index:9999}.krs-lang a{padding:6px 10px;border-radius:999px;text-decoration:none;'
              'color:#334155}.krs-lang a.on{background:#fff;color:#0f172a;box-shadow:0 1px 3px rgba(0,0,0,.18)}'
              '.krs-lang.desk{position:fixed;top:10px;right:14px;background:rgba(255,255,255,.92);box-shadow:0 2px 10px rgba(15,23,42,.15)}'
              '.krs-lang.desk a.on{background:#0f766e;color:#fff}.krs-lang.mob{margin-left:8px;flex:none}'
              '.krs-lang.float{position:fixed;right:10px;bottom:10px;background:#1e293b}.krs-lang.float a{color:#cbd5e1}'
              '.krs-lang.float a.on{background:#fff;color:#111}'
              '@media(max-width:900px){.krs-lang.desk{display:none}}@media(min-width:901px){.krs-lang.mob{display:none}}'
              '@media print{.krs-lang{display:none!important}}</style>')


def toggle_html(lang, kind='desk'):
    nxt = request.full_path if request.query_string else request.path
    th = url_for('set_language', code='th', next=nxt)
    en = url_for('set_language', code='en', next=nxt)
    return (f'<span class="krs-lang {kind}" title="Language">'
            f'<a href="{html.escape(th)}" class="{"on" if lang == "th" else ""}">TH</a>'
            f'<a href="{html.escape(en)}" class="{"on" if lang == "en" else ""}">EN</a></span>')


def register(app, g=None):
    path = os.path.join(app.root_path, 'i18n', 'en.json')
    tr = Translator(path)
    app.extensions['krs_i18n'] = tr

    def current_lang():
        lang = session.get('lang') or request.cookies.get(LANG_COOKIE) or os.environ.get('DEFAULT_LANG', 'th')
        return 'en' if lang == 'en' else 'th'

    @app.route('/lang/<code>')
    def set_language(code):
        code = 'en' if code == 'en' else 'th'
        session['lang'] = code
        nxt = request.args.get('next') or request.referrer or '/'
        if not nxt.startswith('/') or nxt.startswith('//'):
            nxt = '/'
        resp = redirect(nxt)
        resp.set_cookie(LANG_COOKIE, code, max_age=365 * 24 * 3600, samesite='Lax')
        return resp

    @app.context_processor
    def _i18n_ctx():
        return {'current_lang': current_lang()}

    @app.after_request
    def _i18n_after(resp):
        try:
            if resp.direct_passthrough or resp.status_code >= 300 and resp.status_code < 400:
                return resp
            ctype = (resp.mimetype or '')
            lang = current_lang()
            if ctype == 'text/html':
                body = resp.get_data(as_text=True)
                if '<html' not in body[:500].lower() and '<!doctype' not in body[:200].lower():
                    if lang == 'en':
                        tr.load()
                        resp.set_data(tr.page(body))
                    return resp
                if lang == 'en':
                    tr.load()
                    body = tr.page(body)
                # language switch: in the top bar, or floating on standalone pages
                m = re.search(r'(<header class="topbar".*?)(</header>)', body, re.S)
                if m:
                    body = body[:m.start(2)] + toggle_html(lang, 'mob') + body[m.start(2):]
                    extra = TOGGLE_CSS + toggle_html(lang, 'desk')
                else:
                    extra = TOGGLE_CSS + toggle_html(lang, 'float')
                body = re.sub(r'(</body>)', lambda mm: extra + mm.group(1), body, count=1)
                resp.set_data(body)
            elif ctype == 'application/json' and lang == 'en':
                data = resp.get_json(silent=True)
                if isinstance(data, dict):
                    tr.load()
                    changed = False
                    for k in ('message', 'msg', 'error', 'detail', 'status_text', 'label'):
                        if isinstance(data.get(k), str) and TH.search(data[k]):
                            data[k] = tr.text(data[k]); changed = True
                    if changed:
                        resp.set_data(json.dumps(data, ensure_ascii=False))
        except Exception as exc:  # never break a page because of translation
            app.logger.warning('i18n: %s', exc)
        return resp

    return tr
