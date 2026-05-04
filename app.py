import os
from datetime import datetime, date, time, timedelta
import calendar as py_calendar
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from types import SimpleNamespace

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, 'instance')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
os.makedirs(INSTANCE_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'krurakson-dev')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', f"sqlite:///{os.path.join(INSTANCE_DIR, 'krurakson.db')}")
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = UPLOAD_DIR
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB รองรับวิดีโอบทเรียน
ALLOWED_WORKSHEET_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx', 'ppt', 'pptx', 'mp4', 'webm', 'mov'}

def allowed_worksheet_file(filename):
    return filename and '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_WORKSHEET_EXTENSIONS

def save_uploaded_file(file_obj, subdir='worksheets'):
    if not file_obj or not file_obj.filename:
        return '', ''
    if not allowed_worksheet_file(file_obj.filename):
        raise ValueError('รองรับเฉพาะไฟล์ PDF, PNG, JPG, Word, PowerPoint, MP4, WebM, MOV')
    target_dir = os.path.join(app.config['UPLOAD_FOLDER'], subdir)
    os.makedirs(target_dir, exist_ok=True)
    original = file_obj.filename
    safe = secure_filename(original) or 'upload'
    stamp = datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
    filename = f"{stamp}_{safe}"
    file_obj.save(os.path.join(target_dir, filename))
    return f"{subdir}/{filename}", original


def detect_lesson_file_type(file_path):
    ext = (file_path or '').rsplit('.', 1)[-1].lower() if '.' in (file_path or '') else ''
    if ext in {'png', 'jpg', 'jpeg'}:
        return 'image'
    if ext == 'pdf':
        return 'pdf'
    if ext in {'doc', 'docx'}:
        return 'document'
    if ext in {'ppt', 'pptx'}:
        return 'presentation'
    if ext in {'mp4', 'webm', 'mov'}:
        return 'video'
    return 'file'


def youtube_embed_url(url):
    """แปลงลิงก์ YouTube ให้ฝังในหน้าบทเรียนได้"""
    if not url:
        return ''
    u = url.strip()
    if 'youtube.com/embed/' in u:
        return u
    if 'youtu.be/' in u:
        vid = u.split('youtu.be/', 1)[1].split('?', 1)[0].split('&', 1)[0].strip('/')
        return f'https://www.youtube.com/embed/{vid}' if vid else ''
    if 'youtube.com/watch' in u and 'v=' in u:
        vid = u.split('v=', 1)[1].split('&', 1)[0]
        return f'https://www.youtube.com/embed/{vid}' if vid else ''
    if 'youtube.com/shorts/' in u:
        vid = u.split('youtube.com/shorts/', 1)[1].split('?', 1)[0].split('&', 1)[0].strip('/')
        return f'https://www.youtube.com/embed/{vid}' if vid else ''
    return ''

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='student')  # admin teacher student
    citizen_id = db.Column(db.String(20))
    birth_date = db.Column(db.String(20))
    must_change_password = db.Column(db.Boolean, default=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(str(password))

    def check_password(self, password):
        return check_password_hash(self.password_hash, str(password))


class Semester(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    academic_year = db.Column(db.String(20), nullable=False, default='2569')
    term_no = db.Column(db.String(10), nullable=False, default='1')
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    is_active = db.Column(db.Boolean, default=False)

class CalendarImage(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    semester_id = db.Column(db.Integer, db.ForeignKey('semester.id'))
    file_path = db.Column(db.String(500), nullable=False)
    original_name = db.Column(db.String(255), default='')
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    semester = db.relationship('Semester')

class Classroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    is_active = db.Column(db.Boolean, default=True)
    teacher = db.relationship('User')

class TeacherClassroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    teacher = db.relationship('User')
    classroom = db.relationship('Classroom')

class ClassroomStudent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    classroom = db.relationship('Classroom')
    student = db.relationship('User')

class Subject(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    credit = db.Column(db.Float, default=1.0)
    total_periods = db.Column(db.Integer, default=40)
    is_active = db.Column(db.Boolean, default=True)
    teacher = db.relationship('User')

class TeacherSubject(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    teacher = db.relationship('User')
    subject = db.relationship('Subject')

class SubjectClassroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    subject = db.relationship('Subject')
    classroom = db.relationship('Classroom')

class Unit(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    indicators = db.Column(db.Text, default='')
    required_periods = db.Column(db.Integer, default=1)
    subject = db.relationship('Subject')

class Lesson(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    unit_id = db.Column(db.Integer, db.ForeignKey('unit.id'), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    objective = db.Column(db.Text, default='')
    content = db.Column(db.Text, default='')
    media_url = db.Column(db.String(500), default='')
    required_minutes = db.Column(db.Integer, default=50)
    unit = db.relationship('Unit')


class LessonFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lesson_id = db.Column(db.Integer, db.ForeignKey('lesson.id'), nullable=False)
    file_path = db.Column(db.String(500), default='')
    original_file_name = db.Column(db.String(255), default='')
    file_type = db.Column(db.String(50), default='file')  # image, pdf, document, presentation, file
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    lesson = db.relationship('Lesson')

class Worksheet(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lesson_id = db.Column(db.Integer, db.ForeignKey('lesson.id'), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    worksheet_type = db.Column(db.String(50), default='academic') # academic, petanque_score
    description = db.Column(db.Text, default='')
    file_path = db.Column(db.String(500), default='')
    original_file_name = db.Column(db.String(255), default='')
    lesson = db.relationship('Lesson')

class WorksheetQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    worksheet_id = db.Column(db.Integer, db.ForeignKey('worksheet.id'), nullable=False)
    number = db.Column(db.Integer, default=1)
    question_text = db.Column(db.Text, nullable=False)
    answer_type = db.Column(db.String(50), default='text')
    max_score = db.Column(db.Float, default=1)
    worksheet = db.relationship('Worksheet')

class Quiz(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lesson_id = db.Column(db.Integer, db.ForeignKey('lesson.id'), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    pass_percent = db.Column(db.Integer, default=60)
    lesson = db.relationship('Lesson')

class QuizQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    quiz_id = db.Column(db.Integer, db.ForeignKey('quiz.id'), nullable=False)
    question_text = db.Column(db.Text, nullable=False)
    question_type = db.Column(db.String(30), default='choice')
    choices = db.Column(db.Text, default='')  # one choice per line
    correct_answer = db.Column(db.Text, default='')
    score = db.Column(db.Float, default=1)
    quiz = db.relationship('Quiz')

class Assignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    lesson_id = db.Column(db.Integer, db.ForeignKey('lesson.id'), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    due_date = db.Column(db.Date)
    assignment_type = db.Column(db.String(30), default='special')  # special=งานสั่งพิเศษ, lesson=บทเรียนประจำคาบ
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    teacher = db.relationship('User')
    teacher = db.relationship('User')
    subject = db.relationship('Subject')
    classroom = db.relationship('Classroom')
    lesson = db.relationship('Lesson')

class AssignmentStatus(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    status = db.Column(db.String(50), default='งานใหม่')
    lesson_viewed = db.Column(db.Boolean, default=False)
    worksheet_submitted = db.Column(db.Boolean, default=False)
    quiz_submitted = db.Column(db.Boolean, default=False)
    quiz_score = db.Column(db.Float, default=0)
    total_score = db.Column(db.Float, default=0)
    submitted_at = db.Column(db.DateTime)
    assignment = db.relationship('Assignment')
    student = db.relationship('User')

class WorksheetAnswer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('worksheet_question.id'), nullable=False)
    answer_text = db.Column(db.Text, default='')
    file_path = db.Column(db.String(500), default='')
    original_file_name = db.Column(db.String(255), default='')
    score = db.Column(db.Float, default=0)
    question = db.relationship('WorksheetQuestion')

class QuizAnswer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('quiz_question.id'), nullable=False)
    answer_text = db.Column(db.Text, default='')
    is_correct = db.Column(db.Boolean, default=False)
    score = db.Column(db.Float, default=0)

class TeachingSchedule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    weekday = db.Column(db.Integer, nullable=False) # 0=Mon
    period_no = db.Column(db.Integer, nullable=False)
    start_time = db.Column(db.String(5), nullable=False)
    end_time = db.Column(db.String(5), nullable=False)
    room_name = db.Column(db.String(100), default='')
    topic = db.Column(db.String(255), default='')
    lesson_id = db.Column(db.Integer, db.ForeignKey('lesson.id'))
    subject = db.relationship('Subject')
    classroom = db.relationship('Classroom')
    lesson = db.relationship('Lesson')



class CalendarEvent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    event_date = db.Column(db.Date, nullable=False)
    title = db.Column(db.String(255), nullable=False)
    event_type = db.Column(db.String(50), default='กิจกรรมโรงเรียน')
    note = db.Column(db.Text, default='')
    teacher = db.relationship('User')

class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    classroom_id = db.Column(db.Integer, db.ForeignKey('classroom.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(20), default='มา') # มา ขาด ลา มาสาย กิจกรรม
    subject = db.relationship('Subject')
    classroom = db.relationship('Classroom')
    student = db.relationship('User')

class GradeSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    worksheet_weight = db.Column(db.Float, default=30)
    quiz_weight = db.Column(db.Float, default=20)
    attendance_weight = db.Column(db.Float, default=10)
    midterm_weight = db.Column(db.Float, default=20)
    final_weight = db.Column(db.Float, default=20)
    subject = db.relationship('Subject')

class ManualScore(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    midterm = db.Column(db.Float, default=0)
    final = db.Column(db.Float, default=0)
    behavior = db.Column(db.Float, default=0)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def role_required(*roles):
    def deco(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated or current_user.role not in roles:
                flash('ไม่มีสิทธิ์เข้าถึงหน้านี้', 'danger')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return wrapper
    return deco

def teacher_subject_ids():
    if current_user.role == 'admin':
        return [x.id for x in Subject.query.filter_by(is_active=True).all()]
    ids = {x.subject_id for x in TeacherSubject.query.filter_by(teacher_id=current_user.id).all()}
    ids.update(x.id for x in Subject.query.filter_by(teacher_id=current_user.id, is_active=True).all())
    return list(ids)

def teacher_classroom_ids():
    if current_user.role == 'admin':
        return [x.id for x in Classroom.query.filter_by(is_active=True).all()]
    ids = {x.classroom_id for x in TeacherClassroom.query.filter_by(teacher_id=current_user.id).all()}
    ids.update(x.id for x in Classroom.query.filter_by(teacher_id=current_user.id, is_active=True).all())
    return list(ids)

def teacher_filter(model):
    if current_user.role == 'admin':
        return model.query
    if model.__name__ == 'Subject':
        return model.query.filter(model.is_active==True, model.id.in_(teacher_subject_ids() or [-1]))
    if model.__name__ == 'Classroom':
        return model.query.filter(model.is_active==True, model.id.in_(teacher_classroom_ids() or [-1]))
    return model.query.filter_by(teacher_id=current_user.id)


def owns_subject(subject):
    return current_user.role == 'admin' or subject.teacher_id == current_user.id or TeacherSubject.query.filter_by(teacher_id=current_user.id, subject_id=subject.id).first() is not None

def owns_classroom(room):
    return current_user.role == 'admin' or room.teacher_id == current_user.id or TeacherClassroom.query.filter_by(teacher_id=current_user.id, classroom_id=room.id).first() is not None

def owns_unit(unit):
    return owns_subject(unit.subject)

def owns_lesson(lesson):
    return owns_unit(lesson.unit)

def deny_redirect(endpoint='index'):
    flash('ไม่มีสิทธิ์จัดการข้อมูลนี้', 'danger')
    return redirect(url_for(endpoint))

def get_current_schedule():
    if not current_user.is_authenticated or current_user.role != 'teacher':
        return None, None
    now = datetime.now()
    weekday = now.weekday()
    schedules = TeachingSchedule.query.filter_by(teacher_id=current_user.id, weekday=weekday).order_by(TeachingSchedule.start_time).all()
    cur = None
    nxt = None
    now_str = now.strftime('%H:%M')
    for s in schedules:
        if s.start_time <= now_str <= s.end_time:
            cur = s
        if s.start_time > now_str and not nxt:
            nxt = s
    return cur, nxt



def subject_classrooms_for_user(subject_id=None):
    """คืนค่าห้องที่ผูกกับรายวิชา และกรองตามสิทธิ์ครู"""
    q = SubjectClassroom.query
    if subject_id:
        q = q.filter_by(subject_id=subject_id)
    links = q.all()
    rows = []
    allowed_rooms = set(teacher_classroom_ids()) if current_user.is_authenticated and current_user.role != 'admin' else None
    for link in links:
        if allowed_rooms is None or link.classroom_id in allowed_rooms:
            rows.append(link)
    return rows

def get_grade_setting(subject_id):
    setting = GradeSetting.query.filter_by(subject_id=subject_id).first()
    if not setting:
        setting = GradeSetting(subject_id=subject_id)
        db.session.add(setting)
        db.session.commit()
    return setting

def get_manual_score(subject_id, student_id):
    row = ManualScore.query.filter_by(subject_id=subject_id, student_id=student_id).first()
    if not row:
        row = ManualScore(subject_id=subject_id, student_id=student_id)
        db.session.add(row)
        db.session.flush()
    return row

def calculate_grade_row(subject, room, student):
    setting = get_grade_setting(subject.id)
    statuses = AssignmentStatus.query.join(Assignment).filter(
        Assignment.subject_id==subject.id,
        Assignment.classroom_id==room.id,
        AssignmentStatus.student_id==student.id
    ).all()
    quiz_avg = round(sum(x.quiz_score or 0 for x in statuses)/len(statuses),2) if statuses else 0
    complete = sum(1 for x in statuses if x.status=='เรียนจบ')
    completion_percent = (complete / max(1, len(statuses))) * 100
    atts = Attendance.query.filter_by(subject_id=subject.id, classroom_id=room.id, student_id=student.id).all()
    present = sum(1 for a in atts if a.status=='มา')
    absent = sum(1 for a in atts if a.status=='ขาด')
    leave = sum(1 for a in atts if a.status=='ลา')
    late = sum(1 for a in atts if a.status=='มาสาย')
    activity = sum(1 for a in atts if a.status=='กิจกรรม')
    total_att = len(atts)
    bad_units = absent + (late * 0.5)
    attendance_percent = max(0, 100 - (bad_units * 5)) if total_att else 100
    manual = get_manual_score(subject.id, student.id)
    worksheet_score = completion_percent * (setting.worksheet_weight / 100)
    quiz_score = quiz_avg * (setting.quiz_weight / 100)
    attendance_score = attendance_percent * (setting.attendance_weight / 100)
    midterm_score = min(100, manual.midterm or 0) * (setting.midterm_weight / 100)
    final_score = min(100, manual.final or 0) * (setting.final_weight / 100)
    behavior_score = min(100, manual.behavior or 0) * (max(0, 100 - (setting.worksheet_weight+setting.quiz_weight+setting.attendance_weight+setting.midterm_weight+setting.final_weight)) / 100)
    total = min(100, worksheet_score + quiz_score + attendance_score + midterm_score + final_score + behavior_score)
    return {
        'student': student, 'quiz_avg': quiz_avg, 'complete': complete, 'total_assignments': len(statuses),
        'present': present, 'absent': absent, 'leave': leave, 'late': late, 'activity': activity,
        'attendance_percent': round(attendance_percent, 2), 'manual': manual,
        'worksheet_score': round(worksheet_score,2), 'quiz_score': round(quiz_score,2), 'attendance_score': round(attendance_score,2),
        'midterm_score': round(midterm_score,2), 'final_score': round(final_score,2), 'behavior_score': round(behavior_score,2),
        'total': round(total,2), 'grade': grade_from_score(total)
    }

def grade_from_score(score):
    if score >= 80: return '4'
    if score >= 75: return '3.5'
    if score >= 70: return '3'
    if score >= 65: return '2.5'
    if score >= 60: return '2'
    if score >= 55: return '1.5'
    if score >= 50: return '1'
    return '0'

def update_assignment_complete(status):
    assignment = status.assignment
    quiz = Quiz.query.filter_by(lesson_id=assignment.lesson_id).first()
    pass_ok = True
    if quiz:
        pass_ok = status.quiz_submitted and status.quiz_score >= quiz.pass_percent
    if status.lesson_viewed and status.worksheet_submitted and pass_ok:
        status.status = 'เรียนจบ'
    elif status.worksheet_submitted or status.quiz_submitted or status.lesson_viewed:
        status.status = 'กำลังทำ'


def is_school_blocked_day(target_date, teacher_id=None):
    q = CalendarEvent.query.filter_by(event_date=target_date)
    q = q.filter(CalendarEvent.event_type.in_(['วันหยุดราชการ','วันหยุดโรงเรียน']))
    if teacher_id:
        q = q.filter((CalendarEvent.teacher_id==teacher_id) | (CalendarEvent.teacher_id==None))
    return q.first() is not None or target_date.weekday() >= 5

def ordered_lessons_for_subject(subject_id):
    return Lesson.query.join(Unit).filter(Unit.subject_id==subject_id).order_by(Unit.id, Lesson.id).all()

def auto_lesson_for_schedule(schedule_row, target_date=None):
    """ถ้าครูผูกบทเรียนไว้ ใช้บทเรียนนั้นก่อน; ถ้าไม่ผูก ให้เรียงบทเรียนตามวันที่มีเรียนจริงในภาคเรียน"""
    if schedule_row.lesson_id:
        return schedule_row.lesson
    lessons = ordered_lessons_for_subject(schedule_row.subject_id)
    if not lessons:
        return None
    sem = get_active_semester()
    target_date = target_date or date.today()
    start = sem.start_date if sem else target_date
    end = min(target_date, sem.end_date if sem else target_date)
    if end < start:
        return lessons[0]
    cur = start
    count = 0
    while cur <= end:
        if cur.weekday() == schedule_row.weekday and not is_school_blocked_day(cur, schedule_row.teacher_id):
            count += 1
        cur += timedelta(days=1)
    idx = max(0, count - 1)
    return lessons[idx] if idx < len(lessons) else lessons[-1]

def build_student_learning_plan(student, limit=80):
    room_ids = [x.classroom_id for x in ClassroomStudent.query.filter_by(student_id=student.id).all()]
    schedules = TeachingSchedule.query.filter(TeachingSchedule.classroom_id.in_(room_ids or [-1])).order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all()
    sem = get_active_semester()
    if not sem:
        return []
    rows=[]
    cur = sem.start_date
    while cur <= sem.end_date and len(rows) < limit:
        for srow in [x for x in schedules if x.weekday == cur.weekday()]:
            if is_school_blocked_day(cur, srow.teacher_id):
                continue
            les = auto_lesson_for_schedule(srow, cur)
            rows.append(SimpleNamespace(date=cur, schedule=srow, lesson=les))
        cur += timedelta(days=1)
    return rows

def get_or_create_lesson_status(lesson, student):
    # หา/สร้างงานภายในสำหรับบทเรียนประจำคาบ ไม่ให้ปนกับงานพิเศษ
    room_ids = [x.classroom_id for x in ClassroomStudent.query.filter_by(student_id=student.id).all()]
    subject_id = lesson.unit.subject_id
    sched = TeachingSchedule.query.filter(TeachingSchedule.subject_id==subject_id, TeachingSchedule.classroom_id.in_(room_ids or [-1])).first()
    classroom_id = sched.classroom_id if sched else (room_ids[0] if room_ids else None)
    teacher_id = sched.teacher_id if sched else (lesson.unit.subject.teacher_id or User.query.filter_by(role='admin').first().id)
    if not classroom_id:
        return None
    ass = Assignment.query.filter_by(lesson_id=lesson.id, classroom_id=classroom_id, assignment_type='lesson').first()
    if not ass:
        ass = Assignment(teacher_id=teacher_id, subject_id=subject_id, classroom_id=classroom_id, lesson_id=lesson.id, title=f'บทเรียนประจำคาบ: {lesson.title}', assignment_type='lesson')
        db.session.add(ass); db.session.flush()
    st = AssignmentStatus.query.filter_by(assignment_id=ass.id, student_id=student.id).first()
    if not st:
        st = AssignmentStatus(assignment_id=ass.id, student_id=student.id, status='กำลังเรียน', lesson_viewed=True)
        db.session.add(st); db.session.flush()
    else:
        st.lesson_viewed = True
        update_assignment_complete(st)
    db.session.commit()
    return st

@app.route('/')
def index():
    if not current_user.is_authenticated:
        return redirect(url_for('login'))
    if current_user.role == 'student':
        return redirect(url_for('student_dashboard'))
    if current_user.role == 'teacher':
        return redirect(url_for('teacher_dashboard'))
    return redirect(url_for('admin_dashboard'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for('index'))
        flash('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/uploads/<path:filename>')
@login_required
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)



def get_active_semester():
    sem = Semester.query.filter_by(is_active=True).first()
    if not sem:
        sem = Semester.query.order_by(Semester.start_date.desc()).first()
    return sem

def thai_month_name(m):
    return {1:'มกราคม',2:'กุมภาพันธ์',3:'มีนาคม',4:'เมษายน',5:'พฤษภาคม',6:'มิถุนายน',7:'กรกฎาคม',8:'สิงหาคม',9:'กันยายน',10:'ตุลาคม',11:'พฤศจิกายน',12:'ธันวาคม'}.get(m, str(m))

def buddhist_year(d):
    return d.year + 543

def month_range_between(start, end):
    months = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append((y, m))
        m += 1
        if m == 13:
            m = 1
            y += 1
    return months

def build_calendar_dashboard(year=2026, months=(5,6,7,8,9,10), teacher_id=None, semester=None):
    semester = semester or get_active_semester()
    if semester:
        start_date, end_date = semester.start_date, semester.end_date
        month_pairs = month_range_between(start_date, end_date)
    else:
        month_pairs = [(year, m) for m in months]
        start_date, end_date = date(year, min(months), 1), date(year, max(months), 31)

    q = CalendarEvent.query.filter(CalendarEvent.event_date>=start_date, CalendarEvent.event_date<=end_date)
    if teacher_id:
        q = q.filter((CalendarEvent.teacher_id==teacher_id) | (CalendarEvent.teacher_id==None))
    events_by_date = {}
    for e in q.order_by(CalendarEvent.event_date).all():
        events_by_date.setdefault(e.event_date, []).append(e)

    # ใช้วันที่จริงของเครื่อง ไม่บังคับให้กระโดดไปวันเปิดเทอม
    # ถ้าวันจริงอยู่นอกช่วงเดือนที่แสดง จะไม่ไฮไลต์ผิดวัน
    real_today = date.today()
    today = real_today
    blocks = []
    today_position = None
    for y, m in month_pairs:
        weeks = py_calendar.Calendar(firstweekday=6).monthdayscalendar(y, m)
        rows = []
        for wi, week in enumerate(weeks, start=1):
            cells = []
            for di, day_no in enumerate(week):
                d = date(y, m, day_no) if day_no else None
                evs = events_by_date.get(d, []) if d else []
                if d == today:
                    today_position = {
                        'date': d, 'month_name': thai_month_name(m), 'week_no': wi,
                        'weekday_name': ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][di]
                    }
                cells.append({'day': day_no, 'date': d, 'events': evs, 'weekday': di, 'is_today': d == today})
            rows.append(cells)
        blocks.append({'year': y, 'be_year': y+543, 'month': m, 'month_name': thai_month_name(m), 'weeks': rows})
    upcoming_base = today if start_date <= today <= end_date else start_date
    upcoming = [e for day in sorted(events_by_date) for e in events_by_date[day] if day >= upcoming_base][:8]
    if today_position is None:
        today_position = {
            'date': today,
            'month_name': thai_month_name(today.month),
            'week_no': '-',
            'weekday_name': ['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์','อาทิตย์'][today.weekday()],
            'out_of_range': True
        }
    return blocks, upcoming, today_position

@app.route('/admin')
@login_required
@role_required('admin')
def admin_dashboard():
    active_semester = get_active_semester()
    cal_months, upcoming, today_position = build_calendar_dashboard(semester=active_semester)
    return render_template('admin_dashboard.html', users=User.query.count(), teachers=User.query.filter_by(role='teacher').count(), students=User.query.filter_by(role='student').count(), cal_months=cal_months, upcoming=upcoming, today_position=today_position, active_semester=active_semester)

@app.route('/teacher')
@login_required
@role_required('teacher','admin')
def teacher_dashboard():
    cur, nxt = get_current_schedule()
    subjects = teacher_filter(Subject).all()
    classrooms = teacher_filter(Classroom).all()
    assignments = Assignment.query.filter_by(teacher_id=current_user.id).order_by(Assignment.created_at.desc()).limit(5).all() if current_user.role=='teacher' else Assignment.query.order_by(Assignment.created_at.desc()).limit(5).all()
    active_semester = get_active_semester()
    cal_months, upcoming, today_position = build_calendar_dashboard(teacher_id=current_user.id if current_user.role=='teacher' else None, semester=active_semester)
    return render_template('teacher_dashboard.html', cur=cur, nxt=nxt, subjects=subjects, classrooms=classrooms, assignments=assignments, cal_months=cal_months, upcoming=upcoming, today_position=today_position, active_semester=active_semester)

@app.route('/student')
@login_required
@role_required('student')
def student_dashboard():
    statuses = AssignmentStatus.query.filter_by(student_id=current_user.id).join(Assignment).filter(Assignment.assignment_type=='special').order_by(Assignment.created_at.desc()).all()
    room_ids = [x.classroom_id for x in ClassroomStudent.query.filter_by(student_id=current_user.id).all()]
    schedules = TeachingSchedule.query.filter(TeachingSchedule.classroom_id.in_(room_ids or [-1])).order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all()
    today = date.today()
    today_idx = today.weekday()
    today_schedules = [x for x in schedules if x.weekday == today_idx and not is_school_blocked_day(today, x.teacher_id)]
    for row in schedules:
        row.auto_lesson = auto_lesson_for_schedule(row, today)
    for row in today_schedules:
        row.auto_lesson = auto_lesson_for_schedule(row, today)
    learning_plan = build_student_learning_plan(current_user, limit=120)
    return render_template('student_dashboard.html', statuses=statuses, schedules=schedules, today_schedules=today_schedules, learning_plan=learning_plan, day_names=['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์','อาทิตย์'])

@app.route('/users', methods=['GET','POST'])
@login_required
@role_required('admin')
def users():
    if request.method == 'POST':
        u = User(username=request.form['username'], full_name=request.form['full_name'], role=request.form['role'])
        u.set_password(request.form.get('password','1234'))
        db.session.add(u)
        db.session.commit()
        flash('เพิ่มผู้ใช้แล้ว', 'success')
        return redirect(url_for('users'))
    return render_template('users.html', users=User.query.order_by(User.role, User.full_name).all())

@app.route('/users/<int:user_id>/edit', methods=['GET','POST'])
@login_required
@role_required('admin')
def user_edit(user_id):
    u = User.query.get_or_404(user_id)
    if request.method == 'POST':
        u.full_name = request.form['full_name']
        u.username = request.form['username']
        u.role = request.form['role']
        if request.form.get('password'):
            u.set_password(request.form['password'])
        db.session.commit(); flash('แก้ไขผู้ใช้แล้ว', 'success')
        return redirect(url_for('users'))
    return render_template('user_form.html', u=u)

@app.route('/users/<int:user_id>/delete', methods=['POST'])
@login_required
@role_required('admin')
def user_delete(user_id):
    if user_id == current_user.id:
        flash('ไม่สามารถลบบัญชีที่กำลังใช้งานได้', 'danger'); return redirect(url_for('users'))
    u = User.query.get_or_404(user_id)
    db.session.delete(u); db.session.commit(); flash('ลบผู้ใช้แล้ว', 'success')
    return redirect(url_for('users'))


@app.route('/classrooms', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def classrooms():
    if request.method == 'POST':
        teacher_id_raw = request.form.get('teacher_id')
        teacher_id = current_user.id if current_user.role == 'teacher' else (int(teacher_id_raw) if teacher_id_raw else None)
        room = Classroom(name=request.form['name'], teacher_id=teacher_id)
        db.session.add(room); db.session.flush()
        if teacher_id and not TeacherClassroom.query.filter_by(teacher_id=teacher_id, classroom_id=room.id).first():
            db.session.add(TeacherClassroom(teacher_id=teacher_id, classroom_id=room.id))
        db.session.commit()
        flash('สร้างห้องเรียนแล้ว', 'success')
        return redirect(url_for('classrooms'))
    rooms = teacher_filter(Classroom).all()
    return render_template('classrooms.html', rooms=rooms, teachers=User.query.filter_by(role='teacher').all(), teacher_room_links=TeacherClassroom.query.all())

@app.route('/classrooms/<int:classroom_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def classroom_edit(classroom_id):
    room = Classroom.query.get_or_404(classroom_id)
    if not owns_classroom(room): return deny_redirect('classrooms')
    if request.method == 'POST':
        room.name = request.form['name']
        if current_user.role == 'admin':
            teacher_id_raw = request.form.get('teacher_id')
            room.teacher_id = int(teacher_id_raw) if teacher_id_raw else None
        db.session.commit(); flash('แก้ไขห้องเรียนแล้ว', 'success')
        return redirect(url_for('classrooms'))
    return render_template('classroom_form.html', room=room, teachers=User.query.filter_by(role='teacher').all())

def classroom_has_history(classroom_id):
    return any([
        ClassroomStudent.query.filter_by(classroom_id=classroom_id).first(),
        SubjectClassroom.query.filter_by(classroom_id=classroom_id).first(),
        Assignment.query.filter_by(classroom_id=classroom_id).first(),
        TeachingSchedule.query.filter_by(classroom_id=classroom_id).first(),
        Attendance.query.filter_by(classroom_id=classroom_id).first(),
    ])

@app.route('/classrooms/<int:classroom_id>/deactivate', methods=['POST'])
@login_required
@role_required('teacher','admin')
def classroom_deactivate(classroom_id):
    room = Classroom.query.get_or_404(classroom_id)
    if not owns_classroom(room): return deny_redirect('classrooms')
    room.is_active = False
    db.session.commit(); flash('ปิดใช้งานห้องเรียนแล้ว ข้อมูลเดิมยังอยู่ ไม่หาย', 'success')
    return redirect(url_for('classrooms'))

@app.route('/classrooms/<int:classroom_id>/restore', methods=['POST'])
@login_required
@role_required('admin')
def classroom_restore(classroom_id):
    room = Classroom.query.get_or_404(classroom_id)
    room.is_active = True
    db.session.commit(); flash('เปิดใช้งานห้องเรียนแล้ว', 'success')
    return redirect(url_for('classrooms'))

@app.route('/classrooms/<int:classroom_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def classroom_delete(classroom_id):
    room = Classroom.query.get_or_404(classroom_id)
    if not owns_classroom(room): return deny_redirect('classrooms')
    if classroom_has_history(room.id):
        room.is_active = False
        db.session.commit(); flash('ห้องนี้มีนักเรียน/ตาราง/งาน/เช็กชื่อแล้ว จึงเปลี่ยนเป็น “ปิดใช้งาน” แทนการลบเพื่อกันข้อมูลหาย', 'warning')
        return redirect(url_for('classrooms'))
    TeacherClassroom.query.filter_by(classroom_id=room.id).delete()
    db.session.delete(room); db.session.commit(); flash('ลบห้องเรียนที่ไม่มีข้อมูลผูกไว้แล้ว', 'success')
    return redirect(url_for('classrooms'))

@app.route('/classroom_students/<int:link_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def classroom_student_delete(link_id):
    link = ClassroomStudent.query.get_or_404(link_id)
    if not owns_classroom(link.classroom): return deny_redirect('classrooms')
    room_id = link.classroom_id
    db.session.delete(link); db.session.commit(); flash('นำนักเรียนออกจากห้องแล้ว', 'success')
    return redirect(url_for('classroom_students', classroom_id=room_id))


@app.route('/teacher_assignments', methods=['GET','POST'])
@login_required
@role_required('admin')
def teacher_assignments():
    if request.method == 'POST':
        teacher_id = int(request.form['teacher_id'])
        new_subject_ids = {int(x) for x in request.form.getlist('subject_ids') if x}
        new_classroom_ids = {int(x) for x in request.form.getlist('classroom_ids') if x}

        # อัปเดตแบบแทนที่จริง: ติ๊กอะไรไว้ = ครูเห็น/ใช้สิ่งนั้น, เอาติ๊กออก = ถอดสิทธิ์
        old_subject_links = TeacherSubject.query.filter_by(teacher_id=teacher_id).all()
        old_classroom_links = TeacherClassroom.query.filter_by(teacher_id=teacher_id).all()
        for link in old_subject_links:
            if link.subject_id not in new_subject_ids:
                db.session.delete(link)
        for link in old_classroom_links:
            if link.classroom_id not in new_classroom_ids:
                db.session.delete(link)
        for sid in new_subject_ids:
            if not TeacherSubject.query.filter_by(teacher_id=teacher_id, subject_id=sid).first():
                db.session.add(TeacherSubject(teacher_id=teacher_id, subject_id=sid))
        for rid in new_classroom_ids:
            if not TeacherClassroom.query.filter_by(teacher_id=teacher_id, classroom_id=rid).first():
                db.session.add(TeacherClassroom(teacher_id=teacher_id, classroom_id=rid))

        # ตั้งครูหลักให้รายการที่ยังไม่มีครูหลัก เพื่อให้หน้าอื่นแสดงผลชัดเจน
        for sid in new_subject_ids:
            sub = Subject.query.get(sid)
            if sub and not sub.teacher_id:
                sub.teacher_id = teacher_id
        for rid in new_classroom_ids:
            room = Classroom.query.get(rid)
            if room and not room.teacher_id:
                room.teacher_id = teacher_id
        db.session.commit(); flash('บันทึกการมอบหมายแล้ว ครูจะเห็นเฉพาะวิชา/ห้องที่ได้รับมอบหมาย', 'success')
        return redirect(url_for('teacher_assignments', teacher_id=teacher_id))
    selected_teacher_id = request.args.get('teacher_id', type=int)
    teachers = User.query.filter_by(role='teacher').order_by(User.full_name).all()
    subjects = Subject.query.filter_by(is_active=True).order_by(Subject.name).all()
    classrooms = Classroom.query.filter_by(is_active=True).order_by(Classroom.name).all()
    assigned_subject_ids = set()
    assigned_classroom_ids = set()
    assigned_subject_links = []
    assigned_classroom_links = []
    if selected_teacher_id:
        assigned_subject_links = TeacherSubject.query.filter_by(teacher_id=selected_teacher_id).all()
        assigned_classroom_links = TeacherClassroom.query.filter_by(teacher_id=selected_teacher_id).all()
        assigned_subject_ids = {x.subject_id for x in assigned_subject_links}
        assigned_classroom_ids = {x.classroom_id for x in assigned_classroom_links}
    return render_template('teacher_assignments.html', teachers=teachers, subjects=subjects, classrooms=classrooms, selected_teacher_id=selected_teacher_id, assigned_subject_ids=assigned_subject_ids, assigned_classroom_ids=assigned_classroom_ids, assigned_subject_links=assigned_subject_links, assigned_classroom_links=assigned_classroom_links)

@app.route('/teacher_assignments/subject/<int:link_id>/delete', methods=['POST'])
@login_required
@role_required('admin')
def teacher_subject_delete(link_id):
    link = TeacherSubject.query.get_or_404(link_id)
    teacher_id = link.teacher_id
    sub = link.subject
    db.session.delete(link)
    if sub and sub.teacher_id == teacher_id and not TeacherSubject.query.filter(TeacherSubject.teacher_id!=teacher_id, TeacherSubject.subject_id==sub.id).first():
        sub.teacher_id = None
    db.session.commit(); flash('ยกเลิกมอบหมายรายวิชาแล้ว ข้อมูลรายวิชายังอยู่', 'success')
    return redirect(url_for('teacher_assignments', teacher_id=teacher_id))

@app.route('/teacher_assignments/classroom/<int:link_id>/delete', methods=['POST'])
@login_required
@role_required('admin')
def teacher_classroom_delete(link_id):
    link = TeacherClassroom.query.get_or_404(link_id)
    teacher_id = link.teacher_id
    room = link.classroom
    db.session.delete(link)
    if room and room.teacher_id == teacher_id and not TeacherClassroom.query.filter(TeacherClassroom.teacher_id!=teacher_id, TeacherClassroom.classroom_id==room.id).first():
        room.teacher_id = None
    db.session.commit(); flash('ยกเลิกมอบหมายห้องเรียนแล้ว ข้อมูลห้องยังอยู่', 'success')
    return redirect(url_for('teacher_assignments', teacher_id=teacher_id))


@app.route('/classroom/<int:classroom_id>/students', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def classroom_students(classroom_id):
    room = Classroom.query.get_or_404(classroom_id)
    if current_user.role != 'admin' and room.teacher_id != current_user.id:
        flash('ไม่มีสิทธิ์', 'danger'); return redirect(url_for('classrooms'))
    if request.method == 'POST':
        student_id = int(request.form['student_id'])
        if not ClassroomStudent.query.filter_by(classroom_id=room.id, student_id=student_id).first():
            db.session.add(ClassroomStudent(classroom_id=room.id, student_id=student_id))
            db.session.commit()
        return redirect(url_for('classroom_students', classroom_id=room.id))
    links = ClassroomStudent.query.filter_by(classroom_id=room.id).all()
    students = User.query.filter_by(role='student').order_by(User.full_name).all()
    return render_template('classroom_students.html', room=room, links=links, students=students)

@app.route('/import_students', methods=['GET','POST'])
@login_required
@role_required('admin','teacher')
def import_students():
    rooms = teacher_filter(Classroom).all()
    if request.method == 'POST':
        f = request.files.get('excel')
        classroom_id = int(request.form['classroom_id'])
        if not f:
            flash('กรุณาเลือกไฟล์ Excel', 'danger'); return redirect(url_for('import_students'))
        path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename))
        f.save(path)
        wb = load_workbook(path)
        ws = wb.active
        created = updated = 0
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if i == 1: continue
            citizen_id, full_name, birth_date = row[0], row[1], row[2]
            if not citizen_id or not full_name: continue
            citizen_id = str(citizen_id).strip()
            if isinstance(birth_date, datetime):
                password = birth_date.strftime('%d%m%Y')
                birth_text = birth_date.strftime('%d/%m/%Y')
            else:
                birth_text = str(birth_date).strip()
                password = birth_text.replace('/','').replace('-','')
            user = User.query.filter_by(username=citizen_id).first()
            if not user:
                user = User(username=citizen_id, full_name=str(full_name), role='student', citizen_id=citizen_id, birth_date=birth_text, must_change_password=True)
                user.set_password(password)
                db.session.add(user); db.session.flush(); created += 1
            else:
                user.full_name = str(full_name); user.birth_date = birth_text; updated += 1
            if not ClassroomStudent.query.filter_by(classroom_id=classroom_id, student_id=user.id).first():
                db.session.add(ClassroomStudent(classroom_id=classroom_id, student_id=user.id))
        db.session.commit()
        flash(f'นำเข้าเรียบร้อย สร้างใหม่ {created} คน อัปเดต {updated} คน', 'success')
        return redirect(url_for('import_students'))

    return render_template('import_students.html', rooms=rooms)


def parse_excel_date_value(value):
    if value in (None, ''):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    txt = str(value).strip().replace('/', '-').replace('.', '-')
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d-%m-%y'):
        try:
            d = datetime.strptime(txt, fmt).date()
            if d.year > 2400:
                d = date(d.year-543, d.month, d.day)
            elif d.year < 100:
                d = date(d.year+2000, d.month, d.day)
            return d
        except Exception:
            pass
    return None

def birth_password(value):
    d = parse_excel_date_value(value)
    if d:
        return f"{d.day:02d}{d.month:02d}{d.year+543}"
    return str(value).replace('/','').replace('-','').strip() if value else '1234'

def sheet_rows(wb, names):
    for name in names:
        if name in wb.sheetnames:
            ws = wb[name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return [], {}
            headers = [str(x).strip().lower() if x is not None else '' for x in rows[0]]
            return rows[1:], {h:i for i,h in enumerate(headers) if h}
    return [], {}

def cell(row, headers, *names, default=''):
    for n in names:
        i = headers.get(n.lower())
        if i is not None and i < len(row) and row[i] is not None:
            return row[i]
    return default

def get_or_create_teacher(username, full_name=None):
    username = str(username).strip()
    u = User.query.filter_by(username=username).first()
    if not u:
        u = User(username=username, full_name=full_name or username, role='teacher')
        u.set_password('1234')
        db.session.add(u); db.session.flush()
    else:
        u.role = 'teacher'
        if full_name: u.full_name = full_name
    return u

def get_or_create_classroom(name):
    name = str(name).strip()
    room = Classroom.query.filter_by(name=name).first()
    if not room:
        room = Classroom(name=name)
        db.session.add(room); db.session.flush()
    return room

def get_or_create_subject(name, credit=1.0, total_periods=40):
    name = str(name).strip()
    sub = Subject.query.filter_by(name=name).first()
    if not sub:
        sub = Subject(name=name, credit=float(credit or 1), total_periods=int(total_periods or 40))
        db.session.add(sub); db.session.flush()
    return sub

def period_time(period_no):
    times = {1:('08:40','09:30'),2:('09:31','10:20'),3:('10:21','11:10'),4:('11:11','12:00'),5:('13:00','13:50'),6:('13:51','14:40'),7:('14:41','15:30'),8:('15:31','16:00')}
    return times.get(int(period_no or 1), ('08:40','09:30'))

def weekday_value(v):
    txt = str(v).strip()
    mp = {'จันทร์':0,'mon':0,'monday':0,'0':0,'อังคาร':1,'tue':1,'tuesday':1,'1':1,'พุธ':2,'wed':2,'wednesday':2,'2':2,'พฤหัสบดี':3,'พฤหัส':3,'thu':3,'thursday':3,'3':3,'ศุกร์':4,'fri':4,'friday':4,'4':4,'เสาร์':5,'sat':5,'5':5,'อาทิตย์':6,'sun':6,'6':6}
    return mp.get(txt.lower(), mp.get(txt, 0))


IMPORT_MODULES = {
    'teachers': {'title':'นำเข้าครู', 'sheet':'Teachers', 'desc':'เพิ่ม/อัปเดตบัญชีครู'},
    'students': {'title':'นำเข้านักเรียน', 'sheet':'Students', 'desc':'เพิ่ม/อัปเดตนักเรียนและผูกห้องเรียน'},
    'subjects': {'title':'นำเข้ารายวิชา', 'sheet':'Subjects', 'desc':'เพิ่ม/อัปเดตรายวิชา'},
    'classrooms': {'title':'นำเข้าห้องเรียน', 'sheet':'Classrooms', 'desc':'เพิ่ม/อัปเดตห้องเรียน'},
    'teacher_assignments': {'title':'นำเข้าการมอบหมายครู', 'sheet':'TeacherAssignments', 'desc':'กำหนดว่าครูคนไหนสอนวิชา/ห้องไหน'},
    'units': {'title':'นำเข้าหน่วยการเรียนรู้', 'sheet':'Units', 'desc':'เพิ่มหน่วย ตัวชี้วัด และจำนวนคาบ'},
    'lessons': {'title':'นำเข้าบทเรียน', 'sheet':'Lessons', 'desc':'เพิ่มบทเรียน จุดประสงค์ เนื้อหา สื่อ'},
    'worksheets': {'title':'นำเข้าใบงาน', 'sheet':'Worksheets', 'desc':'เพิ่มใบงานตามบทเรียน'},
    'worksheet_questions': {'title':'นำเข้าข้อใบงาน', 'sheet':'WorksheetQuestions', 'desc':'เพิ่มข้อคำถามและคะแนนใบงาน'},
    'quizzes': {'title':'นำเข้าแบบทดสอบ', 'sheet':'Quizzes', 'desc':'เพิ่มแบบทดสอบและเกณฑ์ผ่าน'},
    'quiz_questions': {'title':'นำเข้าข้อสอบ', 'sheet':'QuizQuestions', 'desc':'เพิ่มข้อสอบ ตัวเลือก เฉลย คะแนน'},
    'schedule': {'title':'นำเข้าตารางสอน', 'sheet':'Schedule', 'desc':'เพิ่ม/อัปเดตตารางสอนรายคาบ'},
    'calendar': {'title':'นำเข้าปฏิทิน', 'sheet':'Calendar', 'desc':'เพิ่มวันหยุด กิจกรรม วันทำงาน'},
}
IMPORT_SHEET_ALIASES = {
    'teachers':['Teachers','ครู'],
    'students':['Students','นักเรียน'],
    'subjects':['Subjects','รายวิชา'],
    'classrooms':['Classrooms','ห้องเรียน'],
    'teacher_assignments':['TeacherAssignments','มอบหมายครู'],
    'units':['Units','หน่วย'],
    'lessons':['Lessons','บทเรียน'],
    'worksheets':['Worksheets','ใบงาน'],
    'worksheet_questions':['WorksheetQuestions','ข้อใบงาน'],
    'quizzes':['Quizzes','แบบทดสอบ'],
    'quiz_questions':['QuizQuestions','ข้อสอบ'],
    'schedule':['Schedule','ตารางสอน'],
    'calendar':['Calendar','ปฏิทิน'],
}

@app.route('/imports')
@login_required
@role_required('admin')
def imports_center():
    return render_template('imports.html', modules=IMPORT_MODULES)

@app.route('/imports/<kind>')
@login_required
@role_required('admin')
def import_module_page(kind):
    module = IMPORT_MODULES.get(kind)
    if not module:
        flash('ไม่พบหมวดนำเข้า', 'danger')
        return redirect(url_for('imports_center'))
    return render_template('import_module.html', kind=kind, module=module)

@app.route('/download_import_template/<kind>')
@login_required
@role_required('admin')
def download_import_template_kind(kind):
    if kind not in IMPORT_MODULES:
        flash('ไม่พบไฟล์ตัวอย่าง', 'danger')
        return redirect(url_for('imports_center'))
    path = os.path.join(BASE_DIR, 'import_templates', f'{kind}.xlsx')
    return send_file(path, as_attachment=True)

@app.route('/import_all', methods=['GET','POST'])
@login_required
@role_required('admin')
def import_all():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or not file.filename:
            flash('กรุณาเลือกไฟล์ Excel', 'danger')
            return redirect(url_for('import_all'))
        wb = load_workbook(file)
        kind = request.form.get('kind', '').strip()
        if kind:
            allowed = set(IMPORT_SHEET_ALIASES.get(kind, []))
            for sheet_name in list(wb.sheetnames):
                if sheet_name not in allowed:
                    wb.remove(wb[sheet_name])
        report = []
        # Teachers
        n=0
        rows,h = sheet_rows(wb, ['Teachers','ครู'])
        for r in rows:
            username = cell(r,h,'username','teacher_username','รหัสครู') or cell(r,h,'citizen_id','เลขบัตร')
            full_name = cell(r,h,'full_name','name','ชื่อครู')
            if username and full_name:
                u = get_or_create_teacher(username, full_name)
                if cell(r,h,'password','รหัสผ่าน'):
                    u.set_password(cell(r,h,'password','รหัสผ่าน'))
                n += 1
        report.append(f'ครู {n} รายการ')
        # Classrooms
        n=0
        rows,h = sheet_rows(wb, ['Classrooms','ห้องเรียน'])
        for r in rows:
            name = cell(r,h,'classroom','classroom_name','ห้องเรียน','ห้อง')
            if name:
                room = get_or_create_classroom(name)
                teacher_username = cell(r,h,'teacher_username','ครู')
                if teacher_username:
                    t = User.query.filter_by(username=str(teacher_username).strip()).first()
                    if t:
                        room.teacher_id = t.id
                        if not TeacherClassroom.query.filter_by(teacher_id=t.id, classroom_id=room.id).first():
                            db.session.add(TeacherClassroom(teacher_id=t.id, classroom_id=room.id))
                n += 1
        report.append(f'ห้องเรียน {n} รายการ')
        # Students
        n=0
        rows,h = sheet_rows(wb, ['Students','นักเรียน'])
        for r in rows:
            full_name = cell(r,h,'full_name','name','ชื่อสกุล','ชื่อ-สกุล')
            citizen_id = str(cell(r,h,'citizen_id','เลขบัตร','เลขบัตรประชาชน') or '').strip()
            username = str(cell(r,h,'username','user') or citizen_id or cell(r,h,'student_code','เลขที่','รหัสนักเรียน')).strip()
            if not username or not full_name: continue
            u = User.query.filter_by(username=username).first()
            if not u:
                u = User(username=username, full_name=str(full_name).strip(), role='student', citizen_id=citizen_id, birth_date=str(cell(r,h,'birth_date','วันเกิด') or ''))
                u.set_password(birth_password(cell(r,h,'birth_date','วันเกิด')))
                db.session.add(u); db.session.flush()
            else:
                u.full_name=str(full_name).strip(); u.role='student'; u.citizen_id=citizen_id or u.citizen_id
            classroom_name = cell(r,h,'classroom','ห้องเรียน','ห้อง')
            if classroom_name:
                room = get_or_create_classroom(classroom_name)
                if not ClassroomStudent.query.filter_by(classroom_id=room.id, student_id=u.id).first():
                    db.session.add(ClassroomStudent(classroom_id=room.id, student_id=u.id))
            n += 1
        report.append(f'นักเรียน {n} รายการ')
        # Subjects
        n=0
        rows,h = sheet_rows(wb, ['Subjects','รายวิชา'])
        for r in rows:
            name = cell(r,h,'subject','subject_name','รายวิชา','ชื่อวิชา')
            if name:
                sub = get_or_create_subject(name, cell(r,h,'credit','หน่วยกิต', default=1), cell(r,h,'total_periods','จำนวนคาบ', default=40))
                teacher_username = cell(r,h,'teacher_username','ครู')
                if teacher_username:
                    t = User.query.filter_by(username=str(teacher_username).strip()).first()
                    if t:
                        sub.teacher_id = t.id
                        if not TeacherSubject.query.filter_by(teacher_id=t.id, subject_id=sub.id).first():
                            db.session.add(TeacherSubject(teacher_id=t.id, subject_id=sub.id))
                n += 1
        report.append(f'รายวิชา {n} รายการ')
        # TeacherAssignments
        n=0
        rows,h = sheet_rows(wb, ['TeacherAssignments','มอบหมายครู'])
        for r in rows:
            t = User.query.filter_by(username=str(cell(r,h,'teacher_username','ครู')).strip()).first()
            if not t: continue
            subject_name = cell(r,h,'subject','subject_name','รายวิชา')
            classroom_name = cell(r,h,'classroom','classroom_name','ห้องเรียน')
            if subject_name:
                sub = get_or_create_subject(subject_name)
                if not TeacherSubject.query.filter_by(teacher_id=t.id, subject_id=sub.id).first(): db.session.add(TeacherSubject(teacher_id=t.id, subject_id=sub.id))
                n += 1
            if classroom_name:
                room = get_or_create_classroom(classroom_name)
                if not TeacherClassroom.query.filter_by(teacher_id=t.id, classroom_id=room.id).first(): db.session.add(TeacherClassroom(teacher_id=t.id, classroom_id=room.id))
                n += 1
        report.append(f'มอบหมายครู {n} รายการ')
        # Units
        n=0
        rows,h = sheet_rows(wb, ['Units','หน่วยการเรียนรู้'])
        for r in rows:
            subject_name = cell(r,h,'subject','subject_name','รายวิชา')
            title = cell(r,h,'unit','unit_title','หน่วย','ชื่อหน่วย')
            if subject_name and title:
                sub = get_or_create_subject(subject_name)
                u = Unit.query.filter_by(subject_id=sub.id, title=str(title).strip()).first()
                if not u:
                    u = Unit(subject_id=sub.id, title=str(title).strip())
                    db.session.add(u)
                u.indicators = str(cell(r,h,'indicators','ตัวชี้วัด') or '')
                u.required_periods = int(cell(r,h,'required_periods','จำนวนคาบ', default=1) or 1)
                n += 1
        report.append(f'หน่วยการเรียนรู้ {n} รายการ')
        db.session.flush()
        # Lessons
        n=0
        rows,h = sheet_rows(wb, ['Lessons','บทเรียน'])
        for r in rows:
            subject_name = cell(r,h,'subject','subject_name','รายวิชา')
            unit_title = cell(r,h,'unit','unit_title','หน่วย')
            title = cell(r,h,'lesson','lesson_title','บทเรียน','ชื่อบทเรียน')
            if subject_name and unit_title and title:
                sub = get_or_create_subject(subject_name)
                u = Unit.query.filter_by(subject_id=sub.id, title=str(unit_title).strip()).first() or Unit(subject_id=sub.id, title=str(unit_title).strip())
                db.session.add(u); db.session.flush()
                l = Lesson.query.filter_by(unit_id=u.id, title=str(title).strip()).first()
                if not l:
                    l = Lesson(unit_id=u.id, title=str(title).strip())
                    db.session.add(l)
                l.objective=str(cell(r,h,'objective','จุดประสงค์') or '')
                l.content=str(cell(r,h,'content','เนื้อหา') or '')
                l.media_url=str(cell(r,h,'media_url','สื่อ') or '')
                l.required_minutes=int(cell(r,h,'required_minutes','นาที', default=50) or 50)
                n += 1
        report.append(f'บทเรียน {n} รายการ')
        db.session.flush()
        # Worksheets
        n=0
        rows,h = sheet_rows(wb, ['Worksheets','ใบงาน'])
        for r in rows:
            lesson_title = cell(r,h,'lesson','lesson_title','บทเรียน')
            title = cell(r,h,'worksheet','worksheet_title','ใบงาน','ชื่อใบงาน')
            l = Lesson.query.filter_by(title=str(lesson_title).strip()).first() if lesson_title else None
            if l and title:
                ws = Worksheet.query.filter_by(lesson_id=l.id, title=str(title).strip()).first()
                if not ws:
                    ws = Worksheet(lesson_id=l.id, title=str(title).strip())
                    db.session.add(ws)
                ws.worksheet_type=str(cell(r,h,'worksheet_type','ประเภท', default='text') or 'text')
                ws.description=str(cell(r,h,'description','คำชี้แจง') or '')
                n += 1
        report.append(f'ใบงาน {n} รายการ')
        db.session.flush()
        # WorksheetQuestions
        n=0
        rows,h = sheet_rows(wb, ['WorksheetQuestions','ข้อใบงาน'])
        for r in rows:
            ws_title = cell(r,h,'worksheet','worksheet_title','ใบงาน')
            qtext = cell(r,h,'question_text','คำถาม')
            ws = Worksheet.query.filter_by(title=str(ws_title).strip()).first() if ws_title else None
            if ws and qtext:
                num=int(cell(r,h,'number','ข้อ', default=1) or 1)
                q = WorksheetQuestion.query.filter_by(worksheet_id=ws.id, number=num).first()
                if not q:
                    q = WorksheetQuestion(worksheet_id=ws.id, number=num, question_text=str(qtext))
                    db.session.add(q)
                q.question_text=str(qtext); q.answer_type=str(cell(r,h,'answer_type','ประเภทคำตอบ', default='text') or 'text'); q.max_score=float(cell(r,h,'max_score','คะแนน', default=1) or 1)
                n += 1
        report.append(f'ข้อใบงาน {n} รายการ')
        # Quizzes
        n=0
        rows,h = sheet_rows(wb, ['Quizzes','แบบทดสอบ'])
        for r in rows:
            lesson_title = cell(r,h,'lesson','lesson_title','บทเรียน')
            title = cell(r,h,'quiz','quiz_title','แบบทดสอบ')
            l = Lesson.query.filter_by(title=str(lesson_title).strip()).first() if lesson_title else None
            if l and title:
                qz = Quiz.query.filter_by(lesson_id=l.id, title=str(title).strip()).first()
                if not qz:
                    qz = Quiz(lesson_id=l.id, title=str(title).strip())
                    db.session.add(qz)
                qz.pass_percent=int(cell(r,h,'pass_percent','ผ่านร้อยละ', default=60) or 60)
                n += 1
        report.append(f'แบบทดสอบ {n} รายการ')
        db.session.flush()
        # QuizQuestions
        n=0
        rows,h = sheet_rows(wb, ['QuizQuestions','ข้อสอบ'])
        for r in rows:
            quiz_title = cell(r,h,'quiz','quiz_title','แบบทดสอบ')
            qtext = cell(r,h,'question_text','คำถาม')
            qz = Quiz.query.filter_by(title=str(quiz_title).strip()).first() if quiz_title else None
            if qz and qtext:
                qq = QuizQuestion(quiz_id=qz.id, question_text=str(qtext), question_type=str(cell(r,h,'question_type','ประเภท', default='choice') or 'choice'), choices=str(cell(r,h,'choices','ตัวเลือก') or '').replace('|','\n'), correct_answer=str(cell(r,h,'correct_answer','เฉลย') or ''), score=float(cell(r,h,'score','คะแนน', default=1) or 1))
                db.session.add(qq); n += 1
        report.append(f'ข้อสอบ {n} รายการ')
        # Schedule
        n=0
        rows,h = sheet_rows(wb, ['Schedule','ตารางสอน'])
        for r in rows:
            teacher_username = cell(r,h,'teacher_username','ครู')
            t = User.query.filter_by(username=str(teacher_username).strip()).first() if teacher_username else current_user
            subject_name = cell(r,h,'subject','subject_name','รายวิชา')
            classroom_name = cell(r,h,'classroom','classroom_name','ห้องเรียน')
            if not (t and subject_name and classroom_name): continue
            sub = get_or_create_subject(subject_name); room = get_or_create_classroom(classroom_name)
            period_no = int(cell(r,h,'period_no','คาบ', default=1) or 1)
            st, et = period_time(period_no)
            st = str(cell(r,h,'start_time','เวลาเริ่ม', default=st) or st)[:5]
            et = str(cell(r,h,'end_time','เวลาจบ', default=et) or et)[:5]
            sched = TeachingSchedule.query.filter_by(teacher_id=t.id, weekday=weekday_value(cell(r,h,'weekday','วัน')), period_no=period_no).first()
            if not sched:
                sched = TeachingSchedule(teacher_id=t.id, subject_id=sub.id, classroom_id=room.id, weekday=weekday_value(cell(r,h,'weekday','วัน')), period_no=period_no, start_time=st, end_time=et)
                db.session.add(sched)
            sched.subject_id=sub.id; sched.classroom_id=room.id; sched.start_time=st; sched.end_time=et; sched.room_name=str(cell(r,h,'room_name','สถานที่') or ''); sched.topic=str(cell(r,h,'topic','หัวข้อ') or '')
            n += 1
        report.append(f'ตารางสอน {n} รายการ')
        # Calendar
        n=0
        rows,h = sheet_rows(wb, ['Calendar','ปฏิทิน'])
        for r in rows:
            d = parse_excel_date_value(cell(r,h,'date','event_date','วันที่'))
            title = cell(r,h,'title','รายการ','กิจกรรม')
            if d and title:
                typ = str(cell(r,h,'event_type','ประเภท', default='กิจกรรมโรงเรียน') or 'กิจกรรมโรงเรียน')
                note = str(cell(r,h,'note','หมายเหตุ') or '')
                if not CalendarEvent.query.filter_by(teacher_id=None, event_date=d, title=str(title).strip()).first():
                    db.session.add(CalendarEvent(teacher_id=None, event_date=d, title=str(title).strip(), event_type=typ, note=note))
                n += 1
        report.append(f'ปฏิทิน {n} รายการ')
        db.session.commit()
        flash('นำเข้าข้อมูลสำเร็จ: ' + ' / '.join(report), 'success')
        if kind:
            return redirect(request.referrer or url_for('import_module_page', kind=kind))
        return redirect(url_for('import_all'))
    return render_template('import_all.html')

@app.route('/download_import_template')
@login_required
@role_required('admin')
def download_import_template():
    return send_file(os.path.join(BASE_DIR, 'krurakson_import_template.xlsx'), as_attachment=True)

@app.route('/subjects', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def subjects():
    if request.method == 'POST':
        teacher_id_raw = request.form.get('teacher_id')
        teacher_id = current_user.id if current_user.role == 'teacher' else (int(teacher_id_raw) if teacher_id_raw else None)
        sub = Subject(name=request.form['name'], teacher_id=teacher_id, credit=float(request.form.get('credit',1)), total_periods=int(request.form.get('total_periods',40)))
        db.session.add(sub); db.session.flush()
        if teacher_id and not TeacherSubject.query.filter_by(teacher_id=teacher_id, subject_id=sub.id).first():
            db.session.add(TeacherSubject(teacher_id=teacher_id, subject_id=sub.id))
        db.session.add(GradeSetting(subject_id=sub.id))
        db.session.commit()
        flash('สร้างรายวิชาแล้ว', 'success')
        return redirect(url_for('subjects'))
    return render_template('subjects.html', subjects=teacher_filter(Subject).all(), teachers=User.query.filter_by(role='teacher').all(), subject_links=TeacherSubject.query.all(), classroom_links=SubjectClassroom.query.all())

@app.route('/subjects/<int:subject_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def subject_edit(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    if not owns_subject(subject): return deny_redirect('subjects')
    if request.method == 'POST':
        subject.name = request.form['name']
        subject.credit = float(request.form.get('credit',1))
        subject.total_periods = int(request.form.get('total_periods',40))
        if current_user.role == 'admin':
            teacher_id_raw = request.form.get('teacher_id')
            subject.teacher_id = int(teacher_id_raw) if teacher_id_raw else None
        db.session.commit(); flash('แก้ไขรายวิชาแล้ว', 'success')
        return redirect(url_for('subjects'))
    return render_template('subject_form.html', subject=subject, teachers=User.query.filter_by(role='teacher').all())

def subject_has_history(subject_id):
    return any([
        Unit.query.filter_by(subject_id=subject_id).first(),
        Assignment.query.filter_by(subject_id=subject_id).first(),
        TeachingSchedule.query.filter_by(subject_id=subject_id).first(),
        Attendance.query.filter_by(subject_id=subject_id).first(),
        ManualScore.query.filter_by(subject_id=subject_id).first(),
        SubjectClassroom.query.filter_by(subject_id=subject_id).first(),
    ])

@app.route('/subjects/<int:subject_id>/deactivate', methods=['POST'])
@login_required
@role_required('teacher','admin')
def subject_deactivate(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    if not owns_subject(subject): return deny_redirect('subjects')
    subject.is_active = False
    # ถอดจากตารางสอนใหม่ไม่ได้ แต่เก็บประวัติไว้
    db.session.commit(); flash('ปิดใช้งานรายวิชาแล้ว ข้อมูลเดิมยังอยู่ ไม่หาย', 'success')
    return redirect(url_for('subjects'))

@app.route('/subjects/<int:subject_id>/restore', methods=['POST'])
@login_required
@role_required('admin')
def subject_restore(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    subject.is_active = True
    db.session.commit(); flash('เปิดใช้งานรายวิชาแล้ว', 'success')
    return redirect(url_for('subjects'))

@app.route('/subjects/<int:subject_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def subject_delete(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    if not owns_subject(subject): return deny_redirect('subjects')
    if subject_has_history(subject.id):
        subject.is_active = False
        db.session.commit(); flash('รายวิชานี้มีข้อมูลบทเรียน/งาน/ตาราง/คะแนนแล้ว จึงเปลี่ยนเป็น “ปิดใช้งาน” แทนการลบเพื่อกันข้อมูลหาย', 'warning')
        return redirect(url_for('subjects'))
    TeacherSubject.query.filter_by(subject_id=subject.id).delete()
    GradeSetting.query.filter_by(subject_id=subject.id).delete()
    db.session.delete(subject); db.session.commit(); flash('ลบรายวิชาที่ไม่มีข้อมูลผูกไว้แล้ว', 'success')
    return redirect(url_for('subjects'))


@app.route('/subject/<int:subject_id>', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def subject_detail(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    if not owns_subject(subject): flash('ไม่มีสิทธิ์','danger'); return redirect(url_for('subjects'))
    if request.method == 'POST':
        db.session.add(Unit(subject_id=subject.id, title=request.form['title'], indicators=request.form.get('indicators',''), required_periods=int(request.form.get('required_periods',1))))
        db.session.commit(); return redirect(url_for('subject_detail', subject_id=subject.id))
    units = Unit.query.filter_by(subject_id=subject.id).all()
    completeness = []
    for u in units:
        lessons = Lesson.query.filter_by(unit_id=u.id).all()
        lesson_count = len(lessons)
        worksheets = sum(1 for l in lessons if Worksheet.query.filter_by(lesson_id=l.id).first())
        quizzes = sum(1 for l in lessons if Quiz.query.filter_by(lesson_id=l.id).first())
        percent = min(100, int((lesson_count / max(1,u.required_periods))*100))
        completeness.append((u, lesson_count, worksheets, quizzes, percent, max(0,u.required_periods-lesson_count)))
    subject_rooms = subject_classrooms_for_user(subject.id)
    return render_template('subject_detail.html', subject=subject, units=units, completeness=completeness, subject_rooms=subject_rooms)

@app.route('/unit/<int:unit_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def unit_edit(unit_id):
    unit = Unit.query.get_or_404(unit_id)
    if not owns_unit(unit): return deny_redirect('subjects')
    if request.method == 'POST':
        unit.title = request.form['title']
        unit.indicators = request.form.get('indicators','')
        unit.required_periods = int(request.form.get('required_periods',1))
        db.session.commit(); flash('แก้ไขหน่วยการเรียนรู้แล้ว', 'success')
        return redirect(url_for('subject_detail', subject_id=unit.subject_id))
    return render_template('unit_form.html', unit=unit)

@app.route('/unit/<int:unit_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def unit_delete(unit_id):
    unit = Unit.query.get_or_404(unit_id)
    if not owns_unit(unit): return deny_redirect('subjects')
    subject_id = unit.subject_id
    db.session.delete(unit); db.session.commit(); flash('ลบหน่วยการเรียนรู้แล้ว', 'success')
    return redirect(url_for('subject_detail', subject_id=subject_id))


def add_lesson_uploads(lesson):
    files = request.files.getlist('lesson_files')
    added = 0
    for f in files:
        if not f or not f.filename:
            continue
        try:
            fp, original = save_uploaded_file(f, 'lesson_files')
        except ValueError as e:
            flash(str(e), 'danger')
            continue
        if fp:
            db.session.add(LessonFile(lesson_id=lesson.id, file_path=fp, original_file_name=original, file_type=detect_lesson_file_type(fp)))
            added += 1
    return added

@app.route('/unit/<int:unit_id>/lesson', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def lesson_create(unit_id):
    unit = Unit.query.get_or_404(unit_id)
    subject = unit.subject
    if not owns_subject(subject): flash('ไม่มีสิทธิ์','danger'); return redirect(url_for('subjects'))
    if request.method == 'POST':
        l = Lesson(unit_id=unit.id, title=request.form['title'], objective=request.form.get('objective',''), content=request.form.get('content',''), media_url=request.form.get('media_url',''))
        db.session.add(l); db.session.flush()
        add_lesson_uploads(l)
        db.session.commit()
        flash('สร้างบทเรียนและแนบไฟล์แล้ว นักเรียนจะเห็นทันทีเมื่อเปิดบทเรียน/ตารางสอนที่ผูกบทเรียนนี้', 'success')
        return redirect(url_for('lesson_detail', lesson_id=l.id))
    return render_template('lesson_form.html', unit=unit, lesson=None, files=[])

@app.route('/lesson/<int:lesson_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def lesson_edit(lesson_id):
    lesson = Lesson.query.get_or_404(lesson_id)
    if not owns_lesson(lesson): return deny_redirect('subjects')
    if request.method == 'POST':
        lesson.title = request.form['title']
        lesson.objective = request.form.get('objective','')
        lesson.content = request.form.get('content','')
        lesson.media_url = request.form.get('media_url','')
        add_lesson_uploads(lesson)
        db.session.commit(); flash('แก้ไขบทเรียนแล้ว เนื้อหา/ไฟล์แนบจะแสดงในหน้าบทเรียนทันที', 'success')
        return redirect(url_for('lesson_detail', lesson_id=lesson.id))
    files = LessonFile.query.filter_by(lesson_id=lesson.id).order_by(LessonFile.created_at.desc()).all()
    return render_template('lesson_form.html', unit=lesson.unit, lesson=lesson, files=files)

@app.route('/lesson_file/<int:file_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def lesson_file_delete(file_id):
    lf = LessonFile.query.get_or_404(file_id)
    lesson = lf.lesson
    if not owns_lesson(lesson): return deny_redirect('subjects')
    try:
        abs_path = os.path.join(app.config['UPLOAD_FOLDER'], lf.file_path)
        if os.path.exists(abs_path): os.remove(abs_path)
    except Exception:
        pass
    db.session.delete(lf); db.session.commit(); flash('ลบไฟล์บทเรียนแล้ว', 'success')
    return redirect(url_for('lesson_edit', lesson_id=lesson.id))

@app.route('/lesson/<int:lesson_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def lesson_delete(lesson_id):
    lesson = Lesson.query.get_or_404(lesson_id)
    if not owns_lesson(lesson): return deny_redirect('subjects')
    subject_id = lesson.unit.subject_id
    for lf in LessonFile.query.filter_by(lesson_id=lesson.id).all():
        try:
            abs_path = os.path.join(app.config['UPLOAD_FOLDER'], lf.file_path)
            if os.path.exists(abs_path): os.remove(abs_path)
        except Exception:
            pass
        db.session.delete(lf)
    db.session.delete(lesson); db.session.commit(); flash('ลบบทเรียนแล้ว', 'success')
    return redirect(url_for('subject_detail', subject_id=subject_id))


@app.route('/lesson/<int:lesson_id>')
@login_required
def lesson_detail(lesson_id):
    lesson = Lesson.query.get_or_404(lesson_id)
    worksheets = Worksheet.query.filter_by(lesson_id=lesson.id).all()
    quizzes = Quiz.query.filter_by(lesson_id=lesson.id).all()
    files = LessonFile.query.filter_by(lesson_id=lesson.id).order_by(LessonFile.created_at.desc()).all()
    lesson_status = get_or_create_lesson_status(lesson, current_user) if current_user.role == 'student' else None
    return render_template('lesson_detail.html', lesson=lesson, worksheets=worksheets, quizzes=quizzes, files=files, lesson_status=lesson_status, youtube_embed=youtube_embed_url(lesson.media_url))


@app.route('/subject/<int:subject_id>/worksheets')
@login_required
@role_required('teacher','admin')
def subject_worksheets(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    if not owns_subject(subject): return deny_redirect('subjects')
    items = []
    units = Unit.query.filter_by(subject_id=subject.id).all()
    for unit in units:
        lessons = Lesson.query.filter_by(unit_id=unit.id).all()
        for lesson in lessons:
            for worksheet in Worksheet.query.filter_by(lesson_id=lesson.id).all():
                items.append(SimpleNamespace(unit=unit, lesson=lesson, worksheet=worksheet))
    return render_template('subject_worksheets.html', subject=subject, items=items)

@app.route('/worksheet/<int:worksheet_id>/review')
@login_required
@role_required('teacher','admin')
def worksheet_review(worksheet_id):
    worksheet = Worksheet.query.get_or_404(worksheet_id)
    if not owns_lesson(worksheet.lesson): return deny_redirect('subjects')
    subject = worksheet.lesson.unit.subject
    rows = []
    # หา assignment ที่ผูกกับบทเรียนนี้ แล้วดึงนักเรียนของห้องนั้น ๆ
    assignments = Assignment.query.filter_by(lesson_id=worksheet.lesson_id).all()
    seen = set()
    for assignment in assignments:
        memberships = ClassroomStudent.query.filter_by(classroom_id=assignment.classroom_id).all()
        for m in memberships:
            key = (assignment.id, m.student_id)
            if key in seen:
                continue
            seen.add(key)
            status = AssignmentStatus.query.filter_by(assignment_id=assignment.id, student_id=m.student_id).first()
            answers = WorksheetAnswer.query.filter_by(assignment_id=assignment.id, student_id=m.student_id).join(WorksheetQuestion).filter(WorksheetQuestion.worksheet_id==worksheet.id).order_by(WorksheetQuestion.number).all()
            submitted = bool(status and status.worksheet_submitted) or bool(answers)
            score = sum((a.score or 0) for a in answers)
            rows.append(SimpleNamespace(student=m.student, classroom=assignment.classroom, status=status, answers=answers, submitted=submitted, score=score))
    # ถ้ายังไม่เคยสั่งงาน ให้แสดงข้อความว่างอย่างชัดเจน
    return render_template('worksheet_review.html', worksheet=worksheet, rows=rows)

@app.route('/lesson/<int:lesson_id>/worksheet', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def worksheet_create(lesson_id):
    lesson = Lesson.query.get_or_404(lesson_id)
    if not owns_lesson(lesson): return deny_redirect('subjects')
    if request.method == 'POST':
        ws = Worksheet(lesson_id=lesson.id, title=request.form['title'], worksheet_type=request.form.get('worksheet_type','academic'), description=request.form.get('description',''))
        try:
            fp, original = save_uploaded_file(request.files.get('worksheet_file'), 'worksheets')
            ws.file_path, ws.original_file_name = fp, original
        except ValueError as e:
            flash(str(e), 'danger')
            return redirect(url_for('worksheet_create', lesson_id=lesson.id))
        db.session.add(ws); db.session.flush()
        questions = request.form.get('questions','').splitlines()
        for idx, q in enumerate([x.strip() for x in questions if x.strip()], start=1):
            db.session.add(WorksheetQuestion(worksheet_id=ws.id, number=idx, question_text=q, answer_type='text', max_score=float(request.form.get('default_score', 1) or 1)))
        # petanque template: ระยะ 6.5/7.5/8.5/9.5 x สถานี 1-5
        if ws.worksheet_type == 'petanque_score' and not questions:
            distances = ['6.5 m','7.5 m','8.5 m','9.5 m']
            n=1
            for d in distances:
                for station in range(1,6):
                    db.session.add(WorksheetQuestion(worksheet_id=ws.id, number=n, question_text=f'{d} - สถานีที่ {station}', answer_type='score', max_score=5)); n+=1
        if ws.worksheet_type == 'upload' and not questions:
            db.session.add(WorksheetQuestion(worksheet_id=ws.id, number=1, question_text='แนบไฟล์คำตอบ', answer_type='file', max_score=float(request.form.get('default_score', 1) or 1)))
        db.session.commit(); flash('สร้างใบงานแล้ว', 'success'); return redirect(url_for('lesson_detail', lesson_id=lesson.id))
    return render_template('worksheet_form.html', lesson=lesson, worksheet=None)

@app.route('/worksheet/<int:worksheet_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def worksheet_edit(worksheet_id):
    ws = Worksheet.query.get_or_404(worksheet_id)
    if not owns_lesson(ws.lesson): return deny_redirect('subjects')
    if request.method == 'POST':
        old_type = ws.worksheet_type
        ws.title = request.form['title']
        ws.worksheet_type = request.form.get('worksheet_type','academic')
        ws.description = request.form.get('description','')
        try:
            fp, original = save_uploaded_file(request.files.get('worksheet_file'), 'worksheets')
            if fp:
                ws.file_path, ws.original_file_name = fp, original
        except ValueError as e:
            flash(str(e), 'danger')
            return redirect(url_for('worksheet_edit', worksheet_id=ws.id))
        # ถ้าเปลี่ยนเป็นใบงานเปตอง/อัปโหลดและยังไม่มีข้อ ให้สร้างฟอร์มอัตโนมัติ
        if ws.worksheet_type == 'petanque_score' and old_type != 'petanque_score' and not WorksheetQuestion.query.filter_by(worksheet_id=ws.id).first():
            n=1
            for d in ['6.5 m','7.5 m','8.5 m','9.5 m']:
                for station in range(1,6):
                    db.session.add(WorksheetQuestion(worksheet_id=ws.id, number=n, question_text=f'{d} - สถานีที่ {station}', answer_type='score', max_score=5)); n+=1
        if ws.worksheet_type == 'upload' and not WorksheetQuestion.query.filter_by(worksheet_id=ws.id).first():
            db.session.add(WorksheetQuestion(worksheet_id=ws.id, number=1, question_text='แนบไฟล์คำตอบ', answer_type='file', max_score=1))
        db.session.commit(); flash('แก้ไขใบงานแล้ว', 'success')
        return redirect(url_for('lesson_detail', lesson_id=ws.lesson_id))
    return render_template('worksheet_form.html', lesson=ws.lesson, worksheet=ws)

@app.route('/worksheet/<int:worksheet_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def worksheet_delete(worksheet_id):
    ws = Worksheet.query.get_or_404(worksheet_id)
    if not owns_lesson(ws.lesson): return deny_redirect('subjects')
    lesson_id = ws.lesson_id
    db.session.delete(ws); db.session.commit(); flash('ลบใบงานแล้ว', 'success')
    return redirect(url_for('lesson_detail', lesson_id=lesson_id))

@app.route('/worksheet_question/<int:question_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def worksheet_question_delete(question_id):
    q = WorksheetQuestion.query.get_or_404(question_id)
    if not owns_lesson(q.worksheet.lesson): return deny_redirect('subjects')
    lesson_id = q.worksheet.lesson_id
    db.session.delete(q); db.session.commit(); flash('ลบข้อคำถามใบงานแล้ว', 'success')
    return redirect(url_for('lesson_detail', lesson_id=lesson_id))


@app.route('/lesson/<int:lesson_id>/quiz', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def quiz_create(lesson_id):
    lesson = Lesson.query.get_or_404(lesson_id)
    if request.method == 'POST':
        qz = Quiz(lesson_id=lesson.id, title=request.form['title'], pass_percent=int(request.form.get('pass_percent',60)))
        db.session.add(qz); db.session.commit(); return redirect(url_for('quiz_questions', quiz_id=qz.id))
    return render_template('quiz_form.html', lesson=lesson)

@app.route('/quiz/<int:quiz_id>/questions', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def quiz_questions(quiz_id):
    quiz = Quiz.query.get_or_404(quiz_id)
    if request.method == 'POST':
        question_type = request.form.get('question_type','abcd')
        if question_type == 'abcd':
            labels = ['ก','ข','ค','ง']
            choices = '\n'.join([f"{lb}. {request.form.get('choice_'+lb, '').strip()}" for lb in labels if request.form.get('choice_'+lb, '').strip()])
            correct_answer = request.form.get('correct_letter','ก')
        else:
            choices = request.form.get('choices','')
            correct_answer = request.form.get('correct_answer','')
        db.session.add(QuizQuestion(quiz_id=quiz.id, question_text=request.form['question_text'], question_type=question_type, choices=choices, correct_answer=correct_answer, score=float(request.form.get('score',1))))
        db.session.commit(); flash('เพิ่มข้อสอบแล้ว', 'success'); return redirect(url_for('quiz_questions', quiz_id=quiz.id))
    questions = QuizQuestion.query.filter_by(quiz_id=quiz.id).all()
    return render_template('quiz_questions.html', quiz=quiz, questions=questions)

@app.route('/quiz/<int:quiz_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def quiz_edit(quiz_id):
    quiz = Quiz.query.get_or_404(quiz_id)
    if not owns_lesson(quiz.lesson): return deny_redirect('subjects')
    if request.method == 'POST':
        quiz.title = request.form['title']
        quiz.pass_percent = int(request.form.get('pass_percent',60))
        db.session.commit(); flash('แก้ไขแบบทดสอบแล้ว', 'success')
        return redirect(url_for('lesson_detail', lesson_id=quiz.lesson_id))
    return render_template('quiz_form.html', lesson=quiz.lesson, quiz=quiz)

@app.route('/quiz/<int:quiz_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def quiz_delete(quiz_id):
    quiz = Quiz.query.get_or_404(quiz_id)
    if not owns_lesson(quiz.lesson): return deny_redirect('subjects')
    lesson_id = quiz.lesson_id
    db.session.delete(quiz); db.session.commit(); flash('ลบแบบทดสอบแล้ว', 'success')
    return redirect(url_for('lesson_detail', lesson_id=lesson_id))

@app.route('/quiz_question/<int:question_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def quiz_question_edit(question_id):
    q = QuizQuestion.query.get_or_404(question_id)
    if not owns_lesson(q.quiz.lesson): return deny_redirect('subjects')
    if request.method == 'POST':
        q.question_text = request.form['question_text']
        q.question_type = request.form.get('question_type','abcd')
        if q.question_type == 'abcd':
            labels = ['ก','ข','ค','ง']
            q.choices = '\n'.join([f"{lb}. {request.form.get('choice_'+lb, '').strip()}" for lb in labels if request.form.get('choice_'+lb, '').strip()])
            q.correct_answer = request.form.get('correct_letter','ก')
        else:
            q.choices = request.form.get('choices','')
            q.correct_answer = request.form.get('correct_answer','')
        q.score = float(request.form.get('score',1))
        db.session.commit(); flash('แก้ไขข้อสอบแล้ว', 'success')
        return redirect(url_for('quiz_questions', quiz_id=q.quiz_id))
    return render_template('quiz_question_form.html', quiz=q.quiz, q=q)

@app.route('/quiz_question/<int:question_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def quiz_question_delete(question_id):
    q = QuizQuestion.query.get_or_404(question_id)
    if not owns_lesson(q.quiz.lesson): return deny_redirect('subjects')
    quiz_id = q.quiz_id
    db.session.delete(q); db.session.commit(); flash('ลบข้อสอบแล้ว', 'success')
    return redirect(url_for('quiz_questions', quiz_id=quiz_id))


@app.route('/assign', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def assign():
    subjects = teacher_filter(Subject).all()
    rooms = teacher_filter(Classroom).all()
    lessons = Lesson.query.join(Unit).join(Subject).filter(Subject.id.in_(teacher_subject_ids() or [-1])).all() if current_user.role=='teacher' else Lesson.query.all()
    if request.method == 'POST':
        a = Assignment(teacher_id=current_user.id, subject_id=int(request.form['subject_id']), classroom_id=int(request.form['classroom_id']), lesson_id=int(request.form['lesson_id']), title=request.form['title'], due_date=datetime.strptime(request.form['due_date'],'%Y-%m-%d').date() if request.form.get('due_date') else None)
        db.session.add(a); db.session.flush()
        links = ClassroomStudent.query.filter_by(classroom_id=a.classroom_id).all()
        for link in links:
            db.session.add(AssignmentStatus(assignment_id=a.id, student_id=link.student_id))
        db.session.commit(); flash('สั่งงานให้นักเรียนในห้องแล้ว', 'success'); return redirect(url_for('teacher_dashboard'))
    return render_template('assign.html', subjects=subjects, rooms=rooms, lessons=lessons, assignment=None)

@app.route('/assignments')
@login_required
@role_required('teacher','admin')
def assignments():
    rows = Assignment.query.filter_by(teacher_id=current_user.id).order_by(Assignment.created_at.desc()).all() if current_user.role=='teacher' else Assignment.query.order_by(Assignment.created_at.desc()).all()
    return render_template('assignments.html', assignments=rows)

@app.route('/assignment/<int:assignment_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def assignment_edit(assignment_id):
    a = Assignment.query.get_or_404(assignment_id)
    if current_user.role != 'admin' and a.teacher_id != current_user.id: return deny_redirect('assignments')
    subjects = teacher_filter(Subject).all(); rooms = teacher_filter(Classroom).all()
    lessons = Lesson.query.join(Unit).join(Subject).filter(Subject.id.in_(teacher_subject_ids() or [-1])).all() if current_user.role=='teacher' else Lesson.query.all()
    if request.method == 'POST':
        a.title = request.form['title']
        a.subject_id = int(request.form['subject_id'])
        a.classroom_id = int(request.form['classroom_id'])
        a.lesson_id = int(request.form['lesson_id'])
        a.due_date = datetime.strptime(request.form['due_date'],'%Y-%m-%d').date() if request.form.get('due_date') else None
        db.session.commit(); flash('แก้ไขงานที่สั่งแล้ว', 'success')
        return redirect(url_for('assignments'))
    return render_template('assign.html', subjects=subjects, rooms=rooms, lessons=lessons, assignment=a)

@app.route('/assignment/<int:assignment_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def assignment_delete(assignment_id):
    a = Assignment.query.get_or_404(assignment_id)
    if current_user.role != 'admin' and a.teacher_id != current_user.id: return deny_redirect('assignments')
    db.session.delete(a); db.session.commit(); flash('ลบงานที่สั่งแล้ว', 'success')
    return redirect(url_for('assignments'))


@app.route('/student/assignment/<int:status_id>')
@login_required
@role_required('student')
def student_assignment(status_id):
    status = AssignmentStatus.query.get_or_404(status_id)
    if status.student_id != current_user.id: flash('ไม่มีสิทธิ์','danger'); return redirect(url_for('student_dashboard'))
    status.lesson_viewed = True
    update_assignment_complete(status)
    db.session.commit()
    lesson = status.assignment.lesson
    worksheets = Worksheet.query.filter_by(lesson_id=lesson.id).all()
    quiz = Quiz.query.filter_by(lesson_id=lesson.id).first()
    return render_template('student_assignment.html', status=status, lesson=lesson, worksheets=worksheets, quiz=quiz)

@app.route('/student/assignment/<int:status_id>/worksheet/<int:worksheet_id>', methods=['GET','POST'])
@login_required
@role_required('student')
def do_worksheet(status_id, worksheet_id):
    status = AssignmentStatus.query.get_or_404(status_id)
    worksheet = Worksheet.query.get_or_404(worksheet_id)
    questions = WorksheetQuestion.query.filter_by(worksheet_id=worksheet.id).order_by(WorksheetQuestion.number).all()
    if request.method == 'POST':
        for q in questions:
            ans = request.form.get(f'q_{q.id}','')
            row = WorksheetAnswer.query.filter_by(assignment_id=status.assignment_id, student_id=current_user.id, question_id=q.id).first()
            if not row:
                row = WorksheetAnswer(assignment_id=status.assignment_id, student_id=current_user.id, question_id=q.id)
                db.session.add(row)
            row.answer_text = ans
            file_obj = request.files.get(f'file_{q.id}')
            if file_obj and file_obj.filename:
                try:
                    fp, original = save_uploaded_file(file_obj, 'student_works')
                    row.file_path, row.original_file_name = fp, original
                except ValueError as e:
                    flash(str(e), 'danger')
                    return redirect(url_for('do_worksheet', status_id=status.id, worksheet_id=worksheet.id))
            try: row.score = float(ans) if q.answer_type=='score' and ans else 0
            except: row.score = 0
        status.worksheet_submitted = True; status.submitted_at = datetime.utcnow(); update_assignment_complete(status)
        db.session.commit(); flash('ส่งใบงานแล้ว', 'success'); return redirect(url_for('student_assignment', status_id=status.id))
    answers = WorksheetAnswer.query.filter_by(assignment_id=status.assignment_id, student_id=current_user.id).all()
    old = {a.question_id:a.answer_text for a in answers}
    old_files = {a.question_id:a for a in answers if a.file_path}
    return render_template('do_worksheet.html', status=status, worksheet=worksheet, questions=questions, old=old, old_files=old_files)

@app.route('/student/assignment/<int:status_id>/quiz/<int:quiz_id>', methods=['GET','POST'])
@login_required
@role_required('student')
def do_quiz(status_id, quiz_id):
    status = AssignmentStatus.query.get_or_404(status_id)
    quiz = Quiz.query.get_or_404(quiz_id)
    questions = QuizQuestion.query.filter_by(quiz_id=quiz.id).all()
    if request.method == 'POST':
        total = sum(q.score for q in questions) or 1
        got = 0
        QuizAnswer.query.filter_by(assignment_id=status.assignment_id, student_id=current_user.id).delete()
        for q in questions:
            ans = request.form.get(f'q_{q.id}','')
            if q.question_type == 'abcd':
                correct = ans.strip() == (q.correct_answer or '').strip()
            else:
                correct = ans.strip().lower() == (q.correct_answer or '').strip().lower()
            score = q.score if correct else 0
            got += score
            db.session.add(QuizAnswer(assignment_id=status.assignment_id, student_id=current_user.id, question_id=q.id, answer_text=ans, is_correct=correct, score=score))
        status.quiz_submitted = True
        status.quiz_score = round((got/total)*100,2)
        update_assignment_complete(status)
        db.session.commit(); flash(f'ทำแบบทดสอบแล้ว คะแนน {status.quiz_score}%', 'success'); return redirect(url_for('student_assignment', status_id=status.id))
    return render_template('do_quiz.html', status=status, quiz=quiz, questions=questions)



def default_period_times():
    return {
        1: ('08:40','09:30'),
        2: ('09:31','10:20'),
        3: ('10:21','11:10'),
        4: ('11:11','12:00'),
        5: ('13:00','13:50'),
        6: ('13:51','14:40'),
        7: ('14:41','15:30'),
        8: ('15:31','16:00'),
    }

def thai_weekday_to_int(value):
    txt = str(value or '').strip().lower()
    mapping = {
        '0':0,'จันทร์':0,'จ.':0,'mon':0,'monday':0,
        '1':1,'อังคาร':1,'อ.':1,'tue':1,'tuesday':1,
        '2':2,'พุธ':2,'พ.':2,'wed':2,'wednesday':2,
        '3':3,'พฤหัสบดี':3,'พฤหัส':3,'พฤ.':3,'thu':3,'thursday':3,
        '4':4,'ศุกร์':4,'ศ.':4,'fri':4,'friday':4,
        '5':5,'เสาร์':5,'ส.':5,'sat':5,'saturday':5,
        '6':6,'อาทิตย์':6,'อา.':6,'sun':6,'sunday':6,
    }
    return mapping.get(txt)

def get_or_create_subject_for_teacher(name, teacher_id):
    name = str(name or '').strip()
    if not name:
        return None
    sub = Subject.query.filter_by(name=name).first()
    if not sub:
        sub = Subject(name=name, teacher_id=None, credit=1.0, total_periods=40)
        db.session.add(sub); db.session.flush()
        db.session.add(GradeSetting(subject_id=sub.id))
    if teacher_id and not TeacherSubject.query.filter_by(teacher_id=teacher_id, subject_id=sub.id).first():
        db.session.add(TeacherSubject(teacher_id=teacher_id, subject_id=sub.id))
    return sub

def get_or_create_classroom_for_teacher(name, teacher_id):
    name = str(name or '').strip()
    if not name:
        return None
    room = Classroom.query.filter_by(name=name).first()
    if not room:
        room = Classroom(name=name, teacher_id=None)
        db.session.add(room); db.session.flush()
    if teacher_id and not TeacherClassroom.query.filter_by(teacher_id=teacher_id, classroom_id=room.id).first():
        db.session.add(TeacherClassroom(teacher_id=teacher_id, classroom_id=room.id))
    return room

@app.route('/schedule', methods=['GET','POST'])
@login_required
@role_required('teacher','admin','student')
def schedule():
    teachers = User.query.filter_by(role='teacher').order_by(User.full_name).all()
    day_names = ['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์','อาทิตย์']
    selected_teacher_id = request.args.get('teacher_id', type=int)

    if current_user.role == 'admin':
        active_teacher_id = selected_teacher_id or (teachers[0].id if teachers else None)
        subjects = Subject.query.filter(Subject.id.in_({x.subject_id for x in TeacherSubject.query.filter_by(teacher_id=active_teacher_id).all()} or {-1})).order_by(Subject.name).all() if active_teacher_id else []
        rooms = Classroom.query.filter(Classroom.id.in_({x.classroom_id for x in TeacherClassroom.query.filter_by(teacher_id=active_teacher_id).all()} or {-1})).order_by(Classroom.name).all() if active_teacher_id else []
        schedules_query = TeachingSchedule.query
        if active_teacher_id:
            schedules_query = schedules_query.filter_by(teacher_id=active_teacher_id)
        schedules = schedules_query.order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all()
        teacher_title = User.query.get(active_teacher_id).full_name if active_teacher_id else 'ยังไม่เลือกครู'
    elif current_user.role == 'teacher':
        active_teacher_id = current_user.id
        subjects = teacher_filter(Subject).order_by(Subject.name).all()
        rooms = teacher_filter(Classroom).order_by(Classroom.name).all()
        schedules = TeachingSchedule.query.filter_by(teacher_id=current_user.id).order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all()
        teacher_title = current_user.full_name
    else:
        # นักเรียนเห็นเฉพาะตารางสอนของห้องที่ตนเองอยู่
        student_room_ids = [x.classroom_id for x in ClassroomStudent.query.filter_by(student_id=current_user.id).all()]
        schedules = TeachingSchedule.query.filter(TeachingSchedule.classroom_id.in_(student_room_ids or [-1])).order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all()
        subjects = []
        rooms = []
        active_teacher_id = None
        teacher_title = current_user.full_name

    lessons = Lesson.query.join(Unit).filter(Unit.subject_id.in_([x.id for x in subjects] or [-1])).order_by(Lesson.title).all() if current_user.role != 'student' else []
    edit_row = None
    if request.method == 'POST':
        if current_user.role == 'student':
            return deny_redirect('schedule')
        post_teacher_id = int(request.form.get('teacher_id') or current_user.id)
        if current_user.role != 'admin':
            post_teacher_id = current_user.id
        subject_id = int(request.form['subject_id'])
        classroom_id = int(request.form['classroom_id'])
        # ตารางสอนต้องใช้เฉพาะวิชา/ห้องที่กำหนดให้ครูแล้วเท่านั้น
        if not TeacherSubject.query.filter_by(teacher_id=post_teacher_id, subject_id=subject_id).first():
            flash('ยังไม่ได้มอบหมายวิชานี้ให้ครูคนนี้ กรุณาไปหน้ามอบหมายครูก่อน', 'danger')
            return redirect(url_for('schedule', teacher_id=post_teacher_id))
        if not TeacherClassroom.query.filter_by(teacher_id=post_teacher_id, classroom_id=classroom_id).first():
            flash('ยังไม่ได้มอบหมายห้องนี้ให้ครูคนนี้ กรุณาไปหน้ามอบหมายครูก่อน', 'danger')
            return redirect(url_for('schedule', teacher_id=post_teacher_id))
        period_no = int(request.form['period_no'])
        defaults = default_period_times()
        st, en = defaults.get(period_no, ('08:00','09:00'))
        db.session.add(TeachingSchedule(
            teacher_id=post_teacher_id,
            subject_id=subject_id,
            classroom_id=classroom_id,
            weekday=int(request.form['weekday']),
            period_no=period_no,
            start_time=request.form.get('start_time') or st,
            end_time=request.form.get('end_time') or en,
            room_name=request.form.get('room_name',''),
            topic=request.form.get('topic',''),
            lesson_id=int(request.form['lesson_id']) if request.form.get('lesson_id') else None
        ))
        db.session.commit(); return redirect(url_for('schedule', teacher_id=post_teacher_id))

    periods = list(range(1, 9))
    schedule_grid = {f"{x.weekday}-{x.period_no}": x for x in schedules}
    day_pairs = list(enumerate(day_names[:5]))
    schedules_by_teacher = {}
    if current_user.role == 'admin' and not selected_teacher_id:
        all_rows = TeachingSchedule.query.order_by(TeachingSchedule.teacher_id, TeachingSchedule.weekday, TeachingSchedule.period_no).all()
        for row in all_rows:
            schedules_by_teacher.setdefault(row.teacher_id, []).append(row)
    return render_template('schedule.html', schedules=schedules, subjects=subjects, rooms=rooms, lessons=lessons, edit_row=edit_row, day_names=day_names, periods=periods, schedule_grid=schedule_grid, day_pairs=day_pairs, teachers=teachers, selected_teacher_id=selected_teacher_id, active_teacher_id=active_teacher_id, teacher_title=teacher_title, schedules_by_teacher=schedules_by_teacher)

@app.route('/schedule/import', methods=['POST'])
@login_required
@role_required('teacher','admin')
def schedule_import():
    """นำเข้าตารางสอนจาก Excel
    คอลัมน์ที่รองรับ: day/วัน, period/คาบ, subject/รายวิชา, classroom/ห้อง, room/สถานที่, topic/หัวข้อ
    หรือใช้ลำดับคอลัมน์: วัน, คาบ, รายวิชา, ห้อง, สถานที่, หัวข้อ
    """
    f = request.files.get('excel')
    if not f:
        flash('กรุณาเลือกไฟล์ Excel ตารางสอน', 'danger')
        return redirect(url_for('schedule'))
    path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename))
    f.save(path)
    wb = load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        flash('ไฟล์ Excel ไม่มีข้อมูล', 'danger')
        return redirect(url_for('schedule'))
    headers = [str(x or '').strip().lower() for x in rows[0]]
    def idx(*names, default=None):
        for name in names:
            n = name.lower()
            if n in headers:
                return headers.index(n)
        return default
    i_day = idx('day','วัน','weekday', default=0)
    i_period = idx('period','คาบ','period_no', default=1)
    i_subject = idx('subject','รายวิชา','วิชา', default=2)
    i_classroom = idx('classroom','ห้อง','ชั้น/ห้อง', default=3)
    i_room = idx('room','สถานที่','ห้องเรียน', default=4)
    i_topic = idx('topic','หัวข้อ','เรื่อง', default=5)
    defaults = default_period_times()
    teacher_id = int(request.form.get('teacher_id') or current_user.id)
    if current_user.role != 'admin':
        teacher_id = current_user.id
    created = skipped = 0
    for row in rows[1:]:
        if not row or len(row) <= max(i_day, i_period, i_subject, i_classroom):
            skipped += 1; continue
        weekday = thai_weekday_to_int(row[i_day])
        try:
            period_no = int(row[i_period])
        except Exception:
            period_no = None
        subject_name = row[i_subject]
        classroom_name = row[i_classroom]
        if weekday is None or not period_no or not subject_name or not classroom_name:
            skipped += 1; continue
        sub = get_or_create_subject_for_teacher(subject_name, teacher_id)
        room = get_or_create_classroom_for_teacher(classroom_name, teacher_id)
        if not SubjectClassroom.query.filter_by(subject_id=sub.id, classroom_id=room.id).first():
            db.session.add(SubjectClassroom(subject_id=sub.id, classroom_id=room.id))
        st, en = defaults.get(period_no, ('08:00','09:00'))
        room_name = str(row[i_room] or '') if i_room is not None and i_room < len(row) else ''
        topic = str(row[i_topic] or '') if i_topic is not None and i_topic < len(row) else ''
        old = TeachingSchedule.query.filter_by(teacher_id=teacher_id, weekday=weekday, period_no=period_no, subject_id=sub.id, classroom_id=room.id).first()
        if old:
            old.start_time = st; old.end_time = en; old.room_name = room_name; old.topic = topic
        else:
            db.session.add(TeachingSchedule(teacher_id=teacher_id, subject_id=sub.id, classroom_id=room.id, weekday=weekday, period_no=period_no, start_time=st, end_time=en, room_name=room_name, topic=topic))
            created += 1
    db.session.commit()
    flash(f'นำเข้าตารางสอนแล้ว {created} คาบ ข้าม {skipped} แถว', 'success')
    return redirect(url_for('schedule'))

@app.route('/schedule/<int:schedule_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def schedule_edit(schedule_id):
    row = TeachingSchedule.query.get_or_404(schedule_id)
    if current_user.role != 'admin' and row.teacher_id != current_user.id: return deny_redirect('schedule')
    subjects = teacher_filter(Subject).all(); rooms = teacher_filter(Classroom).all()
    lessons = Lesson.query.join(Unit).join(Subject).filter(Subject.id.in_(teacher_subject_ids() or [-1])).all() if current_user.role=='teacher' else Lesson.query.all()
    if request.method == 'POST':
        row.weekday=int(request.form['weekday']); row.period_no=int(request.form['period_no'])
        row.start_time=request.form['start_time']; row.end_time=request.form['end_time']
        row.subject_id=int(request.form['subject_id']); row.classroom_id=int(request.form['classroom_id'])
        row.lesson_id=int(request.form['lesson_id']) if request.form.get('lesson_id') else None
        row.room_name=request.form.get('room_name',''); row.topic=request.form.get('topic','')
        db.session.commit(); flash('แก้ไขตารางสอนแล้ว', 'success')
        return redirect(url_for('schedule'))
    schedules = TeachingSchedule.query.filter_by(teacher_id=current_user.id).order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all() if current_user.role=='teacher' else TeachingSchedule.query.order_by(TeachingSchedule.weekday, TeachingSchedule.period_no).all()
    day_names = ['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์','อาทิตย์']
    periods = sorted({x.period_no for x in schedules}) or list(range(1, 9))
    period_times = {x.period_no: f"{x.start_time}-{x.end_time}" for x in schedules}
    schedule_grid = {f"{x.weekday}-{x.period_no}": x for x in schedules}
    day_pairs = list(enumerate(day_names[:5]))
    return render_template('schedule.html', schedules=schedules, subjects=subjects, rooms=rooms, lessons=lessons, edit_row=row, day_names=day_names, periods=periods, period_times=period_times, schedule_grid=schedule_grid, day_pairs=day_pairs)

@app.route('/schedule/<int:schedule_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def schedule_delete(schedule_id):
    row = TeachingSchedule.query.get_or_404(schedule_id)
    if current_user.role != 'admin' and row.teacher_id != current_user.id: return deny_redirect('schedule')
    db.session.delete(row); db.session.commit(); flash('ลบคาบสอนแล้ว', 'success')
    return redirect(url_for('schedule'))



def academic_calendar_1_2569_events():
    """ปฏิทินการศึกษา ภาคเรียนที่ 1/2569 ตามภาพตัวอย่างที่ผู้ใช้ให้ไว้
    ใช้ปี ค.ศ. 2026 ในระบบ แต่แสดงผลเป็น พ.ศ. 2569 ในหน้าเว็บได้
    """
    items = [
        # พฤษภาคม 2569
        ('2026-05-01','วันแรงงานแห่งชาติ','วันหยุดราชการ','พฤษภาคม 2569'),
        ('2026-05-04','วันฉัตรมงคล','วันหยุดราชการ','พฤษภาคม 2569'),
        ('2026-05-11','เปิดภาคเรียนที่ 1/2569','กิจกรรมโรงเรียน','พฤษภาคม 2569'),
        ('2026-05-13','วันพืชมงคล','วันหยุดราชการ','พฤษภาคม 2569'),
        ('2026-05-15','ส่งโครงการสอนและแผนการจัดการเรียน','กิจกรรมโรงเรียน','พฤษภาคม 2569'),
        ('2026-05-18','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-19','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-20','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-21','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-22','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-25','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-26','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-27','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-28','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-29','ยื่นคำร้องแก้ 0 ร มส.','กิจกรรมโรงเรียน','18-29 พฤษภาคม 2569'),
        ('2026-05-31','วันวิสาขบูชา','วันหยุดราชการ','พฤษภาคม 2569'),
        # มิถุนายน 2569
        ('2026-06-01','วันหยุดชดเชยวันวิสาขบูชา','วันหยุดราชการ','มิถุนายน 2569'),
        ('2026-06-02','ประกาศผลการแก้ 0 ร มส.','กิจกรรมโรงเรียน','มิถุนายน 2569'),
        ('2026-06-03','วันเฉลิมพระชนมพรรษาสมเด็จพระราชินี','วันหยุดราชการ','มิถุนายน 2569'),
        ('2026-06-08','คลินิก 0 ร มส.','กิจกรรมโรงเรียน','มิถุนายน 2569'),
        ('2026-06-10','ส่ง ปถ.05 ครั้งที่ 1 กำหนดคะแนน','กิจกรรมโรงเรียน','มิถุนายน 2569'),
        ('2026-06-15','PLC รอบที่ 1 กำหนดปัญหา','กิจกรรมโรงเรียน','15-19 มิถุนายน 2569'),
        ('2026-06-16','PLC รอบที่ 1 กำหนดปัญหา','กิจกรรมโรงเรียน','15-19 มิถุนายน 2569'),
        ('2026-06-17','PLC รอบที่ 1 กำหนดปัญหา','กิจกรรมโรงเรียน','15-19 มิถุนายน 2569'),
        ('2026-06-18','PLC รอบที่ 1 กำหนดปัญหา','กิจกรรมโรงเรียน','15-19 มิถุนายน 2569'),
        ('2026-06-19','PLC รอบที่ 1 กำหนดปัญหา','กิจกรรมโรงเรียน','15-19 มิถุนายน 2569'),
        ('2026-06-26','วันสุนทรภู่','กิจกรรมโรงเรียน','หยุดชดเชย 2 กรกฎาคม'),
        # กรกฎาคม 2569
        ('2026-07-01','วันคล้ายวันสถาปนาคณะลูกเสือแห่งชาติ','กิจกรรมโรงเรียน','กรกฎาคม 2569'),
        ('2026-07-06','สำรวจติดตามสอบกลางภาค','กิจกรรมโรงเรียน','กรกฎาคม 2569'),
        ('2026-07-16','สอบกลางภาค','กิจกรรมโรงเรียน','16-17 กรกฎาคม 2569'),
        ('2026-07-17','สอบกลางภาค','กิจกรรมโรงเรียน','16-17 กรกฎาคม 2569'),
        ('2026-07-20','PLC รอบที่ 2 สะท้อนแผน','กิจกรรมโรงเรียน','20-24 กรกฎาคม 2569'),
        ('2026-07-21','PLC รอบที่ 2 สะท้อนแผน','กิจกรรมโรงเรียน','20-24 กรกฎาคม 2569'),
        ('2026-07-22','PLC รอบที่ 2 สะท้อนแผน','กิจกรรมโรงเรียน','20-24 กรกฎาคม 2569'),
        ('2026-07-23','PLC รอบที่ 2 สะท้อนแผน','กิจกรรมโรงเรียน','20-24 กรกฎาคม 2569'),
        ('2026-07-24','ส่ง ปถ.05 ครั้งที่ 2 + คะแนนกลางภาค','กิจกรรมโรงเรียน','กรกฎาคม 2569'),
        ('2026-07-28','วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว','วันหยุดราชการ','กรกฎาคม 2569'),
        ('2026-07-29','วันอาสาฬหบูชา','วันหยุดราชการ','กรกฎาคม 2569'),
        ('2026-07-30','วันเข้าพรรษา','วันหยุดราชการ','กรกฎาคม 2569'),
        # สิงหาคม 2569
        ('2026-08-03','PLC รอบที่ 3 เปิดชั้นเรียน','กิจกรรมโรงเรียน','3-7 สิงหาคม 2569'),
        ('2026-08-04','PLC รอบที่ 3 เปิดชั้นเรียน','กิจกรรมโรงเรียน','3-7 สิงหาคม 2569'),
        ('2026-08-05','PLC รอบที่ 3 เปิดชั้นเรียน','กิจกรรมโรงเรียน','3-7 สิงหาคม 2569'),
        ('2026-08-06','PLC รอบที่ 3 เปิดชั้นเรียน','กิจกรรมโรงเรียน','3-7 สิงหาคม 2569'),
        ('2026-08-07','PLC รอบที่ 3 เปิดชั้นเรียน','กิจกรรมโรงเรียน','3-7 สิงหาคม 2569'),
        ('2026-08-12','วันแม่แห่งชาติ','วันหยุดราชการ','สิงหาคม 2569'),
        ('2026-08-18','สัปดาห์วันวิทยาศาสตร์','กิจกรรมโรงเรียน','สิงหาคม 2569'),
        # กันยายน 2569
        ('2026-09-07','สำรวจข้อสอบปลายภาค','กิจกรรมโรงเรียน','กันยายน 2569'),
        ('2026-09-08','ส่งก่อนปลายเรียน 60% 80%','กิจกรรมโรงเรียน','กันยายน 2569'),
        ('2026-09-09','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-10','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-11','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-14','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-15','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-16','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-17','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-18','แก้ มส.','กิจกรรมโรงเรียน','9-18 กันยายน 2569'),
        ('2026-09-21','ขออนุมัติผล มผ.','กิจกรรมโรงเรียน','กันยายน 2569'),
        ('2026-09-23','สอบปลายภาค','กิจกรรมโรงเรียน','23-25 กันยายน 2569'),
        ('2026-09-24','สอบปลายภาค','กิจกรรมโรงเรียน','23-25 กันยายน 2569'),
        ('2026-09-25','สอบปลายภาค','กิจกรรมโรงเรียน','23-25 กันยายน 2569'),
        # ตุลาคม 2569
        ('2026-10-01','ปิดภาคเรียนที่ 1/2569','กิจกรรมโรงเรียน','ตุลาคม 2569'),
        ('2026-10-09','ส่งผลการเรียน ปถ.05 + bookmark','กิจกรรมโรงเรียน','ตุลาคม 2569'),
        ('2026-10-13','วันคล้ายวันสวรรคต ร.9','วันหยุดราชการ','ตุลาคม 2569'),
        ('2026-10-23','วันปิยมหาราช','วันหยุดราชการ','ตุลาคม 2569'),
        ('2026-10-26','วันออกพรรษา','วันหยุดราชการ','ตุลาคม 2569'),
    ]
    return [(datetime.strptime(d, '%Y-%m-%d').date(), title, typ, note) for d, title, typ, note in items]

def seed_academic_calendar_1_2569(teacher_id=None):
    created = 0
    for event_date, title, event_type, note in academic_calendar_1_2569_events():
        exists = CalendarEvent.query.filter_by(teacher_id=teacher_id, event_date=event_date, title=title, event_type=event_type).first()
        if not exists:
            db.session.add(CalendarEvent(teacher_id=teacher_id, event_date=event_date, title=title, event_type=event_type, note=note))
            created += 1
    db.session.commit()
    return created

def working_day_summary(start=None, end=None, teacher_id=None, semester=None):
    semester = semester or get_active_semester()
    if semester and not start and not end:
        start, end = semester.start_date, semester.end_date
    start = start or date(2026,5,11)
    end = end or date(2026,10,1)
    holidays = set()
    q = CalendarEvent.query.filter(CalendarEvent.event_date>=start, CalendarEvent.event_date<=end, CalendarEvent.event_type.in_(['วันหยุดราชการ','วันหยุดโรงเรียน']))
    if teacher_id:
        q = q.filter((CalendarEvent.teacher_id==teacher_id) | (CalendarEvent.teacher_id==None))
    for e in q.all():
        holidays.add(e.event_date)
    summary = {}
    cur = start
    while cur <= end:
        key = (cur.year, cur.month)
        if key not in summary:
            summary[key] = {'month': f"{thai_month_name(cur.month)} {cur.year+543}", 'working':0, 'holiday':0, 'weekend':0}
        if cur.weekday() >= 5:
            summary[key]['weekend'] += 1
        elif cur in holidays:
            summary[key]['holiday'] += 1
        else:
            summary[key]['working'] += 1
        cur += timedelta(days=1)
    return list(summary.values())


@app.route('/semesters', methods=['GET','POST'])
@login_required
@role_required('admin')
def semesters():
    if request.method == 'POST':
        sem = Semester(
            name=request.form['name'],
            academic_year=request.form.get('academic_year','2569'),
            term_no=request.form.get('term_no','1'),
            start_date=datetime.strptime(request.form['start_date'],'%Y-%m-%d').date(),
            end_date=datetime.strptime(request.form['end_date'],'%Y-%m-%d').date(),
            is_active=bool(request.form.get('is_active'))
        )
        if sem.is_active:
            Semester.query.update({Semester.is_active: False})
        db.session.add(sem); db.session.commit(); flash('เพิ่มภาคเรียนแล้ว', 'success')
        return redirect(url_for('semesters'))
    return render_template('semesters.html', semesters=Semester.query.order_by(Semester.start_date.desc()).all(), edit_row=None)

@app.route('/semesters/<int:semester_id>/edit', methods=['GET','POST'])
@login_required
@role_required('admin')
def semester_edit(semester_id):
    sem = Semester.query.get_or_404(semester_id)
    if request.method == 'POST':
        sem.name=request.form['name']; sem.academic_year=request.form.get('academic_year','')
        sem.term_no=request.form.get('term_no','')
        sem.start_date=datetime.strptime(request.form['start_date'],'%Y-%m-%d').date()
        sem.end_date=datetime.strptime(request.form['end_date'],'%Y-%m-%d').date()
        sem.is_active=bool(request.form.get('is_active'))
        if sem.is_active:
            Semester.query.filter(Semester.id!=sem.id).update({Semester.is_active: False})
        db.session.commit(); flash('แก้ไขภาคเรียนแล้ว', 'success')
        return redirect(url_for('semesters'))
    return render_template('semesters.html', semesters=Semester.query.order_by(Semester.start_date.desc()).all(), edit_row=sem)

@app.route('/semesters/<int:semester_id>/active', methods=['POST'])
@login_required
@role_required('admin')
def semester_active(semester_id):
    Semester.query.update({Semester.is_active: False})
    sem = Semester.query.get_or_404(semester_id); sem.is_active=True
    db.session.commit(); flash(f'เปลี่ยนเป็น {sem.name} แล้ว', 'success')
    return redirect(url_for('semesters'))

@app.route('/semesters/<int:semester_id>/delete', methods=['POST'])
@login_required
@role_required('admin')
def semester_delete(semester_id):
    sem = Semester.query.get_or_404(semester_id)
    db.session.delete(sem); db.session.commit(); flash('ลบภาคเรียนแล้ว', 'success')
    return redirect(url_for('semesters'))

@app.route('/calendar', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def calendar_events():
    semesters = Semester.query.order_by(Semester.start_date.desc()).all()
    selected_semester = Semester.query.get(request.args.get('semester_id', type=int)) if request.args.get('semester_id') else get_active_semester()
    if request.method == 'POST':
        db.session.add(CalendarEvent(
            teacher_id=None if current_user.role=='admin' and request.form.get('global_event') else current_user.id,
            event_date=datetime.strptime(request.form['event_date'],'%Y-%m-%d').date(),
            title=request.form['title'], event_type=request.form.get('event_type','กิจกรรมโรงเรียน'), note=request.form.get('note','')
        ))
        db.session.commit(); flash('เพิ่มรายการปฏิทินแล้ว', 'success'); return redirect(url_for('calendar_events', semester_id=selected_semester.id if selected_semester else None))
    q = CalendarEvent.query if current_user.role=='admin' else CalendarEvent.query.filter((CalendarEvent.teacher_id==current_user.id) | (CalendarEvent.teacher_id==None))
    if selected_semester:
        q = q.filter(CalendarEvent.event_date>=selected_semester.start_date, CalendarEvent.event_date<=selected_semester.end_date)
    events = q.order_by(CalendarEvent.event_date.desc()).all()
    summary = working_day_summary(teacher_id=current_user.id if current_user.role=='teacher' else None, semester=selected_semester)
    images = CalendarImage.query.filter_by(semester_id=selected_semester.id).order_by(CalendarImage.uploaded_at.desc()).all() if selected_semester else []
    return render_template('calendar.html', events=events, event=None, summary=summary, semesters=semesters, selected_semester=selected_semester, images=images)

@app.route('/calendar/import-excel', methods=['POST'])
@login_required
@role_required('teacher','admin')
def calendar_import_excel():
    file = request.files.get('file')
    semester_id = request.form.get('semester_id', type=int)
    sem = Semester.query.get(semester_id) if semester_id else get_active_semester()
    if not file or not file.filename:
        flash('กรุณาเลือกไฟล์ Excel', 'danger'); return redirect(url_for('calendar_events', semester_id=semester_id))
    wb = load_workbook(file)
    ws = wb.active
    created = updated = 0
    # รองรับหัวตาราง: date/event_date, title, event_type, note
    headers = [str(c.value).strip().lower() if c.value else '' for c in ws[1]]
    def col(name, fallback):
        for n in name:
            if n in headers: return headers.index(n)
        return fallback
    date_i = col(['date','event_date','วันที่'], 0)
    title_i = col(['title','รายการ','กิจกรรม'], 1)
    type_i = col(['event_type','ประเภท'], 2)
    note_i = col(['note','หมายเหตุ'], 3)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[date_i] or not row[title_i]:
            continue
        raw_date = row[date_i]
        if isinstance(raw_date, datetime):
            d = raw_date.date()
        elif isinstance(raw_date, date):
            d = raw_date
        else:
            txt = str(raw_date).strip().replace('/', '-')
            parts = txt.split('-')
            if len(parts[0]) == 4:
                d = datetime.strptime(txt, '%Y-%m-%d').date()
            else:
                d = datetime.strptime(txt, '%d-%m-%Y').date()
            if d.year > 2400:
                d = date(d.year-543, d.month, d.day)
        title = str(row[title_i]).strip()
        typ = str(row[type_i]).strip() if len(row) > type_i and row[type_i] else 'กิจกรรมโรงเรียน'
        note = str(row[note_i]).strip() if len(row) > note_i and row[note_i] else ''
        teacher_id = None if current_user.role == 'admin' else current_user.id
        e = CalendarEvent.query.filter_by(teacher_id=teacher_id, event_date=d, title=title).first()
        if e:
            e.event_type = typ; e.note = note; updated += 1
        else:
            db.session.add(CalendarEvent(teacher_id=teacher_id, event_date=d, title=title, event_type=typ, note=note)); created += 1
    db.session.commit(); flash(f'นำเข้าปฏิทินแล้ว เพิ่ม {created} รายการ / อัปเดต {updated} รายการ', 'success')
    return redirect(url_for('calendar_events', semester_id=sem.id if sem else None))

@app.route('/calendar/upload-image', methods=['POST'])
@login_required
@role_required('teacher','admin')
def calendar_upload_image():
    file = request.files.get('image')
    semester_id = request.form.get('semester_id', type=int)
    if not file or not file.filename:
        flash('กรุณาเลือกรูปปฏิทิน', 'danger'); return redirect(url_for('calendar_events', semester_id=semester_id))
    filename = secure_filename(file.filename)
    stamp = datetime.now().strftime('%Y%m%d%H%M%S')
    save_name = f'calendar_{stamp}_{filename}'
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], save_name)
    file.save(save_path)
    db.session.add(CalendarImage(semester_id=semester_id, file_path=save_name, original_name=filename))
    db.session.commit()
    flash('บันทึกรูปปฏิทินไว้เป็นหลักฐานแล้ว หากต้องการอัปเดตข้อมูลอัตโนมัติให้ใช้ไฟล์ Excel นำเข้า', 'success')
    return redirect(url_for('calendar_events', semester_id=semester_id))

@app.route('/calendar/seed-1-2569', methods=['POST'])
@login_required
@role_required('teacher','admin')
def calendar_seed_1_2569():
    # admin สร้างเป็นปฏิทินกลาง ทุกครูเห็น / teacher สร้างเฉพาะตัวเอง
    teacher_id = None if current_user.role == 'admin' else current_user.id
    created = seed_academic_calendar_1_2569(teacher_id=teacher_id)
    flash(f'นำเข้าปฏิทินการศึกษา 1/2569 แล้ว เพิ่มใหม่ {created} รายการ', 'success')
    return redirect(url_for('calendar_events'))

@app.route('/calendar/<int:event_id>/edit', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def calendar_edit(event_id):
    event = CalendarEvent.query.get_or_404(event_id)
    if current_user.role!='admin' and event.teacher_id not in (current_user.id, None): return deny_redirect('calendar_events')
    if request.method == 'POST':
        event.event_date=datetime.strptime(request.form['event_date'],'%Y-%m-%d').date(); event.title=request.form['title']; event.event_type=request.form.get('event_type','กิจกรรมโรงเรียน'); event.note=request.form.get('note','')
        db.session.commit(); flash('แก้ไขปฏิทินแล้ว', 'success'); return redirect(url_for('calendar_events'))
    semesters = Semester.query.order_by(Semester.start_date.desc()).all()
    selected_semester = get_active_semester()
    q = CalendarEvent.query if current_user.role=='admin' else CalendarEvent.query.filter((CalendarEvent.teacher_id==current_user.id) | (CalendarEvent.teacher_id==None))
    if selected_semester:
        q = q.filter(CalendarEvent.event_date>=selected_semester.start_date, CalendarEvent.event_date<=selected_semester.end_date)
    events = q.order_by(CalendarEvent.event_date.desc()).all()
    summary = working_day_summary(teacher_id=current_user.id if current_user.role=='teacher' else None, semester=selected_semester)
    images = CalendarImage.query.filter_by(semester_id=selected_semester.id).order_by(CalendarImage.uploaded_at.desc()).all() if selected_semester else []
    return render_template('calendar.html', events=events, event=event, summary=summary, semesters=semesters, selected_semester=selected_semester, images=images)

@app.route('/calendar/<int:event_id>/delete', methods=['POST'])
@login_required
@role_required('teacher','admin')
def calendar_delete(event_id):
    event=CalendarEvent.query.get_or_404(event_id)
    if current_user.role!='admin' and event.teacher_id != current_user.id: return deny_redirect('calendar_events')
    db.session.delete(event); db.session.commit(); flash('ลบรายการปฏิทินแล้ว', 'success')
    return redirect(url_for('calendar_events'))



@app.route('/records')
@login_required
@role_required('teacher','admin')
def records_center():
    subject_ids = teacher_subject_ids()
    room_ids = teacher_classroom_ids()
    pairs = SubjectClassroom.query.filter(SubjectClassroom.subject_id.in_(subject_ids or [-1]), SubjectClassroom.classroom_id.in_(room_ids or [-1])).all()
    return render_template('records_center.html', pairs=pairs)

@app.route('/attendance/<int:subject_id>/<int:classroom_id>', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def attendance(subject_id, classroom_id):
    subject = Subject.query.get_or_404(subject_id); room = Classroom.query.get_or_404(classroom_id)
    if not owns_subject(subject) or not owns_classroom(room): return deny_redirect('records_center')
    links = ClassroomStudent.query.filter_by(classroom_id=room.id).all()
    att_date = datetime.strptime(request.args.get('date', date.today().isoformat()), '%Y-%m-%d').date()
    if request.method == 'POST':
        att_date = datetime.strptime(request.form['date'], '%Y-%m-%d').date()
        bulk_status = request.form.get('bulk_status')
        for link in links:
            st = bulk_status if bulk_status else request.form.get(f's_{link.student_id}', 'มา')
            row = Attendance.query.filter_by(subject_id=subject.id, classroom_id=room.id, student_id=link.student_id, date=att_date).first()
            if not row:
                row = Attendance(subject_id=subject.id, classroom_id=room.id, student_id=link.student_id, date=att_date)
                db.session.add(row)
            row.status = st
        db.session.commit(); flash('บันทึกเวลาเรียนแล้ว', 'success')
        return redirect(url_for('attendance', subject_id=subject.id, classroom_id=room.id, date=att_date.isoformat()))
    old = {a.student_id:a.status for a in Attendance.query.filter_by(subject_id=subject.id, classroom_id=room.id, date=att_date).all()}
    summary = []
    for link in links:
        atts = Attendance.query.filter_by(subject_id=subject.id, classroom_id=room.id, student_id=link.student_id).all()
        summary.append({'student': link.student, 'มา': sum(1 for a in atts if a.status=='มา'), 'ขาด': sum(1 for a in atts if a.status=='ขาด'), 'ลา': sum(1 for a in atts if a.status=='ลา'), 'มาสาย': sum(1 for a in atts if a.status=='มาสาย'), 'กิจกรรม': sum(1 for a in atts if a.status=='กิจกรรม')})
    return render_template('attendance.html', subject=subject, room=room, links=links, old=old, att_date=att_date, summary=summary)

@app.route('/attendance/<int:subject_id>/<int:classroom_id>/delete-day', methods=['POST'])
@login_required
@role_required('teacher','admin')
def attendance_delete_day(subject_id, classroom_id):
    subject = Subject.query.get_or_404(subject_id); room = Classroom.query.get_or_404(classroom_id)
    if not owns_subject(subject) or not owns_classroom(room): return deny_redirect('records_center')
    att_date = datetime.strptime(request.form['date'], '%Y-%m-%d').date()
    Attendance.query.filter_by(subject_id=subject.id, classroom_id=room.id, date=att_date).delete()
    db.session.commit(); flash('ลบเช็กชื่อของวันที่เลือกแล้ว', 'success')
    return redirect(url_for('attendance', subject_id=subject.id, classroom_id=room.id, date=att_date.isoformat()))

@app.route('/attendance/<int:subject_id>/<int:classroom_id>/export')
@login_required
@role_required('teacher','admin')
def export_attendance(subject_id, classroom_id):
    subject = Subject.query.get_or_404(subject_id); room = Classroom.query.get_or_404(classroom_id)
    if not owns_subject(subject) or not owns_classroom(room): return deny_redirect('records_center')
    wb=Workbook(); ws=wb.active; ws.title='attendance'
    ws.append(['ชื่อ','มา','ขาด','ลา','มาสาย','กิจกรรม'])
    links = ClassroomStudent.query.filter_by(classroom_id=room.id).all()
    for link in links:
        atts = Attendance.query.filter_by(subject_id=subject.id, classroom_id=room.id, student_id=link.student_id).all()
        ws.append([link.student.full_name, sum(1 for a in atts if a.status=='มา'), sum(1 for a in atts if a.status=='ขาด'), sum(1 for a in atts if a.status=='ลา'), sum(1 for a in atts if a.status=='มาสาย'), sum(1 for a in atts if a.status=='กิจกรรม')])
    path=os.path.join(UPLOAD_DIR, f'attendance_{subject_id}_{classroom_id}.xlsx'); wb.save(path)
    return send_file(path, as_attachment=True)

@app.route('/grades/<int:subject_id>/<int:classroom_id>', methods=['GET','POST'])
@login_required
@role_required('teacher','admin')
def grades(subject_id, classroom_id):
    subject = Subject.query.get_or_404(subject_id); room = Classroom.query.get_or_404(classroom_id)
    if not owns_subject(subject) or not owns_classroom(room): return deny_redirect('records_center')
    links = ClassroomStudent.query.filter_by(classroom_id=room.id).all()
    setting = get_grade_setting(subject.id)
    if request.method == 'POST':
        setting.worksheet_weight = float(request.form.get('worksheet_weight', setting.worksheet_weight))
        setting.quiz_weight = float(request.form.get('quiz_weight', setting.quiz_weight))
        setting.attendance_weight = float(request.form.get('attendance_weight', setting.attendance_weight))
        setting.midterm_weight = float(request.form.get('midterm_weight', setting.midterm_weight))
        setting.final_weight = float(request.form.get('final_weight', setting.final_weight))
        for link in links:
            manual = get_manual_score(subject.id, link.student_id)
            manual.midterm = float(request.form.get(f'midterm_{link.student_id}', manual.midterm or 0) or 0)
            manual.final = float(request.form.get(f'final_{link.student_id}', manual.final or 0) or 0)
            manual.behavior = float(request.form.get(f'behavior_{link.student_id}', manual.behavior or 0) or 0)
        db.session.commit(); flash('บันทึกสมุดคะแนนแล้ว', 'success')
        return redirect(url_for('grades', subject_id=subject.id, classroom_id=room.id))
    rows=[]
    for link in links:
        rows.append(calculate_grade_row(subject, room, link.student))
    total_weight = setting.worksheet_weight + setting.quiz_weight + setting.attendance_weight + setting.midterm_weight + setting.final_weight
    return render_template('grades.html', subject=subject, room=room, rows=rows, setting=setting, total_weight=total_weight)

@app.route('/export_grades/<int:subject_id>/<int:classroom_id>')
@login_required
@role_required('teacher','admin')
def export_grades(subject_id, classroom_id):
    subject = Subject.query.get_or_404(subject_id); room = Classroom.query.get_or_404(classroom_id)
    if not owns_subject(subject) or not owns_classroom(room): return deny_redirect('records_center')
    wb=Workbook(); ws=wb.active; ws.title='grades'
    ws.append(['ชื่อ','แบบทดสอบเฉลี่ย','เรียนจบ/งานทั้งหมด','มา','ขาด','ลา','มาสาย','กิจกรรม','กลางภาค','ปลายภาค','จิตพิสัย','คะแนนรวม','เกรด'])
    links = ClassroomStudent.query.filter_by(classroom_id=room.id).all()
    for link in links:
        r=calculate_grade_row(subject, room, link.student)
        ws.append([r['student'].full_name, r['quiz_avg'], f"{r['complete']}/{r['total_assignments']}", r['present'], r['absent'], r['leave'], r['late'], r['activity'], r['manual'].midterm, r['manual'].final, r['manual'].behavior, r['total'], r['grade']])
    path=os.path.join(UPLOAD_DIR, f'grades_{subject_id}_{classroom_id}.xlsx'); wb.save(path)
    return send_file(path, as_attachment=True)

@app.cli.command('init-db')
def init_db_cmd():
    init_db(); print('initialized')


def seed_prem_schedule_for_teacher(t):
    """สร้างรายวิชา ห้องเรียน และตารางสอนตามภาพตัวอย่างที่ผู้ใช้ให้ไว้"""
    period_times = {
        1: ('08:40','09:30'),
        2: ('09:31','10:20'),
        3: ('10:21','11:10'),
        4: ('11:11','12:00'),
        5: ('13:00','13:50'),
        6: ('13:51','14:40'),
        7: ('14:41','15:30'),
        8: ('15:31','16:00'),
    }

    def get_subject(name):
        sub = Subject.query.filter_by(name=name).first()
        if not sub:
            sub = Subject(name=name, teacher_id=None, credit=1.0, total_periods=40)
            db.session.add(sub); db.session.flush()
            db.session.add(GradeSetting(subject_id=sub.id))
        if not TeacherSubject.query.filter_by(teacher_id=t.id, subject_id=sub.id).first():
            db.session.add(TeacherSubject(teacher_id=t.id, subject_id=sub.id))
        return sub

    def get_room(name):
        room = Classroom.query.filter_by(name=name).first()
        if not room:
            room = Classroom(name=name, teacher_id=None)
            db.session.add(room); db.session.flush()
        if not TeacherClassroom.query.filter_by(teacher_id=t.id, classroom_id=room.id).first():
            db.session.add(TeacherClassroom(teacher_id=t.id, classroom_id=room.id))
        return room

    def link_subject_room(sub, room):
        if not SubjectClassroom.query.filter_by(subject_id=sub.id, classroom_id=room.id).first():
            db.session.add(SubjectClassroom(subject_id=sub.id, classroom_id=room.id))

    def add_schedule(weekday, period_no, subject_name, room_name, location='', topic=''):
        sub = get_subject(subject_name)
        room = get_room(room_name)
        link_subject_room(sub, room)
        st, en = period_times[period_no]
        exists = TeachingSchedule.query.filter_by(
            teacher_id=t.id, weekday=weekday, period_no=period_no,
            subject_id=sub.id, classroom_id=room.id
        ).first()
        if not exists:
            db.session.add(TeachingSchedule(
                teacher_id=t.id,
                subject_id=sub.id,
                classroom_id=room.id,
                weekday=weekday,
                period_no=period_no,
                start_time=st,
                end_time=en,
                room_name=location or room_name,
                topic=topic or subject_name
            ))

    # จันทร์
    add_schedule(0, 1, 'ว21104 วิทยาการคำนวณ 1', 'ม.1/2-3', 'ห้อง ม.1/2-3')
    add_schedule(0, 3, 'ว22104 วิทยาการคำนวณ 2', 'ม.2/2-3', 'ห้อง ม.2/2-3')
    add_schedule(0, 4, 'ว22104 วิทยาการคำนวณ 2', 'ม.2/1', 'ห้อง ม.2/1')
    add_schedule(0, 5, 'ว21104 วิทยาการคำนวณ 1', 'ม.1/1', 'ห้อง ม.1/1')
    add_schedule(0, 8, 'บำเพ็ญประโยชน์ PLC', 'PLC', 'PLC', 'บำเพ็ญประโยชน์ / PLC')

    # อังคาร
    add_schedule(1, 1, 'ว23104 วิทยาการคำนวณ 3', 'ม.3/2', 'ห้อง ม.3/2')
    add_schedule(1, 3, 'พ21202 เปตอง 2', 'ม.1/3', 'ห้อง ม.1/3')
    add_schedule(1, 4, 'พ21202 เปตอง 2', 'ม.1/3', 'ห้อง ม.1/3')
    add_schedule(1, 8, 'บำเพ็ญประโยชน์ PLC', 'PLC', 'PLC', 'บำเพ็ญประโยชน์ / PLC')

    # พุธ
    add_schedule(2, 3, 'พ23206 คอมพิวเตอร์ 6', 'ม.3/1', 'ห้อง ม.3/1')
    add_schedule(2, 4, 'พ23206 คอมพิวเตอร์ 6', 'ม.3/1', 'ห้อง ม.3/1')
    add_schedule(2, 6, 'ว31103 วิทยาการคำนวณ', 'ม.4/1', 'ห้อง ม.4/1')
    add_schedule(2, 7, 'ลูกเสือเนตรนารี', 'กิจกรรมรวม', 'กิจกรรมรวม', 'ลูกเสือเนตรนารี')
    add_schedule(2, 8, 'บำเพ็ญประโยชน์ PLC', 'PLC', 'PLC', 'บำเพ็ญประโยชน์ / PLC')

    # พฤหัสบดี
    add_schedule(3, 2, 'ว32104 วิทยาการคำนวณ 2', 'ม.5/1-3', 'ห้อง ม.5/1-3')
    add_schedule(3, 3, 'พ23206 คอมพิวเตอร์ 6', 'ม.3/2', 'ห้อง ม.3/2')
    add_schedule(3, 4, 'พ23206 คอมพิวเตอร์ 6', 'ม.3/2', 'ห้อง ม.3/2')
    add_schedule(3, 7, 'ชุมนุม/คณะสี', 'กิจกรรมรวม', 'กิจกรรมรวม', 'ชุมนุม/คณะสี')
    add_schedule(3, 8, 'บำเพ็ญประโยชน์ PLC', 'PLC', 'PLC', 'บำเพ็ญประโยชน์ / PLC')

    # ศุกร์
    add_schedule(4, 3, 'พ22202 เปตอง 4', 'ม.2/3', 'ห้อง ม.2/3')
    add_schedule(4, 4, 'พ22202 เปตอง 4', 'ม.2/3', 'ห้อง ม.2/3')
    add_schedule(4, 5, 'ว23104 วิทยาการคำนวณ 3', 'ม.3/1', 'ห้อง ม.3/1')
    add_schedule(4, 6, 'ว31103 วิทยาการคำนวณ', 'ม.4/2-3', 'ห้อง ม.4/2-3')
    add_schedule(4, 7, 'อบรมคุณธรรม', 'กิจกรรมรวม', 'กิจกรรมรวม', 'อบรมคุณธรรม')
    add_schedule(4, 8, 'บำเพ็ญประโยชน์ PLC', 'PLC', 'PLC', 'บำเพ็ญประโยชน์ / PLC')

def seed():
    if Semester.query.count() == 0:
        db.session.add(Semester(name='ภาคเรียนที่ 1/2569', academic_year='2569', term_no='1', start_date=date(2026,5,11), end_date=date(2026,10,1), is_active=True))
        db.session.commit()
    if not User.query.filter_by(username='admin').first():
        admin=User(username='admin', full_name='ผู้ดูแลระบบ', role='admin'); admin.set_password('1234'); db.session.add(admin)
    if not User.query.filter_by(username='dekchairukna').first():
        t=User(username='dekchairukna', full_name='นายพศิน พิมพ์คำไหล', role='teacher'); t.set_password('1225'); db.session.add(t); db.session.flush()
        room=Classroom(name='ม.1/1', teacher_id=None); db.session.add(room); db.session.flush(); db.session.add(TeacherClassroom(teacher_id=t.id, classroom_id=room.id))
        sub=Subject(name='วิทยาการคำนวณ', teacher_id=None, credit=1.0, total_periods=40); db.session.add(sub); db.session.flush(); db.session.add(TeacherSubject(teacher_id=t.id, subject_id=sub.id)); db.session.add(GradeSetting(subject_id=sub.id))
        u=Unit(subject_id=sub.id, title='หน่วยที่ 1 อัลกอริทึม', indicators='ว 4.2 เข้าใจและใช้อัลกอริทึม', required_periods=3); db.session.add(u); db.session.flush()
        l=Lesson(unit_id=u.id, title='อัลกอริทึมเบื้องต้น', objective='อธิบายความหมายของอัลกอริทึมได้', content='อัลกอริทึมคือขั้นตอนการแก้ปัญหาอย่างเป็นลำดับ'); db.session.add(l); db.session.flush()
        ws=Worksheet(lesson_id=l.id, title='ใบงานอัลกอริทึม', worksheet_type='text', description='ตอบคำถามตามข้อ'); db.session.add(ws); db.session.flush()
        db.session.add(WorksheetQuestion(worksheet_id=ws.id, number=1, question_text='อัลกอริทึมคืออะไร'))
        qz=Quiz(lesson_id=l.id, title='แบบทดสอบท้ายบท', pass_percent=60); db.session.add(qz); db.session.flush()
        db.session.add(QuizQuestion(quiz_id=qz.id, question_text='อัลกอริทึมหมายถึงอะไร', question_type='abcd', choices='ก. ขั้นตอนแก้ปัญหา\nข. เครื่องพิมพ์\nค. จอภาพ\nง. เมาส์', correct_answer='ก'))
    if not User.query.filter_by(username='teacher').first():
        sample=User(username='teacher', full_name='ครูตัวอย่าง', role='teacher'); sample.set_password('1234'); db.session.add(sample); db.session.flush()
    if not User.query.filter_by(username='student').first():
        s=User(username='student', full_name='นักเรียนตัวอย่าง', role='student'); s.set_password('1234'); db.session.add(s); db.session.flush()
        room=Classroom.query.filter_by(name='ม.1/1').first()
        if room: db.session.add(ClassroomStudent(classroom_id=room.id, student_id=s.id))
    t = User.query.filter_by(username='dekchairukna').first()
    if t:
        seed_prem_schedule_for_teacher(t)
    # ปฏิทินกลาง 1/2569 ตามภาพตัวอย่าง
    seed_academic_calendar_1_2569(teacher_id=None)
    db.session.commit()


def ensure_schema_columns():
    # รองรับการอัปเกรดจากเวอร์ชันเก่าที่มีฐานข้อมูลเดิมอยู่แล้ว
    with db.engine.connect() as conn:
        def cols(table):
            return {row[1] for row in conn.exec_driver_sql(f"PRAGMA table_info({table})").fetchall()}
        if 'is_active' not in cols('subject'):
            conn.exec_driver_sql("ALTER TABLE subject ADD COLUMN is_active BOOLEAN DEFAULT 1")
        if 'is_active' not in cols('classroom'):
            conn.exec_driver_sql("ALTER TABLE classroom ADD COLUMN is_active BOOLEAN DEFAULT 1")
        if 'file_path' not in cols('worksheet'):
            conn.exec_driver_sql("ALTER TABLE worksheet ADD COLUMN file_path VARCHAR(500) DEFAULT ''")
        if 'original_file_name' not in cols('worksheet'):
            conn.exec_driver_sql("ALTER TABLE worksheet ADD COLUMN original_file_name VARCHAR(255) DEFAULT ''")
        if 'file_path' not in cols('worksheet_answer'):
            conn.exec_driver_sql("ALTER TABLE worksheet_answer ADD COLUMN file_path VARCHAR(500) DEFAULT ''")
        if 'original_file_name' not in cols('worksheet_answer'):
            conn.exec_driver_sql("ALTER TABLE worksheet_answer ADD COLUMN original_file_name VARCHAR(255) DEFAULT ''")
        if 'assignment_type' not in cols('assignment'):
            conn.exec_driver_sql("ALTER TABLE assignment ADD COLUMN assignment_type VARCHAR(30) DEFAULT 'special'")
        conn.commit()

def init_db():
    db.create_all(); ensure_schema_columns(); seed()

if __name__ == '__main__':
    with app.app_context():
        init_db()
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 8000)))
