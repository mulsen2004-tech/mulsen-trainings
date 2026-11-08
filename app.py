import os
import io
import base64
import json
import smtplib
import ssl
import secrets
import string
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from datetime import datetime
from functools import wraps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, send_file, jsonify, abort)
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from pptx import Presentation
from PIL import Image
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
try:
    import pymupdf as fitz
    PDF_SUPPORT = True
except ImportError:
    try:
        import fitz
        PDF_SUPPORT = True
    except ImportError:
        PDF_SUPPORT = False
import csv
import re
from docx import Document as DocxDocument
from docx.shared import Pt as DocxPt

try:
    from openai import OpenAI  # type: ignore
    OPENAI_AVAILABLE = True
except ImportError:
    OpenAI = None              # type: ignore
    OPENAI_AVAILABLE = False

# ===== AMHARIC: Azure Speech REST =====
try:
    import requests as _requests
    REQUESTS_AVAILABLE = True
except ImportError:
    _requests = None
    REQUESTS_AVAILABLE = False
# ======================================

try:
    import qrcode  # type: ignore
    from qrcode.constants import ERROR_CORRECT_M  # type: ignore
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

_openai_client = None


# ---------------- APP SETUP ----------------

app = Flask(__name__)
app.config['APP_NAME'] = 'MULSEN Virtual Training App'
app.config['SECRET_KEY'] = 'your-super-secret-key-change-in-production'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///training.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

app.config['MAIL_HOST'] = os.getenv('MAIL_HOST', '')
app.config['MAIL_PORT'] = int(os.getenv('MAIL_PORT', '587'))
app.config['MAIL_USER'] = os.getenv('MAIL_USER', '')
app.config['MAIL_PASS'] = os.getenv('MAIL_PASS', '')
app.config['MAIL_FROM'] = os.getenv('MAIL_FROM', 'no-reply@mulsen.local')
app.config['MAIL_FROM_NAME'] = os.getenv('MAIL_FROM_NAME', 'MULSEN Training')
app.config['MAIL_USE_TLS'] = os.getenv('MAIL_USE_TLS', '1') == '1'
app.config['MAIL_USE_SSL'] = os.getenv('MAIL_USE_SSL', '0') == '1'

# ===== AMHARIC: Azure Speech config =====
app.config['AZURE_SPEECH_KEY'] = os.getenv('AZURE_SPEECH_KEY', '')
app.config['AZURE_SPEECH_REGION'] = os.getenv('AZURE_SPEECH_REGION', 'southafricanorth')
app.config['AZURE_TTS_VOICE_AM'] = os.getenv('AZURE_TTS_VOICE_AM', 'am-ET-MekdesNeural')
# =========================================

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('certificates', exist_ok=True)

db = SQLAlchemy(app)


# ---------------- MODELS ----------------

class Trainer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    trainer_id = db.Column(db.String(5), unique=True, nullable=False, index=True)
    email = db.Column(db.String(255), unique=True, nullable=True, index=True)
    full_name = db.Column(db.String(200), default='')
    department = db.Column(db.String(200), default='')
    position = db.Column(db.String(200), default='')
    password_hash = db.Column(db.String(255), nullable=False)
    name = db.Column(db.String(100), default='Trainee')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def display_name(self):
        return self.full_name or self.name or self.trainer_id


class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)


class Course(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, default='')
    ppt_filename = db.Column(db.String(300))
    source_type = db.Column(db.String(10), default='pptx')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    questions = db.relationship('Question', backref='course', cascade='all, delete-orphan')
    attempts  = db.relationship('Attempt',  backref='course', cascade='all, delete-orphan')
    progress  = db.relationship('Progress', backref='course', cascade='all, delete-orphan')
    page_recordings = db.relationship('PageRecording', backref='course',
                                     cascade='all, delete-orphan')


class Question(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    course_id = db.Column(db.Integer, db.ForeignKey('course.id'), nullable=False)
    question_type = db.Column(db.String(5), default='mc')
    question_text = db.Column(db.Text, nullable=False)
    option_a = db.Column(db.String(500))
    option_b = db.Column(db.String(500))
    option_c = db.Column(db.String(500))
    option_d = db.Column(db.String(500))
    correct_answer = db.Column(db.String(1), nullable=False)


class Attempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    trainer_id = db.Column(db.Integer, db.ForeignKey('trainer.id'),
                           nullable=False, index=True)
    course_id  = db.Column(db.Integer, db.ForeignKey('course.id'),
                           nullable=False, index=True)
    score = db.Column(db.Float, nullable=False)
    total = db.Column(db.Integer, nullable=False)
    percentage = db.Column(db.Float, nullable=False)
    passed = db.Column(db.Boolean, default=False, index=True)
    distinction = db.Column(db.Boolean, default=False)
    completed_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    trainer = db.relationship('Trainer', backref='attempts')


class Progress(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    trainer_id = db.Column(db.Integer, db.ForeignKey('trainer.id'),
                           nullable=False, index=True)
    course_id  = db.Column(db.Integer, db.ForeignKey('course.id'),
                           nullable=False, index=True)
    slide_number = db.Column(db.Integer, nullable=False)
    completed = db.Column(db.Boolean, default=False)
    paused_at = db.Column(db.DateTime, nullable=True)
    last_slide = db.Column(db.Integer, default=1)
    last_seen = db.Column(db.DateTime, nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint('trainer_id', 'course_id', 'slide_number'),)


class InstructorScript(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    course_id = db.Column(db.Integer, db.ForeignKey('course.id'), nullable=False)
    page_number = db.Column(db.Integer, nullable=False)
    narration_text = db.Column(db.Text, nullable=False)
    # ===== AMHARIC: per-page language ('en' or 'am') =====
    language = db.Column(db.String(5), default='en', nullable=False)
    # =====================================================
    __table_args__ = (db.UniqueConstraint('course_id', 'page_number'),)
    course = db.relationship('Course', backref='instructor_scripts')


class PageRecording(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    course_id   = db.Column(db.Integer, db.ForeignKey('course.id'), nullable=False)
    page_number = db.Column(db.Integer, nullable=False)
    filename    = db.Column(db.String(300), nullable=False)
    mimetype    = db.Column(db.String(60), default='audio/webm')
    recorded_by = db.Column(db.String(120))
    created_at  = db.Column(db.DateTime, default=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint('course_id', 'page_number'),)


# ---------------- AUTH DECORATORS ----------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'trainer_id' not in session:
            flash('Please log in first.', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'admin_id' not in session:
            flash('Admin login required.', 'warning')
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated


def any_logged_in_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'admin_id' not in session and 'trainer_id' not in session:
            flash('Please log in first.', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# ---------------- CONTEXT PROCESSOR ----------------

@app.context_processor
def inject_app():
    role = None
    if 'admin_id' in session:
        role = 'admin'
    elif 'trainer_id' in session:
        role = 'trainee'
    return {
        'APP_NAME': app.config['APP_NAME'],
        'current_role': role,
        'current_trainer': session.get('trainer_name'),
        'current_trainer_id': session.get('trainer_id'),
        'current_admin': session.get('admin_username'),
    }


# ---------------- SLIDE CONVERTERS ----------------

def ppt_to_slides(ppt_path):
    prs = Presentation(ppt_path)
    slides_data = []
    for idx, slide in enumerate(prs.slides, start=1):
        text_parts = []
        image_data = None
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    line = ''.join(run.text for run in para.runs)
                    if line.strip():
                        text_parts.append(line)
            if shape.shape_type == 13:
                try:
                    img = Image.open(io.BytesIO(shape.image.blob))
                    if img.mode in ('RGBA', 'P'):
                        img = img.convert('RGB')
                    img.thumbnail((1000, 600))
                    buf = io.BytesIO()
                    img.save(buf, format='JPEG', quality=85)
                    image_data = base64.b64encode(buf.getvalue()).decode('utf-8')
                except Exception as e:
                    print(f"Image extract error: {e}")
        slides_data.append({
            'number': idx,
            'text': '\n'.join(text_parts) if text_parts else ' ',
            'image': image_data,
        })
    return slides_data


def pdf_to_slides(pdf_path):
    doc = fitz.open(pdf_path)
    slides_data = []
    for idx, page in enumerate(doc, start=1):
        text = page.get_text().strip()
        pix = page.get_pixmap(dpi=110)
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        img.thumbnail((1200, 800))
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=85)
        image_data = base64.b64encode(buf.getvalue()).decode('utf-8')
        slides_data.append({
            'number': idx,
            'text': text if text else ' ',
            'image': image_data,
        })
    doc.close()
    return slides_data


def count_slides(course):
    if not course.ppt_filename:
        return None
    path = os.path.join(app.config['UPLOAD_FOLDER'], course.ppt_filename)
    if not os.path.exists(path):
        return None
    try:
        if course.source_type == 'pdf':
            doc = fitz.open(path)
            n = len(doc)
            doc.close()
            return n
        return len(Presentation(path).slides)
    except Exception as e:
        print(f"[count_slides] {e}")
        return None


# ---------------- MANUAL SLIDE COUNT ----------------

MANUAL_COUNTS_FILE = os.path.join(app.config['UPLOAD_FOLDER'], 'manual_slide_counts.json')


def _load_manual_counts():
    if not os.path.exists(MANUAL_COUNTS_FILE):
        return {}
    try:
        with open(MANUAL_COUNTS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_manual_counts(data):
    try:
        with open(MANUAL_COUNTS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f)
    except Exception as e:
        print(f'[manual_counts] save failed: {e}')


def effective_slide_count(course, scripts=None, recs=None):
    n = count_slides(course)
    if n is not None:
        return n
    manual = _load_manual_counts().get(str(course.id))
    if manual:
        return int(manual)
    if scripts is None:
        scripts = InstructorScript.query.filter_by(course_id=course.id).all()
    if scripts:
        return max(s.page_number for s in scripts)
    if recs is None:
        recs = PageRecording.query.filter_by(course_id=course.id).all()
    if recs:
        return max(r.page_number for r in recs)
    return None


def parse_instructor_script(file_path, filename):
    ext = filename.lower().rsplit('.', 1)[-1]
    raw_pages = {}

    if ext == 'csv':
        with open(file_path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    p = int((row.get('page') or '').strip())
                except ValueError:
                    continue
                txt = (row.get('narration') or '').strip()
                if txt:
                    raw_pages.setdefault(p, []).append(txt)
    else:
        if ext == 'docx':
            doc = DocxDocument(file_path)
            lines = [p.text for p in doc.paragraphs]
        else:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

        header_re = re.compile(
            r'^\s*(?:={3,}\s*Page\s+(\d+)\s*={3,}|#*\s*PAGE\s+(\d+)\b)',
            re.IGNORECASE
        )
        current = None
        for line in lines:
            line = line.rstrip('\n')
            m = header_re.match(line)
            if m:
                current = int(m.group(1) or m.group(2))
                raw_pages.setdefault(current, [])
                continue
            if current is not None and line.strip():
                raw_pages[current].append(line.strip())

    result = []
    for p in sorted(raw_pages.keys()):
        deduped = []
        prev = None
        for ln in raw_pages[p]:
            if ln != prev:
                deduped.append(ln)
            prev = ln
        text = '\n'.join(deduped).strip()
        if text:
            result.append({'page': p, 'text': text})
    return result


def _save_instructor_script(course, pages, language='en'):
    # ===== AMHARIC: language stored per page =====
    InstructorScript.query.filter_by(course_id=course.id).delete()
    for p in pages:
        db.session.add(InstructorScript(
            course_id=course.id,
            page_number=p['page'],
            narration_text=p['text'],
            language=language,
        ))
    db.session.commit()
    # ============================================

    for old in os.listdir(TTS_CACHE_DIR):
        if old.startswith(f'course{course.id}_page'):
            try:
                os.remove(os.path.join(TTS_CACHE_DIR, old))
            except OSError:
                pass


# ---------------- TTS + AUDIO ----------------

def get_openai_client():
    global _openai_client
    if not OPENAI_AVAILABLE:
        return None
    if _openai_client is None:
        key = os.getenv('OPENAI_API_KEY')
        if not key:
            return None
        _openai_client = OpenAI(api_key=key)
    return _openai_client


TTS_CACHE_DIR = os.path.join(app.config['UPLOAD_FOLDER'], 'tts_cache')
os.makedirs(TTS_CACHE_DIR, exist_ok=True)

RECORDINGS_DIR = os.path.join(app.config['UPLOAD_FOLDER'], 'recordings')
os.makedirs(RECORDINGS_DIR, exist_ok=True)


# ===== AMHARIC: OpenAI TTS (English) =====
def _openai_tts(text, out_path):
    client = get_openai_client()
    if not client:
        return None

    text = text[:4000] if len(text) > 4000 else text
    if not text.strip():
        return None

    try:
        with client.audio.speech.with_streaming_response.create(
            model='gpt-4o-mini-tts',
            voice='coral',
            input=text,
            instructions='Speak like a professional corporate trainer. '
                         'Clear, engaging, moderate pace, warm and confident tone.',
            response_format='mp3',
        ) as response:
            response.stream_to_file(out_path)
        print(f'[TTS/openai] Generated {out_path} ({os.path.getsize(out_path)} bytes)')
        return out_path
    except Exception as e:
        print(f'[TTS/openai] Generation failed: {e}')
        return None


# ===== AMHARIC: Azure TTS (native Amharic voices) =====
def _azure_tts(text, out_path, voice=None):
    """
    Generate MP3 via Azure Speech REST API.
    Native Amharic voices: am-ET-MekdesNeural (female), am-ET-AmehaNeural (male).
    """
    if not REQUESTS_AVAILABLE:
        print('[TTS/azure] requests library not installed. '
              'Run: pip install requests')
        return None

    key = app.config['AZURE_SPEECH_KEY']
    region = app.config['AZURE_SPEECH_REGION']
    if not key:
        print('[TTS/azure] AZURE_SPEECH_KEY not set. '
              'Set it in .env or environment variables.')
        return None

    voice = voice or app.config['AZURE_TTS_VOICE_AM']

    text = text[:4500] if len(text) > 4500 else text
    if not text.strip():
        return None

    from xml.sax.saxutils import escape as _xml_escape
    safe_text = _xml_escape(text)

    ssml = (
        f'<speak version="1.0" xml:lang="am-ET">'
        f'<voice name="{voice}">'
        f'<prosody rate="0%" pitch="0%">{safe_text}</prosody>'
        f'</voice></speak>'
    )

    url = f'https://{region}.tts.speech.microsoft.com/cognitiveservices/v1'
    headers = {
        'Ocp-Apim-Subscription-Key': key,
        'Content-Type': 'application/ssml+xml',
        'X-Microsoft-OutputFormat': 'audio-24khz-48kbitrate-mono-mp3',
        'User-Agent': 'MULSEN-Training',
    }

    try:
        r = _requests.post(url, headers=headers,
                           data=ssml.encode('utf-8'), timeout=60)
        r.raise_for_status()
        with open(out_path, 'wb') as f:
            f.write(r.content)
        print(f'[TTS/azure] Generated {out_path} '
              f'({os.path.getsize(out_path)} bytes, voice={voice})')
        return out_path
    except Exception as e:
        body = ''
        try:
            body = r.text[:300]
        except Exception:
            pass
        print(f'[TTS/azure] Generation failed: {e} | response: {body}')
        return None


# ===== AMHARIC: unified entry point with language routing =====
def get_or_create_tts_mp3(course_id, page_number, narration_text, lang='en'):
    """
    Cache key includes language so English and Amharic variants of the
    same page can coexist.
    """
    lang = (lang or 'en').lower()
    if lang not in ('en', 'am'):
        lang = 'en'

    cache_path = os.path.join(
        TTS_CACHE_DIR, f'course{course_id}_page{page_number}_{lang}.mp3'
    )
    if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
        return cache_path

    if not narration_text or not narration_text.strip():
        return None

    if lang == 'am':
        return _azure_tts(narration_text, cache_path)

    return _openai_tts(narration_text, cache_path)
# =============================================================


def _audio_duration_seconds(path):
    try:
        if path.lower().endswith('.mp3'):
            bitrate_bps = 128_000
            return int(os.path.getsize(path) * 8 / bitrate_bps)
    except Exception:
        pass
    return None


@app.route('/course/<int:course_id>/slide/<int:slide_num>/audio')
@login_required
def slide_audio(course_id, slide_num):
    course = Course.query.get_or_404(course_id)

    rec = PageRecording.query.filter_by(
        course_id=course.id, page_number=slide_num).first()
    if rec:
        path = os.path.join(RECORDINGS_DIR, rec.filename)
        if os.path.exists(path):
            resp = send_file(path, mimetype=rec.mimetype or 'audio/webm',
                             conditional=True)
            dur = _audio_duration_seconds(path)
            if dur:
                resp.headers['X-Audio-Duration'] = str(dur)
            return resp

    script = InstructorScript.query.filter_by(
        course_id=course.id, page_number=slide_num).first()
    if script:
        # ===== AMHARIC: pass language through =====
        script_lang = (script.language or 'en').lower()
        mp3_path = get_or_create_tts_mp3(
            course.id, slide_num, script.narration_text, lang=script_lang
        )
        # =========================================
        if mp3_path:
            resp = send_file(mp3_path, mimetype='audio/mpeg', conditional=True)
            dur = _audio_duration_seconds(mp3_path)
            if dur:
                resp.headers['X-Audio-Duration'] = str(dur)
            return resp

    return jsonify({
        'fallback': True,
        'text': script.narration_text if script else '',
        'lang': (script.language or 'en') if script else 'en',
    }), 503


# ---------------- PASSWORD + EMAIL HELPERS ----------------

def _gen_password(length=10):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


def _next_trainer_id():
    rows = Trainer.query.with_entities(Trainer.trainer_id).all()
    highest = 0
    for (tid,) in rows:
        if tid is None:
            continue
        tid_s = str(tid).strip()
        if tid_s.isdigit():
            n = int(tid_s)
            if n > highest:
                highest = n
    nxt = highest + 1
    if nxt > 99999:
        used = set()
        for (tid,) in rows:
            if tid and str(tid).isdigit():
                used.add(int(tid))
        for candidate in range(1, 100000):
            if candidate not in used:
                return f'{candidate:05d}'
        raise RuntimeError('Trainer ID space exhausted (100000 trainees)')
    return f'{nxt:05d}'


def _send_mail(to_email, subject, body_text, body_html=None):
    if not app.config['MAIL_HOST']:
        return False, 'MAIL_HOST not configured'
    if not to_email:
        return False, 'no recipient'

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = formataddr((app.config['MAIL_FROM_NAME'], app.config['MAIL_FROM']))
    msg['To'] = to_email
    msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
    if body_html:
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

    try:
        if app.config['MAIL_USE_SSL']:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL(app.config['MAIL_HOST'],
                                  app.config['MAIL_PORT'],
                                  context=ctx, timeout=15) as s:
                if app.config['MAIL_USER']:
                    s.login(app.config['MAIL_USER'], app.config['MAIL_PASS'])
                s.send_message(msg)
        else:
            with smtplib.SMTP(app.config['MAIL_HOST'],
                              app.config['MAIL_PORT'], timeout=15) as s:
                s.ehlo()
                if app.config['MAIL_USE_TLS']:
                    s.starttls(context=ssl.create_default_context())
                    s.ehlo()
                if app.config['MAIL_USER']:
                    s.login(app.config['MAIL_USER'], app.config['MAIL_PASS'])
                s.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)


def _welcome_email_body(full_name, email, password):
    text = (
        f"Hello {full_name},\n\n"
        f"Your MULSEN Virtual Training account has been created.\n\n"
        f"Login URL: {request.url_root.rstrip('/')}/login\n"
        f"Email:    {email}\n"
        f"Password: {password}\n\n"
        f"Please log in and change your password from the settings menu.\n\n"
        f"— MULSEN Training"
    )
    html = f"""
    <div style="font-family:system-ui,sans-serif;max-width:560px;margin:auto;">
      <h2 style="color:#0b6b3a;">Welcome to MULSEN Training</h2>
      <p>Hello <strong>{full_name}</strong>,</p>
      <p>Your account has been created.</p>
      <table style="border-collapse:collapse;margin:1rem 0;">
        <tr><td style="padding:.3rem .6rem;"><b>Login URL</b></td>
            <td style="padding:.3rem .6rem;">{request.url_root.rstrip('/')}/login</td></tr>
        <tr><td style="padding:.3rem .6rem;"><b>Email</b></td>
            <td style="padding:.3rem .6rem;">{email}</td></tr>
        <tr><td style="padding:.3rem .6rem;"><b>Password</b></td>
            <td style="padding:.3rem .6rem;"><code>{password}</code></td></tr>
      </table>
      <p>Please log in and change your password from the settings menu.</p>
      <p style="color:#6b7280;font-size:.85rem;">— MULSEN Training</p>
    </div>
    """
    return text, html


# ---------------- INLINE BASE STYLE (shared by inline pages) ----------------

_INLINE_BASE_STYLE = """
<style>
  *,*::before,*::after{box-sizing:border-box;}
  body{font-family:system-ui,Segoe UI,Roboto,sans-serif;
       background:#f5f7fa;margin:0;padding:2rem;color:#1f2933;
       -webkit-tap-highlight-color:transparent;
       overflow-x:hidden;}
  .wrap{max-width:540px;margin:2rem auto;background:#fff;
        padding:2rem;border-radius:12px;
        box-shadow:0 2px 8px rgba(0,0,0,.06);}
  h1{margin-top:0;}
  label{display:block;font-size:.85rem;font-weight:500;
        margin:1rem 0 .3rem;}
  input{width:100%;padding:.6rem;border:1px solid #d1d5db;
        border-radius:6px;font-size:.95rem;box-sizing:border-box;
        min-height:40px;}
  input:focus{outline:2px solid #0b6b3a;border-color:transparent;}
  button{width:100%;padding:.75rem;margin-top:1.25rem;
         background:#0b6b3a;color:#fff;border:0;border-radius:6px;
         font-size:1rem;font-weight:600;cursor:pointer;
         min-height:44px;font-family:inherit;}
  button:hover{background:#095a31;}
  .muted{color:#6b7280;font-size:.9rem;}
  .req{color:#b91c1c;}
  .footer{text-align:center;margin-top:1rem;font-size:.85rem;color:#6b7280;}
  .footer a{color:#0b6b3a;text-decoration:none;font-weight:500;}
  .ok{background:#dcfce7;color:#166534;padding:1rem;
      border-radius:.5rem;margin-bottom:1rem;line-height:1.6;
      word-wrap:break-word;}
  .err{background:#fee2e2;color:#991b1b;padding:.5rem .75rem;
       border-radius:.4rem;margin-bottom:.5rem;
       word-wrap:break-word;}
  .ok a,.err a{font-weight:600;}

  @media (max-width:640px){
    body{padding:.75rem;}
    .wrap{margin:.75rem auto;padding:1.25rem;border-radius:8px;}
    h1{font-size:1.25rem;}
    input{font-size:16px;padding:.55rem .65rem;}
    button{font-size:1rem;padding:.8rem;}
    .ok{font-size:.88rem;padding:.75rem .85rem;}
    .err{font-size:.85rem;}
  }
  @media (max-width:380px){
    body{padding:.4rem;}
    .wrap{padding:1rem;}
    h1{font-size:1.15rem;}
  }
</style>
"""


# ---------------- SIGN-UP (inline) ----------------

@app.route('/signup', methods=['GET', 'POST'])
def signup_inline():
    error_html = ''
    success_html = ''
    prefilled = {'full_name': '', 'email': '', 'department': '', 'position': ''}

    if request.method == 'POST':
        prefilled['full_name'] = (request.form.get('full_name') or '').strip()
        prefilled['email'] = (request.form.get('email') or '').strip().lower()
        prefilled['department'] = (request.form.get('department') or '').strip()
        prefilled['position'] = (request.form.get('position') or '').strip()

        errors = []
        if not prefilled['email'] or '@' not in prefilled['email']:
            errors.append('A valid email address is required.')
        if not prefilled['full_name']:
            errors.append('Full name is required.')
        if not prefilled['department']:
            errors.append('Department is required.')
        if prefilled['email'] and Trainer.query.filter(
                db.func.lower(Trainer.email) == prefilled['email']).first():
            errors.append('This email is already registered.')

        if not errors:
            password = _gen_password(10)
            try:
                new_id = _next_trainer_id()
                trainee = Trainer(
                    trainer_id=new_id,
                    email=prefilled['email'],
                    full_name=prefilled['full_name'],
                    department=prefilled['department'],
                    position=prefilled['position'],
                    name=(prefilled['full_name'].split()[0][:100]
                          if prefilled['full_name'].split() else 'Trainee'),
                    password_hash=generate_password_hash(password),
                )
                db.session.add(trainee)
                db.session.commit()

                mail_note = ''
                try:
                    body_text, body_html = _welcome_email_body(
                        prefilled['full_name'], prefilled['email'], password)
                    sent, err = _send_mail(
                        prefilled['email'],
                        'Your MULSEN Training account',
                        body_text, body_html)
                    if sent:
                        mail_note = '<br><small>Credentials also emailed.</small>'
                    else:
                        mail_note = ('<br><small style="color:#92400e;">'
                                     f'Email delivery failed: {err}</small>')
                except Exception as e:
                    mail_note = (f'<br><small style="color:#92400e;">'
                                 f'Email error: {e}</small>')

                success_html = f"""
                <div class="ok">
                  <strong>✅ Account created</strong><br>
                  Trainee ID: <strong>{new_id}</strong><br>
                  Email: <strong>{prefilled['email']}</strong><br>
                  Password: <strong>{password}</strong>
                  {mail_note}<br><br>
                  <a href="/login" style="color:#166534;">→ Go to login</a>
                </div>
                """
                prefilled = {k: '' for k in prefilled}
            except Exception as e:
                import traceback
                traceback.print_exc()
                error_html = (
                    f'<div class="err"><strong>Could not create account:</strong> {e}</div>'
                )

        if errors:
            error_html = ''.join(f'<div class="err">{e}</div>' for e in errors)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sign up — MULSEN Training</title>
{_INLINE_BASE_STYLE}
</head>
<body>
<div class="wrap">
  <h1>📝 Create your account</h1>
  <p class="muted">Fill in the form to register. A password will be
     generated and shown to you on the next screen.</p>
  {success_html}{error_html}
  <form method="post">
    <label>Full name <span class="req">*</span></label>
    <input type="text" name="full_name" required value="{prefilled['full_name']}">

    <label>Email address <span class="req">*</span></label>
    <input type="email" name="email" required value="{prefilled['email']}">

    <label>Department <span class="req">*</span></label>
    <input type="text" name="department" required value="{prefilled['department']}">

    <label>Position (optional)</label>
    <input type="text" name="position" value="{prefilled['position']}">

    <button type="submit">Create account</button>
  </form>
  <p class="footer">
    Already registered? <a href="/login">Sign in</a><br>
    Admin? <a href="/admin/login">Admin login</a>
  </p>
</div>
</body>
</html>"""
    return html


# ---------------- ADMIN ADD TRAINEE (inline) ----------------

@app.route('/admin/trainees/new', methods=['GET', 'POST'])
@admin_required
def admin_add_trainee_inline():
    error_html = ''
    success_html = ''
    prefilled = {'full_name': '', 'email': '', 'department': '', 'position': ''}

    if request.method == 'POST':
        prefilled['full_name'] = (request.form.get('full_name') or '').strip()
        prefilled['email'] = (request.form.get('email') or '').strip().lower()
        prefilled['department'] = (request.form.get('department') or '').strip()
        prefilled['position'] = (request.form.get('position') or '').strip()

        errors = []
        if not prefilled['full_name']:
            errors.append('Full name is required.')
        if prefilled['email'] and Trainer.query.filter(
                db.func.lower(Trainer.email) == prefilled['email']).first():
            errors.append(f'Email {prefilled["email"]} already registered.')

        if not errors:
            password = _gen_password(10)
            try:
                new_id = _next_trainer_id()
                trainee = Trainer(
                    trainer_id=new_id,
                    email=prefilled['email'] or None,
                    full_name=prefilled['full_name'],
                    department=prefilled['department'],
                    position=prefilled['position'],
                    name=(prefilled['full_name'].split()[0][:100]
                          if prefilled['full_name'].split() else 'Trainee'),
                    password_hash=generate_password_hash(password),
                )
                db.session.add(trainee)
                db.session.commit()

                mail_status = ''
                if prefilled['email']:
                    try:
                        body_text, body_html = _welcome_email_body(
                            prefilled['full_name'], prefilled['email'], password)
                        sent, err = _send_mail(
                            prefilled['email'],
                            'Your MULSEN Training account',
                            body_text, body_html)
                        mail_status = ('<br><small>Credentials emailed.</small>'
                                       if sent else
                                       ('<br><small style="color:#92400e;">'
                                        f'Email failed: {err}</small>'))
                    except Exception as e:
                        mail_status = (f'<br><small style="color:#92400e;">'
                                       f'Email error: {e}</small>')

                success_html = f"""
                <div class="ok">
                  <strong>✅ Trainee created</strong><br>
                  Full name: <strong>{prefilled['full_name']}</strong><br>
                  ID: <strong>{new_id}</strong><br>
                  Password: <strong>{password}</strong>
                  {mail_status}<br><br>
                  <a href="/admin/trainees/new" style="color:#166534;">→ Add another</a>
                  &nbsp;|&nbsp;
                  <a href="/admin/trainees" style="color:#166534;">→ View all</a>
                </div>
                """
                prefilled = {k: '' for k in prefilled}
            except Exception as e:
                import traceback
                traceback.print_exc()
                error_html = (
                    f'<div class="err"><strong>Could not create trainee:</strong> {e}</div>'
                )

        if errors:
            error_html = ''.join(f'<div class="err">{e}</div>' for e in errors)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Add trainee — MULSEN Admin</title>
{_INLINE_BASE_STYLE}
</head>
<body>
<div class="wrap">
  <h1>➕ Add trainee</h1>
  <p class="muted">A password will be generated and shown below.
     If an email is provided, credentials will be emailed automatically.</p>
  {success_html}{error_html}
  <form method="post">
    <label>Full name <span class="req">*</span></label>
    <input type="text" name="full_name" required value="{prefilled['full_name']}">

    <label>Email (optional)</label>
    <input type="email" name="email" value="{prefilled['email']}">

    <label>Department (optional)</label>
    <input type="text" name="department" value="{prefilled['department']}">

    <label>Position (optional)</label>
    <input type="text" name="position" value="{prefilled['position']}">

    <button type="submit">💾 Save — Add trainee</button>
  </form>
  <p class="footer">
    <a href="/admin/trainees">← Back to trainees list</a>
    &nbsp;|&nbsp;
    <a href="/admin/courses">Admin home</a>
  </p>
</div>
</body>
</html>"""
    return html


# ---------------- SCRIPT TEMPLATE + PASTE / UPLOAD / EDIT / DELETE ----------------

@app.route('/admin/script_template')
@admin_required
def download_script_template():
    doc = DocxDocument()
    doc.add_heading('MULSEN Instructor Script Template', level=1)

    note = doc.add_paragraph()
    run = note.add_run(
        'Instructions:\n'
        '1. Copy the block below and paste it once for each slide.\n'
        '2. Change the number in the "===== Page N =====" header to match '
        'the slide number (1, 2, 3, …).\n'
        '3. Type or paste the narration text underneath the header.\n'
        '4. Save this file, then upload it in Admin → Scripts & Recordings.\n'
        'Do NOT rename or remove the header lines.'
    )
    run.italic = True
    run.font.size = DocxPt(10)

    doc.add_paragraph()

    header = doc.add_paragraph()
    rh = header.add_run('===== Page 1 =====')
    rh.bold = True
    rh.font.size = DocxPt(13)

    doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name='MULSEN_Script_Template.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@app.route('/admin/course/<int:course_id>/script/paste', methods=['POST'])
@admin_required
def paste_instructor_script(course_id):
    course = Course.query.get_or_404(course_id)
    text = request.form.get('script_text', '').strip()
    # ===== AMHARIC: read language from form =====
    language = (request.form.get('language') or 'en').lower()
    if language not in ('en', 'am'):
        language = 'en'
    # ============================================
    if not text:
        flash('Pasted script is empty.', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    tmp_path = os.path.join(
        app.config['UPLOAD_FOLDER'],
        f'tmp_paste_{course_id}_{datetime.now().strftime("%Y%m%d%H%M%S")}.txt'
    )
    with open(tmp_path, 'w', encoding='utf-8') as f:
        f.write(text)

    try:
        pages = parse_instructor_script(tmp_path, 'paste.txt')
    except Exception as e:
        flash(f'Failed to parse pasted script: {e}', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    if not pages:
        flash('No "===== Page N =====" headers found in the pasted text.', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    _save_instructor_script(course, pages, language=language)

    slide_count = effective_slide_count(course)
    lang_label = 'Amharic (አማርኛ)' if language == 'am' else 'English'
    if slide_count is not None and slide_count != len(pages):
        flash(
            f'⚠ {lang_label} script saved with {len(pages)} pages, but the '
            f'slideshow has {slide_count} slides. Pages past {len(pages)} '
            f'will have no narration.',
            'warning'
        )
    else:
        flash(f'✅ {lang_label} script saved: {len(pages)} pages parsed.', 'success')
    return redirect(url_for('admin_scripts', course_id=course.id))


@app.route('/admin/course/<int:course_id>/script', methods=['POST'])
@admin_required
def upload_instructor_script(course_id):
    course = Course.query.get_or_404(course_id)
    f = request.files.get('script_file')
    # ===== AMHARIC: read language from form =====
    language = (request.form.get('language') or 'en').lower()
    if language not in ('en', 'am'):
        language = 'en'
    # ============================================
    if not f or not f.filename:
        flash('No script file uploaded.', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    ext = f.filename.lower().rsplit('.', 1)[-1]
    if ext not in ('docx', 'txt', 'csv'):
        flash('Script must be .docx, .txt, or .csv', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    tmp_path = os.path.join(
        app.config['UPLOAD_FOLDER'],
        f'tmp_script_{course_id}_{datetime.now().strftime("%Y%m%d%H%M%S")}.{ext}'
    )
    f.save(tmp_path)

    try:
        pages = parse_instructor_script(tmp_path, f.filename)
    except Exception as e:
        flash(f'Failed to parse script: {e}', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    if not pages:
        flash('No pages found. Make sure the file uses "===== Page N =====" '
              'or "PAGE N" headers.', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    _save_instructor_script(course, pages, language=language)

    slide_count = effective_slide_count(course)
    lang_label = 'Amharic (አማርኛ)' if language == 'am' else 'English'
    if slide_count is not None and slide_count != len(pages):
        flash(
            f'⚠ {lang_label} script saved: {len(pages)} pages, but the '
            f'slideshow has {slide_count} slides.', 'warning'
        )
    else:
        flash(f'✅ {lang_label} instructor script uploaded: {len(pages)} pages '
              f'parsed. TTS audio will be generated on first playback.', 'success')
    return redirect(url_for('admin_scripts', course_id=course.id))


@app.route('/admin/course/<int:course_id>/script/delete', methods=['POST'])
@admin_required
def delete_instructor_script(course_id):
    course = Course.query.get_or_404(course_id)
    try:
        InstructorScript.query.filter_by(course_id=course.id).delete(
            synchronize_session=False)
        db.session.commit()
        # Cache cleanup handles both _en and _am variants
        for old in os.listdir(TTS_CACHE_DIR):
            if old.startswith(f'course{course.id}_page'):
                try:
                    os.remove(os.path.join(TTS_CACHE_DIR, old))
                except OSError:
                    pass
        flash('🗑 Instructor script deleted. Recordings are untouched.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Failed to delete script: {e}', 'error')
    return redirect(url_for('admin_scripts', course_id=course.id))


@app.route('/admin/course/<int:course_id>/script/page/<int:page_num>/edit',
           methods=['POST'])
@admin_required
def edit_script_page(course_id, page_num):
    course = Course.query.get_or_404(course_id)
    text = (request.form.get('narration_text') or '').strip()
    # ===== AMHARIC: optional language change on edit =====
    new_lang = (request.form.get('language') or '').lower()
    if new_lang not in ('en', 'am'):
        new_lang = None
    # ======================================================
    row = InstructorScript.query.filter_by(
        course_id=course.id, page_number=page_num).first()
    if not row:
        flash(f'No script text for page {page_num}.', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    old_lang = (row.language or 'en').lower()

    if not text:
        db.session.delete(row)
    else:
        row.narration_text = text
        if new_lang:
            row.language = new_lang
    db.session.commit()

    # Clear cache for the old language variant (and the new one if switched)
    langs_to_clear = {old_lang}
    if new_lang:
        langs_to_clear.add(new_lang)
    for lng in langs_to_clear:
        cached = os.path.join(
            TTS_CACHE_DIR, f'course{course.id}_page{page_num}_{lng}.mp3')
        if os.path.exists(cached):
            try:
                os.remove(cached)
            except OSError:
                pass

    flash(f'✅ Page {page_num} updated.', 'success')
    return redirect(url_for('admin_scripts', course_id=course.id))


# ---------------- PAGE RECORDINGS ----------------

@app.route('/admin/course/<int:course_id>/page/<int:page_num>/recording',
           methods=['POST'])
@any_logged_in_required
def upload_page_recording(course_id, page_num):
    course = Course.query.get_or_404(course_id)
    if page_num <= 0:
        return jsonify({'error': 'invalid page'}), 400

    f = request.files.get('audio')
    if not f:
        return jsonify({'error': 'no audio file'}), 400

    mt = (f.mimetype or '').lower()
    if 'webm' in mt:
        ext = 'webm'
    elif 'mp4' in mt or 'm4a' in mt:
        ext = 'mp4'
    elif 'ogg' in mt:
        ext = 'ogg'
    else:
        ext = 'webm'

    stamp = int(datetime.utcnow().timestamp())
    fname = f'course{course.id}_page{page_num}_{stamp}.{ext}'
    fpath = os.path.join(RECORDINGS_DIR, fname)
    f.save(fpath)

    recorded_by = session.get('admin_username') or session.get('trainer_name') or 'unknown'

    rec = PageRecording.query.filter_by(
        course_id=course.id, page_number=page_num).first()
    if rec:
        old_path = os.path.join(RECORDINGS_DIR, rec.filename)
        if os.path.exists(old_path) and rec.filename != fname:
            try:
                os.remove(old_path)
            except OSError:
                pass
        rec.filename = fname
        rec.mimetype = f.mimetype or 'audio/webm'
        rec.recorded_by = recorded_by
        rec.created_at = datetime.utcnow()
    else:
        db.session.add(PageRecording(
            course_id=course.id,
            page_number=page_num,
            filename=fname,
            mimetype=f.mimetype or 'audio/webm',
            recorded_by=recorded_by,
        ))
    db.session.commit()

    return jsonify({'ok': True, 'file': fname})


@app.route('/admin/course/<int:course_id>/page/<int:page_num>/recording',
           methods=['DELETE'])
@any_logged_in_required
def delete_page_recording(course_id, page_num):
    course = Course.query.get_or_404(course_id)
    rec = PageRecording.query.filter_by(
        course_id=course.id, page_number=page_num).first()
    if not rec:
        return jsonify({'ok': True, 'deleted': False})
    path = os.path.join(RECORDINGS_DIR, rec.filename)
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
    db.session.delete(rec)
    db.session.commit()
    return jsonify({'ok': True, 'deleted': True})


@app.route('/admin/course/<int:course_id>/manual_pages', methods=['POST'])
@admin_required
def set_manual_slide_count(course_id):
    course = Course.query.get_or_404(course_id)
    try:
        n = int(request.form.get('page_count', '0'))
    except ValueError:
        n = 0
    if n <= 0 or n > 500:
        flash('Page count must be between 1 and 500.', 'error')
        return redirect(url_for('admin_scripts', course_id=course.id))

    data = _load_manual_counts()
    data[str(course.id)] = n
    _save_manual_counts(data)
    flash(f'✅ Manual page count set to {n} for "{course.name}".', 'success')
    return redirect(url_for('admin_scripts', course_id=course.id))


# ---------------- PROGRESS / PAUSE / RESUME ----------------

@app.route('/course/<int:course_id>/pause', methods=['POST'])
@login_required
def pause_course(course_id):
    data = request.get_json() or {}
    slide_num = int(data.get('slide', 1))
    trainer_id = session['trainer_id']

    row = Progress.query.filter_by(
        trainer_id=trainer_id, course_id=course_id, slide_number=-1
    ).first()
    if not row:
        row = Progress(trainer_id=trainer_id, course_id=course_id,
                       slide_number=-1)
        db.session.add(row)
    row.paused_at = datetime.utcnow()
    row.last_slide = slide_num
    row.last_seen = datetime.utcnow()
    db.session.commit()
    return jsonify({'ok': True, 'slide': slide_num})


@app.route('/course/<int:course_id>/resume-info')
@login_required
def resume_info(course_id):
    row = Progress.query.filter_by(
        trainer_id=session['trainer_id'],
        course_id=course_id,
        slide_number=-1,
    ).first()
    if not row or not row.paused_at:
        return jsonify({'resume': None})
    return jsonify({'resume': {
        'slide': row.last_slide or 1,
        'when': row.paused_at.strftime('%b %d, %Y %H:%M'),
    }})


@app.route('/course/<int:course_id>/resume-clear', methods=['POST'])
@login_required
def resume_clear(course_id):
    row = Progress.query.filter_by(
        trainer_id=session['trainer_id'],
        course_id=course_id,
        slide_number=-1,
    ).first()
    if row:
        db.session.delete(row)
        db.session.commit()
    return jsonify({'ok': True})


@app.route('/course/<int:course_id>/mark-seen', methods=['POST'])
@login_required
def mark_seen(course_id):
    data = request.get_json() or {}
    slide_num = int(data.get('slide', 1))
    if slide_num <= 0:
        return jsonify({'ok': False}), 400

    row = Progress.query.filter_by(
        trainer_id=session['trainer_id'], course_id=course_id,
        slide_number=slide_num
    ).first()
    if not row:
        row = Progress(trainer_id=session['trainer_id'],
                       course_id=course_id, slide_number=slide_num)
        db.session.add(row)
    row.last_seen = datetime.utcnow()
    db.session.commit()
    return jsonify({'ok': True})


# ---------------- CERTIFICATE ----------------

def generate_certificate(trainer_name, course_name, percentage, distinction,
                         attempt_id=None):
    safe_name = secure_filename(
        f"{trainer_name}_{course_name}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
    )
    path = os.path.join('certificates', safe_name)
    c = canvas.Canvas(path, pagesize=landscape(A4))
    width, height = landscape(A4)
    cx, cy = width / 2.0, height / 2.0

    c.setFillColorRGB(0.09, 0.40, 0.20)
    c.rect(0, 0, width, height, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.rect(1.2*cm, 1.2*cm, width - 2.4*cm, height - 2.4*cm, fill=1, stroke=0)
    c.setStrokeColorRGB(0.09, 0.40, 0.20)
    c.setLineWidth(4)
    c.rect(1.5*cm, 1.5*cm, width - 3*cm, height - 3*cm, fill=0, stroke=1)

    green = (0.09, 0.40, 0.20)
    dark  = (0.20, 0.20, 0.20)
    PT_TO_CM = 0.0352778

    lines = [
        ("CERTIFICATE OF COMPLETION", "Helvetica-Bold", 32, green, 0.0),
        ("This certificate is proudly presented to", "Helvetica", 14, dark, 0.9),
        (trainer_name, "Helvetica-Bold", 28, green, 0.7),
        ("for successfully completing the training course",
            "Helvetica", 14, dark, 0.9),
        (course_name, "Helvetica-Bold", 20, green, 0.6),
        (f"with a final score of {percentage:.1f}%", "Helvetica", 14, dark, 0.9),
    ]
    if distinction:
        lines.append(("*** WITH DISTINCTION ***", "Helvetica-Bold", 22,
                      (0.85, 0.60, 0.05), 0.7))

    footer_lines = [
        (f"Issued on {datetime.now().strftime('%B %d, %Y')}",
            "Helvetica-Oblique", 11, (0.35, 0.35, 0.35)),
        ("MULSEN Virtual Training App",
            "Helvetica-Bold", 11, green),
    ]

    def line_pts(items):
        total = 0.0
        for _, _, size, _, gap in items:
            total += gap * 28.3465
            total += size * 1.25
        return total

    main_pt = line_pts(lines)
    footer_pt = line_pts(footer_lines) + 20.0
    offset_pt = footer_pt * 0.35
    top_of_main = cy + (main_pt / 2.0 - offset_pt) * PT_TO_CM * cm

    y = top_of_main
    for text, font, size, color, gap in lines:
        y -= gap * cm
        c.setFont(font, size)
        c.setFillColorRGB(*color)
        c.drawCentredString(cx, y - size * PT_TO_CM * cm * 0.8, text)
        y -= (size * 1.25) * PT_TO_CM * cm

    footer_y = 2.8 * cm + (len(footer_lines) - 1) * 0.55 * cm
    for text, font, size, color in footer_lines:
        c.setFont(font, size)
        c.setFillColorRGB(*color)
        c.drawCentredString(cx, footer_y, text)
        footer_y -= 0.55 * cm

    if QR_AVAILABLE:
        verify_payload = (
            f"MULSEN-CERT|"
            f"id={attempt_id or 'na'}|"
            f"name={trainer_name}|"
            f"course={course_name}|"
            f"score={percentage:.1f}|"
            f"date={datetime.now().strftime('%Y-%m-%d')}"
        )
        try:
            qr = qrcode.QRCode(
                version=None,
                error_correction=ERROR_CORRECT_M,
                box_size=10,
                border=2,
            )
            qr.add_data(verify_payload)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")

            qr_tmp = os.path.join('certificates', f'_qr_{safe_name}.png')
            qr_img.save(qr_tmp)

            qr_size = 2.6 * cm
            qr_x = width - 1.5*cm - qr_size - 0.5*cm
            qr_y = 1.5*cm + 0.5*cm
            c.drawImage(qr_tmp, qr_x, qr_y, width=qr_size, height=qr_size)

            c.setFont("Helvetica", 7)
            c.setFillColorRGB(0.35, 0.35, 0.35)
            c.drawCentredString(qr_x + qr_size/2, qr_y - 0.35*cm,
                                "Scan to verify")

            try:
                os.remove(qr_tmp)
            except OSError:
                pass
        except Exception as e:
            print(f"[cert] QR generation failed: {e}")

    c.save()
    return path


# ---------------- INLINE COURSE VIEW (fully responsive) ----------------

@app.route('/course/<int:course_id>')
@login_required
def course_view(course_id):
    from markupsafe import escape as _escape

    course = Course.query.get_or_404(course_id)
    slides = []
    if course.ppt_filename:
        ppt_path = os.path.join(app.config['UPLOAD_FOLDER'], course.ppt_filename)
        if os.path.exists(ppt_path):
            try:
                if course.source_type == 'pdf':
                    slides = pdf_to_slides(ppt_path)
                else:
                    slides = ppt_to_slides(ppt_path)
            except Exception as e:
                flash(f'Error reading slides: {e}', 'error')

    course_meta = {
        'id': course.id,
        'total': len(slides),
        'hasExam': len(course.questions) > 0,
        'examUrl': url_for('exam', course_id=course.id),
        'dashboardUrl': url_for('trainer_dashboard'),
    }

    slides_json = json.dumps(slides)
    meta_json = json.dumps(course_meta)
    course_title = str(_escape(course.name))

    if not slides:
        return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{course_title}</title>
<style>
body{{font-family:system-ui;padding:2rem;background:#f5f7fa;margin:0;}}
.box{{max-width:800px;margin:2rem auto;background:#fff;padding:2rem;
border-radius:12px;}}
@media (max-width:640px){{
  body{{padding:.75rem;}}
  .box{{margin:.75rem auto;padding:1.25rem;}}
}}
</style></head><body>
<div class="box"><h1>{course_title}</h1>
<p style="color:#92400e;background:#fef3c7;padding:.75rem;
border-radius:.4rem;">No slides found for this course.</p>
<a href="{course_meta['dashboardUrl']}">← Back to dashboard</a>
</div></body></html>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{course_title} — MULSEN Training</title>
<style>
*,*::before,*::after{{box-sizing:border-box;}}
body{{margin:0;padding:1rem;font-family:system-ui,Segoe UI,Roboto,sans-serif;
background:#f5f7fa;color:#1f2933;-webkit-tap-highlight-color:transparent;
overflow-x:hidden;
padding-left:max(1rem, env(safe-area-inset-left));
padding-right:max(1rem, env(safe-area-inset-right));}}
.wrap{{max-width:1100px;margin:0 auto;}}
.topbar{{display:flex;align-items:center;gap:1rem;
justify-content:space-between;margin-bottom:1rem;flex-wrap:wrap;}}
.topbar h1{{margin:0;font-size:1.4rem;word-break:break-word;}}
.topbar a{{color:#2563eb;text-decoration:none;white-space:nowrap;}}
.counter{{color:#6b7280;font-size:.9rem;white-space:nowrap;}}
#resumeBanner{{background:#dbeafe;border:1px solid #93c5fd;
border-radius:.5rem;padding:.85rem 1rem;margin-bottom:1rem;display:none;
justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap;}}
#resumeBanner strong{{color:#1e40af;}}
#resumeBanner .actions{{display:flex;gap:.5rem;flex-wrap:wrap;
min-width:0;flex:1 1 auto;justify-content:flex-end;}}
.stage{{background:#fff;border:1px solid #e5e7eb;border-radius:.6rem;
padding:1rem;}}
#slideCanvas{{min-height:420px;display:flex;align-items:center;
justify-content:center;flex-direction:column;background:#0f172a;color:#e2e8f0;
border-radius:.5rem;padding:1.5rem;text-align:left;}}
#slideCanvas img{{max-width:100%;max-height:600px;display:block;margin:0 auto;}}
#slideCanvas .txt{{white-space:pre-wrap;font-size:1.05rem;
line-height:1.5;max-width:800px;}}
#slideCanvas .txt:lang(am){{line-height:1.7;}}
.progress-wrap{{margin:1rem 0 .25rem;}}
.progress-bar{{width:100%;height:10px;background:#e5e7eb;
border-radius:999px;overflow:hidden;}}
#progressFill{{height:100%;width:0%;background:#2563eb;
transition:width .15s linear;}}
.progress-meta{{display:flex;justify-content:space-between;
font-size:.8rem;color:#6b7280;margin-top:.3rem;}}
.controls{{display:flex;align-items:center;gap:.6rem;margin-top:.6rem;
flex-wrap:wrap;}}
.controls button{{padding:.55rem 1rem;border:0;border-radius:6px;
font-size:.9rem;cursor:pointer;font-weight:500;font-family:inherit;
min-height:40px;}}
#prevBtn, #nextBtn{{background:transparent;color:#0b6b3a;
border:1px solid #0b6b3a;}}
#prevBtn:disabled, #nextBtn:disabled{{opacity:.4;cursor:not-allowed;}}
#playBtn{{background:#0b6b3a;color:#fff;}}
#playBtn:hover{{background:#095a31;}}
#pauseBtn{{background:transparent;color:#0b6b3a;border:1px solid #0b6b3a;}}
.hint{{font-size:.8rem;color:#9ca3af;}}
#resumeBtn{{background:#0b6b3a;color:#fff;padding:.55rem 1rem;border:0;
border-radius:6px;cursor:pointer;font-weight:600;font-size:.9rem;
white-space:nowrap;min-height:40px;flex-shrink:0;font-family:inherit;}}
#startOverBtn{{background:transparent;color:#0b6b3a;border:1px solid #0b6b3a;
padding:.55rem 1rem;border-radius:6px;cursor:pointer;font-weight:600;
font-size:.9rem;white-space:nowrap;min-height:40px;flex-shrink:0;
font-family:inherit;}}
.splash{{position:fixed;inset:0;display:flex;align-items:center;
justify-content:center;flex-direction:column;gap:1rem;
background:rgba(15,23,42,.92);color:#fff;z-index:9999;font-size:1.2rem;}}
.splash .spin{{width:44px;height:44px;border:4px solid #334155;
border-top-color:#60a5fa;border-radius:50%;animation:spin 1s linear infinite;}}
@keyframes spin{{to{{transform:rotate(360deg);}}}}
.toast{{position:fixed;top:70px;right:1rem;z-index:9998;padding:.7rem 1rem;
border-radius:.5rem;box-shadow:0 2px 8px rgba(0,0,0,.15);font-size:.9rem;
max-width:calc(100vw - 2rem);word-wrap:break-word;}}
.toast.ok{{background:#dcfce7;color:#166534;}}
.toast.err{{background:#fee2e2;color:#991b1b;}}

@media (max-width:768px){{
  body{{padding:.55rem;}}
  .topbar h1{{font-size:1.1rem;}}
  .stage{{padding:.75rem;}}
  #slideCanvas{{min-height:280px;padding:1rem;}}
  #slideCanvas .txt{{font-size:.98rem;}}
  .controls{{gap:.45rem;}}
  .controls button{{padding:.5rem .8rem;font-size:.85rem;
                    min-height:40px;flex:1 1 auto;}}
  .hint{{width:100%;order:10;text-align:center;font-size:.75rem;}}
}}

@media (max-width:640px){{
  body{{padding:.35rem;}}
  .topbar{{flex-direction:column;align-items:flex-start;gap:.35rem;}}
  #resumeBanner{{flex-direction:column;align-items:stretch;
                 gap:.55rem;}}
  #resumeBanner .actions{{flex-direction:column;justify-content:stretch;
                          width:100%;}}
  #resumeBtn, #startOverBtn{{width:100%;min-height:44px;
                             font-size:.95rem;white-space:normal;}}
  .stage{{padding:.6rem;border-radius:.5rem;}}
  #slideCanvas{{min-height:220px;padding:.85rem;}}
  #slideCanvas .txt{{font-size:.92rem;}}
  #prevBtn, #nextBtn, #playBtn, #pauseBtn{{
    padding:.55rem .5rem;font-size:.85rem;min-height:44px;
  }}
  .progress-meta{{font-size:.72rem;}}
  .progress-bar{{height:8px;}}
}}

@media (max-width:380px){{
  body{{padding:.25rem;}}
  .controls button{{font-size:.8rem;padding:.5rem .4rem;}}
  #slideCanvas{{min-height:180px;padding:.65rem;}}
  #slideCanvas .txt{{font-size:.88rem;}}
}}

@media (max-height:500px) and (orientation:landscape){{
  #slideCanvas{{min-height:160px;padding:.6rem;}}
  .stage{{padding:.5rem;}}
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <a href="{course_meta['dashboardUrl']}">← Back</a>
    <h1>{course_title}</h1>
    <span class="counter" id="slideCounter"></span>
  </div>

  <div id="resumeBanner">
    <div>
      <strong>⏸ You have a saved spot</strong>
      <span id="resumeWhen" style="color:#6b7280;"></span>
    </div>
    <div class="actions">
      <button type="button" id="resumeBtn">
        ▶ Resume from slide <span id="resumeSlide">?</span>
      </button>
      <button type="button" id="startOverBtn">Start from beginning</button>
    </div>
  </div>

  <div class="stage">
    <div id="slideCanvas"></div>
    <div class="progress-wrap">
      <div class="progress-bar"><div id="progressFill"></div></div>
      <div class="progress-meta">
        <span id="progressLabel">Narration: —</span>
        <span id="progressTime"></span>
      </div>
    </div>
    <div class="controls">
      <button type="button" id="prevBtn">← Prev</button>
      <button type="button" id="playBtn">🔊 Play narration</button>
      <button type="button" id="pauseBtn">⏸ Pause &amp; save</button>
      <button type="button" id="nextBtn" disabled>Next →</button>
      <span class="hint" id="nextHint">Listen to narration to unlock Next</span>
    </div>
  </div>
</div>

<script id="slidesData" type="application/json">{slides_json}</script>
<script id="courseMeta" type="application/json">{meta_json}</script>

<script>
(function(){{
  try {{
    var slides = JSON.parse(document.getElementById('slidesData').textContent);
    var meta = JSON.parse(document.getElementById('courseMeta').textContent);
    if (!slides || !slides.length) return;

    var idx = 0;
    var audio = new Audio();
    var unlocked = {{}};
    var progressTimer = null;
    var currentDuration = 0;
    var startedAt = 0;
    var pausedAccum = 0;
    var pauseStart = 0;
    var isPaused = false;
    var mode = 'idle';
    var courseId = meta.id;

    var canvas = document.getElementById('slideCanvas');
    var counter = document.getElementById('slideCounter');
    var prevBtn = document.getElementById('prevBtn');
    var nextBtn = document.getElementById('nextBtn');
    var playBtn = document.getElementById('playBtn');
    var pauseBtn = document.getElementById('pauseBtn');
    var fill = document.getElementById('progressFill');
    var label = document.getElementById('progressLabel');
    var timeLabel = document.getElementById('progressTime');
    var nextHint = document.getElementById('nextHint');
    var resumeBanner = document.getElementById('resumeBanner');

    function pad(n){{ return n < 10 ? '0'+n : ''+n; }}
    function fmt(s){{ if (!isFinite(s)||s<0) s=0;
      var m=Math.floor(s/60), r=Math.floor(s%60); return m+':'+pad(r); }}
    function slideNum(){{ return slides[idx].number; }}
    function setProgress(p){{ fill.style.width=Math.max(0,Math.min(100,p))+'%'; }}
    function stopTimer(){{ if(progressTimer){{clearInterval(progressTimer);progressTimer=null;}} }}
    function resetProgress(txt){{ setProgress(0); timeLabel.textContent='';
      label.textContent=txt||'Narration: —'; stopTimer(); }}

    function startLoop(dur, onDone){{
      currentDuration = dur && isFinite(dur) && dur>0 ? dur : null;
      startedAt = performance.now();
      pausedAccum = 0; isPaused = false; pauseStart = 0;
      stopTimer();
      progressTimer = setInterval(function(){{
        if (isPaused) return;
        var el = (performance.now()-startedAt-pausedAccum)/1000;
        if (currentDuration){{
          setProgress(el/currentDuration*100);
          timeLabel.textContent = fmt(Math.min(el,currentDuration))+' / '+fmt(currentDuration);
          if (el >= currentDuration){{ stopTimer(); setProgress(100); if(onDone)onDone(); }}
        }} else {{
          setProgress(el/3*100);
          timeLabel.textContent = fmt(el)+' elapsed';
          if (el >= 3){{ stopTimer(); setProgress(100); if(onDone)onDone(); }}
        }}
      }}, 100);
    }}
    function pauseTimer(){{ if (isPaused) return;
      isPaused = true; pauseStart = performance.now(); }}
    function resumeTimer(){{ if (!isPaused) return; isPaused = false;
      if (pauseStart) pausedAccum += performance.now()-pauseStart; pauseStart=0; }}

    function lockNext(t){{ nextBtn.disabled=true;
      nextHint.textContent = t || 'Listen to narration to unlock Next'; }}
    function unlockNext(){{
      unlocked[slideNum()] = true;
      if (idx === slides.length-1){{
        nextBtn.disabled = false;
        nextBtn.textContent = meta.hasExam ? 'Finish → Exam' : 'Finish → Dashboard';
        nextHint.textContent = meta.hasExam ? 'Click to take the exam'
                                            : 'Click to return to dashboard';
      }} else {{
        nextBtn.disabled = false;
        nextBtn.textContent = 'Next →';
        nextHint.textContent = '✅ Ready';
      }}
    }}

    function markSeen(n){{
      try {{
        fetch('/course/'+courseId+'/mark-seen', {{
          method:'POST', headers:{{'Content-Type':'application/json'}},
          body: JSON.stringify({{slide:n}})
        }}).catch(function(){{}});
        fetch('/course/'+courseId+'/progress', {{
          method:'POST', headers:{{'Content-Type':'application/json'}},
          body: JSON.stringify({{slide:n, done:true}})
        }}).catch(function(){{}});
      }} catch(e){{}}
    }}

    function render(auto){{
      if (idx<0 || idx>=slides.length) idx = 0;
      var s = slides[idx];
      canvas.innerHTML = '';
      if (s.image){{
        var img = document.createElement('img');
        img.src = 'data:image/jpeg;base64,'+s.image;
        canvas.appendChild(img);
      }} else if (s.text && s.text.trim()){{
        var d = document.createElement('div');
        d.className = 'txt';
        d.textContent = s.text;
        canvas.appendChild(d);
      }}
      counter.textContent = 'Slide '+s.number+' of '+slides.length;
      prevBtn.disabled = idx===0;
      stopAudio();
      resetProgress('Narration: not started');
      if (idx === slides.length-1){{
        nextBtn.textContent = meta.hasExam ? 'Finish → Exam' : 'Finish → Dashboard';
      }} else {{
        nextBtn.textContent = 'Next →';
      }}
      if (!unlocked[s.number]){{
        lockNext(idx===slides.length-1
          ? (meta.hasExam ? 'Listen to the narration to finish and unlock the exam'
                          : 'Listen to the narration to finish the course')
          : 'Listen to narration to unlock Next');
      }} else {{
        unlockNext();
      }}
      pauseBtn.style.display = 'inline-block';
      markSeen(s.number);
      if (auto){{
        setTimeout(function(){{
          if (!unlocked[s.number]) playNarration();
        }}, 250);
      }}
    }}

    function stopAudio(){{
      try {{ audio.pause(); audio.currentTime=0; }} catch(e){{}}
      audio.removeAttribute('src');
      try {{ audio.load(); }} catch(e){{}}
      if (window.speechSynthesis) window.speechSynthesis.cancel();
      mode='idle'; isPaused=false; pausedAccum=0; stopTimer();
      playBtn.textContent = '🔊 Play narration';
    }}

    function playNarration(){{
      if (mode==='audio' && audio.src && !audio.paused){{
        audio.pause(); pauseTimer(); playBtn.textContent='▶ Resume'; return Promise.resolve();
      }}
      if (mode==='audio' && audio.src && audio.paused){{
        audio.play(); resumeTimer(); playBtn.textContent='⏸ Pause'; return Promise.resolve();
      }}
      if (mode==='tts' && window.speechSynthesis && window.speechSynthesis.speaking
          && !window.speechSynthesis.paused){{
        window.speechSynthesis.pause(); pauseTimer();
        playBtn.textContent='▶ Resume'; return Promise.resolve();
      }}
      if (mode==='tts' && window.speechSynthesis && window.speechSynthesis.paused){{
        window.speechSynthesis.resume(); resumeTimer();
        playBtn.textContent='⏸ Pause'; return Promise.resolve();
      }}
      if (window.speechSynthesis && window.speechSynthesis.speaking){{
        window.speechSynthesis.cancel(); mode='idle';
        playBtn.textContent='🔊 Play narration'; return Promise.resolve();
      }}

      var s = slides[idx];
      playBtn.textContent = '⏳ Loading…';
      resetProgress('Narration: loading…');

      return fetch('/course/'+courseId+'/slide/'+s.number+'/audio?t='+Date.now())
        .then(function(res){{
          if (res.ok){{
            var durH = parseFloat(res.headers.get('X-Audio-Duration')||'0');
            return res.blob().then(function(blob){{
              var url = URL.createObjectURL(blob);
              mode = 'audio';
              audio.src = url;
              audio.onloadedmetadata = function(){{
                var dur = audio.duration && isFinite(audio.duration) ? audio.duration : (durH||0);
                playBtn.textContent = '⏸ Pause';
                startLoop(dur, onFinished);
              }};
              audio.onended = function(){{ onFinished(); }};
              audio.onerror = function(){{
                resetProgress('Narration: error');
                playBtn.textContent = '🔊 Play narration'; mode='idle';
              }};
              return audio.play();
            }});
          }}
          return res.json().then(function(j){{
            var t = (j.text||'').trim() || (s.text||'').trim();
            var fbLang = (j.lang === 'am') ? 'am-ET' : 'en-US';
            if (t && window.speechSynthesis) speakFallback(t, fbLang);
            else {{
              playBtn.textContent = '🔊 Play narration';
              resetProgress('Narration: unavailable'); mode='idle';
            }}
          }}).catch(function(){{
            var t = (s.text||'').trim();
            if (t && window.speechSynthesis) speakFallback(t, 'en-US');
          }});
        }})
        .catch(function(err){{
          console.warn('audio fetch failed', err);
          playBtn.textContent = '🔊 Play narration';
          resetProgress('Narration: error'); mode='idle';
        }});
    }}

    function speakFallback(text, lang){{
      mode = 'tts';
      var u = new SpeechSynthesisUtterance(text);
      u.rate = 1; u.pitch = 1; u.lang = lang || 'en-US';
      try {{
        var wanted = u.lang.split('-')[0].toLowerCase();
        var vs = window.speechSynthesis.getVoices() || [];
        var match = vs.find(function(v){{
          return (v.lang||'').toLowerCase().indexOf(wanted) === 0;
        }});
        if (match) u.voice = match;
      }} catch(e) {{}}
      u.onstart = function(){{
        playBtn.textContent = '⏸ Pause';
        var dur = Math.max(5, u.text.length/14);
        startLoop(dur, onFinished);
      }};
      u.onend = function(){{ onFinished(); }};
      u.onerror = function(){{
        playBtn.textContent = '🔊 Play narration';
        resetProgress('Narration: error'); mode='idle';
      }};
      window.speechSynthesis.speak(u);
    }}

    function onFinished(){{
      playBtn.textContent = '🔊 Play narration';
      setProgress(100); label.textContent='Narration: complete';
      stopTimer(); isPaused=false; mode='idle';
      unlockNext();
      if (idx === slides.length-1){{
        setTimeout(function(){{
          if (meta.hasExam) goExam();
          else window.location.href = meta.dashboardUrl;
        }}, 800);
      }}
    }}

    function goExam(){{
      if (!meta.hasExam){{ window.location.href = meta.dashboardUrl; return; }}
      var sp = document.createElement('div');
      sp.className = 'splash';
      sp.innerHTML = '<div class="spin"></div><div>Opening exam…</div>';
      document.body.appendChild(sp);
      setTimeout(function(){{ window.location.href = meta.examUrl; }}, 700);
    }}

    function savePause(){{
      fetch('/course/'+courseId+'/pause', {{
        method:'POST', headers:{{'Content-Type':'application/json'}},
        body: JSON.stringify({{slide: slideNum()}})
      }}).then(function(r){{ return r.json(); }})
        .then(function(d){{
          label.textContent = 'Paused — progress saved';
          pauseBtn.textContent = '▶ Resume later';
          toast('Progress saved — come back anytime to resume.', 'ok');
        }})
        .catch(function(){{
          toast('Could not save pause state.', 'err');
        }});
    }}

    function toast(msg, cat){{
      var d = document.createElement('div');
      d.className = 'toast ' + cat;
      d.textContent = msg;
      document.body.appendChild(d);
      setTimeout(function(){{ d.remove(); }}, 3500);
    }}

    function checkResume(){{
      fetch('/course/'+courseId+'/resume-info', {{cache:'no-store'}})
        .then(function(r){{ return r.ok ? r.json() : {{resume:null}}; }})
        .then(function(data){{
          if (!data.resume || !data.resume.slide){{ render(false); return; }}
          var targetIdx = -1;
          for (var i=0; i<slides.length; i++){{
            if (slides[i].number === data.resume.slide){{ targetIdx=i; break; }}
          }}
          if (targetIdx < 0){{ render(false); return; }}

          document.getElementById('resumeSlide').textContent = data.resume.slide;
          document.getElementById('resumeWhen').textContent = ' — saved ' + data.resume.when;
          resumeBanner.style.display = 'flex';

          idx = 0;
          render(false);

          document.getElementById('resumeBtn').onclick = function(){{
            idx = targetIdx;
            resumeBanner.style.display = 'none';
            fetch('/course/'+courseId+'/resume-clear', {{method:'POST'}}).catch(function(){{}});
            render(true);
            toast('Resuming from slide '+data.resume.slide, 'ok');
          }};
          document.getElementById('startOverBtn').onclick = function(){{
            idx = 0;
            resumeBanner.style.display = 'none';
            fetch('/course/'+courseId+'/resume-clear', {{method:'POST'}}).catch(function(){{}});
            render(true);
          }};
        }})
        .catch(function(e){{
          console.warn('resume-info failed', e);
          render(false);
        }});
    }}

    prevBtn.onclick = function(){{ if (idx>0){{ idx--; render(false); }} }};
    nextBtn.onclick = function(){{
      if (nextBtn.disabled) return;
      if (idx === slides.length-1) goExam();
      else {{ idx++; render(true); }}
    }};
    playBtn.onclick = playNarration;
    pauseBtn.onclick = savePause;

    document.addEventListener('keydown', function(e){{
      if (e.target && (e.target.tagName==='INPUT'||e.target.tagName==='TEXTAREA')) return;
      if (e.key==='ArrowRight' && !nextBtn.disabled) nextBtn.click();
      if (e.key==='ArrowLeft') prevBtn.click();
      if (e.key===' ' ){{ e.preventDefault(); playBtn.click(); }}
    }});

    checkResume();
  }} catch(err){{
    console.error('course_view init error:', err);
    var d = document.createElement('div');
    d.style.cssText = 'background:#fee2e2;color:#991b1b;padding:1rem;'+
      'border-radius:.5rem;margin:1rem auto;max-width:600px;';
    d.textContent = 'Course view failed to load: ' + err.message;
    document.body.insertBefore(d, document.body.firstChild);
  }}
}})();
</script>
</body>
</html>"""


# ---------------- TRAINEE ROUTES ----------------

@app.route('/')
def index():
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    error_msg = ''
    if request.method == 'POST':
        identifier = request.form.get('identifier', '').strip()
        if not identifier:
            identifier = request.form.get('trainer_id', '').strip()
        password = request.form.get('password', '')

        if not identifier or not password:
            error_msg = 'Email/ID and password are required.'
        else:
            trainer = None
            if '@' in identifier:
                trainer = Trainer.query.filter(
                    db.func.lower(Trainer.email) == identifier.lower()
                ).first()
            if trainer is None:
                if identifier.isdigit() and len(identifier) == 5:
                    trainer = Trainer.query.filter_by(trainer_id=identifier).first()

            if trainer and check_password_hash(trainer.password_hash, password):
                session['trainer_id'] = trainer.id
                session['trainer_name'] = trainer.display_name
                flash(f'Welcome back, {trainer.display_name}!', 'success')
                return redirect(url_for('trainer_dashboard'))
            error_msg = 'Invalid credentials.'

    error_html = (f'<div style="background:#fee2e2;color:#991b1b;'
                  f'padding:.6rem .9rem;border-radius:6px;'
                  f'margin-bottom:1rem;">{error_msg}</div>'
                  if error_msg else '')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Login — MULSEN Training</title>
<style>
  *,*::before,*::after{{box-sizing:border-box;}}
  body{{font-family:system-ui,Segoe UI,Roboto,sans-serif;
        background:#f5f7fa;margin:0;padding:2rem;color:#1f2933;
        -webkit-tap-highlight-color:transparent;}}
  .wrap{{max-width:420px;margin:3rem auto;background:#fff;
        padding:2rem;border-radius:12px;
        box-shadow:0 2px 8px rgba(0,0,0,.06);}}
  h1{{margin-top:0;text-align:center;}}
  label{{display:block;font-size:.85rem;font-weight:500;
        margin:1rem 0 .3rem;}}
  input{{width:100%;padding:.6rem;border:1px solid #d1d5db;
        border-radius:6px;font-size:.95rem;box-sizing:border-box;
        min-height:42px;}}
  input:focus{{outline:2px solid #0b6b3a;border-color:transparent;}}
  .submit{{width:100%;padding:.75rem;margin-top:1.25rem;
          background:#0b6b3a;color:#fff;border:0;border-radius:6px;
          font-size:1rem;font-weight:600;cursor:pointer;
          min-height:44px;font-family:inherit;}}
  .submit:hover{{background:#095a31;}}
  .signup-btn{{display:block;text-align:center;padding:.9rem 1rem;
              background:#0b6b3a;color:#ffffff;font-weight:700;
              font-size:1rem;text-decoration:none;border-radius:8px;
              letter-spacing:.3px;min-height:44px;}}
  .signup-btn:hover{{background:#095a31;}}
  .sep{{border:0;border-top:1px solid #e5e7eb;margin:1.5rem 0;}}
  .footer{{text-align:center;margin-top:1rem;font-size:.85rem;
          color:#6b7280;}}
  .footer a{{color:#0b6b3a;text-decoration:none;font-weight:500;}}

  @media (max-width:640px){{
    body{{padding:.75rem;}}
    .wrap{{margin:.75rem auto;padding:1.25rem;border-radius:8px;}}
    h1{{font-size:1.25rem;}}
    input{{font-size:16px;padding:.55rem .65rem;}}
  }}
  @media (max-width:380px){{
    body{{padding:.4rem;}}
    .wrap{{padding:1rem;}}
  }}
</style>
</head>
<body>
<div class="wrap">
  <h1>👋 Trainee Login</h1>

  {error_html}

  <form method="post">
    <label for="identifier">Email (or 5-digit ID)</label>
    <input id="identifier" type="text" name="identifier"
           autocomplete="username" required autofocus
           placeholder="you@company.com">

    <label for="password">Password</label>
    <input id="password" type="password" name="password"
           autocomplete="current-password" required>

    <button class="submit" type="submit">Sign in</button>
  </form>

  <hr class="sep">

  <a href="/signup" class="signup-btn">
    ✍ SIGN UP — Create new account
  </a>

  <p class="footer">
    Admin? <a href="/admin/login">Admin login →</a>
  </p>
</div>
</body>
</html>"""


@app.route('/register', methods=['GET', 'POST'])
def register():
    return signup_inline()


@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login'))


@app.route('/trainer')
@login_required
def trainer_dashboard():
    trainer = Trainer.query.get(session['trainer_id'])
    courses = Course.query.order_by(Course.id.desc()).all()
    attempts = Attempt.query.filter_by(trainer_id=trainer.id) \
                            .order_by(Attempt.completed_at.desc()).all()

    status = {}
    for c in courses:
        cs_attempts = [a for a in attempts if a.course_id == c.id]
        best = max((a.percentage for a in cs_attempts), default=0)
        progress_rows = Progress.query.filter_by(
            trainer_id=trainer.id, course_id=c.id).all()
        status[c.id] = {
            'attempts_used': len(cs_attempts),
            'best': best,
            'passed': any(a.passed for a in cs_attempts),
            'completed_slides': sum(1 for p in progress_rows
                                    if p.completed and p.slide_number > 0),
        }

    return render_template(
        'trainer_dashboard.html',
        trainer=trainer,
        courses=courses,
        attempts=attempts,
        status=status,
        materials=courses,
    )


@app.route('/change_password', methods=['GET', 'POST'])
@login_required
def change_password():
    if request.method == 'POST':
        old = request.form.get('old_password', '')
        new = request.form.get('new_password', '')
        confirm = request.form.get('confirm_password', '')
        trainer = Trainer.query.get(session['trainer_id'])
        if not check_password_hash(trainer.password_hash, old):
            flash('Current password is incorrect.', 'error')
        elif len(new) < 6:
            flash('New password must be at least 6 characters.', 'error')
        elif new != confirm:
            flash('Passwords do not match.', 'error')
        else:
            trainer.password_hash = generate_password_hash(new)
            db.session.commit()

            if trainer.email:
                _send_mail(
                    trainer.email,
                    'Your MULSEN password was changed',
                    f'Hello {trainer.display_name},\n\n'
                    f'Your password was just changed. If this was not you, '
                    f'contact your administrator.\n\n— MULSEN Training',
                )
            flash('Password changed successfully.', 'success')
            return redirect(url_for('trainer_dashboard'))
    return render_template('change_password.html')


@app.route('/course/<int:course_id>/progress', methods=['POST'])
@login_required
def save_progress(course_id):
    data = request.get_json() or {}
    slide_number = int(data.get('slide', 0))
    done = bool(data.get('done', False))
    if slide_number <= 0:
        return jsonify({'ok': False}), 400
    row = Progress.query.filter_by(
        trainer_id=session['trainer_id'],
        course_id=course_id,
        slide_number=slide_number).first()
    if not row:
        row = Progress(trainer_id=session['trainer_id'],
                       course_id=course_id,
                       slide_number=slide_number)
        db.session.add(row)
    row.completed = done
    db.session.commit()
    return jsonify({'ok': True})


@app.route('/exam/<int:course_id>', methods=['GET', 'POST'])
@login_required
def exam(course_id):
    course = Course.query.get_or_404(course_id)
    questions = course.questions
    if not questions:
        flash('This course has no exam questions yet.', 'warning')
        return redirect(url_for('trainer_dashboard'))

    prior_attempts = Attempt.query.filter_by(
        trainer_id=session['trainer_id'], course_id=course.id).count()
    if prior_attempts >= 2:
        flash('You have used all 2 attempts for this course.', 'error')
        return redirect(url_for('trainer_dashboard'))

    if request.method == 'POST':
        correct = 0
        for q in questions:
            answer = request.form.get(f'q_{q.id}', '').upper()
            if answer == q.correct_answer.upper():
                correct += 1
        total = len(questions)
        percentage = (correct / total) * 100
        passed = percentage > 70
        distinction = percentage > 90

        attempt = Attempt(
            trainer_id=session['trainer_id'],
            course_id=course.id,
            score=correct, total=total,
            percentage=percentage,
            passed=passed, distinction=distinction,
        )
        db.session.add(attempt)
        db.session.commit()

        if passed:
            flash(f'Congratulations! You scored {percentage:.1f}%!', 'success')
        else:
            flash(f'You scored {percentage:.1f}%. You need >70% to pass.', 'warning')
        return redirect(url_for('result', attempt_id=attempt.id))

    return render_template('exam.html', course=course, questions=questions,
                           attempt_number=prior_attempts + 1)


@app.route('/result/<int:attempt_id>')
@login_required
def result(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)
    if attempt.trainer_id != session['trainer_id']:
        flash('Unauthorized.', 'error')
        return redirect(url_for('trainer_dashboard'))
    return render_template('certificate.html', attempt=attempt)


@app.route('/certificate/<int:attempt_id>')
@login_required
def download_certificate(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)
    if attempt.trainer_id != session['trainer_id'] or not attempt.passed:
        flash('Certificate not available.', 'error')
        return redirect(url_for('trainer_dashboard'))
    trainer = attempt.trainer
    path = generate_certificate(trainer.display_name, attempt.course.name,
                                attempt.percentage, attempt.distinction,
                                attempt_id=attempt.id)
    return send_file(path, as_attachment=True,
                     download_name=f"Certificate_{attempt.course.name}.pdf")


# ---------------- ADMIN ROUTES ----------------

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        admin = Admin.query.filter_by(username=username).first()
        if admin and check_password_hash(admin.password_hash, password):
            session['admin_id'] = admin.id
            session['admin_username'] = admin.username
            flash('Welcome Admin!', 'success')
            return redirect(url_for('admin_courses'))
        flash('Invalid admin credentials.', 'error')
    return render_template('admin_login.html')


@app.route('/admin')
@admin_required
def admin_dashboard():
    return redirect(url_for('admin_courses'))


@app.route('/admin/courses')
@admin_required
def admin_courses():
    courses = Course.query.order_by(Course.id.desc()).all()

    course_audio = {}
    for c in courses:
        scripts = InstructorScript.query.filter_by(course_id=c.id).all()
        recs = PageRecording.query.filter_by(course_id=c.id).all()
        am_pages = sum(1 for s in scripts if (s.language or 'en') == 'am')
        en_pages = len(scripts) - am_pages
        course_audio[c.id] = {
            'script_pages': len(scripts),
            'am_pages': am_pages,
            'en_pages': en_pages,
            'recorded_pages': len(recs),
            'slide_count': effective_slide_count(c, scripts, recs),
        }

    return render_template('admin_courses.html',
                           courses=courses,
                           course_audio=course_audio)


@app.route('/admin/exams')
@app.route('/admin/exams/<int:course_id>')
@admin_required
def admin_exams(course_id=None):
    courses = Course.query.order_by(Course.id.desc()).all()
    selected = None
    if course_id:
        selected = Course.query.get_or_404(course_id)
    elif courses:
        selected = courses[0]
    return render_template('admin_exams.html',
                           courses=courses,
                           selected=selected)


@app.route('/admin/scripts')
@app.route('/admin/scripts/<int:course_id>')
@admin_required
def admin_scripts(course_id=None):
    courses = Course.query.order_by(Course.id.desc()).all()
    selected = None
    if course_id:
        selected = Course.query.get_or_404(course_id)
    elif courses:
        selected = courses[0]

    course_audio = {}
    for c in courses:
        scripts = InstructorScript.query.filter_by(course_id=c.id).all()
        recs = PageRecording.query.filter_by(course_id=c.id).all()
        am_pages = sum(1 for s in scripts if (s.language or 'en') == 'am')
        en_pages = len(scripts) - am_pages
        course_audio[c.id] = {
            'script_pages': len(scripts),
            'am_pages': am_pages,
            'en_pages': en_pages,
            'recorded_pages': len(recs),
            'slide_count': effective_slide_count(c, scripts, recs),
        }

    selected_scripts = []
    selected_recordings = {}
    if selected:
        selected_scripts = (InstructorScript.query
                            .filter_by(course_id=selected.id)
                            .order_by(InstructorScript.page_number)
                            .all())
        for r in PageRecording.query.filter_by(course_id=selected.id).all():
            selected_recordings[r.page_number] = r

        if course_audio[selected.id]['slide_count'] is None:
            if selected_scripts:
                course_audio[selected.id]['slide_count'] = max(
                    s.page_number for s in selected_scripts)
            elif selected_recordings:
                course_audio[selected.id]['slide_count'] = max(
                    selected_recordings.keys())

    return render_template('admin_scripts.html',
                           courses=courses,
                           selected=selected,
                           course_audio=course_audio,
                           selected_scripts=selected_scripts,
                           selected_recordings=selected_recordings)


@app.route('/admin/trainees')
@admin_required
def admin_trainees():
    trainers = Trainer.query.order_by(Trainer.trainer_id).all()
    return render_template('admin_trainees.html', trainers=trainers)


@app.route('/admin/records')
@admin_required
def admin_records():
    attempts = Attempt.query.order_by(Attempt.completed_at.desc()).all()
    trainers = Trainer.query.order_by(Trainer.trainer_id).all()
    courses  = Course.query.order_by(Course.id).all()

    summary = []
    for t in trainers:
        rows = []
        for c in courses:
            tas = [a for a in attempts
                   if a.trainer_id == t.id and a.course_id == c.id]
            if tas:
                best = max(tas, key=lambda x: x.percentage)
                rows.append({
                    'course': c,
                    'attempts': len(tas),
                    'best': best.percentage,
                    'passed': best.passed,
                    'distinction': best.distinction,
                    'last': max(a.completed_at for a in tas),
                    'best_attempt_id': best.id,
                })
        total_courses = len(rows)
        passed_courses = sum(1 for r in rows if r['passed'])
        distinctions = sum(1 for r in rows if r['distinction'])
        avg_best = (sum(r['best'] for r in rows) / total_courses) if total_courses else 0
        summary.append({
            'trainer': t,
            'rows': rows,
            'total_courses': total_courses,
            'passed_courses': passed_courses,
            'distinctions': distinctions,
            'avg_best': avg_best,
        })

    recent = Attempt.query.order_by(Attempt.completed_at.desc()).limit(200).all()

    return render_template('admin_records.html',
                           summary=summary,
                           recent=recent,
                           trainers=trainers,
                           courses=courses)


@app.route('/admin/records/export.csv')
@admin_required
def export_records_csv():
    trainers = Trainer.query.order_by(Trainer.trainer_id).all()
    courses  = Course.query.order_by(Course.id).all()
    attempts = Attempt.query.all()

    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(['Trainer ID', 'Full Name', 'Email', 'Department', 'Position',
                'Course', 'Attempts', 'Best %', 'Result', 'Distinction',
                'Last Attempt'])

    for t in trainers:
        for c in courses:
            tas = [a for a in attempts
                   if a.trainer_id == t.id and a.course_id == c.id]
            if not tas:
                continue
            best = max(tas, key=lambda x: x.percentage)
            w.writerow([
                t.trainer_id,
                t.display_name,
                t.email or '',
                t.department or '',
                t.position or '',
                c.name,
                len(tas),
                f'{best.percentage:.1f}',
                'Passed' if best.passed else 'Not passed',
                'Yes' if best.distinction else 'No',
                max(a.completed_at for a in tas).strftime('%Y-%m-%d %H:%M'),
            ])

    buf.seek(0)
    return send_file(
        io.BytesIO(buf.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=f'training_report_{datetime.now().strftime("%Y%m%d_%H%M")}.csv'
    )


# --- course management ---

@app.route('/admin/course/create', methods=['POST'])
@admin_required
def create_course():
    name = request.form.get('name', '').strip()
    description = request.form.get('description', '').strip()
    if not name:
        flash('Course name required.', 'error')
        return redirect(url_for('admin_courses'))

    existing = Course.query.filter(
        db.func.lower(Course.name) == name.lower()
    ).first()
    if existing:
        flash(f'A course named "{name}" already exists. '
              f'Edit it instead of creating a duplicate.', 'error')
        return redirect(url_for('admin_courses'))

    course = Course(name=name, description=description)

    ppt_file = request.files.get('ppt_file')
    if ppt_file and ppt_file.filename:
        fname_lower = ppt_file.filename.lower()
        if fname_lower.endswith('.pptx'):
            course.source_type = 'pptx'
        elif fname_lower.endswith('.pdf'):
            course.source_type = 'pdf'
        else:
            flash('Only .pptx or .pdf files are accepted.', 'error')
            return redirect(url_for('admin_courses'))

        filename = secure_filename(
            f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{ppt_file.filename}")
        ppt_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        ppt_file.save(ppt_path)
        course.ppt_filename = filename

    db.session.add(course)
    db.session.commit()
    flash(f'✅ Course "{name}" created successfully.', 'success')
    return redirect(url_for('admin_courses'))


@app.route('/admin/course/<int:course_id>/edit', methods=['POST'])
@admin_required
def edit_course(course_id):
    course = Course.query.get_or_404(course_id)
    name = request.form.get('name', '').strip()
    description = request.form.get('description', '').strip()
    if not name:
        flash('Course name cannot be empty.', 'error')
        return redirect(url_for('admin_courses'))

    dup = Course.query.filter(
        db.func.lower(Course.name) == name.lower(),
        Course.id != course.id
    ).first()
    if dup:
        flash(f'Another course named "{name}" already exists.', 'error')
        return redirect(url_for('admin_courses'))

    course.name = name
    course.description = description

    ppt_file = request.files.get('ppt_file')
    if ppt_file and ppt_file.filename:
        fname_lower = ppt_file.filename.lower()
        if fname_lower.endswith('.pptx'):
            course.source_type = 'pptx'
        elif fname_lower.endswith('.pdf'):
            course.source_type = 'pdf'
        else:
            flash('Only .pptx or .pdf files are accepted.', 'error')
            return redirect(url_for('admin_courses'))

        if course.ppt_filename:
            old = os.path.join(app.config['UPLOAD_FOLDER'], course.ppt_filename)
            if os.path.exists(old):
                try:
                    os.remove(old)
                except OSError:
                    pass

        filename = secure_filename(
            f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{ppt_file.filename}")
        ppt_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        course.ppt_filename = filename

    db.session.commit()
    flash(f'✅ Course "{course.name}" updated.', 'success')
    return redirect(url_for('admin_courses'))


@app.route('/admin/course/<int:course_id>/delete', methods=['POST'])
@admin_required
def delete_course(course_id):
    course = Course.query.get_or_404(course_id)
    course_name = course.name
    try:
        Attempt.query.filter_by(course_id=course.id).delete(synchronize_session=False)
        Progress.query.filter_by(course_id=course.id).delete(synchronize_session=False)
        Question.query.filter_by(course_id=course.id).delete(synchronize_session=False)
        InstructorScript.query.filter_by(course_id=course.id).delete(synchronize_session=False)

        recs = PageRecording.query.filter_by(course_id=course.id).all()
        rec_filenames = [r.filename for r in recs]
        PageRecording.query.filter_by(course_id=course.id).delete(synchronize_session=False)

        ppt_filename = course.ppt_filename
        db.session.delete(course)
        db.session.commit()

        if ppt_filename:
            p = os.path.join(app.config['UPLOAD_FOLDER'], ppt_filename)
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError as e:
                print(f'[delete_course] failed to remove {p}: {e}')

        for fname in rec_filenames:
            p = os.path.join(RECORDINGS_DIR, fname)
            try:
                if os.path.exists(p):
                    os.remove(p)
            except OSError as e:
                print(f'[delete_course] failed to remove {p}: {e}')

        for old in os.listdir(TTS_CACHE_DIR):
            if old.startswith(f'course{course_id}_page'):
                try:
                    os.remove(os.path.join(TTS_CACHE_DIR, old))
                except OSError:
                    pass

        manual = _load_manual_counts()
        if str(course_id) in manual:
            del manual[str(course_id)]
            _save_manual_counts(manual)

        flash(f'🗑 Course "{course_name}" deleted.', 'success')
    except Exception as e:
        db.session.rollback()
        print(f'[delete_course] failed: {e}')
        flash(f'Could not delete course: {e}', 'error')

    return redirect(url_for('admin_courses'))


# --- question management ---

@app.route('/admin/course/<int:course_id>/question', methods=['POST'])
@admin_required
def add_question(course_id):
    course = Course.query.get_or_404(course_id)
    qtype = request.form.get('question_type', 'mc')

    if qtype == 'tf':
        correct = request.form.get('correct_answer_tf', 'T').upper()
        if correct not in ('T', 'F'):
            correct = 'T'
        opt_a = opt_b = opt_c = opt_d = None
    else:
        correct = request.form.get('correct_answer', 'A').upper()
        opt_a = request.form.get('option_a', '').strip() or None
        opt_b = request.form.get('option_b', '').strip() or None
        opt_c = request.form.get('option_c', '').strip() or None
        opt_d = request.form.get('option_d', '').strip() or None

    text = request.form.get('question_text', '').strip()
    if not text:
        flash('Question text required.', 'error')
        return redirect(url_for('admin_exams', course_id=course.id))

    q = Question(
        course_id=course.id,
        question_type=qtype,
        question_text=text,
        option_a=opt_a, option_b=opt_b, option_c=opt_c, option_d=opt_d,
        correct_answer=correct,
    )
    db.session.add(q)
    db.session.commit()
    flash(f'✅ Question added to "{course.name}".', 'success')
    return redirect(url_for('admin_exams', course_id=course.id))


@app.route('/admin/question/<int:q_id>/edit', methods=['POST'])
@admin_required
def edit_question(q_id):
    q = Question.query.get_or_404(q_id)
    qtype = request.form.get('question_type', q.question_type)

    text = request.form.get('question_text', '').strip()
    if not text:
        flash('Question text required.', 'error')
        return redirect(url_for('admin_exams', course_id=q.course_id))

    q.question_text = text
    q.question_type = qtype

    if qtype == 'tf':
        correct = request.form.get('correct_answer_tf', 'T').upper()
        q.correct_answer = correct if correct in ('T', 'F') else 'T'
        q.option_a = q.option_b = q.option_c = q.option_d = None
    else:
        q.correct_answer = (request.form.get('correct_answer', 'A') or 'A').upper()
        q.option_a = request.form.get('option_a', '').strip() or None
        q.option_b = request.form.get('option_b', '').strip() or None
        q.option_c = request.form.get('option_c', '').strip() or None
        q.option_d = request.form.get('option_d', '').strip() or None

    db.session.commit()
    flash('✅ Question updated.', 'success')
    return redirect(url_for('admin_exams', course_id=q.course_id))


@app.route('/admin/question/<int:q_id>/delete', methods=['POST'])
@admin_required
def delete_question(q_id):
    q = Question.query.get_or_404(q_id)
    course_id = q.course_id
    db.session.delete(q)
    db.session.commit()
    flash('🗑 Question deleted.', 'success')
    return redirect(url_for('admin_exams', course_id=course_id))


# --- trainee management ---

@app.route('/admin/trainer/create', methods=['POST'])
@admin_required
def create_trainer():
    email = (request.form.get('email') or '').strip().lower()
    full_name = (request.form.get('full_name') or '').strip()
    department = (request.form.get('department') or '').strip()
    position = (request.form.get('position') or '').strip()

    if not full_name:
        flash('Full name is required.', 'error')
        return redirect(url_for('admin_trainees'))

    if email and Trainer.query.filter(
            db.func.lower(Trainer.email) == email).first():
        flash(f'Email {email} already registered.', 'error')
        return redirect(url_for('admin_trainees'))

    try:
        new_id = _next_trainer_id()
        password = _gen_password(10)
        trainee = Trainer(
            trainer_id=new_id,
            email=email or None,
            full_name=full_name,
            department=department or '',
            position=position or '',
            name=(full_name.split()[0][:100] if full_name.split() else 'Trainee'),
            password_hash=generate_password_hash(password),
        )
        db.session.add(trainee)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        import traceback
        traceback.print_exc()
        flash(f'Could not create trainee: {e}', 'error')
        return redirect(url_for('admin_trainees'))

    message = f'✅ Trainee "{full_name}" created (ID {new_id}, password {password}).'
    if email:
        try:
            body_text, body_html = _welcome_email_body(full_name, email, password)
            sent, err = _send_mail(email, 'Your MULSEN Training account',
                                   body_text, body_html)
            if sent:
                message += ' Password emailed.'
            else:
                message += f' ⚠ Email delivery failed ({err}) — give password manually.'
        except Exception as e:
            message += ' ⚠ Email delivery error — give password manually.'

    flash(message, 'success')
    return redirect(url_for('admin_trainees'))


@app.route('/admin/trainer/<int:trainer_pk>/reset', methods=['POST'])
@admin_required
def reset_password(trainer_pk):
    trainer = Trainer.query.get_or_404(trainer_pk)
    new_pw = _gen_password(10)
    trainer.password_hash = generate_password_hash(new_pw)
    db.session.commit()

    message = f'🔑 Password reset for "{trainer.display_name}" — new password: {new_pw}'
    if trainer.email:
        body_text, body_html = _welcome_email_body(
            trainer.display_name, trainer.email, new_pw)
        sent, err = _send_mail(trainer.email,
                               'Your MULSEN password has been reset',
                               body_text, body_html)
        if sent:
            message += ' (emailed)'
        else:
            message += f' ⚠ email delivery failed ({err})'
    flash(message, 'success')
    return redirect(url_for('admin_trainees'))


@app.route('/admin/trainer/<int:trainer_pk>/delete', methods=['POST'])
@admin_required
def delete_trainer(trainer_pk):
    trainer = Trainer.query.get_or_404(trainer_pk)
    try:
        Attempt.query.filter_by(trainer_id=trainer.id).delete(synchronize_session=False)
        Progress.query.filter_by(trainer_id=trainer.id).delete(synchronize_session=False)
        db.session.delete(trainer)
        db.session.commit()
        flash('🗑 Trainee deleted.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Could not delete trainee: {e}', 'error')
    return redirect(url_for('admin_trainees'))


@app.route('/admin/certificate/<int:attempt_id>')
@admin_required
def admin_download_certificate(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)
    if not attempt.passed:
        flash('Certificate not available — trainee did not pass.', 'error')
        return redirect(url_for('admin_records'))
    trainer = attempt.trainer
    path = generate_certificate(trainer.display_name, attempt.course.name,
                                attempt.percentage, attempt.distinction,
                                attempt_id=attempt.id)
    inline = request.args.get('inline') == '1'
    return send_file(
        path,
        mimetype='application/pdf',
        as_attachment=not inline,
        download_name=f"Certificate_{trainer.trainer_id}_{attempt.course.name}.pdf",
    )


@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_id', None)
    session.pop('admin_username', None)
    flash('Admin logged out.', 'info')
    return redirect(url_for('admin_login'))


# ---------------- DB INIT + LIGHT MIGRATIONS ----------------

def _light_migrate():
    try:
        from sqlalchemy import inspect, text
        insp = inspect(db.engine)
        if 'trainer' not in insp.get_table_names():
            return

        existing = {c['name'] for c in insp.get_columns('trainer')}
        wanted = {
            'email':      "VARCHAR(255)",
            'full_name':  "VARCHAR(200) DEFAULT ''",
            'department': "VARCHAR(200) DEFAULT ''",
            'position':   "VARCHAR(200) DEFAULT ''",
        }
        with db.engine.connect() as conn:
            for col, typ in wanted.items():
                if col not in existing:
                    try:
                        conn.execute(text(f"ALTER TABLE trainer ADD COLUMN {col} {typ}"))
                        conn.commit()
                        print(f"[migrate] added column trainer.{col}")
                    except Exception as e:
                        print(f"[migrate] failed to add trainer.{col}: {e}")

        if 'progress' in insp.get_table_names():
            existing_p = {c['name'] for c in insp.get_columns('progress')}
            wanted_p = {
                'paused_at':   "DATETIME",
                'last_slide':  "INTEGER DEFAULT 1",
                'last_seen':   "DATETIME",
            }
            with db.engine.connect() as conn:
                for col, typ in wanted_p.items():
                    if col not in existing_p:
                        try:
                            conn.execute(text(
                                f"ALTER TABLE progress ADD COLUMN {col} {typ}"))
                            conn.commit()
                            print(f"[migrate] added column progress.{col}")
                        except Exception as e:
                            print(f"[migrate] failed to add progress.{col}: {e}")

        # ===== AMHARIC: ensure language column exists on instructor_script =====
        if 'instructor_script' in insp.get_table_names():
            existing_s = {c['name'] for c in insp.get_columns('instructor_script')}
            if 'language' not in existing_s:
                try:
                    with db.engine.connect() as conn:
                        conn.execute(text(
                            "ALTER TABLE instructor_script ADD COLUMN "
                            "language VARCHAR(5) DEFAULT 'en' NOT NULL"))
                        conn.commit()
                    print("[migrate] added column instructor_script.language")
                except Exception as e:
                    print(f"[migrate] failed to add instructor_script.language: {e}")
        # ======================================================================
    except Exception as e:
        print(f"[migrate] skipped: {e}")


def init_db():
    with app.app_context():
        db.create_all()
        _light_migrate()
        if not Admin.query.filter_by(username='admin').first():
            db.session.add(Admin(
                username='admin',
                password_hash=generate_password_hash('admin123'),
            ))
            db.session.commit()
            print("Default admin created: admin / admin123")


with app.app_context():
    try:
        db.create_all()
        _light_migrate()
        if not Admin.query.filter_by(username='admin').first():
            db.session.add(Admin(
                username='admin',
                password_hash=generate_password_hash('admin123'),
            ))
            db.session.commit()
            print("Default admin created: admin / admin123")
    except Exception as _e:
        print(f"[bootstrap] init skipped: {_e}")


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)