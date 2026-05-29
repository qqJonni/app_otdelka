import os
import sqlite3
import secrets
import io
from datetime import datetime, date, timedelta
from functools import wraps

from flask import (Flask, render_template, redirect, url_for, request,
                   flash, session, g, send_file, abort)
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')

AVATAR_FOLDER = os.path.join(app.root_path, 'static', 'avatars')
ALLOWED_AVATAR_EXT = {'jpg', 'jpeg', 'png'}
MAX_AVATAR_SIZE = 2 * 1024 * 1024  # 2 МБ

login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Пожалуйста, войдите в систему.'

DATABASE = 'otdelka.db'

DEFAULT_STAGES = [
    'Стяжка пола',
    'Штукатурка стен',
    'Шпатлёвка стен и откосов',
    'Грунтовка',
    'Укладка напольного покрытия',
    'Поклейка обоев / покраска стен',
    'Установка межкомнатных дверей',
    'Установка плинтусов',
    'Сантехническая разводка и установка приборов',
    'Электромонтажные работы',
    'Финишная уборка',
]


# ─── Database ────────────────────────────────────────────────────────────────

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
        db.execute('PRAGMA journal_mode=WAL')
        db.execute('PRAGMA foreign_keys=ON')
    return db


@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()


def query_db(sql, args=(), one=False):
    cur = get_db().execute(sql, args)
    rv = cur.fetchall()
    return (rv[0] if rv else None) if one else rv


def execute_db(sql, args=()):
    db = get_db()
    cur = db.execute(sql, args)
    db.commit()
    return cur.lastrowid


def get_setting(key, default=None):
    row = query_db('SELECT value FROM settings WHERE key=?', [key], one=True)
    return row['value'] if row else default


def set_setting(key, value):
    existing = query_db('SELECT id FROM settings WHERE key=?', [key], one=True)
    if existing:
        execute_db('UPDATE settings SET value=? WHERE key=?', [str(value), key])
    else:
        execute_db('INSERT INTO settings (key,value) VALUES (?,?)', [key, str(value)])


def run_migrations(db):
    """Apply schema migrations safely on existing databases."""
    # stages_template.work_description
    cols = [r[1] for r in db.execute("PRAGMA table_info(stages_template)").fetchall()]
    if 'work_description' not in cols:
        db.execute("ALTER TABLE stages_template ADD COLUMN work_description TEXT")

    # apartment_stages.work_description
    cols = [r[1] for r in db.execute("PRAGMA table_info(apartment_stages)").fetchall()]
    if 'work_description' not in cols:
        db.execute("ALTER TABLE apartment_stages ADD COLUMN work_description TEXT")

    # users.avatar
    cols = [r[1] for r in db.execute("PRAGMA table_info(users)").fetchall()]
    if 'avatar' not in cols:
        db.execute("ALTER TABLE users ADD COLUMN avatar TEXT")

    # apartments.created_at
    cols = [r[1] for r in db.execute("PRAGMA table_info(apartments)").fetchall()]
    if 'created_at' not in cols:
        db.execute("ALTER TABLE apartments ADD COLUMN created_at TEXT")
        # Заполняем существующие строки текущим временем
        db.execute("UPDATE apartments SET created_at=datetime('now') WHERE created_at IS NULL")

    # settings table (kept for future use)
    db.execute('''CREATE TABLE IF NOT EXISTS settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        key TEXT UNIQUE NOT NULL,
        value TEXT NOT NULL
    )''')

    db.commit()

    # Создать папку для аватаров
    os.makedirs(AVATAR_FOLDER, exist_ok=True)


def init_db():
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    db.execute('PRAGMA foreign_keys=ON')

    db.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','foreman','guest')),
            full_name TEXT NOT NULL DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS buildings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT
        );

        CREATE TABLE IF NOT EXISTS apartments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            building_id INTEGER NOT NULL REFERENCES buildings(id) ON DELETE CASCADE,
            number TEXT NOT NULL,
            plan_start_date TEXT,
            plan_end_date TEXT,
            foreman_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
            created_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS stages_template (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            default_order INTEGER NOT NULL DEFAULT 0,
            work_description TEXT
        );

        CREATE TABLE IF NOT EXISTS apartment_stages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            apartment_id INTEGER NOT NULL REFERENCES apartments(id) ON DELETE CASCADE,
            stage_name TEXT NOT NULL,
            order_num INTEGER NOT NULL DEFAULT 0,
            volume_sqm REAL,
            deadline TEXT,
            status TEXT NOT NULL DEFAULT 'not_started'
                CHECK(status IN ('not_started','in_progress','done','overdue')),
            started_at TEXT,
            completed_at TEXT,
            work_description TEXT
        );

        CREATE TABLE IF NOT EXISTS guest_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            building_id INTEGER NOT NULL REFERENCES buildings(id) ON DELETE CASCADE,
            token TEXT UNIQUE NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            key TEXT UNIQUE NOT NULL,
            value TEXT NOT NULL
        );
    ''')
    db.commit()

    # Run migrations for existing databases (adds columns if missing)
    run_migrations(db)

    # Seed only if empty
    count = db.execute('SELECT COUNT(*) FROM users').fetchone()[0]
    if count > 0:
        db.close()
        return

    # Users
    users = [
        ('admin', generate_password_hash('admin123'), 'admin', 'Администратор'),
        ('foreman1', generate_password_hash('123456'), 'foreman', 'Иванов Иван'),
        ('foreman2', generate_password_hash('123456'), 'foreman', 'Петров Пётр'),
    ]
    db.executemany(
        'INSERT INTO users (username,password_hash,role,full_name) VALUES (?,?,?,?)',
        users
    )

    # Stages template with sample descriptions
    stage_descriptions = {
        'Стяжка пола': 'Выравнивание основания, заливка стяжки М300. Толщина 50–70 мм. Выдержка 28 дней до полного набора прочности.',
        'Штукатурка стен': 'Машинная или ручная штукатурка по маякам. Слой 10–20 мм. Материал — гипсовая или цементно-песчаная смесь.',
        'Шпатлёвка стен и откосов': 'Финишная шпатлёвка в 2 слоя. Шлифовка между слоями. Обязательная обработка откосов и углов.',
        'Грунтовка': 'Грунтование всех поверхностей перед финишными работами. Глубокое проникновение. Дать высохнуть 4–6 часов.',
        'Укладка напольного покрытия': 'Укладка плитки, ламината или иного покрытия согласно проекту. Соблюдать зазоры и уровень.',
        'Поклейка обоев / покраска стен': 'Поклейка обоев встык или покраска в 2 слоя. Соблюдать направление рисунка и перекрытие швов.',
        'Установка межкомнатных дверей': 'Монтаж коробок, навешивание полотен, установка наличников и фурнитуры.',
        'Установка плинтусов': 'Монтаж напольных и потолочных плинтусов. Подрезка углов. Фиксация на клей или дюбели.',
        'Сантехническая разводка и установка приборов': 'Подключение смесителей, унитаза, ванны/душевой кабины. Проверка на протечки.',
        'Электромонтажные работы': 'Установка розеток, выключателей, светильников. Подключение щитка. Прозвонка цепей.',
        'Финишная уборка': 'Влажная уборка всех поверхностей. Удаление строительного мусора. Мытьё окон.',
    }
    for i, name in enumerate(DEFAULT_STAGES):
        db.execute(
            'INSERT INTO stages_template (name,default_order,work_description) VALUES (?,?,?)',
            (name, i + 1, stage_descriptions.get(name, ''))
        )

    # Building
    db.execute("INSERT INTO buildings (name,address) VALUES ('Солнечный','ул. Солнечная, 1')")
    building_id = db.execute('SELECT last_insert_rowid()').fetchone()[0]

    foreman1_id = db.execute("SELECT id FROM users WHERE username='foreman1'").fetchone()[0]
    foreman2_id = db.execute("SELECT id FROM users WHERE username='foreman2'").fetchone()[0]

    today = date.today()
    stage_days = 3  # default
    apts = [
        (building_id, '1', str(today), str(today + timedelta(days=60)), foreman1_id),
        (building_id, '2', str(today + timedelta(days=5)), str(today + timedelta(days=75)), foreman1_id),
        (building_id, '3', str(today + timedelta(days=10)), str(today + timedelta(days=80)), foreman2_id),
    ]
    for apt_row in apts:
        db.execute(
            'INSERT INTO apartments (building_id,number,plan_start_date,plan_end_date,foreman_id) VALUES (?,?,?,?,?)',
            apt_row
        )
        apt_id = db.execute('SELECT last_insert_rowid()').fetchone()[0]
        plan_start_str = apt_row[2]
        plan_start_dt = date.fromisoformat(plan_start_str) if plan_start_str else None

        for i, stage_name in enumerate(DEFAULT_STAGES):
            order = i + 1
            if plan_start_dt:
                deadline = str(plan_start_dt + timedelta(days=order * stage_days))
            else:
                deadline = None
            desc = stage_descriptions.get(stage_name, '')
            db.execute(
                '''INSERT INTO apartment_stages
                   (apartment_id,stage_name,order_num,volume_sqm,deadline,status,work_description)
                   VALUES (?,?,?,?,?,?,?)''',
                (apt_id, stage_name, order, 30.0, deadline, 'not_started', desc)
            )

    db.commit()
    db.close()


# ─── User model ──────────────────────────────────────────────────────────────

class User(UserMixin):
    def __init__(self, row):
        self.id = row['id']
        self.username = row['username']
        self.role = row['role']
        self.full_name = row['full_name']
        self.avatar = row['avatar'] if 'avatar' in row.keys() else None

    @property
    def initials(self):
        name = self.full_name or self.username
        parts = name.strip().split()
        if len(parts) >= 2:
            return (parts[0][0] + parts[1][0]).upper()
        return name[:2].upper() if len(name) >= 2 else name[:1].upper()


@login_manager.user_loader
def load_user(user_id):
    row = query_db('SELECT * FROM users WHERE id=?', [user_id], one=True)
    return User(row) if row else None


# ─── Helpers ─────────────────────────────────────────────────────────────────

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != 'admin':
            abort(403)
        return f(*args, **kwargs)
    return decorated


def foreman_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role not in ('admin', 'foreman'):
            abort(403)
        return f(*args, **kwargs)
    return decorated


def refresh_overdue():
    """Mark stages as overdue only when deadline is set and passed and not done."""
    today = str(date.today())
    execute_db(
        """UPDATE apartment_stages SET status='overdue'
           WHERE status='not_started'
             AND deadline IS NOT NULL AND deadline != ''
             AND deadline < ?""",
        [today]
    )
    execute_db(
        """UPDATE apartment_stages SET status='overdue'
           WHERE status='in_progress'
             AND deadline IS NOT NULL AND deadline != ''
             AND deadline < ?""",
        [today]
    )


STATUS_LABEL = {
    'not_started': ('Не начат', 'secondary'),
    'in_progress': ('В работе', 'warning'),
    'done': ('Завершён', 'success'),
    'overdue': ('Просрочен', 'danger'),
}

app.jinja_env.globals['STATUS_LABEL'] = STATUS_LABEL
app.jinja_env.globals['now'] = datetime.now


# ─── Auth routes ─────────────────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return _redirect_by_role(current_user.role)

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        row = query_db('SELECT * FROM users WHERE username=?', [username], one=True)
        if row and check_password_hash(row['password_hash'], password):
            user = User(row)
            login_user(user)
            return _redirect_by_role(user.role)
        flash('Неверный логин или пароль.', 'danger')

    return render_template('login.html')


def _redirect_by_role(role):
    if role == 'admin':
        return redirect(url_for('admin_dashboard'))
    if role == 'foreman':
        return redirect(url_for('foreman_dashboard'))
    return redirect(url_for('login'))


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ─── Guest token route ────────────────────────────────────────────────────────

@app.route('/guest/<token>')
def guest_view(token):
    gt = query_db('SELECT * FROM guest_tokens WHERE token=?', [token], one=True)
    if not gt:
        abort(404)
    refresh_overdue()
    building = query_db('SELECT * FROM buildings WHERE id=?', [gt['building_id']], one=True)
    apartments = query_db(
        'SELECT * FROM apartments WHERE building_id=? ORDER BY number', [gt['building_id']]
    )
    apt_data = []
    for apt in apartments:
        stages = query_db(
            'SELECT * FROM apartment_stages WHERE apartment_id=? ORDER BY order_num',
            [apt['id']]
        )
        total = len(stages)
        done = sum(1 for s in stages if s['status'] == 'done')
        pct = int(done / total * 100) if total else 0
        apt_data.append({'apt': apt, 'stages': stages, 'pct': pct})

    return render_template('guest/view.html',
                           building=building, apt_data=apt_data, token=token)


# ─── Admin: Dashboard ────────────────────────────────────────────────────────

@app.route('/admin/')
@login_required
@admin_required
def admin_dashboard():
    refresh_overdue()
    total = query_db('SELECT COUNT(*) as c FROM apartments', one=True)['c']
    in_progress = query_db(
        "SELECT COUNT(DISTINCT apartment_id) as c FROM apartment_stages WHERE status='in_progress'",
        one=True
    )['c']
    done_apts = query_db(
        """SELECT COUNT(*) as c FROM apartments a
           WHERE NOT EXISTS (
               SELECT 1 FROM apartment_stages s
               WHERE s.apartment_id=a.id AND s.status != 'done'
           ) AND EXISTS (SELECT 1 FROM apartment_stages s WHERE s.apartment_id=a.id)""",
        one=True
    )['c']
    overdue = query_db(
        "SELECT COUNT(DISTINCT apartment_id) as c FROM apartment_stages WHERE status='overdue'",
        one=True
    )['c']
    buildings = query_db('SELECT * FROM buildings ORDER BY name')
    return render_template('admin/dashboard.html',
                           total=total, in_progress=in_progress,
                           done_apts=done_apts, overdue=overdue,
                           buildings=buildings)


# ─── Admin: Buildings ────────────────────────────────────────────────────────

@app.route('/admin/buildings')
@login_required
@admin_required
def admin_buildings():
    buildings = query_db('SELECT * FROM buildings ORDER BY name')
    return render_template('admin/buildings.html', buildings=buildings)


@app.route('/admin/buildings/add', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_building_add():
    if request.method == 'POST':
        name = request.form['name'].strip()
        address = request.form.get('address', '').strip()
        if name:
            execute_db('INSERT INTO buildings (name,address) VALUES (?,?)', [name, address])
            flash('ЖК создан.', 'success')
            return redirect(url_for('admin_buildings'))
    return render_template('admin/building_form.html', building=None)


@app.route('/admin/buildings/<int:bid>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_building_edit(bid):
    building = query_db('SELECT * FROM buildings WHERE id=?', [bid], one=True)
    if not building:
        abort(404)
    if request.method == 'POST':
        name = request.form['name'].strip()
        address = request.form.get('address', '').strip()
        execute_db('UPDATE buildings SET name=?,address=? WHERE id=?', [name, address, bid])
        flash('ЖК обновлён.', 'success')
        return redirect(url_for('admin_buildings'))
    return render_template('admin/building_form.html', building=building)


@app.route('/admin/buildings/<int:bid>/delete', methods=['POST'])
@login_required
@admin_required
def admin_building_delete(bid):
    execute_db('DELETE FROM buildings WHERE id=?', [bid])
    flash('ЖК удалён.', 'success')
    return redirect(url_for('admin_buildings'))


# ─── Admin: Apartments ───────────────────────────────────────────────────────

@app.route('/admin/buildings/<int:bid>/apartments')
@login_required
@admin_required
def admin_apartments(bid):
    building = query_db('SELECT * FROM buildings WHERE id=?', [bid], one=True)
    if not building:
        abort(404)
    apartments = query_db(
        '''SELECT a.*, u.full_name as foreman_name
           FROM apartments a LEFT JOIN users u ON a.foreman_id=u.id
           WHERE a.building_id=? ORDER BY a.number''',
        [bid]
    )
    apt_stats = {}
    for apt in apartments:
        stages = query_db(
            'SELECT status FROM apartment_stages WHERE apartment_id=?', [apt['id']]
        )
        total = len(stages)
        done = sum(1 for s in stages if s['status'] == 'done')
        apt_stats[apt['id']] = {'total': total, 'done': done,
                                 'pct': int(done / total * 100) if total else 0}
    return render_template('admin/apartments.html',
                           building=building, apartments=apartments, apt_stats=apt_stats)


@app.route('/admin/buildings/<int:bid>/apartments/add', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_apartment_add(bid):
    building = query_db('SELECT * FROM buildings WHERE id=?', [bid], one=True)
    foremen  = query_db("SELECT * FROM users WHERE role='foreman' ORDER BY full_name")

    if request.method == 'POST':
        number     = request.form['number'].strip()
        plan_start = request.form.get('plan_start_date') or None
        plan_end   = request.form.get('plan_end_date') or None
        foreman_id = request.form.get('foreman_id') or None

        apt_id = execute_db(
            '''INSERT INTO apartments
               (building_id, number, plan_start_date, plan_end_date, foreman_id, created_at)
               VALUES (?,?,?,?,?,datetime('now'))''',
            [bid, number, plan_start, plan_end, foreman_id]
        )

        templates = query_db('SELECT * FROM stages_template ORDER BY default_order')
        added = 0
        for t in templates:
            if not request.form.get(f'stage_sel_{t["id"]}'):
                continue  # чекбокс не отмечен

            # Дедлайн — только из поля формы (NULL если не заполнено)
            deadline_raw = request.form.get(f'stage_deadline_{t["id"]}', '').strip()
            deadline = deadline_raw if deadline_raw else None

            volume_raw = request.form.get(f'stage_volume_{t["id"]}', '').strip()
            try:
                volume = float(volume_raw) if volume_raw else None
            except ValueError:
                volume = None

            execute_db(
                '''INSERT INTO apartment_stages
                   (apartment_id, stage_name, order_num, volume_sqm, deadline, status, work_description)
                   VALUES (?,?,?,?,?,?,?)''',
                [apt_id, t['name'], t['default_order'], volume, deadline,
                 'not_started', t['work_description'] or '']
            )
            added += 1

        flash(f'Квартира создана, добавлено {added} этапов.', 'success')
        return redirect(url_for('admin_apartments', bid=bid))

    templates = query_db('SELECT * FROM stages_template ORDER BY default_order')
    return render_template('admin/apartment_form.html',
                           building=building, apartment=None, foremen=foremen,
                           templates=templates)


@app.route('/admin/apartments/<int:apt_id>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_apartment_edit(apt_id):
    apt = query_db('SELECT * FROM apartments WHERE id=?', [apt_id], one=True)
    if not apt:
        abort(404)
    building = query_db('SELECT * FROM buildings WHERE id=?', [apt['building_id']], one=True)
    foremen = query_db("SELECT * FROM users WHERE role='foreman' ORDER BY full_name")
    if request.method == 'POST':
        number = request.form['number'].strip()
        plan_start = request.form.get('plan_start_date') or None
        plan_end = request.form.get('plan_end_date') or None
        foreman_id = request.form.get('foreman_id') or None
        execute_db(
            'UPDATE apartments SET number=?,plan_start_date=?,plan_end_date=?,foreman_id=? WHERE id=?',
            [number, plan_start, plan_end, foreman_id, apt_id]
        )
        flash('Квартира обновлена. Дедлайны этапов не изменены.', 'success')
        return redirect(url_for('admin_apartments', bid=apt['building_id']))
    return render_template('admin/apartment_form.html',
                           building=building, apartment=apt, foremen=foremen)


@app.route('/admin/apartments/<int:apt_id>/delete', methods=['POST'])
@login_required
@admin_required
def admin_apartment_delete(apt_id):
    apt = query_db('SELECT * FROM apartments WHERE id=?', [apt_id], one=True)
    if not apt:
        abort(404)
    bid = apt['building_id']
    execute_db('DELETE FROM apartments WHERE id=?', [apt_id])
    flash('Квартира удалена.', 'success')
    return redirect(url_for('admin_apartments', bid=bid))


# ─── Admin: Stages of apartment ──────────────────────────────────────────────

@app.route('/admin/apartments/<int:apt_id>/stages')
@login_required
@admin_required
def admin_stages(apt_id):
    refresh_overdue()
    apt = query_db('SELECT * FROM apartments WHERE id=?', [apt_id], one=True)
    if not apt:
        abort(404)
    building = query_db('SELECT * FROM buildings WHERE id=?', [apt['building_id']], one=True)
    stages = query_db(
        'SELECT * FROM apartment_stages WHERE apartment_id=? ORDER BY order_num', [apt_id]
    )
    # Этапы справочника, которых ещё нет в квартире — для модалки добавления
    existing_names = {s['stage_name'] for s in stages}
    available = query_db('SELECT * FROM stages_template ORDER BY default_order')
    available = [t for t in available if t['name'] not in existing_names]
    return render_template('admin/stages.html', apt=apt, building=building,
                           stages=stages, available=available)


@app.route('/admin/apartments/<int:apt_id>/stages/add', methods=['POST'])
@login_required
@admin_required
def admin_stage_add(apt_id):
    """Добавить один этап в существующую квартиру (через модальное окно)."""
    apt = query_db('SELECT * FROM apartments WHERE id=?', [apt_id], one=True)
    if not apt:
        abort(404)
    stage_name = request.form.get('stage_name', '').strip()
    deadline   = request.form.get('deadline') or None
    volume     = request.form.get('volume_sqm') or None
    if volume:
        try:
            volume = float(volume)
        except ValueError:
            volume = None

    if stage_name:
        max_ord = query_db(
            'SELECT COALESCE(MAX(order_num), 0) as m FROM apartment_stages WHERE apartment_id=?',
            [apt_id], one=True
        )['m']
        # Берём work_description из справочника, если есть
        tpl = query_db('SELECT work_description FROM stages_template WHERE name=?',
                       [stage_name], one=True)
        work_desc = tpl['work_description'] if tpl and tpl['work_description'] else ''
        execute_db(
            '''INSERT INTO apartment_stages
               (apartment_id, stage_name, order_num, volume_sqm, deadline, status, work_description)
               VALUES (?,?,?,?,?,?,?)''',
            [apt_id, stage_name, max_ord + 1, volume, deadline, 'not_started', work_desc]
        )
        flash(f'Этап «{stage_name}» добавлен в квартиру.', 'success')
    return redirect(url_for('admin_stages', apt_id=apt_id))


@app.route('/admin/stages/<int:stage_id>/delete', methods=['POST'])
@login_required
@admin_required
def admin_stage_delete(stage_id):
    """Удалить этап из квартиры."""
    stage = query_db('SELECT * FROM apartment_stages WHERE id=?', [stage_id], one=True)
    if not stage:
        abort(404)
    apt_id = stage['apartment_id']
    execute_db('DELETE FROM apartment_stages WHERE id=?', [stage_id])
    flash(f'Этап «{stage["stage_name"]}» удалён из квартиры.', 'success')
    return redirect(url_for('admin_stages', apt_id=apt_id))


@app.route('/admin/stages/<int:stage_id>/move/<direction>', methods=['POST'])
@login_required
@admin_required
def admin_stage_move(stage_id, direction):
    """Переместить этап вверх (up) или вниз (down), обменяв order_num с соседом."""
    stage = query_db('SELECT * FROM apartment_stages WHERE id=?', [stage_id], one=True)
    if not stage or direction not in ('up', 'down'):
        abort(404)
    apt_id = stage['apartment_id']
    cur_order = stage['order_num']

    if direction == 'up':
        neighbor = query_db(
            '''SELECT * FROM apartment_stages
               WHERE apartment_id=? AND order_num < ?
               ORDER BY order_num DESC LIMIT 1''',
            [apt_id, cur_order], one=True
        )
    else:
        neighbor = query_db(
            '''SELECT * FROM apartment_stages
               WHERE apartment_id=? AND order_num > ?
               ORDER BY order_num ASC LIMIT 1''',
            [apt_id, cur_order], one=True
        )

    if neighbor:
        execute_db('UPDATE apartment_stages SET order_num=? WHERE id=?',
                   [neighbor['order_num'], stage_id])
        execute_db('UPDATE apartment_stages SET order_num=? WHERE id=?',
                   [cur_order, neighbor['id']])
    return redirect(url_for('admin_stages', apt_id=apt_id))


@app.route('/admin/stages/<int:stage_id>/edit', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_stage_edit(stage_id):
    stage = query_db('SELECT * FROM apartment_stages WHERE id=?', [stage_id], one=True)
    if not stage:
        abort(404)
    apt = query_db('SELECT * FROM apartments WHERE id=?', [stage['apartment_id']], one=True)
    building = query_db('SELECT * FROM buildings WHERE id=?', [apt['building_id']], one=True)
    if request.method == 'POST':
        volume = request.form.get('volume_sqm') or None
        deadline = request.form.get('deadline') or None
        work_desc = request.form.get('work_description', '').strip()
        if volume:
            volume = float(volume)
        execute_db(
            'UPDATE apartment_stages SET volume_sqm=?,deadline=?,work_description=? WHERE id=?',
            [volume, deadline, work_desc, stage_id]
        )
        flash('Этап обновлён.', 'success')
        return redirect(url_for('admin_stages', apt_id=stage['apartment_id']))
    return render_template('admin/stage_edit.html', stage=stage, apt=apt, building=building)


# ─── Admin: Stages template ──────────────────────────────────────────────────

@app.route('/admin/stages-template')
@login_required
@admin_required
def admin_stages_template():
    templates = query_db('SELECT * FROM stages_template ORDER BY default_order')
    return render_template('admin/stages_template.html', templates=templates)


@app.route('/admin/stages-template/add', methods=['POST'])
@login_required
@admin_required
def admin_stages_template_add():
    name = request.form.get('name', '').strip()
    if name:
        max_order = query_db('SELECT MAX(default_order) as m FROM stages_template', one=True)['m'] or 0
        execute_db('INSERT INTO stages_template (name,default_order,work_description) VALUES (?,?,?)',
                   [name, max_order + 1, ''])
        flash('Этап добавлен в справочник.', 'success')
    return redirect(url_for('admin_stages_template'))


@app.route('/admin/stages-template/<int:tid>/delete', methods=['POST'])
@login_required
@admin_required
def admin_stages_template_delete(tid):
    execute_db('DELETE FROM stages_template WHERE id=?', [tid])
    flash('Этап удалён из справочника.', 'success')
    return redirect(url_for('admin_stages_template'))


@app.route('/admin/stages-template/<int:tid>/update', methods=['POST'])
@login_required
@admin_required
def admin_stages_template_update(tid):
    name = request.form.get('name', '').strip()
    work_desc = request.form.get('work_description', '').strip()
    if name:
        execute_db(
            'UPDATE stages_template SET name=?,work_description=? WHERE id=?',
            [name, work_desc, tid]
        )
        flash('Этап обновлён.', 'success')
    return redirect(url_for('admin_stages_template'))


# ─── Profile (все роли) ───────────────────────────────────────────────────────

@app.route('/profile')
@login_required
def profile():
    return render_template('profile.html')


@app.route('/profile/avatar', methods=['POST'])
@login_required
def profile_avatar():
    f = request.files.get('avatar')
    if not f or f.filename == '':
        flash('Выберите файл.', 'warning')
        return redirect(url_for('profile'))

    ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else ''
    if ext not in ALLOWED_AVATAR_EXT:
        flash('Допустимые форматы: JPG, JPEG, PNG.', 'danger')
        return redirect(url_for('profile'))

    data = f.read()
    if len(data) > MAX_AVATAR_SIZE:
        flash('Файл слишком большой (максимум 2 МБ).', 'danger')
        return redirect(url_for('profile'))

    os.makedirs(AVATAR_FOLDER, exist_ok=True)

    # Удаляем старый аватар (любое расширение)
    for old_ext in ALLOWED_AVATAR_EXT:
        old_path = os.path.join(AVATAR_FOLDER, f'user_{current_user.id}.{old_ext}')
        if os.path.exists(old_path):
            os.remove(old_path)

    # Сохраняем новый
    filename = f'user_{current_user.id}.{ext}'
    filepath = os.path.join(AVATAR_FOLDER, filename)
    with open(filepath, 'wb') as out:
        out.write(data)

    execute_db('UPDATE users SET avatar=? WHERE id=?', [filename, current_user.id])
    # Обновляем объект в сессии
    current_user.avatar = filename
    flash('Аватар обновлён.', 'success')
    return redirect(url_for('profile'))


@app.route('/profile/login', methods=['POST'])
@login_required
def profile_login():
    if current_user.role == 'guest':
        abort(403)
    new_login = request.form.get('new_username', '').strip()
    if len(new_login) < 3:
        flash('Логин должен содержать минимум 3 символа.', 'danger')
        return redirect(url_for('profile'))
    existing = query_db('SELECT id FROM users WHERE username=? AND id!=?',
                        [new_login, current_user.id], one=True)
    if existing:
        flash('Этот логин уже используется другим пользователем.', 'danger')
        return redirect(url_for('profile'))
    execute_db('UPDATE users SET username=? WHERE id=?', [new_login, current_user.id])
    current_user.username = new_login
    flash('Логин успешно изменён.', 'success')
    return redirect(url_for('profile'))


@app.route('/profile/password', methods=['POST'])
@login_required
def profile_password():
    current_pw  = request.form.get('current_password', '')
    new_pw      = request.form.get('new_password', '')
    confirm_pw  = request.form.get('confirm_password', '')

    row = query_db('SELECT password_hash FROM users WHERE id=?', [current_user.id], one=True)
    if not check_password_hash(row['password_hash'], current_pw):
        flash('Текущий пароль введён неверно.', 'danger')
        return redirect(url_for('profile'))
    if len(new_pw) < 4:
        flash('Новый пароль должен содержать минимум 4 символа.', 'danger')
        return redirect(url_for('profile'))
    if new_pw != confirm_pw:
        flash('Новый пароль и подтверждение не совпадают.', 'danger')
        return redirect(url_for('profile'))

    execute_db('UPDATE users SET password_hash=? WHERE id=?',
               [generate_password_hash(new_pw), current_user.id])
    flash('Пароль успешно изменён.', 'success')
    return redirect(url_for('profile'))


# ─── Admin: Guest tokens ──────────────────────────────────────────────────────

@app.route('/admin/guest-tokens')
@login_required
@admin_required
def admin_guest_tokens():
    tokens = query_db(
        '''SELECT gt.*, b.name as building_name
           FROM guest_tokens gt JOIN buildings b ON gt.building_id=b.id
           ORDER BY gt.created_at DESC'''
    )
    buildings = query_db('SELECT * FROM buildings ORDER BY name')
    return render_template('admin/guest_tokens.html', tokens=tokens, buildings=buildings)


@app.route('/admin/guest-tokens/create', methods=['POST'])
@login_required
@admin_required
def admin_guest_token_create():
    building_id = request.form.get('building_id')
    if building_id:
        token = secrets.token_urlsafe(24)
        execute_db(
            'INSERT INTO guest_tokens (building_id,token,created_at) VALUES (?,?,?)',
            [building_id, token, datetime.now().isoformat()]
        )
        flash(f'Ссылка создана: /guest/{token}', 'success')
    return redirect(url_for('admin_guest_tokens'))


@app.route('/admin/guest-tokens/<int:tid>/delete', methods=['POST'])
@login_required
@admin_required
def admin_guest_token_delete(tid):
    execute_db('DELETE FROM guest_tokens WHERE id=?', [tid])
    flash('Гостевой доступ удалён.', 'success')
    return redirect(url_for('admin_guest_tokens'))


# ─── Admin: Export ────────────────────────────────────────────────────────────

@app.route('/admin/export', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_export():
    buildings = query_db('SELECT * FROM buildings ORDER BY name')
    if request.method == 'POST':
        building_id = request.form.get('building_id')
        date_from = request.form.get('date_from')
        date_to = request.form.get('date_to')

        rows = query_db(
            '''SELECT a.number, s.stage_name, s.volume_sqm, s.completed_at
               FROM apartment_stages s
               JOIN apartments a ON s.apartment_id=a.id
               WHERE a.building_id=? AND s.status='done'
               AND s.completed_at IS NOT NULL
               AND substr(s.completed_at,1,10) >= ?
               AND substr(s.completed_at,1,10) <= ?
               ORDER BY a.number, s.order_num''',
            [building_id, date_from, date_to]
        )

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Выполненные работы'

        header_fill = PatternFill(fill_type='solid', fgColor='4472C4')
        header_font_white = Font(bold=True, color='FFFFFF')

        headers = ['Номер квартиры', 'Этап', 'Объём (кв.м)', 'Дата завершения']
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')

        for row_idx, row in enumerate(rows, 2):
            ws.cell(row=row_idx, column=1, value=row['number'])
            ws.cell(row=row_idx, column=2, value=row['stage_name'])
            ws.cell(row=row_idx, column=3, value=row['volume_sqm'])
            completed = row['completed_at']
            if completed:
                completed = completed[:10]
            ws.cell(row=row_idx, column=4, value=completed)

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 4

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, download_name='выполненные_работы.xlsx',
                         as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    return render_template('admin/export.html', buildings=buildings)


# ─── Admin: Users ────────────────────────────────────────────────────────────

@app.route('/admin/users')
@login_required
@admin_required
def admin_users():
    users = query_db("SELECT * FROM users ORDER BY role, username")
    return render_template('admin/users.html', users=users)


@app.route('/admin/users/add', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_user_add():
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password'].strip()
        role = request.form['role']
        full_name = request.form.get('full_name', '').strip()
        if username and password and role:
            try:
                execute_db(
                    'INSERT INTO users (username,password_hash,role,full_name) VALUES (?,?,?,?)',
                    [username, generate_password_hash(password), role, full_name]
                )
                flash('Пользователь создан.', 'success')
                return redirect(url_for('admin_users'))
            except Exception:
                flash('Пользователь с таким логином уже существует.', 'danger')
    return render_template('admin/user_form.html')


@app.route('/admin/users/<int:uid>/delete', methods=['POST'])
@login_required
@admin_required
def admin_user_delete(uid):
    if uid == current_user.id:
        flash('Нельзя удалить собственную учётную запись.', 'danger')
    else:
        execute_db('DELETE FROM users WHERE id=?', [uid])
        flash('Пользователь удалён.', 'success')
    return redirect(url_for('admin_users'))


# ─── Foreman routes ───────────────────────────────────────────────────────────

@app.route('/foreman/')
@login_required
@foreman_required
def foreman_dashboard():
    refresh_overdue()
    if current_user.role == 'admin':
        return redirect(url_for('admin_dashboard'))

    # Сортировка: последняя созданная — первой (новые вверху)
    apartments = query_db(
        '''SELECT a.*, b.name as building_name, b.address as building_address
           FROM apartments a JOIN buildings b ON a.building_id=b.id
           WHERE a.foreman_id=?
           ORDER BY COALESCE(a.created_at, '') DESC, a.id DESC''',
        [current_user.id]
    )

    apt_data = []
    for apt in apartments:
        stages = query_db(
            'SELECT * FROM apartment_stages WHERE apartment_id=? ORDER BY order_num',
            [apt['id']]
        )
        total = len(stages)
        done  = sum(1 for s in stages if s['status'] == 'done')
        active = any(s['status'] in ('in_progress', 'overdue') for s in stages)
        pct   = int(done / total * 100) if total else 0
        apt_data.append({
            'apt':    apt,
            'stages': stages,
            'total':  total,
            'done':   done,
            'active': active,
            'pct':    pct,
        })

    # Сводная статистика
    stats = {
        'total':       len(apt_data),
        'in_progress': sum(1 for d in apt_data if d['active'] and d['done'] < d['total']),
        'completed':   sum(1 for d in apt_data if d['total'] > 0 and d['done'] == d['total']),
    }

    return render_template('foreman/dashboard.html', apt_data=apt_data, stats=stats)


@app.route('/foreman/stages/<int:stage_id>/action', methods=['POST'])
@login_required
@foreman_required
def foreman_stage_action(stage_id):
    stage = query_db('SELECT * FROM apartment_stages WHERE id=?', [stage_id], one=True)
    if not stage:
        abort(404)
    apt = query_db('SELECT * FROM apartments WHERE id=?', [stage['apartment_id']], one=True)
    if current_user.role != 'admin' and apt['foreman_id'] != current_user.id:
        abort(403)

    now = datetime.now().isoformat()
    if stage['status'] == 'not_started':
        execute_db(
            "UPDATE apartment_stages SET status='in_progress', started_at=? WHERE id=?",
            [now, stage_id]
        )
    elif stage['status'] in ('in_progress', 'overdue'):
        execute_db(
            "UPDATE apartment_stages SET status='done', completed_at=? WHERE id=?",
            [now, stage_id]
        )
    return redirect(url_for('foreman_dashboard'))


# ─── Error handlers ───────────────────────────────────────────────────────────

@app.errorhandler(403)
def forbidden(e):
    return render_template('errors/403.html'), 403


@app.errorhandler(404)
def not_found(e):
    return render_template('errors/404.html'), 404


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    init_db()
    app.run(debug=True, host='0.0.0.0', port=8080, use_reloader=False)
