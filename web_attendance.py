import flet as ft
import sqlite3
import hashlib
from datetime import date
import os
import json
import base64
import io

# --- IMPORTACIÓN DE LIBRERÍAS EXTERNAS CON MANEJO DE ERRORES ---
try:
    import pandas as pd
except ImportError:
    pd = None
    print("⚠️ ADVERTENCIA: 'pandas' no está instalado.")

try:
    import openpyxl
except ImportError:
    print("⚠️ ADVERTENCIA: 'openpyxl' no está instalado.")

try:
    import xlsxwriter
except ImportError:
    print("⚠️ ADVERTENCIA: 'xlsxwriter' no está instalado.")

try:
    import firebase_admin
    from firebase_admin import credentials, firestore
except ImportError:
    firebase_admin = None
    firestore = None
    print("⚠️ ADVERTENCIA: 'firebase-admin' no está instalado. Se usará SQLite localmente.")

# ======================================================================
# 1. LÓGICA DE BASE DE DATOS (Híbrido: Firestore Cloud / SQLite Local)
# ======================================================================

DB_NAME = 'asistencia_alumnos.db'
db_firestore = None
app_id = "asistencia-unsam-app"

def get_sqlite_conn():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def init_backend():
    # 1. Inicializar SQLite siempre como respaldo
    init_sqlite_db()

    # 2. Intentar Inicializar Firestore
    global db_firestore, app_id
    firebase_config_str = os.environ.get('__firebase_config', None)
    app_id = os.environ.get('__app_id', 'local-dev')

    if firebase_admin and firebase_config_str:
        try:
            if not firebase_admin._apps:
                cred_dict = json.loads(firebase_config_str)
                cred = credentials.Certificate(cred_dict) if "private_key" in cred_dict else credentials.ApplicationDefault()
                firebase_admin.initialize_app(cred)
            db_firestore = firestore.client()
            print(f"✅ MODO NUBE: Conectado a Firestore ({app_id})")
            return True
        except Exception as e:
            print(f"⚠️ Error Firestore: {e}. Usando SQLite.")
    else:
        print("ℹ️ MODO LOCAL: Usando SQLite")
    return False

def init_sqlite_db():
    conn = get_sqlite_conn()
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS Usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, password TEXT NOT NULL, role TEXT NOT NULL)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Ciclos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE, activo INTEGER DEFAULT 0)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Cursos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, ciclo_id INTEGER, FOREIGN KEY (ciclo_id) REFERENCES Ciclos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Alumnos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, nombre TEXT NOT NULL, dni TEXT, observaciones TEXT, tutor_nombre TEXT, tutor_telefono TEXT, UNIQUE(curso_id, nombre), FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Asistencia (id INTEGER PRIMARY KEY AUTOINCREMENT, alumno_id INTEGER NOT NULL, fecha TEXT NOT NULL, status TEXT NOT NULL, UNIQUE(alumno_id, fecha), FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Requisitos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, descripcion TEXT NOT NULL, FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Requisitos_Cumplidos (requisito_id INTEGER NOT NULL, alumno_id INTEGER NOT NULL, PRIMARY KEY (requisito_id, alumno_id), FOREIGN KEY (requisito_id) REFERENCES Requisitos(id) ON DELETE CASCADE, FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)")
    
    # Migraciones simples
    for col in ["dni", "observaciones", "tutor_nombre", "tutor_telefono"]:
        try: cursor.execute(f"ALTER TABLE Alumnos ADD COLUMN {col} TEXT")
        except: pass

    cursor.execute("SELECT COUNT(*) FROM Usuarios")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", ("admin", hash_password("admin"), "admin"))
    
    cursor.execute("SELECT COUNT(*) FROM Ciclos")
    if cursor.fetchone()[0] == 0:
        anio = str(date.today().year)
        cursor.execute("INSERT INTO Ciclos (nombre, activo) VALUES (?, 1)", (anio,))
        cid = cursor.lastrowid
        cursor.execute("UPDATE Cursos SET ciclo_id = ? WHERE ciclo_id IS NULL", (cid,))

    conn.commit()
    conn.close()

def get_col(name):
    return db_firestore.collection('artifacts').document(app_id).collection('public').document('data').collection(name)

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# ======================================================================
# FUNCIONES CRUD
# ======================================================================

def authenticate_user(username, password):
    pwd = hash_password(password)
    if db_firestore:
        try:
            q = get_col('Usuarios').where('username', '==', username).limit(1).stream()
            doc = next(q, None)
            if doc and doc.to_dict().get('password') == pwd:
                return (True, doc.to_dict().get('role'))
            if not doc and username=="admin" and password=="admin":
                 if not next(get_col('Usuarios').limit(1).stream(), None):
                     add_user("admin", "admin", "admin")
                     return (True, "admin")
        except: pass
    else:
        conn = get_sqlite_conn()
        u = conn.execute("SELECT role FROM Usuarios WHERE username=? AND password=?", (username, pwd)).fetchone()
        conn.close()
        if u: return (True, u['role'])
    return (False, None)

def get_users():
    if db_firestore: return [{**d.to_dict(), 'id': d.id} for d in get_col('Usuarios').stream()]
    conn = get_sqlite_conn()
    res = [dict(r) for r in conn.execute("SELECT * FROM Usuarios").fetchall()]
    conn.close()
    return res

def add_user(u, p, r):
    hp = hash_password(p)
    if db_firestore:
        if next(get_col('Usuarios').where('username', '==', u).limit(1).stream(), None): return False
        get_col('Usuarios').add({'username': u, 'password': hp, 'role': r})
        return True
    try:
        conn = get_sqlite_conn()
        conn.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", (u, hp, r))
        conn.commit(); conn.close()
        return True
    except: return False

def delete_user(uid):
    if db_firestore: get_col('Usuarios').document(uid).delete()
    else:
        conn = get_sqlite_conn()
        conn.execute("DELETE FROM Usuarios WHERE id=?", (uid,))
        conn.commit(); conn.close()

def get_ciclos():
    if db_firestore:
        return [{**d.to_dict(), 'id': d.id} for d in get_col('Ciclos').order_by('nombre', direction=firestore.Query.DESCENDING).stream()]
    conn = get_sqlite_conn()
    res = [dict(r) for r in conn.execute("SELECT * FROM Ciclos ORDER BY nombre DESC").fetchall()]
    conn.close()
    return res

def get_ciclo_activo():
    if db_firestore:
        doc = next(get_col('Ciclos').where('activo', '==', 1).limit(1).stream(), None)
        return {**doc.to_dict(), 'id': doc.id} if doc else None
    conn = get_sqlite_conn()
    res = conn.execute("SELECT * FROM Ciclos WHERE activo=1").fetchone()
    conn.close()
    return dict(res) if res else None

def add_ciclo(nombre):
    if db_firestore:
        batch = db_firestore.batch()
        for d in get_col('Ciclos').where('activo', '==', 1).stream(): batch.update(d.reference, {'activo': 0})
        get_col('Ciclos').add({'nombre': nombre, 'activo': 1})
        batch.commit()
    else:
        conn = get_sqlite_conn()
        conn.execute("UPDATE Ciclos SET activo=0")
        conn.execute("INSERT INTO Ciclos (nombre, activo) VALUES (?, 1)", (nombre,))
        conn.commit(); conn.close()

def activar_ciclo(cid):
    if db_firestore:
        batch = db_firestore.batch()
        for d in get_col('Ciclos').where('activo', '==', 1).stream(): batch.update(d.reference, {'activo': 0})
        batch.update(get_col('Ciclos').document(cid), {'activo': 1})
        batch.commit()
    else:
        conn = get_sqlite_conn()
        conn.execute("UPDATE Ciclos SET activo=0")
        conn.execute("UPDATE Ciclos SET activo=1 WHERE id=?", (cid,))
        conn.commit(); conn.close()

def get_cursos():
    ciclo = get_ciclo_activo()
    if not ciclo: return []
    cid = ciclo['id']
    if db_firestore:
        return [{**d.to_dict(), 'id': d.id} for d in get_col('Cursos').where('ciclo_id', '==', cid).order_by('nombre').stream()]
    conn = get_sqlite_conn()
    res = [dict(r) for r in conn.execute("SELECT * FROM Cursos WHERE ciclo_id=? ORDER BY nombre", (cid,)).fetchall()]
    conn.close()
    return res

def add_curso(nombre):
    ciclo = get_ciclo_activo()
    if not ciclo: return False
    if db_firestore:
        if next(get_col('Cursos').where('ciclo_id', '==', ciclo['id']).where('nombre', '==', nombre).limit(1).stream(), None): return False
        get_col('Cursos').add({'nombre': nombre, 'ciclo_id': ciclo['id']})
        return True
    try:
        conn = get_sqlite_conn()
        conn.execute("INSERT INTO Cursos (nombre, ciclo_id) VALUES (?, ?)", (nombre, ciclo['id']))
        conn.commit(); conn.close()
        return True
    except: return False

def delete_curso(cid):
    if db_firestore: get_col('Cursos').document(cid).delete()
    else:
        conn = get_sqlite_conn()
        conn.execute("DELETE FROM Cursos WHERE id=?", (cid,))
        conn.commit(); conn.close()

def get_alumnos(curso_id):
    if db_firestore:
        return [{**d.to_dict(), 'id': d.id} for d in get_col('Alumnos').where('curso_id', '==', curso_id).order_by('nombre').stream()]
    conn = get_sqlite_conn()
    res = [dict(r) for r in conn.execute("SELECT * FROM Alumnos WHERE curso_id=? ORDER BY nombre", (curso_id,)).fetchall()]
    conn.close()
    return res

def get_alumno_by_id(aid):
    if db_firestore:
        doc = get_col('Alumnos').document(aid).get()
        if doc.exists:
            d = doc.to_dict()
            c_doc = get_col('Cursos').document(d['curso_id']).get()
            d['curso_nombre'] = c_doc.to_dict().get('nombre', '?') if c_doc.exists else '?'
            return {**d, 'id': doc.id}
    else:
        conn = get_sqlite_conn()
        row = conn.execute("SELECT a.*, c.nombre as curso_nombre FROM Alumnos a JOIN Cursos c ON a.curso_id = c.id WHERE a.id=?", (aid,)).fetchone()
        conn.close()
        return dict(row) if row else None
    return None

def add_alumno(curso_id, nombre, dni, obs, t_n, t_t):
    data = {'curso_id': curso_id, 'nombre': nombre, 'dni': dni, 'observaciones': obs, 'tutor_nombre': t_n, 'tutor_telefono': t_t}
    if db_firestore: get_col('Alumnos').add(data)
    else:
        conn = get_sqlite_conn()
        conn.execute("INSERT INTO Alumnos (curso_id, nombre, dni, observaciones, tutor_nombre, tutor_telefono) VALUES (?,?,?,?,?,?)", (curso_id, nombre, dni, obs, t_n, t_t))
        conn.commit(); conn.close()

def update_alumno(aid, nombre, dni, obs, t_n, t_t):
    data = {'nombre': nombre, 'dni': dni, 'observaciones': obs, 'tutor_nombre': t_n, 'tutor_telefono': t_t}
    if db_firestore: get_col('Alumnos').document(aid).update(data)
    else:
        conn = get_sqlite_conn()
        conn.execute("UPDATE Alumnos SET nombre=?, dni=?, observaciones=?, tutor_nombre=?, tutor_telefono=? WHERE id=?", (nombre, dni, obs, t_n, t_t, aid))
        conn.commit(); conn.close()

def delete_alumno(aid):
    if db_firestore: get_col('Alumnos').document(aid).delete()
    else:
        conn = get_sqlite_conn()
        conn.execute("DELETE FROM Alumnos WHERE id=?", (aid,))
        conn.commit(); conn.close()

def search_students(term):
    term = term.lower()
    results = []
    if db_firestore:
        try:
            act_ciclo = get_ciclo_activo()
            if not act_ciclo: return []
            act_cursos = {c['id']: c['nombre'] for c in get_cursos()}
            for doc in get_col('Alumnos').stream():
                d = doc.to_dict()
                if d.get('curso_id') in act_cursos:
                    if term in d.get('nombre','').lower() or term in d.get('dni',''):
                        d['id'] = doc.id
                        d['curso_nombre'] = act_cursos[d['curso_id']]
                        d['ciclo_nombre'] = act_ciclo['nombre']
                        results.append(d)
        except: pass
    else:
        conn = get_sqlite_conn()
        rows = conn.execute("SELECT a.*, c.nombre as curso_nombre, ci.nombre as ciclo_nombre FROM Alumnos a JOIN Cursos c ON a.curso_id = c.id JOIN Ciclos ci ON c.ciclo_id = ci.id WHERE (lower(a.nombre) LIKE ? OR a.dni LIKE ?) AND ci.activo=1", (f"%{term}%", f"%{term}%")).fetchall()
        conn.close()
        results = [dict(r) for r in rows]
    return sorted(results, key=lambda x: x['nombre'])

def register_asistencia(aid, cid, fecha, status):
    if db_firestore:
        doc_id = f"{aid}_{fecha}"
        get_col('Asistencia').document(doc_id).set({'alumno_id': aid, 'curso_id': cid, 'fecha': fecha, 'status': status}, merge=True)
    else:
        conn = get_sqlite_conn()
        conn.execute("INSERT OR REPLACE INTO Asistencia (alumno_id, fecha, status) VALUES (?, ?, ?)", (aid, fecha, status))
        conn.commit(); conn.close()

def get_asistencia_diaria(curso_id, fecha):
    if db_firestore:
        try:
            rows = get_col('Asistencia').where('curso_id', '==', curso_id).where('fecha', '==', fecha).stream()
            return {r.to_dict()['alumno_id']: r.to_dict()['status'] for r in rows}
        except: 
            rows = get_col('Asistencia').where('fecha', '==', fecha).stream()
            return {r.to_dict()['alumno_id']: r.to_dict()['status'] for r in rows}
    else:
        conn = get_sqlite_conn()
        rows = conn.execute("SELECT a.id, asis.status FROM Alumnos a LEFT JOIN Asistencia asis ON a.id = asis.alumno_id AND asis.fecha = ? WHERE a.curso_id = ?", (fecha, curso_id)).fetchall()
        conn.close()
        return {r['id']: r['status'] for r in rows if r['status']}

def get_report_data(curso_id, start, end):
    alumnos = get_alumnos(curso_id)
    report = []
    asis_map = {} 
    
    if db_firestore:
        try:
            q = get_col('Asistencia').where('curso_id', '==', curso_id).where('fecha', '>=', start).where('fecha', '<=', end).stream()
            for doc in q:
                d = doc.to_dict()
                aid = d['alumno_id']
                if aid not in asis_map: asis_map[aid] = []
                asis_map[aid].append(d['status'])
        except: pass
    else:
        conn = get_sqlite_conn()
        rows = conn.execute("SELECT alumno_id, status FROM Asistencia WHERE fecha >= ? AND fecha <= ? AND alumno_id IN (SELECT id FROM Alumnos WHERE curso_id=?)", (start, end, curso_id)).fetchall()
        conn.close()
        for r in rows:
            aid = r['alumno_id']
            if aid not in asis_map: asis_map[aid] = []
            asis_map[aid].append(r['status'])

    for a in alumnos:
        statuses = asis_map.get(a['id'], [])
        counts = {k: statuses.count(k) for k in ['P','T','A','J','S','N']}
        faltas = counts['A'] + counts['S'] + (counts['T'] * 0.25)
        total = counts['P'] + counts['T'] + counts['A'] + counts['J'] + counts['S']
        pct = (faltas/total*100) if total > 0 else 0
        report.append({
            'nombre': a['nombre'], 'dni': a.get('dni','-'), 'tutor': a.get('tutor_nombre', '-'), 'tel': a.get('tutor_telefono', '-'),
            'p': counts['P'], 't': counts['T'], 'a': counts['A'], 'j': counts['J'], 's': counts['S'], 'faltas': faltas, 'pct': round(pct, 1)
        })
    return report

def crud_req(action, **kwargs):
    if db_firestore:
        col = get_col('Requisitos')
        if action=='get': return [{**d.to_dict(), 'id': d.id} for d in col.where('curso_id', '==', kwargs['cid']).stream()]
        if action=='add': col.add({'curso_id': kwargs['cid'], 'descripcion': kwargs['desc']})
        if action=='del': col.document(kwargs['rid']).delete()
    else:
        conn = get_sqlite_conn()
        if action=='get': 
            res = [dict(r) for r in conn.execute("SELECT * FROM Requisitos WHERE curso_id=?", (kwargs['cid'],)).fetchall()]
            conn.close(); return res
        if action=='add': conn.execute("INSERT INTO Requisitos (curso_id, descripcion) VALUES (?, ?)", (kwargs['cid'], kwargs['desc']))
        if action=='del': conn.execute("DELETE FROM Requisitos WHERE id=?", (kwargs['rid'],))
        conn.commit(); conn.close()

def crud_req_done(action, **kwargs):
    if db_firestore:
        col = get_col('Requisitos_Cumplidos')
        if action=='get': return {d.to_dict()['alumno_id'] for d in col.where('requisito_id', '==', kwargs['rid']).stream()}
        if action=='toggle':
            did = f"{kwargs['rid']}_{kwargs['aid']}"
            if kwargs['val']: col.document(did).set({'requisito_id': kwargs['rid'], 'alumno_id': kwargs['aid']})
            else: col.document(did).delete()
    else:
        conn = get_sqlite_conn()
        if action=='get':
            res = {r['alumno_id'] for r in conn.execute("SELECT alumno_id FROM Requisitos_Cumplidos WHERE requisito_id=?", (kwargs['rid'],)).fetchall()}
            conn.close(); return res
        if action=='toggle':
            if kwargs['val']: conn.execute("INSERT OR IGNORE INTO Requisitos_Cumplidos (requisito_id, alumno_id) VALUES (?, ?)", (kwargs['rid'], kwargs['aid']))
            else: conn.execute("DELETE FROM Requisitos_Cumplidos WHERE requisito_id=? AND alumno_id=?", (kwargs['rid'], kwargs['aid']))
        conn.commit(); conn.close()

def get_requisitos(cid): return crud_req('get', cid=cid)
def add_requisito(cid, desc): crud_req('add', cid=cid, desc=desc)
def delete_requisito(rid): crud_req('del', rid=rid)
def get_cumplimientos(rid): return crud_req_done('get', rid=rid)
def toggle_cumplimiento(rid, aid, val): crud_req_done('toggle', rid=rid, aid=aid, val=val)

def get_student_req_status(aid, cid):
    reqs = get_requisitos(cid)
    res = []
    for r in reqs:
        done = get_cumplimientos(r['id'])
        res.append({'desc': r['descripcion'], 'ok': aid in done})
    return res

# ======================================================================
# 2. INTERFAZ GRÁFICA (Flet - Strings only)
# ======================================================================

def main(page: ft.Page):
    page.title = "Asistencia UNSAM"
    page.theme_mode = "light"
    page.padding = 0
    init_backend()

    state = {"role": None, "username": None, "curso_id": None, "curso_nombre": None, "search": "", "st_view": None, "st_edit": None}

    def show_snack(m, c="green"):
        page.snack_bar = ft.SnackBar(ft.Text(m), bgcolor=c)
        page.snack_bar.open = True
        page.update()

    def login_view():
        user = ft.TextField(label="Usuario", width=300, bgcolor="white")
        pwd = ft.TextField(label="Clave", password=True, width=300, bgcolor="white")
        def login(e):
            ok, role = authenticate_user(user.value, pwd.value)
            if ok: state["role"], state["username"] = role, user.value; page.go("/dashboard")
            else: show_snack("Datos incorrectos", "red")
        
        logo = ft.Image(src="logo_unsam.png", width=200) if os.path.exists("logo_unsam.png") else ft.Icon("school", size=80, color="blue")
        
        return ft.View("/", [
            ft.Container(content=ft.Column([
                logo, ft.Text("Sistema de Asistencia", size=24, weight="bold"),
                ft.Text("UNSAM", size=16, color="grey"), ft.Divider(height=20, color="transparent"),
                user, pwd, ft.ElevatedButton("ENTRAR", on_click=login, width=300, height=50, bgcolor="blue", color="white")
            ], horizontal_alignment="center"), alignment=ft.alignment.center, expand=True, bgcolor="#f0f2f5")])

    def dashboard_view():
        ciclo = get_ciclo_activo()
        c_nombre = ciclo['nombre'] if ciclo else "Sin Ciclo"
        search = ft.TextField(hint_text="Buscar...", expand=True, bgcolor="white")
        def do_search(e): 
            if search.value: state["search"]=search.value; page.go("/search")
        
        cursos_col = ft.Column(scroll="auto", expand=True)
        def load():
            cursos_col.controls.clear()
            for c in get_cursos():
                cursos_col.controls.append(ft.Container(
                    content=ft.Row([
                        ft.Icon("book", color="blue"),
                        ft.Text(c['nombre'], weight="bold", size=16, expand=True),
                        ft.IconButton("arrow_forward", on_click=lambda e, cid=c['id'], cn=c['nombre']: go_curso(cid, cn)),
                        ft.IconButton("delete", icon_color="red", on_click=lambda e, cid=c['id']: (delete_curso(cid), load())) if state["role"]=='admin' else ft.Container()
                    ], alignment="spaceBetween"),
                    padding=15, bgcolor="white", border_radius=10, shadow=ft.BoxShadow(blur_radius=2, color="black"), margin=5
                ))
            page.update()
        
        def go_curso(cid, cn): state["curso_id"]=cid; state["curso_nombre"]=cn; page.go("/curso")
        def add_c(e): page.go("/form_curso") if ciclo else show_snack("Falta Ciclo Activo", "red")
        load()
        admin = ft.IconButton("settings", icon_color="white", on_click=lambda _: page.go("/admin")) if state["role"]=='admin' else ft.Container()
        return ft.View("/dashboard", [ft.AppBar(title=ft.Text("Panel Principal"), bgcolor="blue", color="white", actions=[admin, ft.IconButton("logout", icon_color="white", on_click=lambda _: page.go("/"))]), ft.Container(content=ft.Column([ft.Text(f"Ciclo: {c_nombre}", color="blue", weight="bold"), ft.Row([search, ft.IconButton("search", on_click=do_search)]), ft.Row([ft.Text("Cursos", size=20, weight="bold"), ft.IconButton("add_circle", icon_color="green", icon_size=30, on_click=add_c)], alignment="spaceBetween"), cursos_col]), padding=15, bgcolor="#f0f2f5", expand=True)])

    def curso_view():
        col = ft.Column(scroll="auto", expand=True)
        def load():
            col.controls.clear()
            for a in get_alumnos(state["curso_id"]):
                col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(a['nombre']), subtitle=ft.Text(f"DNI: {a.get('dni','-')}"), trailing=ft.PopupMenuButton(items=[ft.PopupMenuItem("Editar", on_click=lambda e, aid=a['id']: (state.update({"st_edit": aid}), page.go("/form_student"))), ft.PopupMenuItem("Borrar", on_click=lambda e, aid=a['id']: (delete_alumno(aid), load()))])), bgcolor="white", border_radius=10, margin=2))
            page.update()
        load()
        return ft.View("/curso", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(state["curso_nombre"]), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.Row([ft.ElevatedButton("Asistencia", icon="check", expand=True, on_click=lambda _: page.go("/asistencia")), ft.ElevatedButton("Pedidos", icon="list", expand=True, on_click=lambda _: page.go("/pedidos")), ft.ElevatedButton("Reportes", icon="bar_chart", expand=True, on_click=lambda _: page.go("/reportes"))]), ft.Divider(), ft.Row([ft.Text("Alumnos", size=18, weight="bold"), ft.IconButton("person_add", icon_color="green", on_click=lambda _: (state.update({"st_edit": None}), page.go("/form_student")))], alignment="spaceBetween"), col]), padding=15, bgcolor="#f0f2f5", expand=True)])

    def asistencia_view():
        dp = ft.TextField(label="Fecha", value=date.today().isoformat(), bgcolor="white")
        col = ft.Column(scroll="auto", expand=True); vals = {}
        def load(e=None):
            try: 
                if date.fromisoformat(dp.value).weekday() >= 5: show_snack("⚠️ Fin de semana", "orange")
            except: pass
            ex = get_asistencia_diaria(state["curso_id"], dp.value); col.controls.clear(); vals.clear()
            for a in get_alumnos(state["curso_id"]):
                dd = ft.Dropdown(options=[ft.dropdown.Option(x) for x in ["P","T","A","J","S","N"]], value=ex.get(a['id'], "P"), width=80, bgcolor="white")
                vals[a['id']] = dd
                col.controls.append(ft.Container(content=ft.Row([ft.Text(a['nombre'], expand=True), dd]), padding=5, bgcolor="white", border_radius=5, margin=2))
            page.update()
        def save(e):
            try:
                d = date.fromisoformat(dp.value)
                if d > date.today() or d.weekday() >= 5: return show_snack("Fecha inválida", "red")
            except: return show_snack("Error fecha", "red")
            for aid, dd in vals.items(): register_asistencia(aid, state["curso_id"], dp.value, dd.value)
            show_snack("Guardado"); page.go("/curso")
        load()
        return ft.View("/asistencia", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Asistencia"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.Row([dp, ft.IconButton("refresh", on_click=load)]), ft.ElevatedButton("GUARDAR", on_click=save, bgcolor="green", color="white", width=float("inf")), ft.Divider(), col]), padding=15, bgcolor="#f0f2f5", expand=True)])

    def reportes_view():
        d1 = ft.TextField(label="Desde", value=date.today().replace(day=1).isoformat(), width=130, bgcolor="white"); d2 = ft.TextField(label="Hasta", value=date.today().isoformat(), width=130, bgcolor="white")
        table_cont = ft.Column(scroll="auto", expand=True)
        def gen(e):
            data = get_report_data(state["curso_id"], d1.value, d2.value); rows = []
            for d in data:
                c = "red" if d['faltas']>=25 else ("orange" if d['faltas']>=15 else "black")
                rows.append(ft.DataRow(cells=[ft.DataCell(ft.Text(d['nombre'], color=c)), ft.DataCell(ft.Text(str(d['p']))), ft.DataCell(ft.Text(str(d['t']))), ft.DataCell(ft.Text(str(d['a']))), ft.DataCell(ft.Text(str(d['j']))), ft.DataCell(ft.Text(str(d['s']))), ft.DataCell(ft.Text(f"{d['faltas']}", color=c, weight="bold")), ft.DataCell(ft.Text(f"{d['pct']}%", color=c))]))
            dt = ft.DataTable(columns=[ft.DataColumn(ft.Text("Alumno")), ft.DataColumn(ft.Text("P"), numeric=True), ft.DataColumn(ft.Text("T"), numeric=True), ft.DataColumn(ft.Text("A"), numeric=True), ft.DataColumn(ft.Text("J"), numeric=True), ft.DataColumn(ft.Text("S"), numeric=True), ft.DataColumn(ft.Text("F"), numeric=True), ft.DataColumn(ft.Text("%"), numeric=True)], rows=rows, bgcolor="white", border_radius=10, column_spacing=15)
            table_cont.controls = [ft.Row([dt], scroll="always")]; page.update()
        def export(e):
            if not pd: return show_snack("Falta pandas", "red")
            data = get_report_data(state["curso_id"], d1.value, d2.value)
            if not data: return show_snack("Sin datos", "orange")
            df = pd.DataFrame(data).rename(columns={'nombre':'Alumno', 'p':'Pres', 't':'Tarde', 'a':'Aus', 'j':'Just', 's':'Susp', 'faltas':'Total', 'pct':'%'})
            output = io.BytesIO(); df.to_excel(output, index=False, engine='xlsxwriter'); b64 = base64.b64encode(output.getvalue()).decode()
            page.launch_url(f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}", web_window_name="reporte.xlsx")
        return ft.View("/reportes", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Reportes"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.Row([d1, d2, ft.ElevatedButton("Ver", on_click=gen)]), ft.ElevatedButton("Excel", icon="download", on_click=export, bgcolor="green", color="white"), ft.Divider(), table_cont]), padding=15, bgcolor="#f0f2f5", expand=True)])

    # --- OTRAS VISTAS ---
    def search_view():
        term = state["search"]; res = search_students(term); col = ft.Column(scroll="auto")
        if not res: col.controls.append(ft.Text("Sin resultados"))
        else:
            for r in res: col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(r['nombre']), subtitle=ft.Text(f"Curso: {r['curso_nombre']}"), on_click=lambda e, s=r: (state.update({"st_view": s['id'], "curso_id": s['curso_id']}), page.go("/student_detail"))), bgcolor="white", border_radius=10, margin=2))
        return ft.View("/search", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(f"Busqueda: {term}"), bgcolor="blue", color="white"), ft.Container(content=col, padding=15, bgcolor="#f0f2f5", expand=True)])

    def student_detail_view():
        aid = state["st_view"]; s = get_alumno_by_id(aid)
        if not s: return ft.View("/error", [ft.Text("Error")])
        req_col = ft.Column()
        for r in get_student_req_status(aid, s['curso_id']): req_col.controls.append(ft.Row([ft.Icon("check" if r['ok'] else "close", color="green" if r['ok'] else "red"), ft.Text(r['desc'])]))
        card = ft.Container(content=ft.Column([ft.Text(s['nombre'], size=24, weight="bold"), ft.Text(f"Curso: {s['curso_nombre']}"), ft.Text(f"DNI: {s.get('dni','-')}"), ft.Divider(), ft.Text("Tutor:", weight="bold"), ft.Text(f"{s.get('tutor_nombre','-')} ({s.get('tutor_telefono','-')})"), ft.Divider(), ft.Text("Obs:", weight="bold"), ft.Text(s.get('observaciones','-'), italic=True), ft.Divider(), ft.Text("Papeles:", weight="bold"), req_col]), padding=20, bgcolor="white", border_radius=15, shadow=ft.BoxShadow(blur_radius=5, color="black"))
        return ft.View("/student_detail", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/search")), title=ft.Text("Ficha"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([card], scroll="auto"), padding=20, bgcolor="#f0f2f5", expand=True)])

    def form_student_view():
        is_edit = state["st_edit"] is not None
        nm = ft.TextField(label="Nombre", bgcolor="white"); dni = ft.TextField(label="DNI", bgcolor="white"); obs = ft.TextField(label="Obs", multiline=True, bgcolor="white"); tn = ft.TextField(label="Tutor", bgcolor="white"); tt = ft.TextField(label="Tel Tutor", bgcolor="white")
        if is_edit:
            d = get_alumno_by_id(state["st_edit"]); nm.value=d.get('nombre',''); dni.value=d.get('dni',''); obs.value=d.get('observaciones',''); tn.value=d.get('tutor_nombre',''); tt.value=d.get('tutor_telefono','')
        def save(e):
            if nm.value:
                if is_edit: update_alumno(state["st_edit"], nm.value, dni.value, obs.value, tn.value, tt.value)
                else: add_alumno(state["curso_id"], nm.value, dni.value, obs.value, tn.value, tt.value)
                page.go("/curso")
        return ft.View("/form_student", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Alumno"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([nm, dni, obs, tn, tt, ft.ElevatedButton("Guardar", on_click=save, bgcolor="green", color="white")]), padding=20, bgcolor="#f0f2f5", expand=True)])

    def form_curso_view():
        tf = ft.TextField(label="Nombre", bgcolor="white")
        def save(e): 
            if add_curso(tf.value): page.go("/dashboard")
            else: show_snack("Error", "red")
        return ft.View("/form_curso", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Nuevo Curso"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([tf, ft.ElevatedButton("Crear", on_click=save)]), padding=20, bgcolor="#f0f2f5", expand=True)])

    def pedidos_view():
        dd = ft.Dropdown(label="Pedido", expand=True, bgcolor="white", on_change=lambda e: lc()); col = ft.Column(scroll="auto", expand=True); rm = {}
        def lr():
            rs = get_requisitos(state["curso_id"]); rm.clear(); dd.options.clear()
            for r in rs: rm[r['descripcion']] = r['id']; dd.options.append(ft.dropdown.Option(r['descripcion']))
            if rs: dd.value = rs[0]['descripcion']
            page.update(); lc()
        def lc():
            col.controls.clear()
            if not dd.value: return
            rid = rm[dd.value]; done = get_cumplimientos(rid)
            for a in get_alumnos(state["curso_id"]): col.controls.append(ft.Container(content=ft.Checkbox(label=a['nombre'], value=(a['id'] in done), on_change=lambda e, aid=a['id'], rid=rid: toggle_cumplimiento(rid, aid, e.control.value)), bgcolor="white", padding=5, border_radius=5, margin=2))
            page.update()
        def add(e): page.go("/form_req")
        def dele(e): 
            if dd.value: delete_requisito(rm[dd.value]); lr()
        lr()
        return ft.View("/pedidos", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Pedidos"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.Row([dd, ft.IconButton("add", on_click=add), ft.IconButton("delete", icon_color="red", on_click=dele)]), ft.Divider(), col]), padding=15, bgcolor="#f0f2f5", expand=True)])

    def form_req_view():
        tf = ft.TextField(label="Descripción", bgcolor="white")
        def save(e):
            if tf.value: add_requisito(state["curso_id"], tf.value); page.go("/pedidos")
        return ft.View("/form_req", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/pedidos")), title=ft.Text("Nuevo Pedido"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([tf, ft.ElevatedButton("Crear", on_click=save)]), padding=20, bgcolor="#f0f2f5", expand=True)])

    def admin_view():
        return ft.View("/admin", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Admin"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.ListTile(leading=ft.Icon("calendar_month"), title=ft.Text("Ciclos"), on_click=lambda _: page.go("/ciclos")), ft.ListTile(leading=ft.Icon("people"), title=ft.Text("Usuarios"), on_click=lambda _: page.go("/users"))]), padding=20, bgcolor="#f0f2f5", expand=True)])

    def ciclos_view():
        tf = ft.TextField(label="Año", expand=True, bgcolor="white"); col = ft.Column(scroll="auto", expand=True)
        def ld():
            col.controls.clear()
            for c in get_ciclos():
                act = c['activo']==1
                col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("check" if act else "circle", color="green" if act else "grey"), title=ft.Text(c['nombre']), trailing=ft.ElevatedButton("Activar", on_click=lambda e, cid=c['id']: (activar_ciclo(cid), ld())) if not act else None), bgcolor="white", border_radius=10, margin=2))
            page.update()
        def add(e): 
            if tf.value: add_ciclo(tf.value); tf.value=""; ld()
        ld()
        return ft.View("/ciclos", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/admin")), title=ft.Text("Ciclos"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.Row([tf, ft.IconButton("add", on_click=add)]), ft.Divider(), col]), padding=20, bgcolor="#f0f2f5", expand=True)])

    def users_view():
        u = ft.TextField(label="User", expand=True, bgcolor="white"); p = ft.TextField(label="Pass", password=True, expand=True, bgcolor="white"); r = ft.Dropdown(options=[ft.dropdown.Option("preceptor"), ft.dropdown.Option("admin")], value="preceptor", width=100, bgcolor="white"); col = ft.Column()
        def ld():
            col.controls.clear()
            for us in get_users():
                col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(us['username']), subtitle=ft.Text(us['role']), trailing=ft.IconButton("delete", icon_color="red", on_click=lambda e, uid=us['id']: (delete_user(uid), ld())) if us['username']!=state['username'] else None), bgcolor="white", border_radius=10, margin=2))
            page.update()
        def add(e): 
            if add_user(u.value, p.value, r.value): u.value=""; p.value=""; ld()
        ld()
        return ft.View("/users", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/admin")), title=ft.Text("Usuarios"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([ft.Row([u, p, r, ft.IconButton("add", on_click=add)]), ft.Divider(), col]), padding=20, bgcolor="#f0f2f5", expand=True)])

    # --- ROUTER ---
    def router(route):
        page.views.clear()
        views = {
            "/": login_view, "/dashboard": dashboard_view, "/curso": curso_view, "/asistencia": asistencia_view,
            "/pedidos": pedidos_view, "/form_req": form_req_view, "/reportes": reportes_view,
            "/search": search_view, "/student_detail": student_detail_view, "/form_student": form_student_view,
            "/form_curso": form_curso_view, "/admin": admin_view, "/ciclos": ciclos_view, "/users": users_view
        }
        if page.route in views: page.views.append(views[page.route]())
        else: page.views.append(login_view())
        page.update()

    page.on_route_change = router
    page.on_view_pop = lambda view: page.go(page.views[-2].route)
    page.go("/")

if __name__ == "__main__":
    # Detección de entorno: Si hay PORT, es Nube. Si no, es Local.
    port_env = os.environ.get("PORT")
    
    if port_env:
        # MODO NUBE (Render/Railway): Escuchar en 0.0.0.0
        ft.app(target=main, view=ft.AppView.WEB_BROWSER, port=int(port_env), host="0.0.0.0", web_renderer="html")
    else:
        # MODO LOCAL (Tu PC): Escuchar en localhost (127.0.0.1)
        # Esto permite que Windows abra el navegador correctamente
        ft.app(target=main, view=ft.AppView.WEB_BROWSER, port=8000, web_renderer="html")
