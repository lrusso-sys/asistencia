import flet as ft
import sqlite3
import hashlib
from datetime import date
import os

# Intentamos importar pandas y openpyxl
try:
    import pandas as pd
except ImportError:
    pd = None
    print("⚠️ ADVERTENCIA: 'pandas' no está instalado. La exportación fallará.")

try:
    import openpyxl
except ImportError:
    print("⚠️ ADVERTENCIA: 'openpyxl' no está instalado. La exportación a Excel fallará.")

# ======================================================================
# 1. LÓGICA DE BASE DE DATOS (Backend Completo)
# ======================================================================

DB_NAME = 'asistencia_alumnos.db'

def get_db_connection():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;") 
    return conn

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""CREATE TABLE IF NOT EXISTS Usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, password TEXT NOT NULL, role TEXT NOT NULL)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS Cursos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS Alumnos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, nombre TEXT NOT NULL, dni TEXT, observaciones TEXT, UNIQUE(curso_id, nombre), FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS Asistencia (id INTEGER PRIMARY KEY AUTOINCREMENT, alumno_id INTEGER NOT NULL, fecha TEXT NOT NULL, status TEXT NOT NULL, UNIQUE(alumno_id, fecha), FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS Requisitos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, descripcion TEXT NOT NULL, FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS Requisitos_Cumplidos (requisito_id INTEGER NOT NULL, alumno_id INTEGER NOT NULL, PRIMARY KEY (requisito_id, alumno_id), FOREIGN KEY (requisito_id) REFERENCES Requisitos(id) ON DELETE CASCADE, FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)""")

    try: cursor.execute("ALTER TABLE Alumnos ADD COLUMN dni TEXT")
    except: pass
    try: cursor.execute("ALTER TABLE Alumnos ADD COLUMN observaciones TEXT")
    except: pass
    
    # Tabla Ciclos
    cursor.execute("""CREATE TABLE IF NOT EXISTS Ciclos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE, activo INTEGER DEFAULT 0)""")
    try: cursor.execute("ALTER TABLE Cursos ADD COLUMN ciclo_id INTEGER REFERENCES Ciclos(id) ON DELETE CASCADE")
    except: pass

    # Inicializar Ciclo por defecto
    cursor.execute("SELECT COUNT(*) FROM Ciclos")
    if cursor.fetchone()[0] == 0:
        anio = str(date.today().year)
        cursor.execute("INSERT INTO Ciclos (nombre, activo) VALUES (?, 1)", (anio,))
        cid = cursor.lastrowid
        cursor.execute("UPDATE Cursos SET ciclo_id = ? WHERE ciclo_id IS NULL", (cid,))

    # Admin por defecto
    cursor.execute("SELECT COUNT(*) FROM Usuarios")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", ("admin", hash_password("admin"), "admin"))
    
    conn.commit()
    conn.close()

# --- Funciones CRUD ---

def authenticate_user(username, password):
    conn = get_db_connection()
    user = conn.execute("SELECT * FROM Usuarios WHERE username = ? AND password = ?", (username, hash_password(password))).fetchone()
    conn.close()
    return (True, user['role']) if user else (False, None)

def get_users():
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Usuarios").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_user(u, p, r):
    try:
        conn = get_db_connection()
        conn.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", (u, hash_password(p), r))
        conn.commit()
        return True
    except: return False
    finally: conn.close()

def delete_user(uid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Usuarios WHERE id = ?", (uid,))
    conn.commit()
    conn.close()

# --- Ciclos ---
def get_ciclos():
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Ciclos ORDER BY nombre DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_ciclo_activo():
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM Ciclos WHERE activo = 1").fetchone()
    conn.close()
    return dict(row) if row else None

def add_ciclo(nombre):
    try:
        conn = get_db_connection()
        conn.execute("UPDATE Ciclos SET activo = 0")
        conn.execute("INSERT INTO Ciclos (nombre, activo) VALUES (?, 1)", (nombre,))
        conn.commit()
        return True
    except: return False
    finally: conn.close()

def activar_ciclo(cid):
    conn = get_db_connection()
    conn.execute("UPDATE Ciclos SET activo = 0")
    conn.execute("UPDATE Ciclos SET activo = 1 WHERE id = ?", (cid,))
    conn.commit()
    conn.close()

# --- Cursos ---
def get_cursos():
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT c.* FROM Cursos c
        JOIN Ciclos ci ON c.ciclo_id = ci.id
        WHERE ci.activo = 1
        ORDER BY c.nombre
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_curso(nombre):
    try:
        conn = get_db_connection()
        ciclo = conn.execute("SELECT id FROM Ciclos WHERE activo = 1").fetchone()
        if not ciclo: return False
        conn.execute("INSERT INTO Cursos (nombre, ciclo_id) VALUES (?, ?)", (nombre, ciclo['id']))
        conn.commit()
        return True
    except: return False
    finally: conn.close()

def delete_curso(cid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Cursos WHERE id = ?", (cid,))
    conn.commit()
    conn.close()

# --- Alumnos ---
def get_alumnos(curso_id):
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Alumnos WHERE curso_id = ? ORDER BY nombre", (curso_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_alumno_by_id(aid):
    conn = get_db_connection()
    row = conn.execute("SELECT a.*, c.nombre as curso_nombre FROM Alumnos a JOIN Cursos c ON a.curso_id = c.id WHERE a.id = ?", (aid,)).fetchone()
    conn.close()
    return dict(row) if row else None

def add_alumno(curso_id, nombre, dni, obs):
    try:
        conn = get_db_connection()
        conn.execute("INSERT INTO Alumnos (curso_id, nombre, dni, observaciones) VALUES (?, ?, ?, ?)", (curso_id, nombre, dni, obs))
        conn.commit()
        return True
    except: return False
    finally: conn.close()

def update_alumno(aid, nombre, dni, obs):
    conn = get_db_connection()
    conn.execute("UPDATE Alumnos SET nombre=?, dni=?, observaciones=? WHERE id=?", (nombre, dni, obs, aid))
    conn.commit()
    conn.close()

def delete_alumno(aid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Alumnos WHERE id=?", (aid,))
    conn.commit()
    conn.close()

def search_students(term):
    conn = get_db_connection()
    term = f"%{term}%"
    query = """
        SELECT a.*, c.nombre as curso_nombre, ci.nombre as ciclo_nombre
        FROM Alumnos a 
        JOIN Cursos c ON a.curso_id = c.id
        JOIN Ciclos ci ON c.ciclo_id = ci.id
        WHERE (a.nombre LIKE ? OR a.dni LIKE ?) AND ci.activo = 1
        ORDER BY a.nombre
    """
    rows = conn.execute(query, (term, term)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

# --- Asistencia ---
def get_asistencia_diaria(curso_id, fecha):
    conn = get_db_connection()
    rows = conn.execute("SELECT a.id, asis.status FROM Alumnos a LEFT JOIN Asistencia asis ON a.id = asis.alumno_id AND asis.fecha = ? WHERE a.curso_id = ?", (fecha, curso_id)).fetchall()
    conn.close()
    return {r['id']: r['status'] for r in rows if r['status']}

def register_asistencia(aid, fecha, status):
    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO Asistencia (alumno_id, fecha, status) VALUES (?, ?, ?)", (aid, fecha, status))
    conn.commit()
    conn.close()

def get_report_data(curso_id, start, end):
    conn = get_db_connection()
    alumnos = conn.execute("SELECT * FROM Alumnos WHERE curso_id = ?", (curso_id,)).fetchall()
    data = []
    for a in alumnos:
        rows = conn.execute("SELECT status, COUNT(*) as c FROM Asistencia WHERE alumno_id=? AND fecha>=? AND fecha<=? GROUP BY status", (a['id'], start, end)).fetchall()
        counts = {r['status']: r['c'] for r in rows}
        p, t, aus, j, s = counts.get('P',0), counts.get('T',0), counts.get('A',0), counts.get('J',0), counts.get('S',0)
        faltas = aus + s + (t * 0.25)
        total = p + t + aus + j + s
        pct = (faltas/total*100) if total > 0 else 0
        data.append({'nombre': a['nombre'], 'dni': a['dni'], 'p': p, 't': t, 'a': aus, 'j': j, 's': s, 'faltas': faltas, 'pct': round(pct,1)})
    conn.close()
    return data

# --- Requisitos ---
def get_requisitos(curso_id):
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Requisitos WHERE curso_id = ?", (curso_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_requisito(curso_id, desc):
    conn = get_db_connection()
    conn.execute("INSERT INTO Requisitos (curso_id, descripcion) VALUES (?, ?)", (curso_id, desc))
    conn.commit()
    conn.close()

def delete_requisito(rid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Requisitos WHERE id = ?", (rid,))
    conn.commit()
    conn.close()

def get_cumplimientos(rid):
    conn = get_db_connection()
    rows = conn.execute("SELECT alumno_id FROM Requisitos_Cumplidos WHERE requisito_id = ?", (rid,)).fetchall()
    conn.close()
    return {r['alumno_id'] for r in rows}

def toggle_cumplimiento(rid, aid, val):
    conn = get_db_connection()
    if val: conn.execute("INSERT OR IGNORE INTO Requisitos_Cumplidos (requisito_id, alumno_id) VALUES (?, ?)", (rid, aid))
    else: conn.execute("DELETE FROM Requisitos_Cumplidos WHERE requisito_id = ? AND alumno_id = ?", (rid, aid))
    conn.commit()
    conn.close()

def get_student_req_status(aid, cid):
    conn = get_db_connection()
    reqs = conn.execute("SELECT * FROM Requisitos WHERE curso_id=?",(cid,)).fetchall()
    done = conn.execute("SELECT requisito_id FROM Requisitos_Cumplidos WHERE alumno_id=?",(aid,)).fetchall()
    done_ids = {r['requisito_id'] for r in done}
    conn.close()
    return [{'desc': r['descripcion'], 'ok': r['id'] in done_ids} for r in reqs]

# ======================================================================
# 2. INTERFAZ GRÁFICA WEB (Flet)
# ======================================================================

def main(page: ft.Page):
    page.title = "Sistema de Asistencia UNSAM"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 0
    COLOR_PRIMARY = "blue"
    COLOR_BG = "#F0F2F5"
    
    init_db()

    state = {
        "role": None,
        "username": None,
        "curso_id": None,
        "curso_nombre": None,
        "search_term": None,
        "student_id_view": None,
        "student_id_edit": None
    }

    def show_snack(msg, color="green"):
        page.snack_bar = ft.SnackBar(ft.Text(msg), bgcolor=color)
        page.snack_bar.open = True
        page.update()

    # --- VISTAS ---

    def login_view():
        user = ft.TextField(label="Usuario", width=300, bgcolor="white")
        pwd = ft.TextField(label="Contraseña", password=True, width=300, bgcolor="white")
        
        def login_click(e):
            ok, role = authenticate_user(user.value, pwd.value)
            if ok:
                state["role"] = role
                state["username"] = user.value
                page.go("/dashboard")
            else:
                show_snack("Credenciales incorrectas", "red")

        logo_widget = ft.Image(src="logo_unsam.png", width=200) if os.path.exists("logo_unsam.png") else ft.Icon("school", size=80, color=COLOR_PRIMARY)

        return ft.View("/", [
            ft.Container(
                content=ft.Column([
                    logo_widget,
                    ft.Text("Sistema de Asistencia", size=24, weight="bold"),
                    ft.Text("UNSAM", size=16, color="grey"),
                    ft.Divider(height=20, color="transparent"),
                    user, pwd,
                    ft.ElevatedButton("INGRESAR", on_click=login_click, width=300, height=50, bgcolor=COLOR_PRIMARY, color="white"),
                ], horizontal_alignment="center"),
                alignment=ft.alignment.center, expand=True, bgcolor=COLOR_BG
            )
        ])

    def dashboard_view():
        ciclo_activo = get_ciclo_activo()
        nombre_ciclo = ciclo_activo['nombre'] if ciclo_activo else "Sin Ciclo Activo"

        search_field = ft.TextField(hint_text="Buscar alumno (Nombre/DNI)...", expand=True, bgcolor="white", border_radius=10)
        
        def go_search(e):
            if search_field.value:
                state["search_term"] = search_field.value
                page.go("/search")

        cursos_col = ft.Column(scroll="auto", expand=True)
        
        def load_cursos():
            cursos_col.controls.clear()
            for c in get_cursos():
                card = ft.Container(
                    content=ft.Row([
                        ft.Icon("book", color=COLOR_PRIMARY),
                        ft.Text(c['nombre'], weight="bold", size=16, expand=True),
                        ft.IconButton("arrow_forward", on_click=lambda e, cid=c['id'], cn=c['nombre']: ir_curso(cid, cn)),
                        ft.IconButton("delete", icon_color="red", on_click=lambda e, cid=c['id']: del_c(cid)) if state["role"] == 'admin' else ft.Container()
                    ], alignment="spaceBetween"),
                    padding=15,
                    bgcolor="white",
                    border_radius=10,
                    shadow=ft.BoxShadow(blur_radius=5, color="black12")
                )
                cursos_col.controls.append(card)
            if not cursos_col.controls:
                cursos_col.controls.append(ft.Text(f"No hay cursos en el Ciclo {nombre_ciclo}.", italic=True))
            page.update()

        def ir_curso(cid, cn):
            state["curso_id"] = cid
            state["curso_nombre"] = cn
            page.go("/curso")

        def del_c(cid):
            delete_curso(cid)
            load_cursos()

        def go_add_curso(e):
            if not ciclo_activo:
                show_snack("Debes crear/activar un ciclo lectivo primero.", "red")
            else:
                page.go("/form_curso")

        load_cursos()

        admin_btn = ft.IconButton("settings", tooltip="Admin", icon_color="white", on_click=lambda _: page.go("/admin")) if state["role"] == 'admin' else ft.Container()

        return ft.View("/dashboard", [
            ft.AppBar(title=ft.Text("Panel Principal"), bgcolor=COLOR_PRIMARY, color="white", 
                      actions=[admin_btn, ft.IconButton("logout", icon_color="white", on_click=lambda _: page.go("/"))]),
            ft.Container(
                content=ft.Column([
                    ft.Container(content=ft.Text(f"Ciclo Lectivo: {nombre_ciclo}", weight="bold", color=COLOR_PRIMARY), padding=5),
                    ft.Container(
                        content=ft.Row([search_field, ft.IconButton("search", on_click=go_search)]),
                        padding=10, bgcolor="white", border_radius=10
                    ),
                    ft.Row([
                        ft.Text("Mis Cursos", size=20, weight="bold"), 
                        ft.IconButton("add_circle", icon_color="green", icon_size=30, on_click=go_add_curso)
                    ], alignment="spaceBetween"),
                    cursos_col
                ]),
                padding=15, expand=True, bgcolor=COLOR_BG
            )
        ])

    def admin_view():
        return ft.View("/admin", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")),
                      title=ft.Text("Administración"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(
                content=ft.Column([
                    ft.Text("Seleccione una opción:", size=18, weight="bold"),
                    ft.Divider(),
                    ft.ListTile(leading=ft.Icon("calendar_month", color=COLOR_PRIMARY), title=ft.Text("Ciclos Lectivos"), subtitle=ft.Text("Crear, cerrar y cambiar años escolares"), on_click=lambda _: page.go("/ciclos")),
                    ft.ListTile(leading=ft.Icon("people", color=COLOR_PRIMARY), title=ft.Text("Usuarios"), subtitle=ft.Text("Gestionar preceptores y admins"), on_click=lambda _: page.go("/users"))
                ]),
                padding=20, bgcolor=COLOR_BG, expand=True
            )
        ])

    def ciclos_view():
        tf_new = ft.TextField(label="Año / Nombre Ciclo", expand=True, bgcolor="white")
        list_col = ft.Column(scroll="auto", expand=True)

        def load():
            list_col.controls.clear()
            ciclos = get_ciclos()
            for c in ciclos:
                is_active = c['activo'] == 1
                icon = "check_circle" if is_active else "circle_outlined"
                color = "green" if is_active else "grey"
                trailing = ft.Container()
                if not is_active:
                    trailing = ft.ElevatedButton("Activar", on_click=lambda e, cid=c['id']: activate(cid))
                else:
                    trailing = ft.Text("ACTIVO", color="green", weight="bold")

                card = ft.Container(content=ft.ListTile(leading=ft.Icon(icon, color=color), title=ft.Text(c['nombre'], weight="bold"), trailing=trailing), bgcolor="white", border_radius=10, margin=2)
                list_col.controls.append(card)
            page.update()

        def add(e):
            if tf_new.value:
                add_ciclo(tf_new.value); tf_new.value = ""; load(); show_snack("Ciclo creado y activado")
        
        def activate(cid):
            activar_ciclo(cid); load(); show_snack("Ciclo cambiado correctamente")

        load()
        return ft.View("/ciclos", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/admin")),
                      title=ft.Text("Gestión Ciclos Lectivos"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(content=ft.Column([ft.Text("Crear Nuevo Ciclo (Cierra el actual)", weight="bold"), ft.Row([tf_new, ft.IconButton("add_circle", icon_color="green", icon_size=40, on_click=add)]), ft.Divider(), ft.Text("Historial de Ciclos", weight="bold"), list_col]), padding=20, bgcolor=COLOR_BG, expand=True)
        ])

    def users_view():
        users_col = ft.Column()
        def load():
            users_col.controls.clear()
            for u in get_users():
                users_col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(u['username']), subtitle=ft.Text(u['role']), trailing=ft.IconButton("delete", icon_color="red", on_click=lambda e, uid=u['id']: rem(uid)) if u['username'] != state['username'] else None), bgcolor="white", border_radius=10, margin=2))
            page.update()
        def add(e): 
            if add_user(u_tf.value, p_tf.value, r_dd.value): u_tf.value=""; p_tf.value=""; load()
        def rem(uid): delete_user(uid); load()
        u_tf = ft.TextField(label="Usuario", expand=True, bgcolor="white"); p_tf = ft.TextField(label="Clave", password=True, expand=True, bgcolor="white")
        r_dd = ft.Dropdown(options=[ft.dropdown.Option("preceptor"), ft.dropdown.Option("admin")], value="preceptor", width=100, bgcolor="white")
        load()
        return ft.View("/users", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/admin")), title=ft.Text("Gestión Usuarios"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(content=ft.Column([ft.Row([u_tf, p_tf, r_dd, ft.IconButton("add", on_click=add)]), ft.Divider(), users_col]), padding=15, bgcolor=COLOR_BG, expand=True)
        ])

    def search_view():
        term = state["search_term"]
        results_col = ft.Column(scroll="auto", expand=True)
        res = search_students(term)
        if not res: results_col.controls.append(ft.Text("No se encontraron resultados."))
        else:
            for r in res:
                card = ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(r['nombre'], weight="bold"), subtitle=ft.Text(f"DNI: {r['dni'] or '-'} | Curso: {r['curso_nombre']}"), on_click=lambda e, s=r: go_detail(s)), bgcolor="white", border_radius=10, margin=ft.margin.only(bottom=5))
                results_col.controls.append(card)
        def go_detail(s): state["student_id_view"] = s['id']; state["curso_id"] = s['curso_id']; page.go("/student_detail")
        return ft.View("/search", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(f"Búsqueda: {term}"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=results_col, padding=15, expand=True, bgcolor=COLOR_BG)])

    def student_detail_view():
        aid = state["student_id_view"]; student = get_alumno_by_id(aid); reqs = get_student_req_status(aid, student['curso_id'])
        req_col = ft.Column()
        for r in reqs: req_col.controls.append(ft.Row([ft.Icon("check_circle" if r['ok'] else "cancel", color="green" if r['ok'] else "red"), ft.Text(r['desc'])]))
        card = ft.Container(content=ft.Column([ft.Text(student['nombre'], size=24, weight="bold"), ft.Text(f"Curso: {student['curso_nombre']}", size=16), ft.Text(f"DNI: {student['dni'] or 'No registrado'}", size=16), ft.Divider(), ft.Text("Observaciones:", weight="bold"), ft.Text(student['observaciones'] or "-", italic=True), ft.Divider(), ft.Text("Documentación:", weight="bold"), req_col]), padding=20, bgcolor="white", border_radius=15, shadow=ft.BoxShadow(blur_radius=10, color="black12"))
        return ft.View("/student_detail", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/search")), title=ft.Text("Ficha de Alumno"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([card], scroll="auto"), padding=20, expand=True, bgcolor=COLOR_BG)])

    def form_curso_view():
        tf = ft.TextField(label="Nombre del Curso", bgcolor="white")
        def save(e): 
            if add_curso(tf.value): page.go("/dashboard")
            else: show_snack("Error", "red")
        return ft.View("/form_curso", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Nuevo Curso"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([tf, ft.ElevatedButton("Guardar", on_click=save)]), padding=20, bgcolor=COLOR_BG, expand=True)])

    def curso_view():
        alumnos_col = ft.Column(scroll="auto", expand=True)
        def load_alumnos():
            alumnos_col.controls.clear()
            for a in get_alumnos(state["curso_id"]):
                alumnos_col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(a['nombre']), subtitle=ft.Text(f"DNI: {a['dni']}"), trailing=ft.PopupMenuButton(items=[ft.PopupMenuItem(text="Editar", on_click=lambda e, aid=a['id']: go_edit(aid)), ft.PopupMenuItem(text="Eliminar", on_click=lambda e, aid=a['id']: del_s(aid))])), bgcolor="white", border_radius=10, margin=2))
            page.update()
        def go_edit(aid): state["student_id_edit"] = aid; page.go("/form_student")
        def go_add(e): state["student_id_edit"] = None; page.go("/form_student")
        def del_s(aid): delete_alumno(aid); load_alumnos()
        load_alumnos()
        return ft.View("/curso", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(state["curso_nombre"]), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([ft.Row([ft.ElevatedButton("Asistencia", icon="check", on_click=lambda _: page.go("/asistencia"), expand=True), ft.ElevatedButton("Pedidos", icon="list", on_click=lambda _: page.go("/pedidos"), expand=True), ft.ElevatedButton("Reportes", icon="bar_chart", on_click=lambda _: page.go("/reportes"), expand=True)]), ft.Divider(), ft.Row([ft.Text("Alumnos", size=18, weight="bold"), ft.IconButton("person_add", icon_color="green", on_click=go_add)], alignment="spaceBetween"), alumnos_col]), padding=15, expand=True, bgcolor=COLOR_BG)])

    def form_student_view():
        is_edit = state["student_id_edit"] is not None
        name = ft.TextField(label="Nombre", bgcolor="white"); dni = ft.TextField(label="DNI", bgcolor="white"); obs = ft.TextField(label="Observaciones", multiline=True, bgcolor="white")
        if is_edit:
            data = get_alumno_by_id(state["student_id_edit"])
            name.value = data['nombre']; dni.value = data['dni']; obs.value = data['observaciones']
        def save(e):
            if name.value:
                if is_edit: update_alumno(state["student_id_edit"], name.value, dni.value, obs.value)
                else: add_alumno(state["curso_id"], name.value, dni.value, obs.value)
                page.go("/curso")
        return ft.View("/form_student", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Alumno"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([name, dni, obs, ft.ElevatedButton("Guardar", on_click=save)]), padding=20, bgcolor=COLOR_BG, expand=True)])

    def asistencia_view():
        date_pk = ft.TextField(label="Fecha (AAAA-MM-DD)", value=date.today().isoformat(), bgcolor="white")
        list_view = ft.Column(scroll="auto", expand=True)
        status_vars = {} 
        def load_list(e=None):
            try:
                if date.fromisoformat(date_pk.value).weekday() >= 5: show_snack("⚠️ Fin de semana", "orange")
            except: pass
            existing = get_asistencia_diaria(state["curso_id"], date_pk.value)
            list_view.controls.clear(); status_vars.clear()
            for a in get_alumnos(state["curso_id"]):
                dd = ft.Dropdown(options=[ft.dropdown.Option(x) for x in ["P", "T", "A", "J", "S", "N"]], value=existing.get(a['id'], "P"), width=80, bgcolor="white")
                status_vars[a['id']] = dd
                list_view.controls.append(ft.Container(content=ft.Row([ft.Text(a['nombre'], expand=True), dd]), padding=5, bgcolor="white", border_radius=5, margin=2))
            page.update()
        def save(e):
            try:
                d = date.fromisoformat(date_pk.value)
                if d > date.today(): show_snack("Futuro no permitido", "red"); return
                if d.weekday() >= 5: show_snack("Fin de semana no permitido", "red"); return
            except: return
            for aid, dd in status_vars.items(): register_asistencia(aid, date_pk.value, dd.value)
            show_snack("Guardado"); page.go("/curso")
        load_list()
        return ft.View("/asistencia", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Asistencia"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([ft.Row([date_pk, ft.IconButton("refresh", on_click=load_list)]), ft.ElevatedButton("GUARDAR", on_click=save, bgcolor="green", color="white", width=float("inf")), ft.Divider(), list_view]), padding=15, bgcolor=COLOR_BG, expand=True)])

    def pedidos_view():
        req_dd = ft.Dropdown(label="Seleccionar Pedido", expand=True, bgcolor="white", on_change=lambda e: load_checks())
        list_view = ft.Column(scroll="auto", expand=True); req_map = {} 
        def load_reqs():
            reqs = get_requisitos(state["curso_id"]); req_map.clear(); req_dd.options.clear()
            for r in reqs: req_map[r['descripcion']] = r['id']; req_dd.options.append(ft.dropdown.Option(r['descripcion']))
            if reqs: req_dd.value = reqs[0]['descripcion']
            page.update(); load_checks()
        def load_checks():
            list_view.controls.clear()
            if not req_dd.value: return
            rid = req_map[req_dd.value]; done = get_cumplimientos(rid)
            for a in get_alumnos(state["curso_id"]):
                def on_change(e, aid=a['id'], rid=rid): toggle_cumplimiento(rid, aid, e.control.value)
                list_view.controls.append(ft.Container(content=ft.Checkbox(label=a['nombre'], value=(a['id'] in done), on_change=on_change), bgcolor="white", padding=5, border_radius=5, margin=2))
            page.update()
        def go_add(e): page.go("/form_requisito")
        def del_r(e): 
            if req_dd.value: delete_requisito(req_map[req_dd.value]); load_reqs()
        load_reqs()
        return ft.View("/pedidos", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Pedidos"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([ft.Row([req_dd, ft.IconButton("add", on_click=go_add), ft.IconButton("delete", icon_color="red", on_click=del_r)]), ft.Divider(), list_view]), padding=15, bgcolor=COLOR_BG, expand=True)])

    def form_requisito_view():
        tf = ft.TextField(label="Descripción", bgcolor="white")
        def save(e):
            if tf.value: add_requisito(state["curso_id"], tf.value); page.go("/pedidos")
        return ft.View("/form_requisito", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/pedidos")), title=ft.Text("Nuevo Pedido"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(content=ft.Column([tf, ft.ElevatedButton("Crear", on_click=save)]), padding=20, bgcolor=COLOR_BG, expand=True)])

    def reportes_view():
        d1 = ft.TextField(label="Desde", value=date.today().replace(day=1).isoformat(), width=120, bgcolor="white"); d2 = ft.TextField(label="Hasta", value=date.today().isoformat(), width=120, bgcolor="white")
        table_cont = ft.Column(scroll="auto", expand=True)
        
        def gen(e):
            data = get_report_data(state["curso_id"], d1.value, d2.value); rows = []
            for d in data:
                rows.append(ft.DataRow(cells=[ft.DataCell(ft.Text(d['nombre'])), ft.DataCell(ft.Text(str(d['p']))), ft.DataCell(ft.Text(str(d['t']))), ft.DataCell(ft.Text(str(d['a']))), ft.DataCell(ft.Text(str(d['j']))), ft.DataCell(ft.Text(str(d['s']))), ft.DataCell(ft.Text(f"{d['faltas']}", color="red" if d['pct']>15 else "black")), ft.DataCell(ft.Text(f"{d['pct']}%", color="red" if d['pct']>25 else "black"))]))
            table = ft.DataTable(columns=[ft.DataColumn(ft.Text("Alumno")), ft.DataColumn(ft.Text("P"), numeric=True), ft.DataColumn(ft.Text("T"), numeric=True), ft.DataColumn(ft.Text("A"), numeric=True), ft.DataColumn(ft.Text("J"), numeric=True), ft.DataColumn(ft.Text("S"), numeric=True), ft.DataColumn(ft.Text("Faltas"), numeric=True), ft.DataColumn(ft.Text("% Aus."), numeric=True)], rows=rows, bgcolor="white", border_radius=10, column_spacing=20)
            table_cont.controls = [ft.Row([table], scroll="always")]; page.update()
        
        def export(e):
            print("Iniciando exportación...") # Log en consola
            if not pd: 
                show_snack("Error: 'pandas' no instalado", "red")
                return
            try:
                data = get_report_data(state["curso_id"], d1.value, d2.value)
                if not data:
                    show_snack("No hay datos para exportar.", "orange")
                    return
                df = pd.DataFrame(data)
                df = df.rename(columns={'nombre': 'Alumno', 'dni': 'DNI', 'p': 'Presentes', 't': 'Tardes', 'a': 'Ausentes', 'j': 'Justificadas', 's': 'Suspensiones', 'faltas': 'Total Faltas', 'pct': '% Ausentismo'})
                
                fname = f"reporte_curso_{state['curso_id']}_{date.today()}.xlsx"
                df.to_excel(fname, index=False)
                print(f"Archivo guardado: {fname}") # Log en consola
                show_snack(f"Archivo guardado localmente: {fname}", "green")
            except ImportError as ie:
                print(f"Error import: {ie}")
                show_snack("Falta 'openpyxl'. Ejecuta 'pip install openpyxl'", "red")
            except Exception as ex:
                print(f"Error general: {ex}")
                show_snack(f"Error al exportar: {ex}", "red")

        return ft.View("/reportes", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Reportes"), bgcolor=COLOR_PRIMARY, color="white"), ft.Container(content=ft.Column([ft.Row([d1, d2, ft.ElevatedButton("Ver", on_click=gen)]), ft.ElevatedButton("Exportar Excel", icon="download", on_click=export), ft.Divider(), table_cont]), padding=15, bgcolor=COLOR_BG, expand=True)])

    # --- ROUTER ---
    def route_change(route):
        page.views.clear()
        routes = {
            "/": login_view,
            "/dashboard": dashboard_view,
            "/admin": admin_view,
            "/ciclos": ciclos_view,
            "/users": users_view,
            "/search": search_view,
            "/student_detail": student_detail_view,
            "/form_curso": form_curso_view,
            "/curso": curso_view,
            "/form_student": form_student_view,
            "/asistencia": asistencia_view,
            "/pedidos": pedidos_view,
            "/form_requisito": form_requisito_view,
            "/reportes": reportes_view
        }
        if page.route in routes: page.views.append(routes[page.route]())
        else: page.views.append(login_view())
        page.update()

    def view_pop(view):
        page.views.pop(); top_view = page.views[-1]; page.go(top_view.route)

    page.on_route_change = route_change; page.on_view_pop = view_pop; page.go("/")

if __name__ == "__main__":
    # Obtener el puerto del entorno (necesario para Render/Heroku)
    port = int(os.environ.get("PORT", 8000))
    ft.app(target=main, view=ft.WEB_BROWSER, port=port)