import flet as ft
import sqlite3
import hashlib
from datetime import date
import os
import base64
import io

# --- LIBRERÍAS OPCIONALES ---
try:
    import pandas as pd
except ImportError:
    pd = None
    print("⚠️ Pandas no instalado. La exportación a Excel estará deshabilitada.")

try:
    import xlsxwriter
except ImportError:
    print("⚠️ XlsxWriter no instalado. La exportación a Excel estará deshabilitada.")

# ======================================================================
# 1. BASE DE DATOS (SQLite Local)
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
    
    cursor.execute("CREATE TABLE IF NOT EXISTS Usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, password TEXT NOT NULL, role TEXT NOT NULL)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Ciclos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE, activo INTEGER DEFAULT 0)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Cursos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, ciclo_id INTEGER, FOREIGN KEY (ciclo_id) REFERENCES Ciclos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Alumnos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, nombre TEXT NOT NULL, dni TEXT, observaciones TEXT, tutor_nombre TEXT, tutor_telefono TEXT, UNIQUE(curso_id, nombre), FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Asistencia (id INTEGER PRIMARY KEY AUTOINCREMENT, alumno_id INTEGER NOT NULL, fecha TEXT NOT NULL, status TEXT NOT NULL, UNIQUE(alumno_id, fecha), FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Requisitos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, descripcion TEXT NOT NULL, FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)")
    cursor.execute("CREATE TABLE IF NOT EXISTS Requisitos_Cumplidos (requisito_id INTEGER NOT NULL, alumno_id INTEGER NOT NULL, PRIMARY KEY (requisito_id, alumno_id), FOREIGN KEY (requisito_id) REFERENCES Requisitos(id) ON DELETE CASCADE, FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)")

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

# ======================================================================
# FUNCIONES CRUD
# ======================================================================

def authenticate_user(username, password):
    conn = get_db_connection()
    pwd = hash_password(password)
    user = conn.execute("SELECT * FROM Usuarios WHERE username = ? AND password = ?", (username, pwd)).fetchone()
    conn.close()
    if user: return True, user['role']
    return False, None

def get_ciclo_activo():
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM Ciclos WHERE activo = 1").fetchone()
    conn.close()
    return dict(row) if row else None

def get_curso_by_id(cid):
    conn = get_db_connection()
    row = conn.execute("SELECT c.*, ci.nombre as ciclo_nombre FROM Cursos c JOIN Ciclos ci ON c.ciclo_id = ci.id WHERE c.id = ?", (cid,)).fetchone()
    conn.close()
    return dict(row) if row else None

def get_cursos():
    ciclo = get_ciclo_activo()
    if not ciclo: return []
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Cursos WHERE ciclo_id = ? ORDER BY nombre", (ciclo['id'],)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_curso(nombre):
    ciclo = get_ciclo_activo()
    if not ciclo: return False
    try:
        conn = get_db_connection()
        conn.execute("INSERT INTO Cursos (nombre, ciclo_id) VALUES (?, ?)", (nombre, ciclo['id']))
        conn.commit(); conn.close()
        return True
    except: return False

def delete_curso(cid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Cursos WHERE id = ?", (cid,))
    conn.commit(); conn.close()

def get_alumnos(curso_id):
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Alumnos WHERE curso_id = ? ORDER BY nombre", (curso_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_alumno(cid, nombre, dni, obs, t_n, t_t):
    try:
        conn = get_db_connection()
        conn.execute("INSERT INTO Alumnos (curso_id, nombre, dni, observaciones, tutor_nombre, tutor_telefono) VALUES (?,?,?,?,?,?)", (cid, nombre, dni, obs, t_n, t_t))
        conn.commit(); conn.close()
        return True
    except: return False

def update_alumno(aid, nombre, dni, obs, t_n, t_t):
    conn = get_db_connection()
    conn.execute("UPDATE Alumnos SET nombre=?, dni=?, observaciones=?, tutor_nombre=?, tutor_telefono=? WHERE id=?", (nombre, dni, obs, t_n, t_t, aid))
    conn.commit(); conn.close()

def delete_alumno(aid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Alumnos WHERE id=?", (aid,))
    conn.commit(); conn.close()

def get_alumno_by_id(aid):
    conn = get_db_connection()
    row = conn.execute("SELECT * FROM Alumnos WHERE id = ?", (aid,)).fetchone()
    conn.close()
    return dict(row) if row else None

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

def get_asistencia_diaria(curso_id, fecha):
    conn = get_db_connection()
    rows = conn.execute("SELECT alumno_id, status FROM Asistencia WHERE fecha = ? AND alumno_id IN (SELECT id FROM Alumnos WHERE curso_id=?)", (fecha, curso_id)).fetchall()
    conn.close()
    return {r['alumno_id']: r['status'] for r in rows}

def register_asistencia(aid, cid, fecha, status):
    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO Asistencia (alumno_id, fecha, status) VALUES (?, ?, ?)", (aid, fecha, status))
    conn.commit(); conn.close()

def get_student_attendance_history(aid):
    conn = get_db_connection()
    rows = conn.execute("SELECT fecha, status FROM Asistencia WHERE alumno_id = ? ORDER BY fecha DESC", (aid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_student_stats(aid):
    conn = get_db_connection()
    rows = conn.execute("SELECT status FROM Asistencia WHERE alumno_id = ?", (aid,)).fetchall()
    conn.close()
    statuses = [r['status'] for r in rows]
    counts = {k: statuses.count(k) for k in ['P','T','A','J','S','N']}
    faltas = counts['A'] + counts['S'] + (counts['T'] * 0.25)
    total = counts['P'] + counts['T'] + counts['A'] + counts['J'] + counts['S']
    pct = (faltas/total*100) if total > 0 else 0
    return {
        'presentes': counts['P'], 'tardes': counts['T'], 'ausentes': counts['A'],
        'justificadas': counts['J'], 'suspensiones': counts['S'],
        'total_faltas': faltas, 'porcentaje': round(pct, 1),
        'total_registros': total
    }

def get_report_data(curso_id, start, end):
    conn = get_db_connection()
    alumnos = conn.execute("SELECT * FROM Alumnos WHERE curso_id=?", (curso_id,)).fetchall()
    asistencias = conn.execute("SELECT alumno_id, status FROM Asistencia WHERE fecha >= ? AND fecha <= ? AND alumno_id IN (SELECT id FROM Alumnos WHERE curso_id=?)", (start, end, curso_id)).fetchall()
    conn.close()
    
    asis_map = {} 
    for r in asistencias:
        aid = r['alumno_id']
        if aid not in asis_map: asis_map[aid] = []
        asis_map[aid].append(r['status'])
        
    report = []
    for a in alumnos:
        statuses = asis_map.get(a['id'], [])
        counts = {k: statuses.count(k) for k in ['P','T','A','J','S','N']}
        faltas = counts['A'] + counts['S'] + (counts['T'] * 0.25)
        total = counts['P'] + counts['T'] + counts['A'] + counts['J'] + counts['S']
        pct = (faltas/total*100) if total > 0 else 0
        report.append({
            'nombre': a['nombre'], 'dni': a['dni'], 
            'p': counts['P'], 't': counts['T'], 'a': counts['A'], 
            'j': counts['J'], 's': counts['S'], 
            'faltas': faltas, 'pct': round(pct, 1)
        })
    return report

def get_users():
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Usuarios").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_user(u, p, r):
    try:
        conn = get_db_connection()
        conn.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", (u, hash_password(p), r))
        conn.commit(); conn.close()
        return True
    except: return False

def delete_user(uid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Usuarios WHERE id = ?", (uid,))
    conn.commit(); conn.close()

def get_ciclos():
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Ciclos ORDER BY nombre DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_ciclo(nombre):
    try:
        conn = get_db_connection()
        conn.execute("UPDATE Ciclos SET activo = 0")
        conn.execute("INSERT INTO Ciclos (nombre, activo) VALUES (?, 1)", (nombre,))
        conn.commit(); conn.close()
        return True
    except: return False

def activar_ciclo(cid):
    conn = get_db_connection()
    conn.execute("UPDATE Ciclos SET activo = 0")
    conn.execute("UPDATE Ciclos SET activo = 1 WHERE id = ?", (cid,))
    conn.commit(); conn.close()

def get_requisitos(cid):
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Requisitos WHERE curso_id=?", (cid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_requisito(cid, desc):
    conn = get_db_connection()
    conn.execute("INSERT INTO Requisitos (curso_id, descripcion) VALUES (?, ?)", (cid, desc))
    conn.commit(); conn.close()

def delete_requisito(rid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Requisitos WHERE id=?", (rid,))
    conn.commit(); conn.close()

def get_cumplimientos(rid):
    conn = get_db_connection()
    rows = conn.execute("SELECT alumno_id FROM Requisitos_Cumplidos WHERE requisito_id=?", (rid,)).fetchall()
    conn.close()
    return {r['alumno_id'] for r in rows}

def toggle_cumplimiento(rid, aid, val):
    conn = get_db_connection()
    if val: conn.execute("INSERT OR IGNORE INTO Requisitos_Cumplidos (requisito_id, alumno_id) VALUES (?, ?)", (rid, aid))
    else: conn.execute("DELETE FROM Requisitos_Cumplidos WHERE requisito_id=? AND alumno_id=?", (rid, aid))
    conn.commit(); conn.close()

def get_student_req_status(aid, cid):
    reqs = get_requisitos(cid)
    res = []
    for r in reqs:
        done = get_cumplimientos(r['id'])
        res.append({'id': r['id'], 'desc': r['descripcion'], 'ok': aid in done})
    return res

# ======================================================================
# 2. INTERFAZ GRÁFICA (Flet - Diseño Premium)
# ======================================================================

def main(page: ft.Page):
    page.title = "Sistema de Asistencia UNSAM"
    page.theme_mode = "light"
    page.padding = 0
    
    # Colores Hex
    PRIMARY = "#3F51B5"
    SECONDARY = "#1A237E"
    BG_COLOR = "#F0F0F0"
    CARD_COLOR = "#FFFFFF"
    DANGER = "#E53935"
    SUCCESS = "#43A047"
    
    init_db()
    
    state = {
        "role": None, "username": None, 
        "curso_id": None, "curso_nombre": None, 
        "search": "", "st_view": None, "st_edit": None
    }

    def show_snack(m, c=SUCCESS):
        page.snack_bar = ft.SnackBar(ft.Text(m), bgcolor=c)
        page.snack_bar.open = True
        page.update()

    def create_card(content, padding=15, on_click=None):
        return ft.Container(
            content=content, padding=padding, bgcolor=CARD_COLOR, border_radius=8,
            shadow=ft.BoxShadow(blur_radius=5, color="#00000030", offset=ft.Offset(0, 2)),
            margin=ft.margin.only(bottom=10), on_click=on_click
        )

    # --- VISTAS ---

    def login_view():
        user = ft.TextField(label="Usuario", width=300, bgcolor="white", border_radius=8, border_color=PRIMARY)
        pwd = ft.TextField(label="Clave", password=True, width=300, bgcolor="white", border_radius=8, border_color=PRIMARY)
        
        def login(e):
            ok, role = authenticate_user(user.value, pwd.value)
            if ok:
                state["role"], state["username"] = role, user.value
                page.go("/dashboard")
            else:
                show_snack("Datos incorrectos", DANGER)
        
        return ft.View("/", [
            ft.Container(
                content=ft.Column([
                    ft.Icon("school", size=80, color=PRIMARY),
                    ft.Text("Sistema de Asistencia", size=28, weight="bold", color=SECONDARY),
                    ft.Text("UNSAM", size=18, color="grey"),
                    ft.Divider(height=30, color="transparent"),
                    ft.Container(
                        content=ft.Column([user, ft.Container(height=10), pwd, ft.Container(height=20), ft.ElevatedButton("INGRESAR", on_click=login, width=300, height=50, bgcolor=PRIMARY, color="white")]),
                        padding=40, bgcolor="white", border_radius=20,
                        shadow=ft.BoxShadow(blur_radius=20, color="#0000001A")
                    ),
                    ft.Container(height=20),
                    ft.Text("Admin Default: admin / admin", size=12, color="grey")
                ], horizontal_alignment="center"),
                alignment=ft.alignment.center, expand=True, bgcolor=BG_COLOR
            )
        ])

    def dashboard_view():
        ciclo = get_ciclo_activo()
        c_nombre = ciclo['nombre'] if ciclo else "Sin Ciclo Activo"
        search = ft.TextField(hint_text="Buscar alumno...", expand=True, bgcolor="white", border_radius=20, border_color="transparent")
        
        def do_search(e): 
            if search.value: state["search"] = search.value; page.go("/search")
        search.on_submit = do_search
        
        cursos_col = ft.Column(scroll="auto", expand=True)
        
        def load():
            cursos_col.controls.clear()
            cursos = get_cursos()
            if not cursos: cursos_col.controls.append(ft.Text("No hay cursos activos.", italic=True, color="grey"))
            
            for c in cursos:
                def create_click(cid, cn): return lambda e: go_curso(cid, cn)
                def create_del(cid): return lambda e: (delete_curso(cid), load())

                action_row = ft.Row([
                    ft.IconButton("arrow_forward", icon_color=PRIMARY, on_click=create_click(c['id'], c['nombre'])),
                ])
                if state["role"] == 'admin':
                    action_row.controls.append(ft.IconButton("delete", icon_color=DANGER, on_click=create_del(c['id'])))

                cursos_col.controls.append(create_card(
                    content=ft.Row([
                        ft.Row([ft.Container(content=ft.Icon("class", color="white"), bgcolor=PRIMARY, border_radius=10, padding=10), ft.Text(c['nombre'], weight="bold", size=18, color=SECONDARY)]),
                        action_row
                    ], alignment="spaceBetween")
                ))
            page.update()
        
        def go_curso(cid, cn): state["curso_id"]=cid; state["curso_nombre"]=cn; page.go("/curso")
        def add_c(e): 
            if ciclo: page.go("/form_curso")
            else: show_snack("Falta Ciclo Activo", DANGER)

        load()
        admin_btn = ft.IconButton("settings", icon_color="white", on_click=lambda _: page.go("/admin")) if state["role"] == 'admin' else ft.Container()
        
        return ft.View("/dashboard", [
            ft.AppBar(title=ft.Text("Panel Principal"), bgcolor=PRIMARY, color="white", center_title=True, actions=[admin_btn, ft.IconButton("logout", icon_color="white", on_click=lambda _: page.go("/"))]),
            ft.Container(content=ft.Column([
                ft.Container(content=ft.Row([ft.Text(f"Ciclo: {c_nombre}", color=PRIMARY, weight="bold"), ft.Container(content=search, width=300)], alignment="spaceBetween"), padding=ft.padding.only(bottom=20)),
                ft.Row([ft.Text("Mis Cursos", size=24, weight="bold", color=SECONDARY), ft.ElevatedButton("Nuevo Curso", icon="add", bgcolor=SUCCESS, color="white", on_click=add_c)], alignment="spaceBetween"),
                ft.Container(height=10), cursos_col
            ]), padding=30, bgcolor=BG_COLOR, expand=True)
        ])

    def curso_view():
        col = ft.Column(scroll="auto", expand=True)
        def load():
            col.controls.clear()
            alumnos = get_alumnos(state["curso_id"])
            if not alumnos: col.controls.append(ft.Text("No hay alumnos.", italic=True, color="grey"))
            for a in alumnos:
                def go_det(aid, cid): state["st_view"] = aid; state["curso_id"] = cid; page.go("/student_detail")
                def edit_clk(aid): return lambda e: (state.update({"st_edit": aid}), page.go("/form_student"))
                def del_clk(aid): return lambda e: (delete_alumno(aid), load())
                col.controls.append(create_card(content=ft.ListTile(leading=ft.CircleAvatar(content=ft.Text(a['nombre'][0]), bgcolor="#E3F2FD", color=PRIMARY), title=ft.Text(a['nombre'], weight="bold"), subtitle=ft.Text(f"DNI: {a.get('dni','-')}"), on_click=lambda e, s=a: go_det(s['id'], state["curso_id"]), trailing=ft.PopupMenuButton(icon="more_vert", items=[ft.PopupMenuItem("Editar", icon="edit", on_click=edit_clk(a['id'])), ft.PopupMenuItem("Borrar", icon="delete", on_click=del_clk(a['id']))])), padding=0))
            page.update()
        load()
        return ft.View("/curso", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(state["curso_nombre"]), bgcolor=PRIMARY, color="white", center_title=True),
            ft.Container(content=ft.Column([
                ft.Container(content=ft.Row([
                    ft.ElevatedButton("Asistencia", icon="check_circle", height=50, on_click=lambda _: page.go("/asistencia"), bgcolor="#3949AB", color="white", expand=True),
                    ft.ElevatedButton("Pedidos", icon="assignment", height=50, on_click=lambda _: page.go("/pedidos"), bgcolor="#F57C00", color="white", expand=True),
                    ft.ElevatedButton("Reportes", icon="bar_chart", height=50, on_click=lambda _: page.go("/reportes"), bgcolor="#00897B", color="white", expand=True)
                ], spacing=10), padding=ft.padding.only(bottom=20)),
                ft.Row([ft.Text("Alumnos", size=22, weight="bold", color=SECONDARY), ft.IconButton("person_add", icon_color="white", bgcolor=SUCCESS, on_click=lambda _: (state.update({"st_edit": None}), page.go("/form_student")))], alignment="spaceBetween"),
                ft.Container(height=10), col
            ]), padding=20, bgcolor=BG_COLOR, expand=True)
        ])

    def asistencia_view():
        dp = ft.TextField(label="Fecha (AAAA-MM-DD)", value=date.today().isoformat(), bgcolor="white", border_radius=10)
        col = ft.Column(scroll="auto", expand=True); vals = {}
        def load(e=None):
            try: 
                if date.fromisoformat(dp.value).weekday() >= 5: show_snack("⚠️ Es fin de semana.", "orange")
            except: show_snack("Fecha inválida.", DANGER); return
            ex = get_asistencia_diaria(state["curso_id"], dp.value); col.controls.clear(); vals.clear()
            for a in get_alumnos(state["curso_id"]):
                dd = ft.Dropdown(options=[ft.dropdown.Option(x) for x in ["P","T","A","J","S","N"]], value=ex.get(a['id'], "P"), width=80, bgcolor="white", border_radius=8, content_padding=10)
                vals[a['id']] = dd
                col.controls.append(create_card(content=ft.Row([ft.Text(a['nombre'], weight="bold", size=16, expand=True), dd], alignment="spaceBetween"), padding=10))
            page.update()
        def save(e):
            try:
                d = date.fromisoformat(dp.value)
                if d > date.today(): return show_snack("Fecha futura", DANGER)
                if d.weekday() >= 5: return show_snack("Es fin de semana", DANGER)
            except: return show_snack("Fecha inválida", DANGER)
            for aid, dd in vals.items(): register_asistencia(aid, state["curso_id"], dp.value, dd.value)
            show_snack("Guardado"); page.go("/curso")
        load()
        return ft.View("/asistencia", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Tomar Asistencia"), bgcolor=PRIMARY, color="white"),
            ft.Container(content=ft.Column([create_card(ft.Row([dp, ft.IconButton("refresh", on_click=load, icon_color=PRIMARY)], alignment="center")), ft.ElevatedButton("GUARDAR CAMBIOS", on_click=save, bgcolor=SUCCESS, color="white", height=50, width=float("inf")), ft.Container(height=10), col]), padding=20, bgcolor=BG_COLOR, expand=True)
        ])

    def reportes_view():
        d1 = ft.TextField(label="Desde", value=date.today().replace(day=1).isoformat(), width=130, bgcolor="white", border_radius=8)
        d2 = ft.TextField(label="Hasta", value=date.today().isoformat(), width=130, bgcolor="white", border_radius=8)
        table_cont = ft.Column(scroll="auto", expand=True)
        def gen(e):
            data = get_report_data(state["curso_id"], d1.value, d2.value); rows = []
            for d in data:
                c = DANGER if d['faltas']>=25 else "black"
                rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(d['nombre'], color=c, weight="bold")), ft.DataCell(ft.Text(str(d['p']))), ft.DataCell(ft.Text(str(d['t']))), 
                    ft.DataCell(ft.Text(str(d['a']))), ft.DataCell(ft.Text(str(d['j']))), ft.DataCell(ft.Text(str(d['s']))),
                    ft.DataCell(ft.Container(content=ft.Text(f"{d['faltas']}", color="white", weight="bold"), bgcolor=c if c==DANGER else "transparent", padding=5, border_radius=5)),
                    ft.DataCell(ft.Text(f"{d['pct']}%", color=c, weight="bold"))
                ]))
            dt = ft.DataTable(columns=[ft.DataColumn(ft.Text("Alumno")), ft.DataColumn(ft.Text("P"), numeric=True), ft.DataColumn(ft.Text("T"), numeric=True), ft.DataColumn(ft.Text("A"), numeric=True), ft.DataColumn(ft.Text("J"), numeric=True), ft.DataColumn(ft.Text("S"), numeric=True), ft.DataColumn(ft.Text("Faltas"), numeric=True), ft.DataColumn(ft.Text("% Aus."), numeric=True)], rows=rows, bgcolor="white", border_radius=10, column_spacing=15, heading_row_color="#E3F2FD", heading_row_height=40)
            table_cont.controls = [create_card(ft.Row([dt], scroll="always"), padding=0)]; page.update()
        
        def export(e):
            if not pd: return show_snack("Falta pandas", DANGER)
            data = get_report_data(state["curso_id"], d1.value, d2.value)
            if not data: return show_snack("Sin datos", "orange")
            df = pd.DataFrame(data).rename(columns={'nombre':'Alumno', 'p':'Pres', 't':'Tarde', 'a':'Aus', 'j':'Just', 's':'Susp', 'faltas':'Total', 'pct':'%'})
            output = io.BytesIO(); df.to_excel(output, index=False, engine='xlsxwriter'); b64 = base64.b64encode(output.getvalue()).decode()
            page.launch_url(f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}", web_window_name="reporte.xlsx")

        return ft.View("/reportes", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Reportes"), bgcolor=PRIMARY, color="white"),
            ft.Container(content=ft.Column([create_card(ft.Row([d1, d2, ft.ElevatedButton("VER TABLA", on_click=gen, bgcolor=PRIMARY, color="white", height=45)], alignment="center")), ft.ElevatedButton("DESCARGAR EXCEL", icon="download", on_click=export, bgcolor="green", color="white", width=float("inf"), height=45), ft.Container(height=10), table_cont]), padding=20, bgcolor=BG_COLOR, expand=True)
        ])

    def search_view():
        term = state["search"]; res = search_students(term); col = ft.Column(scroll="auto")
        if not res: col.controls.append(ft.Text("Sin resultados", color="grey", size=16))
        else:
            for r in res:
                def go_det(s): state["st_view"] = s['id']; state["curso_id"] = s['curso_id']; page.go("/student_detail")
                col.controls.append(create_card(content=ft.ListTile(leading=ft.Icon("person", color=PRIMARY, size=30), title=ft.Text(r['nombre'], weight="bold"), subtitle=ft.Text(f"Curso: {r['curso_nombre']} ({r['ciclo_nombre']})"), on_click=lambda e, s=r: go_det(s), trailing=ft.Icon("chevron_right", color="grey"))))
        return ft.View("/search", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(f"Búsqueda: {term}"), bgcolor=PRIMARY, color="white"), ft.Container(content=col, padding=20, bgcolor=BG_COLOR, expand=True)])

    def export_student_ficha(page, alumno, curso_data, stats, requisitos):
        if not pd: return show_snack("Error: 'pandas' no instalado.", DANGER)
        try:
            ficha_data = [["ALUMNO", alumno.get('nombre','')], ["DNI", alumno.get('dni','')], ["CURSO", curso_data.get('nombre','')], ["CICLO", curso_data.get('ciclo_nombre','')], ["TUTOR", alumno.get('tutor_nombre','')], ["TEL", alumno.get('tutor_telefono','')]]
            stats_data = [["Faltas (Eq.)", stats['total_faltas']], ["% Ausentismo", f"{stats['porcentaje']}%"], ["Presentes", stats['presentes']], ["Ausentes", stats['ausentes']], ["Justificadas", stats['justificadas']], ["Suspensiones", stats['suspensiones']]]
            req_data = [["Requisito", "Cumplido"]] + [[r['desc'], 'SÍ' if r['ok'] else 'NO'] for r in requisitos]
            
            history = get_student_attendance_history(alumno['id'])
            df_hist = pd.DataFrame(history).rename(columns={'fecha':'Fecha', 'status':'Estado'}).drop(columns=['status'], errors='ignore') if history else pd.DataFrame([["Sin registros"]], columns=['Info'])
            
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            pd.DataFrame(ficha_data, columns=['Campo', 'Valor']).to_excel(writer, sheet_name='Ficha', index=False)
            pd.DataFrame(stats_data, columns=['Concepto', 'Valor']).to_excel(writer, sheet_name='Estadísticas', index=False)
            pd.DataFrame(req_data[1:], columns=req_data[0]).to_excel(writer, sheet_name='Documentación', index=False)
            df_hist.to_excel(writer, sheet_name='Historial', index=False)
            writer.close(); output.seek(0)
            
            b64 = base64.b64encode(output.getvalue()).decode()
            fname = f"ficha_{alumno['nombre'].replace(' ', '_')}.xlsx"
            page.launch_url(f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}", web_window_name=fname)
            show_snack("Ficha descargada.", SUCCESS)
        except Exception as e: show_snack(f"Error: {e}", DANGER)

    def student_detail_view():
        aid = state["st_view"]; s = get_alumno_by_id(aid)
        if not s: return ft.View("/error", [ft.Text("Error: Alumno no encontrado")])
        curso_data = get_curso_by_id(s['curso_id'])
        stats = get_student_stats(aid)
        reqs = get_student_req_status(aid, s['curso_id'])
        
        def stat_box(l, v, c=SECONDARY): return ft.Container(content=ft.Column([ft.Text(v, size=20, weight="bold", color=c), ft.Text(l, size=12, color="grey")], horizontal_alignment="center"), padding=10, bgcolor="white", border_radius=5, expand=True, alignment=ft.alignment.center)
        stat_row = ft.Row([stat_box("Faltas", str(stats['total_faltas']), "red" if stats['total_faltas']>20 else SECONDARY), stat_box("% Aus.", f"{stats['porcentaje']}%"), stat_box("Pres.", str(stats['presentes']), "green")], spacing=10)
        req_col = ft.Column()
        for r in reqs: req_col.controls.append(ft.Row([ft.Icon("check_circle" if r['ok'] else "cancel", color=SUCCESS if r['ok'] else DANGER), ft.Text(r['desc'])]))
        
        card = create_card(content=ft.Column([
            ft.Row([ft.Icon("person_pin", size=50, color=PRIMARY), ft.Column([ft.Text(s['nombre'], size=24, weight="bold"), ft.Text(f"DNI: {s.get('dni','-')}", color="grey")])]),
            ft.Divider(), ft.Text("Estadísticas", weight="bold", color=PRIMARY), stat_row,
            ft.Divider(), ft.Text("Contacto", weight="bold", color=SECONDARY), ft.ListTile(leading=ft.Icon("phone"), title=ft.Text(f"{s.get('tutor_nombre','-')}"), subtitle=ft.Text(f"{s.get('tutor_telefono','-')}")),
            ft.Text("Obs", weight="bold", color=SECONDARY), ft.Container(content=ft.Text(s.get('observaciones','-'), italic=True), padding=10, bgcolor="#F5F5F5", border_radius=5, width=float("inf")),
            ft.Container(height=10), ft.Text("Papeles", weight="bold", color=SECONDARY), req_col,
            ft.Container(height=20), ft.ElevatedButton("DESCARGAR FICHA EXCEL", icon="download", on_click=lambda e: export_student_ficha(page, s, curso_data, stats, reqs), bgcolor="#00897B", color="white", width=float("inf"), disabled=(pd is None))
        ]), padding=25)
        return ft.View("/student_detail", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/search")), title=ft.Text("Ficha"), bgcolor=PRIMARY, color="white"), ft.Container(content=ft.Column([card], scroll="auto"), padding=20, bgcolor=BG_COLOR, expand=True)])

    def pedidos_view():
        dd = ft.Dropdown(label="Pedido", expand=True, bgcolor="white", on_change=lambda e: lc(), border_radius=8)
        col = ft.Column(scroll="auto", expand=True); rm = {}
        def lr():
            rs = get_requisitos(state["curso_id"]); rm.clear(); dd.options.clear()
            for r in rs: rm[r['descripcion']] = r['id']; dd.options.append(ft.dropdown.Option(r['descripcion']))
            if rs: dd.value = rs[0]['descripcion']
            page.update(); lc()
        def lc():
            col.controls.clear()
            if not dd.value: return
            rid = rm[dd.value]; done = get_cumplimientos(rid)
            for a in get_alumnos(state["curso_id"]):
                def on_chg(e, aid=a['id'], rid=rid): toggle_cumplimiento(rid, aid, e.control.value)
                col.controls.append(create_card(content=ft.Checkbox(label=a['nombre'], value=(a['id'] in done), on_change=lambda e, aid=a['id'], rid=rid: on_chg(e, aid, rid)), padding=10))
            page.update()
        def add(e): page.go("/form_req")
        def dele(e): 
            if dd.value: delete_requisito(rm[dd.value]); lr(); show_snack("Eliminado", DANGER)
        lr()
        return ft.View("/pedidos", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Documentación"), bgcolor=PRIMARY, color="white"), ft.Container(content=ft.Column([create_card(ft.Row([dd, ft.IconButton("add", on_click=add, icon_color=PRIMARY), ft.IconButton("delete", icon_color=DANGER, on_click=dele)])), ft.Divider(color="transparent"), ft.Text("Marcar entregas:", weight="bold"), col]), padding=20, bgcolor=BG_COLOR, expand=True)])

    def form_student_view():
        is_edit = state["st_edit"] is not None
        nm = ft.TextField(label="Nombre", bgcolor="white", border_radius=8); dni = ft.TextField(label="DNI", bgcolor="white", border_radius=8); obs = ft.TextField(label="Obs", multiline=True, bgcolor="white", border_radius=8); tn = ft.TextField(label="Tutor", bgcolor="white", border_radius=8); tt = ft.TextField(label="Tel Tutor", bgcolor="white", border_radius=8)
        if is_edit:
            d = get_alumno_by_id(state["st_edit"]); nm.value=d.get('nombre',''); dni.value=d.get('dni',''); obs.value=d.get('observaciones',''); tn.value=d.get('tutor_nombre',''); tt.value=d.get('tutor_telefono','')
        def save(e):
            if nm.value:
                if is_edit: update_alumno(state["st_edit"], nm.value, dni.value, obs.value, tn.value, tt.value)
                else: add_alumno(state["curso_id"], nm.value, dni.value, obs.value, tn.value, tt.value)
                page.go("/curso")
        return ft.View("/form_student", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Ficha del Alumno"), bgcolor=PRIMARY, color="white"), ft.Container(content=create_card(ft.Column([ft.Text("Datos Alumno", size=18, weight="bold"), nm, dni, ft.Text("Datos Tutor", size=18, weight="bold"), tn, tt, ft.Text("Observaciones", size=18, weight="bold"), obs, ft.ElevatedButton("GUARDAR", on_click=save, bgcolor=SUCCESS, color="white", height=45, width=float("inf"))])), padding=20, bgcolor=BG_COLOR, expand=True)])

    def form_curso_view():
        tf = ft.TextField(label="Nombre Curso", bgcolor="white", border_radius=8)
        def save(e): 
            if add_curso(tf.value): page.go("/dashboard")
            else: show_snack("Error", DANGER)
        return ft.View("/form_curso", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Nuevo Curso"), bgcolor=PRIMARY, color="white"), ft.Container(content=create_card(ft.Column([ft.Text("Nombre del curso:", color="grey"), tf, ft.ElevatedButton("CREAR", on_click=save, bgcolor=SUCCESS, color="white", height=45, width=float("inf"))])), padding=20, bgcolor=BG_COLOR, expand=True)])

    def form_req_view():
        tf = ft.TextField(label="Descripción", bgcolor="white", border_radius=8)
        def save(e):
            if tf.value: add_requisito(state["curso_id"], tf.value); page.go("/pedidos")
        return ft.View("/form_req", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/pedidos")), title=ft.Text("Nuevo Requisito"), bgcolor=PRIMARY, color="white"), ft.Container(content=create_card(ft.Column([ft.Text("Documento a solicitar:", color="grey"), tf, ft.ElevatedButton("CREAR", on_click=save, bgcolor=SUCCESS, color="white", height=45, width=float("inf"))])), padding=20, bgcolor=BG_COLOR, expand=True)])

    def admin_view():
        if state["role"]!='admin': return ft.View("/admin", [ft.AppBar(title=ft.Text("Error"), bgcolor=DANGER), ft.Text("Acceso Denegado")])
        return ft.View("/admin", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Admin"), bgcolor=PRIMARY, color="white"), ft.Container(content=ft.Column([create_card(ft.ListTile(leading=ft.Icon("calendar_month", color=PRIMARY), title=ft.Text("Ciclos Lectivos"), on_click=lambda _: page.go("/ciclos"))), create_card(ft.ListTile(leading=ft.Icon("people", color=PRIMARY), title=ft.Text("Usuarios"), on_click=lambda _: page.go("/users")))]), padding=20, bgcolor=BG_COLOR, expand=True)])

    def ciclos_view():
        tf = ft.TextField(label="Año", expand=True, bgcolor="white", border_radius=8); col = ft.Column(scroll="auto", expand=True)
        def ld():
            col.controls.clear()
            for c in get_ciclos():
                act = c['activo']==1
                tr = ft.ElevatedButton("Activar", on_click=lambda e, cid=c['id']: (activar_ciclo(cid), ld()), bgcolor="orange", color="white") if not act else ft.Text("ACTIVO", color="green", weight="bold")
                col.controls.append(create_card(ft.ListTile(leading=ft.Icon("check_circle" if act else "circle", color="green" if act else "grey"), title=ft.Text(c['nombre'], weight="bold"), trailing=tr), padding=0))
            page.update()
        def add(e): 
            if tf.value: add_ciclo(tf.value); tf.value=""; ld()
        ld()
        return ft.View("/ciclos", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/admin")), title=ft.Text("Ciclos"), bgcolor=PRIMARY, color="white"), ft.Container(content=ft.Column([create_card(ft.Row([tf, ft.IconButton("add", on_click=add)])), ft.Container(height=20), col]), padding=20, bgcolor=BG_COLOR, expand=True)])

    def users_view():
        u = ft.TextField(label="User", expand=True, bgcolor="white", border_radius=8); p = ft.TextField(label="Pass", password=True, expand=True, bgcolor="white", border_radius=8); r = ft.Dropdown(options=[ft.dropdown.Option("preceptor"), ft.dropdown.Option("admin")], value="preceptor", width=100, bgcolor="white", border_radius=8); col = ft.Column()
        def ld():
            col.controls.clear()
            for us in get_users():
                tr = ft.IconButton("delete", icon_color=DANGER, on_click=lambda e, uid=us['id']: (delete_user(uid), ld())) if us['username']!=state['username'] else None
                col.controls.append(create_card(ft.ListTile(leading=ft.Icon("person", color=PRIMARY), title=ft.Text(us['username']), subtitle=ft.Text(us['role']), trailing=tr), padding=0))
            page.update()
        def add(e): 
            if add_user(u.value, p.value, r.value): u.value=""; p.value=""; ld()
        ld()
        return ft.View("/users", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/admin")), title=ft.Text("Usuarios"), bgcolor=PRIMARY, color="white"), ft.Container(content=ft.Column([create_card(ft.Row([u, p, r, ft.IconButton("add_circle", icon_color="green", icon_size=40, on_click=add)])), ft.Container(height=20), col]), padding=20, bgcolor=BG_COLOR, expand=True)])

    # --- ROUTER ---
    def router(route):
        page.views.clear()
        views = {
            "/": login_view, "/dashboard": dashboard_view, "/curso": curso_view, "/asistencia": asistencia_view,
            "/pedidos": pedidos_view, "/form_req": form_req_view, "/reportes": reportes_view,
            "/search": search_view, "/student_detail": student_detail_view, "/form_student": form_student_view,
            "/form_curso": form_curso_view, "/admin": admin_view, "/ciclos": ciclos_view, "/users": users_view
        }
        if state["role"] is None and page.route != "/": page.route = "/"
        if page.route in views: page.views.append(views[page.route]())
        else: page.views.append(login_view())
        page.update()

    page.on_route_change = router
    page.on_view_pop = lambda view: page.go(page.views[-2].route)
    page.go("/")

if __name__ == "__main__":
    # CONFIGURACIÓN DE ARRANQUE ROBUSTA
    port_env = os.environ.get("PORT")
    
    if port_env:
        # NUBE: Usar el puerto del entorno y host 0.0.0.0
        # IMPORTANTE: No usar 'view=' para modo headless
        print(f"🚀 Iniciando en NUBE (Puerto {port_env})")
        ft.app(target=main, port=int(port_env), host="0.0.0.0", web_renderer="html")
    else:
        # LOCAL: Usar puerto 8550 y localhost
        print("🏠 Iniciando en LOCAL (localhost:8550)")
        ft.app(target=main, view=ft.WEB_BROWSER, port=8550, host="127.0.0.1")
