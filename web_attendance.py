import flet as ft
import sqlite3
import hashlib
from datetime import date, datetime
import os
import base64
import io
import threading

# --- IMPORTACIÓN DE LIBRERÍAS EXTERNAS ---
try:
    import pandas as pd
except ImportError:
    pd = None
    print("⚠️ Pandas no instalado.")

try:
    import xlsxwriter
except ImportError:
    print("⚠️ XlsxWriter no instalado.")

# ======================================================================
# CAPA 1: UTILIDADES Y VALIDACIONES
# ======================================================================

class Validator:
    @staticmethod
    def is_weekend(date_str: str) -> bool:
        """Devuelve True si la fecha es Sábado o Domingo."""
        try:
            d = date.fromisoformat(date_str)
            return d.weekday() >= 5
        except ValueError:
            return False

    @staticmethod
    def is_future_date(date_str: str) -> bool:
        """Devuelve True si la fecha es posterior a hoy."""
        try:
            d = date.fromisoformat(date_str)
            return d > date.today()
        except ValueError:
            return False

    @staticmethod
    def is_valid_text(text: str, min_len: int = 1) -> bool:
        return text is not None and len(text.strip()) >= min_len

class Security:
    @staticmethod
    def hash_password(password: str) -> str:
        # En un entorno real productivo, se recomienda usar bcrypt.
        # Aquí usamos SHA256 por compatibilidad con librerías estándar.
        return hashlib.sha256(password.encode()).hexdigest()

# ======================================================================
# CAPA 2: GESTIÓN DE BASE DE DATOS (Database Manager)
# ======================================================================

class DatabaseManager:
    def __init__(self, db_name='asistencia_alumnos.db'):
        self.db_name = db_name
        self.lock = threading.Lock() # Para seguridad en hilos
        self._init_db()

    def get_connection(self):
        conn = sqlite3.connect(self.db_name, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn

    def _init_db(self):
        with self.lock:
            conn = self.get_connection()
            cursor = conn.cursor()
            
            # Definición de Esquema
            queries = [
                "CREATE TABLE IF NOT EXISTS Usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, password TEXT NOT NULL, role TEXT NOT NULL)",
                "CREATE TABLE IF NOT EXISTS Ciclos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL UNIQUE, activo INTEGER DEFAULT 0)",
                "CREATE TABLE IF NOT EXISTS Cursos (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, ciclo_id INTEGER, FOREIGN KEY (ciclo_id) REFERENCES Ciclos(id) ON DELETE CASCADE)",
                "CREATE TABLE IF NOT EXISTS Alumnos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, nombre TEXT NOT NULL, dni TEXT, observaciones TEXT, tutor_nombre TEXT, tutor_telefono TEXT, UNIQUE(curso_id, nombre), FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)",
                "CREATE TABLE IF NOT EXISTS Asistencia (id INTEGER PRIMARY KEY AUTOINCREMENT, alumno_id INTEGER NOT NULL, fecha TEXT NOT NULL, status TEXT NOT NULL, UNIQUE(alumno_id, fecha), FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)",
                "CREATE TABLE IF NOT EXISTS Requisitos (id INTEGER PRIMARY KEY AUTOINCREMENT, curso_id INTEGER NOT NULL, descripcion TEXT NOT NULL, FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE)",
                "CREATE TABLE IF NOT EXISTS Requisitos_Cumplidos (requisito_id INTEGER NOT NULL, alumno_id INTEGER NOT NULL, PRIMARY KEY (requisito_id, alumno_id), FOREIGN KEY (requisito_id) REFERENCES Requisitos(id) ON DELETE CASCADE, FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE)"
            ]
            
            for q in queries:
                cursor.execute(q)

            # Migraciones (Columnas nuevas)
            for col in ["dni", "observaciones", "tutor_nombre", "tutor_telefono"]:
                try: cursor.execute(f"ALTER TABLE Alumnos ADD COLUMN {col} TEXT")
                except: pass

            # Datos Semilla (Seed Data)
            cursor.execute("SELECT COUNT(*) FROM Usuarios")
            if cursor.fetchone()[0] == 0:
                cursor.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", ("admin", Security.hash_password("admin"), "admin"))
            
            cursor.execute("SELECT COUNT(*) FROM Ciclos")
            if cursor.fetchone()[0] == 0:
                anio = str(date.today().year)
                cursor.execute("INSERT INTO Ciclos (nombre, activo) VALUES (?, 1)", (anio,))
                cid = cursor.lastrowid
                cursor.execute("UPDATE Cursos SET ciclo_id = ? WHERE ciclo_id IS NULL", (cid,))

            conn.commit()
            conn.close()

    # --- Métodos CRUD Genéricos ---
    
    def fetch_all(self, query, params=()):
        with self.lock:
            conn = self.get_connection()
            try:
                cursor = conn.execute(query, params)
                return [dict(row) for row in cursor.fetchall()]
            finally:
                conn.close()

    def fetch_one(self, query, params=()):
        with self.lock:
            conn = self.get_connection()
            try:
                cursor = conn.execute(query, params)
                row = cursor.fetchone()
                return dict(row) if row else None
            finally:
                conn.close()

    def execute_query(self, query, params=()):
        with self.lock:
            conn = self.get_connection()
            try:
                conn.execute(query, params)
                conn.commit()
                return True
            except sqlite3.Error as e:
                print(f"DB Error: {e}")
                return False
            finally:
                conn.close()

    # --- Métodos Específicos de Negocio ---

    def authenticate(self, username, password):
        user = self.fetch_one("SELECT * FROM Usuarios WHERE username = ?", (username,))
        if user and user['password'] == Security.hash_password(password):
            return user
        return None

    def get_ciclo_activo(self):
        return self.fetch_one("SELECT * FROM Ciclos WHERE activo = 1")

    def get_cursos_activos(self):
        ciclo = self.get_ciclo_activo()
        if not ciclo: return []
        return self.fetch_all("SELECT * FROM Cursos WHERE ciclo_id = ? ORDER BY nombre", (ciclo['id'],))

    def get_alumnos_curso(self, curso_id):
        return self.fetch_all("SELECT * FROM Alumnos WHERE curso_id = ? ORDER BY nombre", (curso_id,))

    def get_asistencia_fecha(self, curso_id, fecha):
        rows = self.fetch_all("SELECT alumno_id, status FROM Asistencia WHERE fecha = ? AND alumno_id IN (SELECT id FROM Alumnos WHERE curso_id=?)", (fecha, curso_id))
        return {row['alumno_id']: row['status'] for row in rows}

    def registrar_asistencia(self, alumno_id, fecha, status):
        return self.execute_query("INSERT OR REPLACE INTO Asistencia (alumno_id, fecha, status) VALUES (?, ?, ?)", (alumno_id, fecha, status))

    def get_reporte_curso(self, curso_id, start_date, end_date):
        alumnos = self.get_alumnos_curso(curso_id)
        # Optimización: Una sola query para todas las asistencias del rango
        asistencias = self.fetch_all("""
            SELECT alumno_id, status 
            FROM Asistencia 
            WHERE fecha >= ? AND fecha <= ? 
            AND alumno_id IN (SELECT id FROM Alumnos WHERE curso_id=?)
        """, (start_date, end_date, curso_id))

        # Procesamiento en memoria (más rápido que N queries)
        asis_map = {}
        for r in asistencias:
            if r['alumno_id'] not in asis_map: asis_map[r['alumno_id']] = []
            asis_map[r['alumno_id']].append(r['status'])

        reporte = []
        for a in alumnos:
            statuses = asis_map.get(a['id'], [])
            counts = {k: statuses.count(k) for k in ['P','T','A','J','S','N']}
            faltas = counts['A'] + counts['S'] + (counts['T'] * 0.25)
            total = sum(counts[k] for k in ['P','T','A','J','S'])
            pct = (faltas / total * 100) if total > 0 else 0
            
            # Datos extendidos para la ficha
            reporte.append({
                'id': a['id'],
                'nombre': a['nombre'], 
                'dni': a.get('dni', '-'),
                'tutor_nombre': a.get('tutor_nombre', '-'),
                'tutor_telefono': a.get('tutor_telefono', '-'),
                'observaciones': a.get('observaciones', ''),
                'p': counts['P'], 't': counts['T'], 'a': counts['A'], 
                'j': counts['J'], 's': counts['S'], 
                'faltas': faltas, 'pct': round(pct, 1),
                'total_registros': total
            })
        return reporte
    
    def get_historial_alumno(self, alumno_id):
        return self.fetch_all("SELECT fecha, status FROM Asistencia WHERE alumno_id = ? ORDER BY fecha DESC", (alumno_id,))

    def search_alumnos(self, term):
        term = f"%{term}%"
        return self.fetch_all("""
            SELECT a.*, c.nombre as curso_nombre, ci.nombre as ciclo_nombre 
            FROM Alumnos a 
            JOIN Cursos c ON a.curso_id = c.id 
            JOIN Ciclos ci ON c.ciclo_id = ci.id
            WHERE (a.nombre LIKE ? OR a.dni LIKE ?) AND ci.activo = 1
            ORDER BY a.nombre
        """, (term, term))

    # Métodos Requisitos
    def get_requisitos_estado(self, alumno_id, curso_id):
        reqs = self.fetch_all("SELECT * FROM Requisitos WHERE curso_id = ?", (curso_id,))
        cumplidos_raw = self.fetch_all("SELECT requisito_id FROM Requisitos_Cumplidos WHERE alumno_id = ?", (alumno_id,))
        cumplidos_ids = {r['requisito_id'] for r in cumplidos_raw}
        
        result = []
        for r in reqs:
            result.append({
                'id': r['id'],
                'desc': r['descripcion'],
                'ok': r['id'] in cumplidos_ids
            })
        return result

# Instancia Global del Gestor de BD
db = DatabaseManager()

# ======================================================================
# CAPA 3: INTERFAZ DE USUARIO (Vistas y Componentes)
# ======================================================================

# --- Configuración de Estilo ---
THEME = {
    "primary": "#3F51B5",
    "secondary": "#1A237E",
    "bg": "#F5F7FB",
    "card": "#FFFFFF",
    "danger": "#E53935",
    "success": "#43A047",
    "warning": "#FB8C00"
}

def create_card(content, padding=15, on_click=None):
    return ft.Container(
        content=content, 
        padding=padding, 
        bgcolor=THEME["card"], 
        border_radius=8,
        shadow=ft.BoxShadow(blur_radius=5, color="#00000030", offset=ft.Offset(0, 2)),
        margin=ft.margin.only(bottom=10), 
        on_click=on_click,
        animate=ft.animation.Animation(200, "easeOut")
    )

def show_snack(page, message, color=THEME["success"]):
    page.snack_bar = ft.SnackBar(ft.Text(message), bgcolor=color)
    page.snack_bar.open = True
    page.update()

# --- VISTA: LOGIN ---
def view_login(page: ft.Page):
    user_input = ft.TextField(label="Usuario", width=300, bgcolor="white", border_radius=8, prefix_icon="person")
    pass_input = ft.TextField(label="Contraseña", password=True, width=300, bgcolor="white", border_radius=8, prefix_icon="lock", can_reveal_password=True)

    def login_action(e):
        user = db.authenticate(user_input.value, pass_input.value)
        if user:
            # Guardar sesión en el estado de la página (no global)
            page.session.set("user", user)
            page.go("/dashboard")
        else:
            show_snack(page, "Credenciales incorrectas", THEME["danger"])

    content = ft.Container(
        content=ft.Column([
            ft.Icon("school", size=80, color=THEME["primary"]),
            ft.Text("Sistema de Asistencia", size=28, weight="bold", color=THEME["secondary"]),
            ft.Text("UNSAM", size=18, color="grey"),
            ft.Divider(height=30, color="transparent"),
            ft.Container(
                content=ft.Column([
                    user_input, ft.Container(height=10), pass_input, ft.Container(height=20),
                    ft.ElevatedButton("INGRESAR", on_click=login_action, width=300, height=50, bgcolor=THEME["primary"], color="white")
                ], horizontal_alignment="center"),
                padding=40, bgcolor="white", border_radius=20,
                shadow=ft.BoxShadow(blur_radius=20, color="#0000001A")
            )
        ], horizontal_alignment="center"),
        alignment=ft.alignment.center, expand=True, bgcolor=THEME["bg"]
    )
    return ft.View("/", [content])

# --- VISTA: DASHBOARD ---
def view_dashboard(page: ft.Page):
    user = page.session.get("user")
    if not user: return view_login(page) # Protección

    ciclo = db.get_ciclo_activo()
    ciclo_txt = ciclo['nombre'] if ciclo else "Sin Ciclo Activo"
    
    search_input = ft.TextField(hint_text="Buscar alumno...", expand=True, bgcolor="white", border_radius=20, border_color="transparent")
    
    def search_action(e):
        if search_input.value:
            page.session.set("search_term", search_input.value)
            page.go("/search")

    search_input.on_submit = search_action

    cursos_grid = ft.Column(scroll="auto", expand=True)

    def load_cursos():
        cursos_grid.controls.clear()
        cursos = db.get_cursos_activos()
        if not cursos:
            cursos_grid.controls.append(ft.Text("No hay cursos activos o creados.", italic=True, color="grey"))
        
        for c in cursos:
            # Closure para capturar variables en loop
            def on_click_curso(e, cid=c['id'], cname=c['nombre']):
                page.session.set("curso_id", cid)
                page.session.set("curso_nombre", cname)
                page.go("/curso")
            
            def on_delete_curso(e, cid=c['id']):
                if db.execute_query("DELETE FROM Cursos WHERE id=?", (cid,)):
                    load_cursos()
                    page.update()

            actions_row = [ft.IconButton("arrow_forward", icon_color=THEME["primary"], on_click=on_click_curso)]
            if user['role'] == 'admin':
                actions_row.append(ft.IconButton("delete", icon_color=THEME["danger"], on_click=on_delete_curso))

            card = create_card(ft.Row([
                ft.Row([
                    ft.Container(content=ft.Icon("class", color="white"), bgcolor=THEME["primary"], border_radius=10, padding=10),
                    ft.Text(c['nombre'], weight="bold", size=18, color=THEME["secondary"])
                ]),
                ft.Row(actions_row)
            ], alignment="spaceBetween"))
            cursos_grid.controls.append(card)
        page.update()

    load_cursos()

    header_actions = [ft.IconButton("logout", icon_color="white", on_click=lambda _: page.go("/"))]
    if user['role'] == 'admin':
        header_actions.insert(0, ft.IconButton("settings", icon_color="white", on_click=lambda _: page.go("/admin")))

    return ft.View("/dashboard", [
        ft.AppBar(title=ft.Text("Panel Principal"), bgcolor=THEME["primary"], color="white", center_title=True, actions=header_actions),
        ft.Container(content=ft.Column([
            ft.Container(content=ft.Row([
                ft.Text(f"Ciclo: {ciclo_txt}", color=THEME["primary"], weight="bold"),
                ft.Container(content=search_input, width=300)
            ], alignment="spaceBetween"), padding=ft.padding.only(bottom=20)),
            ft.Row([
                ft.Text("Mis Cursos", size=24, weight="bold", color=THEME["secondary"]),
                ft.ElevatedButton("Nuevo Curso", icon="add", bgcolor=THEME["success"], color="white", 
                                  on_click=lambda _: page.go("/form_curso") if ciclo else show_snack(page, "Falta ciclo activo", THEME["danger"]))
            ], alignment="spaceBetween"),
            ft.Container(height=10),
            cursos_grid
        ]), padding=30, bgcolor=THEME["bg"], expand=True)
    ])

# --- VISTA: DETALLE DE CURSO ---
def view_curso(page: ft.Page):
    curso_id = page.session.get("curso_id")
    curso_nombre = page.session.get("curso_nombre")
    if not curso_id: return view_dashboard(page)
    user_role = page.session.get("user")['role']

    alumnos_list = ft.Column(scroll="auto", expand=True)

    def load_alumnos():
        alumnos_list.controls.clear()
        alumnos = db.get_alumnos_curso(curso_id)
        
        if not alumnos:
            alumnos_list.controls.append(ft.Text("No hay alumnos matriculados.", italic=True, color="grey"))
        
        for a in alumnos:
            def on_detail(e, aid=a['id']):
                page.session.set("alumno_id", aid)
                page.go("/student_detail")
            
            def on_edit(e, aid=a['id']):
                page.session.set("alumno_id_edit", aid)
                page.go("/form_student")
            
            def on_delete(e, aid=a['id']):
                db.delete_alumno(aid)
                load_alumnos()
                page.update()

            menu_items = [ft.PopupMenuItem("Editar", icon="edit", on_click=on_edit)]
            if user_role == 'admin':
                menu_items.append(ft.PopupMenuItem("Borrar", icon="delete", on_click=on_delete))

            card = create_card(ft.ListTile(
                leading=ft.CircleAvatar(content=ft.Text(a['nombre'][0]), bgcolor="#E3F2FD", color=THEME["primary"]),
                title=ft.Text(a['nombre'], weight="bold"),
                subtitle=ft.Text(f"DNI: {a['dni'] or '-'}"),
                on_click=on_detail,
                trailing=ft.PopupMenuButton(icon="more_vert", items=menu_items)
            ), padding=0)
            alumnos_list.controls.append(card)
        page.update()

    load_alumnos()

    return ft.View("/curso", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), 
                  title=ft.Text(curso_nombre), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=ft.Column([
            ft.Row([
                ft.ElevatedButton("Asistencia", icon="check_circle", on_click=lambda _: page.go("/asistencia"), bgcolor=THEME["primary"], color="white", expand=True),
                ft.ElevatedButton("Documentos", icon="assignment", on_click=lambda _: page.go("/pedidos"), bgcolor=THEME["warning"], color="white", expand=True),
                ft.ElevatedButton("Reportes", icon="bar_chart", on_click=lambda _: page.go("/reportes"), bgcolor="#00897B", color="white", expand=True)
            ]),
            ft.Divider(),
            ft.Row([
                ft.Text("Alumnos", size=20, weight="bold", color=THEME["secondary"]),
                ft.IconButton("person_add", icon_color="white", bgcolor=THEME["success"], 
                              on_click=lambda _: (page.session.set("alumno_id_edit", None), page.go("/form_student")))
            ], alignment="spaceBetween"),
            alumnos_list
        ]), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

# --- VISTA: TOMAR ASISTENCIA ---
def view_asistencia(page: ft.Page):
    curso_id = page.session.get("curso_id")
    date_input = ft.TextField(label="Fecha", value=date.today().isoformat(), bgcolor="white", border_radius=10)
    list_col = ft.Column(scroll="auto", expand=True)
    inputs_map = {}

    def load_status(e=None):
        fecha = date_input.value
        # Validación
        if Validator.is_future_date(fecha):
            show_snack(page, "No se puede registrar asistencia futura", THEME["danger"])
            return
        if Validator.is_weekend(fecha):
            show_snack(page, "Advertencia: Es fin de semana", THEME["warning"])

        saved_data = db.get_asistencia_fecha(curso_id, fecha)
        alumnos = db.get_alumnos_curso(curso_id)
        
        list_col.controls.clear()
        inputs_map.clear()
        
        for a in alumnos:
            status = saved_data.get(a['id'], "P")
            dd = ft.Dropdown(
                options=[ft.dropdown.Option(x) for x in ["P","T","A","J","S","N"]],
                value=status, width=100, bgcolor="white", border_radius=8
            )
            inputs_map[a['id']] = dd
            list_col.controls.append(create_card(
                ft.Row([ft.Text(a['nombre'], weight="bold", size=16), dd], alignment="spaceBetween"), 
                padding=10
            ))
        page.update()

    def save_all(e):
        fecha = date_input.value
        if Validator.is_future_date(fecha):
             show_snack(page, "Error: Fecha futura", THEME["danger"])
             return
             
        count = 0
        for aid, dd in inputs_map.items():
            db.registrar_asistencia(aid, fecha, dd.value)
            count += 1
        show_snack(page, f"Guardados {count} registros.")
        page.go("/curso")

    load_status()

    return ft.View("/asistencia", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), 
                  title=ft.Text("Tomar Asistencia"), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=ft.Column([
            create_card(ft.Row([date_input, ft.IconButton("refresh", icon_color=THEME["primary"], on_click=load_status)])),
            ft.ElevatedButton("GUARDAR TODO", on_click=save_all, bgcolor=THEME["success"], color="white", height=50, width=float("inf")),
            ft.Container(height=10),
            list_col
        ]), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

# --- VISTA: REPORTES Y EXPORTACIÓN ---
def view_reportes(page: ft.Page):
    curso_id = page.session.get("curso_id")
    # Fechas por defecto: Inicio de año hasta hoy
    d_start = ft.TextField(label="Desde", value=date.today().replace(month=1, day=1).isoformat(), width=150, bgcolor="white")
    d_end = ft.TextField(label="Hasta", value=date.today().isoformat(), width=150, bgcolor="white")
    table_container = ft.Column(scroll="auto", expand=True)

    def generate_report(e=None):
        data = db.get_reporte_curso(curso_id, d_start.value, d_end.value)
        rows = []
        for d in data:
            color = THEME["danger"] if d['faltas'] >= 25 else ("black" if d['faltas'] < 15 else THEME["warning"])
            rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(d['nombre'], color=color, weight="bold")),
                ft.DataCell(ft.Text(str(d['p']))),
                ft.DataCell(ft.Text(str(d['t']))),
                ft.DataCell(ft.Text(str(d['a']))),
                ft.DataCell(ft.Text(str(d['j']))),
                ft.DataCell(ft.Text(str(d['s']))),
                ft.DataCell(ft.Container(content=ft.Text(f"{d['faltas']}", color="white", weight="bold"), bgcolor=color if color != "black" else "grey", padding=5, border_radius=5)),
                ft.DataCell(ft.Text(f"{d['pct']}%"))
            ]))
        
        dt = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Alumno")),
                ft.DataColumn(ft.Text("P"), numeric=True),
                ft.DataColumn(ft.Text("T"), numeric=True),
                ft.DataColumn(ft.Text("A"), numeric=True),
                ft.DataColumn(ft.Text("J"), numeric=True),
                ft.DataColumn(ft.Text("S"), numeric=True),
                ft.DataColumn(ft.Text("Faltas"), numeric=True),
                ft.DataColumn(ft.Text("%"), numeric=True),
            ],
            rows=rows, bgcolor="white", border_radius=10, column_spacing=15, heading_row_color="#E3F2FD"
        )
        table_container.controls = [create_card(ft.Row([dt], scroll="always"), padding=0)]
        page.update()

    def export_excel(e):
        if not pd or not xlsxwriter:
            show_snack(page, "Librerías de Excel no instaladas", THEME["danger"])
            return
        
        data = db.get_reporte_curso(curso_id, d_start.value, d_end.value)
        if not data:
            show_snack(page, "Sin datos para exportar", THEME["warning"])
            return

        # Preparar DataFrame
        df = pd.DataFrame(data)
        # Limpiar columnas internas
        df = df.drop(columns=['id', 'tutor_nombre', 'tutor_telefono', 'observaciones', 'total_registros'], errors='ignore')
        df = df.rename(columns={'nombre':'Alumno', 'dni':'DNI', 'p':'Pres.', 't':'Tardes', 'a':'Aus.', 'j':'Just.', 's':'Susp.', 'faltas':'Total Faltas', 'pct':'% Ausentismo'})

        # Generar Buffer
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        
        b64 = base64.b64encode(output.read()).decode()
        filename = f"reporte_curso_{curso_id}.xlsx"
        page.launch_url(f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}", web_window_name=filename)
        show_snack(page, "Descarga iniciada", THEME["success"])

    return ft.View("/reportes", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Reportes"), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=ft.Column([
            create_card(ft.Row([d_start, d_end, ft.ElevatedButton("VER", on_click=generate_report, bgcolor=THEME["primary"], color="white")], alignment="center")),
            ft.ElevatedButton("DESCARGAR EXCEL", icon="download", bgcolor=THEME["success"], color="white", width=float("inf"), on_click=export_excel),
            ft.Container(height=10),
            table_container
        ]), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

# --- VISTA: DETALLE ALUMNO ---
def view_student_detail(page: ft.Page):
    aid = page.session.get("alumno_id")
    curso_id = page.session.get("curso_id")
    if not aid: return view_dashboard(page)
    
    # Obtener datos usando un solo rango amplio para estadísticas anuales
    report_data = db.get_reporte_curso(curso_id, "2000-01-01", "2100-12-31")
    # Filtrar el alumno (ineficiente si son miles, pero ok para escuela)
    stats = next((s for s in report_data if s['id'] == aid), None)
    student_info = db.fetch_one("SELECT * FROM Alumnos WHERE id=?", (aid,))
    
    # Obtener Requisitos
    reqs = db.get_requisitos_estado(aid, curso_id)

    # Widget de Estadísticas
    def stat_box(label, val, color="black"):
        return ft.Container(
            content=ft.Column([
                ft.Text(str(val), size=22, weight="bold", color=color),
                ft.Text(label, size=12, color="grey")
            ], horizontal_alignment="center"),
            padding=10, bgcolor="white", border_radius=8, expand=True, alignment=ft.alignment.center,
            border=ft.border.all(1, "#EEEEEE")
        )
    
    stat_row = ft.Row([
        stat_box("Faltas", stats['faltas'], THEME["danger"] if stats['faltas'] > 20 else "black"),
        stat_box("Ausentismo", f"{stats['pct']}%"),
        stat_box("Presentes", stats['p'], THEME["success"])
    ], spacing=10)

    # Widget de Requisitos
    req_list = ft.Column()
    for r in reqs:
        icon = "check_circle" if r['ok'] else "cancel"
        color = THEME["success"] if r['ok'] else THEME["danger"]
        req_list.controls.append(ft.Row([ft.Icon(icon, color=color), ft.Text(r['desc'])]))

    def export_ficha(e):
        if not pd: return show_snack(page, "Falta pandas", THEME["danger"])
        
        # Generar Excel Multitabla
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        
        # Hoja 1: Datos
        data_ficha = [
            ["Nombre", student_info['nombre']], ["DNI", student_info['dni']],
            ["Tutor", student_info['tutor_nombre']], ["Teléfono", student_info['tutor_telefono']],
            ["Obs", student_info['observaciones']]
        ]
        pd.DataFrame(data_ficha, columns=["Campo", "Valor"]).to_excel(writer, sheet_name="Ficha", index=False)
        
        # Hoja 2: Stats
        pd.DataFrame([stats]).to_excel(writer, sheet_name="Estadisticas", index=False)
        
        # Hoja 3: Historial
        hist = db.get_historial_alumno(aid)
        pd.DataFrame([dict(h) for h in hist]).to_excel(writer, sheet_name="Historial", index=False)
        
        writer.close()
        output.seek(0)
        b64 = base64.b64encode(output.read()).decode()
        page.launch_url(f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}", web_window_name=f"ficha_{aid}.xlsx")

    content = create_card(ft.Column([
        ft.Row([
            ft.Icon("person", size=50, color=THEME["primary"]),
            ft.Column([
                ft.Text(student_info['nombre'], size=24, weight="bold"),
                ft.Text(f"DNI: {student_info['dni'] or '-'}", color="grey")
            ])
        ]),
        ft.Divider(),
        ft.Text("Estadísticas Anuales", weight="bold", color=THEME["primary"]),
        stat_row,
        ft.Divider(),
        ft.Text("Información de Contacto", weight="bold"),
        ft.Text(f"Tutor: {student_info['tutor_nombre'] or '-'} | Tel: {student_info['tutor_telefono'] or '-'}"),
        ft.Text("Observaciones:", weight="bold", size=12),
        ft.Text(student_info['observaciones'] or "-", italic=True),
        ft.Divider(),
        ft.Text("Documentación", weight="bold"),
        req_list,
        ft.Container(height=20),
        ft.ElevatedButton("DESCARGAR FICHA COMPLETA", icon="download", bgcolor="#00897B", color="white", width=float("inf"), on_click=export_ficha)
    ]), padding=25)

    return ft.View("/student_detail", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Ficha Alumno"), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=ft.Column([content], scroll="auto"), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

# --- VISTA: FORMULARIO ALUMNO ---
def view_form_student(page: ft.Page):
    curso_id = page.session.get("curso_id")
    aid_edit = page.session.get("alumno_id_edit")
    is_edit = aid_edit is not None
    
    nm = ft.TextField(label="Nombre Completo", bgcolor="white")
    dni = ft.TextField(label="DNI", bgcolor="white")
    tn = ft.TextField(label="Nombre Tutor", bgcolor="white")
    tt = ft.TextField(label="Teléfono Tutor", bgcolor="white")
    obs = ft.TextField(label="Observaciones", multiline=True, bgcolor="white")

    if is_edit:
        d = db.fetch_one("SELECT * FROM Alumnos WHERE id=?", (aid_edit,))
        if d:
            nm.value = d['nombre']; dni.value = d['dni']; obs.value = d['observaciones']
            tn.value = d['tutor_nombre']; tt.value = d['tutor_telefono']

    def save(e):
        if not nm.value:
            show_snack(page, "Nombre obligatorio", THEME["danger"])
            return
        
        if is_edit:
            db.execute_query("UPDATE Alumnos SET nombre=?, dni=?, observaciones=?, tutor_nombre=?, tutor_telefono=? WHERE id=?", (nm.value, dni.value, obs.value, tn.value, tt.value, aid_edit))
        else:
            if not db.execute_query("INSERT INTO Alumnos (curso_id, nombre, dni, observaciones, tutor_nombre, tutor_telefono) VALUES (?,?,?,?,?,?)", (curso_id, nm.value, dni.value, obs.value, tn.value, tt.value)):
                show_snack(page, "Error: Nombre duplicado", THEME["danger"])
                return
        
        show_snack(page, "Guardado correctamente")
        page.go("/curso")

    return ft.View("/form_student", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Editar Alumno" if is_edit else "Nuevo Alumno"), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=create_card(ft.Column([
            nm, dni, ft.Divider(), tn, tt, ft.Divider(), obs,
            ft.ElevatedButton("GUARDAR", on_click=save, bgcolor=THEME["success"], color="white", width=float("inf"))
        ])), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

# --- OTRAS VISTAS (Simplificadas para completar el flujo) ---
def view_form_curso(page: ft.Page):
    tf = ft.TextField(label="Nombre Curso", bgcolor="white")
    def save(e):
        ciclo = db.get_ciclo_activo()
        if not ciclo: return show_snack(page, "No hay ciclo activo", THEME["danger"])
        if db.execute_query("INSERT INTO Cursos (nombre, ciclo_id) VALUES (?, ?)", (tf.value, ciclo['id'])):
            page.go("/dashboard")
        else: show_snack(page, "Error al crear", THEME["danger"])
    return ft.View("/form_curso", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Nuevo Curso"), bgcolor=THEME["primary"], color="white"), ft.Container(content=create_card(ft.Column([tf, ft.ElevatedButton("Crear", on_click=save, bgcolor=THEME["success"], color="white")])), padding=20, bgcolor=THEME["bg"], expand=True)])

def view_pedidos(page: ft.Page):
    # Lógica de pedidos
    curso_id = page.session.get("curso_id")
    req_dd = ft.Dropdown(label="Requisito", expand=True, bgcolor="white")
    list_col = ft.Column(scroll="auto", expand=True)
    
    def load_checks(e=None):
        list_col.controls.clear()
        if not req_dd.value: return
        rid = int(req_dd.value)
        cumplidos = {r['alumno_id'] for r in db.fetch_all("SELECT alumno_id FROM Requisitos_Cumplidos WHERE requisito_id=?", (rid,))}
        
        for a in db.get_alumnos_curso(curso_id):
            def on_chg(e, aid=a['id'], rid=rid):
                if e.control.value: db.execute_query("INSERT OR IGNORE INTO Requisitos_Cumplidos (requisito_id, alumno_id) VALUES (?, ?)", (rid, aid))
                else: db.execute_query("DELETE FROM Requisitos_Cumplidos WHERE requisito_id=? AND alumno_id=?", (rid, aid))
            
            list_col.controls.append(create_card(ft.Checkbox(label=a['nombre'], value=(a['id'] in cumplidos), on_change=on_chg), padding=10))
        page.update()

    def load_dd():
        reqs = db.fetch_all("SELECT * FROM Requisitos WHERE curso_id=?", (curso_id,))
        req_dd.options = [ft.dropdown.Option(key=str(r['id']), text=r['descripcion']) for r in reqs]
        if reqs: 
            req_dd.value = str(reqs[0]['id'])
            load_checks()
        page.update()

    def add_req(e): page.go("/form_req")
    
    load_dd()
    return ft.View("/pedidos", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Documentación"), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=ft.Column([
            create_card(ft.Row([req_dd, ft.IconButton("add", icon_color=THEME["primary"], on_click=add_req), ft.IconButton("refresh", on_click=lambda e: load_dd())])),
            list_col
        ]), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

def view_form_req(page: ft.Page):
    tf = ft.TextField(label="Descripción", bgcolor="white")
    def save(e):
        if db.execute_query("INSERT INTO Requisitos (curso_id, descripcion) VALUES (?, ?)", (page.session.get("curso_id"), tf.value)):
            page.go("/pedidos")
    return ft.View("/form_req", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/pedidos")), title=ft.Text("Nuevo Requisito"), bgcolor=THEME["primary"], color="white"), ft.Container(content=create_card(ft.Column([tf, ft.ElevatedButton("Guardar", on_click=save)])), padding=20, bgcolor=THEME["bg"], expand=True)])

def view_search(page: ft.Page):
    term = page.session.get("search_term")
    res = db.search_alumnos(term)
    col = ft.Column(scroll="auto", expand=True)
    
    if not res: col.controls.append(ft.Text("Sin resultados"))
    else:
        for r in res:
            def on_clk(e, aid=r['id'], cid=r['curso_id'], cname=r['curso_nombre']):
                page.session.set("alumno_id", aid)
                page.session.set("curso_id", cid)
                page.session.set("curso_nombre", cname)
                page.go("/student_detail")
            
            col.controls.append(create_card(ft.ListTile(
                leading=ft.Icon("person", color=THEME["primary"]),
                title=ft.Text(r['nombre'], weight="bold"),
                subtitle=ft.Text(f"{r['curso_nombre']} - DNI: {r['dni']}"),
                on_click=on_clk
            )))
    
    return ft.View("/search", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(f"Búsqueda: {term}"), bgcolor=THEME["primary"], color="white"), ft.Container(content=col, padding=20, bgcolor=THEME["bg"], expand=True)])

def view_admin(page: ft.Page):
    # Panel admin simplificado
    return ft.View("/admin", [
        ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Admin"), bgcolor=THEME["primary"], color="white"),
        ft.Container(content=ft.Column([
            create_card(ft.ListTile(leading=ft.Icon("calendar_month"), title=ft.Text("Gestión de Ciclos"), on_click=lambda _: show_snack(page, "Funcionalidad simplificada para demo", THEME["warning"]))),
            create_card(ft.ListTile(leading=ft.Icon("people"), title=ft.Text("Gestión de Usuarios"), on_click=lambda _: show_snack(page, "Funcionalidad simplificada para demo", THEME["warning"])))
        ]), padding=20, bgcolor=THEME["bg"], expand=True)
    ])

# ======================================================================
# CONTROLADOR PRINCIPAL (Router)
# ======================================================================

def main(page: ft.Page):
    page.title = "Asistencia UNSAM"
    page.theme_mode = "light"
    page.padding = 0

    # Mapa de rutas
    routes = {
        "/": view_login,
        "/dashboard": view_dashboard,
        "/curso": view_curso,
        "/asistencia": view_asistencia,
        "/reportes": view_reportes,
        "/student_detail": view_student_detail,
        "/form_student": view_form_student,
        "/form_curso": view_form_curso,
        "/pedidos": view_pedidos,
        "/form_req": view_form_req,
        "/search": view_search,
        "/admin": view_admin
    }

    def route_change(route):
        page.views.clear()
        # Verificar Auth (básico)
        if page.route != "/" and not page.session.get("user"):
            page.go("/")
            return

        view_fn = routes.get(page.route)
        if view_fn:
            page.views.append(view_fn(page))
        else:
            page.views.append(view_login(page))
        page.update()

    def view_pop(view):
        page.views.pop()
        top_view = page.views[-1]
        page.go(top_view.route)

    page.on_route_change = route_change
    page.on_view_pop = view_pop
    page.go("/")

if __name__ == "__main__":
    # Configuración de puerto para Nube vs Local
    port_env = os.environ.get("PORT")
    if port_env:
        ft.app(target=main, view=ft.AppView.WEB_BROWSER, port=int(port_env), host="0.0.0.0")
    else:
        ft.app(target=main, view=ft.AppView.WEB_BROWSER, port=8550)
