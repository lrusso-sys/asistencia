import customtkinter as ctk
import sqlite3
import hashlib
from datetime import date
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import os

# Intentamos importar pandas
try:
    import pandas as pd
except ImportError:
    pd = None

# ======================================================================
# CONFIGURACIÓN VISUAL (ESTILOS)
# ======================================================================
COLOR_PRIMARY = "#3B8ED0"    # Azul estándar CTK
COLOR_SUCCESS = "#2CC985"    # Verde menta moderno
COLOR_DANGER  = "#E74C3C"    # Rojo suave
COLOR_WARN    = "#F39C12"    # Naranja
COLOR_CARD    = ("#EBECF0", "#2B2B2B") # Color de tarjetas (Light/Dark)
COLOR_TEXT    = ("#1A1A1A", "#FFFFFF")

FONT_TITLE = ("Roboto Medium", 24)
FONT_HEADER = ("Roboto Medium", 18)
FONT_NORMAL = ("Roboto", 14)
FONT_SMALL = ("Roboto", 12)

# ======================================================================
# 1. Lógica de Base de Datos y Autenticación
# ======================================================================

DB_NAME = 'asistencia_alumnos.db'

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;") 
    return conn

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # --- Tablas Base ---
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'preceptor'))
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Cursos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Alumnos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            curso_id INTEGER NOT NULL,
            nombre TEXT NOT NULL,
            dni TEXT,
            observaciones TEXT,
            UNIQUE(curso_id, nombre),
            FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Asistencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            alumno_id INTEGER NOT NULL,
            fecha TEXT NOT NULL,
            status TEXT NOT NULL,
            UNIQUE(alumno_id, fecha),
            FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Feriados (
            fecha TEXT PRIMARY KEY,
            descripcion TEXT
        )
    """)

    # --- Nuevas Tablas para Requisitos (Pedidos) ---
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Requisitos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            curso_id INTEGER NOT NULL,
            descripcion TEXT NOT NULL,
            FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Requisitos_Cumplidos (
            requisito_id INTEGER NOT NULL,
            alumno_id INTEGER NOT NULL,
            PRIMARY KEY (requisito_id, alumno_id),
            FOREIGN KEY (requisito_id) REFERENCES Requisitos(id) ON DELETE CASCADE,
            FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE
        )
    """)
    
    # --- Migración de Esquema ---
    try: cursor.execute("ALTER TABLE Alumnos ADD COLUMN dni TEXT")
    except sqlite3.OperationalError: pass 
    try: cursor.execute("ALTER TABLE Alumnos ADD COLUMN observaciones TEXT")
    except sqlite3.OperationalError: pass 

    # Crear Admin Default
    cursor.execute("SELECT COUNT(*) FROM Usuarios")
    if cursor.fetchone()[0] == 0:
        admin_pass = hash_password("admin")
        cursor.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", 
                       ("admin", admin_pass, "admin"))
        print("Usuario admin creado.")

    conn.commit()
    conn.close()

# --- Auth ---
def authenticate_user(username, password):
    conn = get_db_connection()
    pwd_hash = hash_password(password)
    user = conn.execute("SELECT * FROM Usuarios WHERE username = ? AND password = ?", (username, pwd_hash)).fetchone()
    conn.close()
    if user: return True, user['role']
    return False, None

def add_user(username, password, role):
    conn = get_db_connection()
    try:
        pwd_hash = hash_password(password)
        conn.execute("INSERT INTO Usuarios (username, password, role) VALUES (?, ?, ?)", (username, pwd_hash, role))
        conn.commit()
        return True, "Usuario creado."
    except sqlite3.IntegrityError: return False, "Usuario ya existe."
    finally: conn.close()

def get_users():
    conn = get_db_connection()
    users = conn.execute("SELECT id, username, role FROM Usuarios ORDER BY username").fetchall()
    conn.close()
    return [dict(u) for u in users]

def delete_user(uid):
    conn = get_db_connection()
    try:
        conn.execute("DELETE FROM Usuarios WHERE id = ?", (uid,))
        conn.commit()
        return True, "Eliminado"
    except Exception as e: return False, str(e)
    finally: conn.close()

# --- Cursos ---
def add_curso(nombre):
    conn = get_db_connection()
    try:
        conn.execute("INSERT INTO Cursos (nombre) VALUES (?)", (nombre,))
        conn.commit()
        return True, "Curso agregado."
    except: return False, "Curso ya existe."
    finally: conn.close()

def get_cursos():
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Cursos ORDER BY nombre").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def delete_curso(cid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Cursos WHERE id = ?", (cid,))
    conn.commit()
    conn.close()
    return True, "Eliminado"

# --- Alumnos ---
def add_alumno(curso_id, nombre, dni="", obs=""):
    conn = get_db_connection()
    try:
        conn.execute("INSERT INTO Alumnos (curso_id, nombre, dni, observaciones) VALUES (?, ?, ?, ?)", 
                     (curso_id, nombre, dni, obs))
        conn.commit()
        return True, "Alumno agregado."
    except sqlite3.IntegrityError: return False, "Alumno ya existe en este curso."
    finally: conn.close()

def get_alumnos_by_curso(curso_id):
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Alumnos WHERE curso_id = ? ORDER BY nombre", (curso_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def update_alumno(aid, nombre, curso_id, dni, obs):
    conn = get_db_connection()
    try:
        conn.execute("UPDATE Alumnos SET nombre=?, curso_id=?, dni=?, observaciones=? WHERE id=?", 
                     (nombre, curso_id, dni, obs, aid))
        conn.commit()
        return True, "Actualizado."
    except: return False, "Error al actualizar."
    finally: conn.close()

def delete_alumno(aid):
    conn = get_db_connection()
    conn.execute("DELETE FROM Alumnos WHERE id=?", (aid,))
    conn.commit()
    conn.close()

# --- BÚSQUEDA Y DETALLES ---
def search_students(term):
    conn = get_db_connection()
    term = f"%{term}%"
    query = """
        SELECT a.*, c.nombre as curso_nombre 
        FROM Alumnos a
        JOIN Cursos c ON a.curso_id = c.id
        WHERE a.nombre LIKE ? OR a.dni LIKE ?
        ORDER BY a.nombre
    """
    rows = conn.execute(query, (term, term)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_student_req_status(aid, cid):
    conn = get_db_connection()
    reqs = conn.execute("SELECT * FROM Requisitos WHERE curso_id=?",(cid,)).fetchall()
    done = conn.execute("SELECT requisito_id FROM Requisitos_Cumplidos WHERE alumno_id=?",(aid,)).fetchall()
    done_ids = {r['requisito_id'] for r in done}
    conn.close()
    
    res = []
    for r in reqs:
        res.append({
            'desc': r['descripcion'],
            'ok': r['id'] in done_ids
        })
    return res

# --- Asistencia ---
def register_asistencia(aid, fecha, status):
    conn = get_db_connection()
    conn.execute("INSERT OR REPLACE INTO Asistencia (alumno_id, fecha, status) VALUES (?, ?, ?)", (aid, fecha, status))
    conn.commit()
    conn.close()
    return True

def get_asistencia_diaria_curso(curso_id, fecha):
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT a.id, asis.status FROM Alumnos a
        LEFT JOIN Asistencia asis ON a.id = asis.alumno_id AND asis.fecha = ?
        WHERE a.curso_id = ?
    """, (fecha, curso_id)).fetchall()
    conn.close()
    return {row['id']: row['status'] for row in rows if row['status']}

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
        data.append({
            'nombre': a['nombre'], 'dni': a['dni'], 'presentes': p, 'tardes': t, 'ausentes': aus,
            'justificadas': j, 'faltas': faltas, 'pct': round(pct,1)
        })
    conn.close()
    return data

# --- Lógica de Requisitos ---
def add_requisito(curso_id, descripcion):
    conn = get_db_connection()
    conn.execute("INSERT INTO Requisitos (curso_id, descripcion) VALUES (?, ?)", (curso_id, descripcion))
    conn.commit()
    conn.close()

def get_requisitos(curso_id):
    conn = get_db_connection()
    rows = conn.execute("SELECT * FROM Requisitos WHERE curso_id = ?", (curso_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def delete_requisito(req_id):
    conn = get_db_connection()
    conn.execute("DELETE FROM Requisitos WHERE id = ?", (req_id,))
    conn.commit()
    conn.close()

def toggle_requisito_alumno(req_id, aid, entregado):
    conn = get_db_connection()
    if entregado:
        conn.execute("INSERT OR IGNORE INTO Requisitos_Cumplidos (requisito_id, alumno_id) VALUES (?, ?)", (req_id, aid))
    else:
        conn.execute("DELETE FROM Requisitos_Cumplidos WHERE requisito_id = ? AND alumno_id = ?", (req_id, aid))
    conn.commit()
    conn.close()

def get_cumplimientos(req_id):
    conn = get_db_connection()
    rows = conn.execute("SELECT alumno_id FROM Requisitos_Cumplidos WHERE requisito_id = ?", (req_id,)).fetchall()
    conn.close()
    return {r['alumno_id'] for r in rows}


# ======================================================================
# 2. Interfaz Gráfica Moderna
# ======================================================================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Asistencia - UNSAM")
        self.geometry("1100x750")
        
        # Configuración de tema
        ctk.set_appearance_mode("System") 
        ctk.set_default_color_theme("blue") 
        
        self.current_user_role = None 
        self.current_user_name = None
        self.current_curso_id = None
        self.current_curso_nombre = ""

        # Contenedor principal
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        init_db()
        self.load_frames()
        self.show_frame("LoginPanel")

    def load_frames(self):
        for F in (LoginPanel, MainCoursePanel, CourseManagementPanel, DailyAttendancePanel, ReportPanel, AdminUserPanel, RequirementsPanel):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        if hasattr(frame, 'load_data'): frame.load_data()

    def show_message(self, title, msg, is_error=False):
        if is_error: messagebox.showerror(title, msg)
        else: messagebox.showinfo(title, msg)
    
    def confirm_action(self, title, msg):
        return messagebox.askyesno(title, msg)

    def logout(self):
        self.current_user_role = None
        self.show_frame("LoginPanel")

# --- LOGIN ---
class LoginPanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        # Panel centrado tipo tarjeta
        card = ctk.CTkFrame(self, width=400, corner_radius=20, fg_color=COLOR_CARD)
        card.place(relx=0.5, rely=0.5, anchor="center")

        try:
            if os.path.exists("logo_unsam.png"):
                img = ctk.CTkImage(Image.open("logo_unsam.png"), size=(250, 120))
                ctk.CTkLabel(card, image=img, text="").pack(pady=(40, 20))
            else:
                ctk.CTkLabel(card, text="UNSAM", font=("Roboto", 40, "bold")).pack(pady=(40, 10))
        except: pass

        ctk.CTkLabel(card, text="Acceso al Sistema", font=FONT_HEADER, text_color="gray").pack(pady=(0, 20))

        self.user = ctk.CTkEntry(card, placeholder_text="Usuario", width=280, height=40, font=FONT_NORMAL)
        self.user.pack(pady=10)
        
        self.pwd = ctk.CTkEntry(card, placeholder_text="Contraseña", show="*", width=280, height=40, font=FONT_NORMAL)
        self.pwd.pack(pady=10)
        
        ctk.CTkButton(card, text="INGRESAR", command=self.login, width=280, height=45, 
                      font=("Roboto", 14, "bold"), fg_color=COLOR_PRIMARY).pack(pady=(20, 40))

    def login(self):
        u, p = self.user.get(), self.pwd.get()
        ok, role = authenticate_user(u, p)
        if ok:
            self.controller.current_user_role = role
            self.controller.current_user_name = u
            self.user.delete(0, 'end')
            self.pwd.delete(0, 'end')
            self.controller.show_frame("MainCoursePanel")
        else:
            self.controller.show_message("Error", "Credenciales inválidas", True)

# --- FICHA DE ALUMNO ---
class StudentDetailWindow(ctk.CTkToplevel):
    def __init__(self, parent, student_data):
        super().__init__(parent)
        self.title(f"Ficha: {student_data['nombre']}")
        self.geometry("600x600")
        self.configure(fg_color=("gray95", "gray10"))
        
        # Header Azul
        header = ctk.CTkFrame(self, fg_color=COLOR_PRIMARY, corner_radius=0, height=100)
        header.pack(fill="x")
        
        ctk.CTkLabel(header, text=student_data['nombre'], font=("Roboto", 26, "bold"), text_color="white").place(relx=0.05, rely=0.3)
        ctk.CTkLabel(header, text=f"Curso: {student_data['curso_nombre']}", font=("Roboto", 16), text_color="white").place(relx=0.05, rely=0.7)

        # Cuerpo
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Info DNI
        if student_data['dni']:
            row = ctk.CTkFrame(body, fg_color=COLOR_CARD, corner_radius=10)
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text="🪪 DNI:", font=FONT_NORMAL, width=100).pack(side="left", padx=10, pady=10)
            ctk.CTkLabel(row, text=student_data['dni'], font=("Roboto", 14, "bold")).pack(side="left", padx=10)
        
        # Info Obs
        if student_data['observaciones']:
            row = ctk.CTkFrame(body, fg_color=COLOR_CARD, corner_radius=10)
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row, text="📝 Notas:", font=FONT_NORMAL, width=100).pack(side="left", padx=10, pady=10, anchor="n")
            ctk.CTkLabel(row, text=student_data['observaciones'], wraplength=400, justify="left").pack(side="left", padx=10, pady=10)

        ctk.CTkLabel(body, text="Estado de Documentación", font=FONT_HEADER).pack(pady=(20, 10), anchor="w")

        req_frame = ctk.CTkScrollableFrame(body, fg_color="transparent")
        req_frame.pack(fill="both", expand=True)
        
        status_list = get_student_req_status(student_data['id'], student_data['curso_id'])
        
        if not status_list:
            ctk.CTkLabel(req_frame, text="No hay pedidos registrados.").pack(pady=20)
        else:
            for s in status_list:
                card = ctk.CTkFrame(req_frame, fg_color=COLOR_CARD, corner_radius=10)
                card.pack(fill="x", pady=5)
                
                icon = "✅" if s['ok'] else "⏳"
                color = COLOR_SUCCESS if s['ok'] else COLOR_DANGER
                text_st = "ENTREGADO" if s['ok'] else "PENDIENTE"
                
                ctk.CTkLabel(card, text=icon, font=("Arial", 20)).pack(side="left", padx=15, pady=10)
                ctk.CTkLabel(card, text=s['desc'], font=("Roboto", 14, "bold")).pack(side="left", padx=5)
                ctk.CTkLabel(card, text=text_st, text_color=color, font=("Roboto", 12, "bold")).pack(side="right", padx=15)

# --- DIALOGO ALUMNO ---
class StudentDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, callback, current_data=None, cursos_list=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x550")
        self.callback = callback
        
        ctk.CTkLabel(self, text=title, font=FONT_HEADER).pack(pady=20)
        
        self.ent_name = ctk.CTkEntry(self, placeholder_text="Nombre Completo", width=300, height=35)
        self.ent_name.pack(pady=10)
        
        self.ent_dni = ctk.CTkEntry(self, placeholder_text="DNI (Opcional)", width=300, height=35)
        self.ent_dni.pack(pady=10)
        
        ctk.CTkLabel(self, text="Observaciones:").pack(pady=(10,0))
        self.ent_obs = ctk.CTkTextbox(self, width=300, height=100)
        self.ent_obs.pack(pady=5)

        self.combo_curso = None
        self.curso_map = {}
        if cursos_list:
            ctk.CTkLabel(self, text="Cambiar Curso:").pack(pady=(10,0))
            self.combo_curso = ctk.CTkComboBox(self, width=300, height=35)
            self.combo_curso.pack(pady=5)
            self.curso_map = {c['nombre']: c['id'] for c in cursos_list}
            self.combo_curso.configure(values=list(self.curso_map.keys()))

        if current_data:
            self.ent_name.insert(0, current_data['nombre'])
            if current_data['dni']: self.ent_dni.insert(0, current_data['dni'])
            if current_data['observaciones']: self.ent_obs.insert("0.0", current_data['observaciones'])
            if self.combo_curso:
                curr_c_name = next((n for n, i in self.curso_map.items() if i == current_data['curso_id']), "")
                self.combo_curso.set(curr_c_name)

        ctk.CTkButton(self, text="GUARDAR", command=self.save, width=300, height=45, fg_color=COLOR_SUCCESS).pack(pady=30)

    def save(self):
        data = {
            'nombre': self.ent_name.get().strip(),
            'dni': self.ent_dni.get().strip(),
            'obs': self.ent_obs.get("0.0", "end").strip()
        }
        if self.combo_curso:
            data['curso_id'] = self.curso_map.get(self.combo_curso.get())
        
        if data['nombre']:
            self.callback(data)
            self.destroy()

# ======================================================================
# PANELES PRINCIPALES (DASHBOARD)
# ======================================================================

class MainCoursePanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        # Header
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=30, pady=20)
        
        ctk.CTkLabel(head, text="Panel Principal", font=FONT_TITLE).pack(side="left")
        ctk.CTkButton(head, text="Cerrar Sesión", fg_color="transparent", border_width=1, border_color="gray", 
                      text_color=("black", "white"), command=controller.logout).pack(side="right")
        self.btn_admin = ctk.CTkButton(head, text="👥 Usuarios", fg_color=COLOR_PRIMARY, command=lambda: controller.show_frame("AdminUserPanel"))
        
        # Buscador
        search_card = ctk.CTkFrame(self, fg_color=COLOR_CARD, height=60)
        search_card.pack(fill="x", padx=30, pady=(0, 20))
        
        ctk.CTkLabel(search_card, text="🔍", font=("Arial", 20)).pack(side="left", padx=(20, 10))
        self.search_entry = ctk.CTkEntry(search_card, placeholder_text="Buscar alumno por Nombre o DNI...", 
                                         height=40, font=FONT_NORMAL, border_width=0, fg_color="transparent")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        self.search_entry.bind("<Return>", lambda e: self.perform_search())
        
        ctk.CTkButton(search_card, text="BUSCAR", width=100, command=self.perform_search, fg_color=COLOR_PRIMARY).pack(side="right", padx=20)

        # Contenido
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=30, pady=(0, 30))
        
        # Titulo Lista
        row_act = ctk.CTkFrame(content, fg_color="transparent")
        row_act.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(row_act, text="Mis Cursos", font=FONT_HEADER).pack(side="left")
        self.btn_add = ctk.CTkButton(row_act, text="+ Nuevo Curso", fg_color=COLOR_SUCCESS, command=self.add_course)
        self.btn_add.pack(side="right")

        self.scroll = ctk.CTkScrollableFrame(content, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True)

    def perform_search(self):
        term = self.search_entry.get().strip()
        if not term: return
        results = search_students(term)
        if not results: self.controller.show_message("Búsqueda", f"No se encontró: '{term}'", True)
        elif len(results) == 1: StudentDetailWindow(self, results[0])
        else: self.show_search_results_dialog(results)

    def show_search_results_dialog(self, results):
        d = ctk.CTkToplevel(self)
        d.title("Resultados")
        d.geometry("400x500")
        ctk.CTkLabel(d, text="Seleccione Alumno", font=FONT_HEADER).pack(pady=10)
        scroll = ctk.CTkScrollableFrame(d)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        for r in results:
            t = f"{r['nombre']}\n{r['curso_nombre']} - DNI: {r['dni']}"
            ctk.CTkButton(scroll, text=t, anchor="w", fg_color=COLOR_CARD, text_color=COLOR_TEXT, hover_color=COLOR_PRIMARY,
                          command=lambda s=r: [StudentDetailWindow(self, s), d.destroy()]).pack(fill="x", pady=5)

    def load_data(self):
        if self.controller.current_user_role == 'admin':
            self.btn_admin.pack(side="right", padx=10)
            self.btn_add.configure(state="normal")
        else:
            self.btn_admin.pack_forget()
            # self.btn_add.configure(state="disabled") # Opcional
            
        for w in self.scroll.winfo_children(): w.destroy()
        
        # Grid de Cursos (estilo Cards)
        cursos = get_cursos()
        for c in cursos:
            card = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD, corner_radius=15, height=60)
            card.pack(fill="x", pady=8)
            
            # Nombre Curso
            ctk.CTkLabel(card, text=c['nombre'], font=("Roboto", 16, "bold")).pack(side="left", padx=20, pady=15)
            
            # Botones Accion
            if self.controller.current_user_role == 'admin':
                ctk.CTkButton(card, text="🗑️", width=40, height=35, fg_color=COLOR_DANGER, 
                              command=lambda cid=c['id']: self.delete(cid)).pack(side="right", padx=(5, 20))
            
            ctk.CTkButton(card, text="ENTRAR ➡️", width=100, height=35, fg_color=COLOR_PRIMARY, 
                          command=lambda cid=c['id'], cn=c['nombre']: self.select(cid, cn)).pack(side="right", padx=5)

    def add_course(self):
        d = ctk.CTkInputDialog(text="Nombre del Curso:", title="Nuevo")
        n = d.get_input()
        if n: add_curso(n) and self.load_data()

    def delete(self, cid):
        if self.controller.confirm_action("Eliminar", "Se borrará TODO el curso."):
            delete_curso(cid)
            self.load_data()

    def select(self, cid, cn):
        self.controller.current_curso_id = cid
        self.controller.current_curso_nombre = cn
        self.controller.show_frame("CourseManagementPanel")

class CourseManagementPanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        # Top Bar
        top = ctk.CTkFrame(self, fg_color="transparent", height=50)
        top.pack(fill="x", padx=30, pady=20)
        
        ctk.CTkButton(top, text="⬅️ Volver", width=80, fg_color="transparent", border_width=1, text_color=("black", "white"),
                      command=lambda: controller.show_frame("MainCoursePanel")).pack(side="left")
        
        self.lbl_title = ctk.CTkLabel(top, text="Curso", font=FONT_TITLE)
        self.lbl_title.pack(side="left", padx=20)

        # Layout Split: Left (List) - Right (Actions)
        split = ctk.CTkFrame(self, fg_color="transparent")
        split.pack(fill="both", expand=True, padx=30, pady=(0, 30))

        # Lista Alumnos
        left = ctk.CTkFrame(split, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=(0, 20))
        
        ctk.CTkLabel(left, text="Listado de Alumnos", font=FONT_HEADER, text_color="gray").pack(anchor="w", pady=(0, 10))
        self.list_frame = ctk.CTkScrollableFrame(left, fg_color="transparent")
        self.list_frame.pack(fill="both", expand=True)

        # Sidebar Acciones
        right = ctk.CTkFrame(split, width=250, corner_radius=20, fg_color=COLOR_CARD)
        right.pack(side="right", fill="y")
        
        ctk.CTkLabel(right, text="Herramientas", font=("Roboto", 16, "bold")).pack(pady=20)
        
        self.create_sidebar_btn(right, "✅  Tomar Asistencia", lambda: controller.show_frame("DailyAttendancePanel"), COLOR_PRIMARY)
        self.create_sidebar_btn(right, "📋  Pedidos / Document.", lambda: controller.show_frame("RequirementsPanel"), COLOR_WARN)
        self.create_sidebar_btn(right, "📊  Informes y Estadíst.", lambda: controller.show_frame("ReportPanel"), "#8E44AD") # Violeta
        ctk.CTkFrame(right, height=2, fg_color="gray").pack(fill="x", padx=20, pady=20)
        self.create_sidebar_btn(right, "➕  Agregar Alumno", self.add_student_dialog, COLOR_SUCCESS)

    def create_sidebar_btn(self, parent, text, cmd, color):
        ctk.CTkButton(parent, text=text, command=cmd, fg_color=color, height=45, anchor="w", 
                      font=("Roboto", 13, "bold")).pack(fill="x", padx=20, pady=8)

    def load_data(self):
        self.lbl_title.configure(text=f"{self.controller.current_curso_nombre}")
        for w in self.list_frame.winfo_children(): w.destroy()
        
        alumnos = get_alumnos_by_curso(self.controller.current_curso_id)
        if not alumnos: ctk.CTkLabel(self.list_frame, text="No hay alumnos registrados.").pack(pady=20)

        for a in alumnos:
            card = ctk.CTkFrame(self.list_frame, fg_color=COLOR_CARD, corner_radius=10)
            card.pack(fill="x", pady=5)
            
            info = a['nombre']
            sub = f"DNI: {a['dni']}" if a['dni'] else "Sin DNI"
            
            # Info Block
            info_frame = ctk.CTkFrame(card, fg_color="transparent")
            info_frame.pack(side="left", padx=15, pady=10)
            ctk.CTkLabel(info_frame, text=info, font=("Roboto", 14, "bold")).pack(anchor="w")
            ctk.CTkLabel(info_frame, text=sub, font=("Roboto", 11), text_color="gray").pack(anchor="w")

            # Actions
            if self.controller.current_user_role == 'admin':
                ctk.CTkButton(card, text="🗑️", width=30, fg_color="transparent", text_color=COLOR_DANGER, hover_color=COLOR_CARD,
                              command=lambda x=a['id']: self.delete_s(x)).pack(side="right", padx=5)

            ctk.CTkButton(card, text="📝 Editar", width=80, fg_color=COLOR_PRIMARY, height=30,
                          command=lambda x=a: self.edit_student(x)).pack(side="right", padx=5)

    def add_student_dialog(self):
        def cb(data):
            add_alumno(self.controller.current_curso_id, data['nombre'], data['dni'], data['obs'])
            self.load_data()
        StudentDialog(self, "Nuevo Alumno", cb)

    def edit_student(self, a_data):
        def cb(data):
            update_alumno(a_data['id'], data['nombre'], data['curso_id'], data['dni'], data['obs'])
            self.load_data()
        StudentDialog(self, "Editar Alumno", cb, current_data=a_data, cursos_list=get_cursos())

    def delete_s(self, aid):
        if self.controller.confirm_action("Borrar", "Se borrará el alumno."):
            delete_alumno(aid)
            self.load_data()

# --- REQUISITOS ---
class RequirementsPanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=30, pady=20)
        ctk.CTkButton(head, text="⬅️ Volver", width=80, fg_color="transparent", border_width=1, text_color=("black", "white"),
                      command=lambda: controller.show_frame("CourseManagementPanel")).pack(side="left")
        ctk.CTkLabel(head, text="Gestión de Pedidos", font=FONT_TITLE).pack(side="left", padx=20)

        # Control Bar
        bar = ctk.CTkFrame(self, fg_color=COLOR_CARD, height=60)
        bar.pack(fill="x", padx=30, pady=(0, 20))
        
        ctk.CTkLabel(bar, text="Pedido:", font=FONT_NORMAL).pack(side="left", padx=20)
        self.combo_req = ctk.CTkComboBox(bar, width=250, command=self.on_req_change)
        self.combo_req.pack(side="left", padx=5)
        
        ctk.CTkButton(bar, text="+ Nuevo", command=self.add_req, width=80, fg_color=COLOR_SUCCESS).pack(side="left", padx=10)
        ctk.CTkButton(bar, text="Borrar Actual", command=self.del_req, width=100, fg_color=COLOR_DANGER).pack(side="right", padx=20)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=30, pady=(0, 30))

        self.req_map = {} 
        self.current_req_id = None

    def load_data(self):
        reqs = get_requisitos(self.controller.current_curso_id)
        self.req_map = {r['descripcion']: r['id'] for r in reqs}
        vals = list(self.req_map.keys())
        self.combo_req.configure(values=vals)
        if vals:
            self.combo_req.set(vals[0])
            self.current_req_id = self.req_map[vals[0]]
            self.load_checklist()
        else:
            self.combo_req.set("")
            self.current_req_id = None
            for w in self.scroll.winfo_children(): w.destroy()
            ctk.CTkLabel(self.scroll, text="Crea un pedido nuevo para comenzar.").pack(pady=50)

    def on_req_change(self, choice):
        self.current_req_id = self.req_map[choice]
        self.load_checklist()

    def add_req(self):
        d = ctk.CTkInputDialog(text="Nombre del Pedido:", title="Nuevo")
        txt = d.get_input()
        if txt:
            add_requisito(self.controller.current_curso_id, txt)
            self.load_data()
            self.combo_req.set(txt)
            self.on_req_change(txt)

    def del_req(self):
        if not self.current_req_id: return
        if self.controller.confirm_action("Borrar", "Se eliminará el pedido."):
            delete_requisito(self.current_req_id)
            self.load_data()

    def load_checklist(self):
        for w in self.scroll.winfo_children(): w.destroy()
        if not self.current_req_id: return

        alumnos = get_alumnos_by_curso(self.controller.current_curso_id)
        cumplidos = get_cumplimientos(self.current_req_id) 

        for a in alumnos:
            row = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD, corner_radius=8)
            row.pack(fill="x", pady=4)
            
            ctk.CTkLabel(row, text=a['nombre'], font=FONT_NORMAL).pack(side="left", padx=15, pady=8)
            
            var = ctk.BooleanVar(value=(a['id'] in cumplidos))
            def on_toggle(aid=a['id'], v=var): toggle_requisito_alumno(self.current_req_id, aid, v.get())

            cb = ctk.CTkCheckBox(row, text="ENTREGADO", variable=var, command=on_toggle, 
                                 fg_color=COLOR_SUCCESS, hover_color="#2ECC71")
            cb.pack(side="right", padx=15)

class DailyAttendancePanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.vars = {}
        
        # Header
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=30, pady=20)
        ctk.CTkButton(head, text="⬅️ Volver", width=80, fg_color="transparent", border_width=1, text_color=("black", "white"),
                      command=lambda: controller.show_frame("CourseManagementPanel")).pack(side="left")
        ctk.CTkLabel(head, text="Asistencia Diaria", font=FONT_TITLE).pack(side="left", padx=20)
        
        # Control Date
        ctrl = ctk.CTkFrame(self, fg_color=COLOR_CARD, height=60)
        ctrl.pack(fill="x", padx=30, pady=(0, 20))
        
        ctk.CTkLabel(ctrl, text="Fecha (AAAA-MM-DD):").pack(side="left", padx=20)
        self.date_entry = ctk.CTkEntry(ctrl, width=120)
        self.date_entry.insert(0, date.today().isoformat())
        self.date_entry.pack(side="left", padx=5)
        ctk.CTkButton(ctrl, text="Cargar Datos", command=self.load_list).pack(side="left", padx=15)
        
        ctk.CTkButton(ctrl, text="💾 GUARDAR CAMBIOS", fg_color=COLOR_SUCCESS, command=self.save, font=("Roboto", 13, "bold")).pack(side="right", padx=20)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=30, pady=(0, 30))

    def load_data(self):
        self.date_entry.delete(0, 'end')
        self.date_entry.insert(0, date.today().isoformat())
        self.load_list()

    def load_list(self):
        for w in self.scroll.winfo_children(): w.destroy()
        self.vars = {}
        dt = self.date_entry.get()
        
        try:
            if date.fromisoformat(dt) > date.today():
                self.controller.show_message("Error", "No puedes cargar fechas futuras", True)
                return
        except: return

        existing = get_asistencia_diaria_curso(self.controller.current_curso_id, dt)
        alumnos = get_alumnos_by_curso(self.controller.current_curso_id)

        # Header Columns
        h = ctk.CTkFrame(self.scroll, fg_color="transparent")
        h.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(h, text="Alumno", font=("Roboto", 12, "bold"), anchor="w").pack(side="left", padx=15, fill="x", expand=True)
        opts = ["P", "T", "A", "J", "S", "N"]
        for o in opts: ctk.CTkLabel(h, text=o, width=50, font=("Roboto", 12, "bold")).pack(side="left")

        for a in alumnos:
            row = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD, corner_radius=8)
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(row, text=a['nombre'], font=FONT_NORMAL, anchor="w").pack(side="left", padx=15, fill="x", expand=True)
            
            v = ctk.StringVar(value=existing.get(a['id'], 'P'))
            self.vars[a['id']] = v
            
            for o in opts:
                ctk.CTkRadioButton(row, text="", variable=v, value=o, width=50, border_width_checked=6, border_width_unchecked=2).pack(side="left")

    def save(self):
        dt = self.date_entry.get()
        c = 0
        for aid, v in self.vars.items():
            register_asistencia(aid, dt, v.get())
            c+=1
        self.controller.show_message("Guardado", f"{c} registros guardados.")
        self.controller.show_frame("CourseManagementPanel")

class ReportPanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=30, pady=20)
        ctk.CTkButton(head, text="⬅️ Volver", width=80, fg_color="transparent", border_width=1, text_color=("black", "white"),
                      command=lambda: controller.show_frame("CourseManagementPanel")).pack(side="left")
        ctk.CTkLabel(head, text="Informes y Estadísticas", font=FONT_TITLE).pack(side="left", padx=20)
        
        f = ctk.CTkFrame(self, fg_color=COLOR_CARD, height=60)
        f.pack(fill="x", padx=30, pady=(0, 20))
        
        ctk.CTkLabel(f, text="Desde:").pack(side="left", padx=10)
        self.d1 = ctk.CTkEntry(f, width=100); self.d1.pack(side="left", padx=5)
        
        ctk.CTkLabel(f, text="Hasta:").pack(side="left", padx=10)
        self.d2 = ctk.CTkEntry(f, width=100); self.d2.pack(side="left", padx=5)
        
        ctk.CTkButton(f, text="Generar Tabla", command=self.gen, fg_color=COLOR_PRIMARY).pack(side="left", padx=20)
        ctk.CTkButton(f, text="Descargar Excel", command=self.excel, fg_color="#27AE60").pack(side="right", padx=20)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=30, pady=(0, 30))

    def load_data(self):
        self.d1.delete(0,'end'); self.d1.insert(0, date.today().replace(day=1).isoformat())
        self.d2.delete(0,'end'); self.d2.insert(0, date.today().isoformat())
        for w in self.scroll.winfo_children(): w.destroy()

    def gen(self):
        for w in self.scroll.winfo_children(): w.destroy()
        data = get_report_data(self.controller.current_curso_id, self.d1.get(), self.d2.get())
        
        # Header Table
        h = ctk.CTkFrame(self.scroll, fg_color="gray", height=30)
        h.pack(fill="x")
        cols = [("Alumno", 200), ("Pres", 60), ("Tard", 60), ("Aus", 60), ("Just", 60), ("Faltas", 80), ("%", 80)]
        for c, w in cols: ctk.CTkLabel(h, text=c, width=w, font=("Roboto", 11, "bold"), text_color="white").pack(side="left", padx=1)

        for d in data:
            row = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD, corner_radius=5)
            row.pack(fill="x", pady=2)
            
            ctk.CTkLabel(row, text=d['nombre'], width=200, anchor="w").pack(side="left", padx=5)
            vals = [d['presentes'], d['tardes'], d['ausentes'], d['justificadas']]
            for v in vals: ctk.CTkLabel(row, text=str(v), width=60).pack(side="left", padx=1)
            
            # Highlight Faltas
            ctk.CTkLabel(row, text=f"{d['faltas']:.2f}", width=80, font=("Roboto", 12, "bold")).pack(side="left", padx=1)
            
            color = "red" if d['pct'] > 20 else ("orange" if d['pct'] > 10 else "green")
            ctk.CTkLabel(row, text=f"{d['pct']}%", width=80, text_color=color, font=("Roboto", 12, "bold")).pack(side="left", padx=1)

    def excel(self):
        if not pd: return self.controller.show_message("Error", "Instala pandas", True)
        data = get_report_data(self.controller.current_curso_id, self.d1.get(), self.d2.get())
        fn = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if fn:
            pd.DataFrame(data).to_excel(fn)
            self.controller.show_message("Éxito", "Exportado")

class AdminUserPanel(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=30, pady=20)
        ctk.CTkButton(head, text="⬅️ Volver", width=80, fg_color="transparent", border_width=1, text_color=("black", "white"),
                      command=lambda: controller.show_frame("MainCoursePanel")).pack(side="left")
        ctk.CTkLabel(head, text="Gestión de Usuarios", font=FONT_TITLE).pack(side="left", padx=20)
        
        form = ctk.CTkFrame(self, fg_color=COLOR_CARD, height=60)
        form.pack(fill="x", padx=30, pady=(0, 20))
        
        self.u = ctk.CTkEntry(form, placeholder_text="Usuario"); self.u.pack(side="left", padx=10, pady=10)
        self.p = ctk.CTkEntry(form, placeholder_text="Contraseña"); self.p.pack(side="left", padx=10, pady=10)
        self.r = ctk.CTkComboBox(form, values=["preceptor", "admin"]); self.r.pack(side="left", padx=10, pady=10)
        ctk.CTkButton(form, text="Crear Usuario", command=self.add, fg_color=COLOR_SUCCESS).pack(side="left", padx=20)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=30, pady=(0, 30))

    def load_data(self):
        for w in self.scroll.winfo_children(): w.destroy()
        for u in get_users():
            row = ctk.CTkFrame(self.scroll, fg_color=COLOR_CARD, corner_radius=10)
            row.pack(fill="x", pady=5)
            
            icon = "👮" if u['role'] == 'admin' else "🧑‍🏫"
            ctk.CTkLabel(row, text=f"{icon} {u['username']}", font=FONT_NORMAL).pack(side="left", padx=20, pady=10)
            ctk.CTkLabel(row, text=u['role'].upper(), font=("Roboto", 10, "bold"), text_color="gray").pack(side="left", padx=10)
            
            if u['username'] != self.controller.current_user_name:
                ctk.CTkButton(row, text="Eliminar", width=80, height=30, fg_color=COLOR_DANGER, 
                              command=lambda uid=u['id']: self.rem(uid)).pack(side="right", padx=20)

    def add(self):
        add_user(self.u.get(), self.p.get(), self.r.get())
        self.load_data()
    def rem(self, uid):
        delete_user(uid)
        self.load_data()

if __name__ == "__main__":
    app = App()
    app.mainloop()