import flet as ft
import hashlib
from datetime import date
import os
import json
import base64
import time # Para el backoff en las llamadas a la API

# Importación de librerías externas
try:
    import pandas as pd
except ImportError:
    pd = None
    print("⚠️ ADVERTENCIA: 'pandas' no está instalado. La exportación fallará.")

try:
    import openpyxl
except ImportError:
    print("⚠️ ADVERTENCIA: 'openpyxl' no está instalado. La exportación a Excel fallará.")

# Importación de Firebase Admin SDK
try:
    import firebase_admin
    from firebase_admin import credentials, firestore
except ImportError:
    firebase_admin = None
    firestore = None
    print("⚠️ ADVERTENCIA: 'firebase-admin' no está instalado. La persistencia fallará.")

# ======================================================================
# 1. LÓGICA DE BASE DE DATOS (Backend con Firestore)
# ======================================================================

db = None
app_id = "default-app-id" # Default for local testing

def init_firestore():
    """Inicializa la conexión a Firestore usando credenciales inyectadas o mock para desarrollo local."""
    global db, app_id

    # 1. Obtener la configuración y App ID del entorno
    firebase_config_str = os.environ.get('__firebase_config', None)
    app_id = os.environ.get('__app_id', 'local-asistencia-app')

    if firebase_admin is None:
        print("ERROR: firebase-admin no está instalado. No se puede conectar a Firestore.")
        return False

    if firebase_admin._apps:
        # Ya está inicializado, solo obtener la referencia
        db = firestore.client()
        print(f"Firestore ya inicializado para App ID: {app_id}")
        return True

    try:
        if firebase_config_str:
            # Modo de Despliegue (Cloud Canvas)
            firebase_config = json.loads(firebase_config_str)
            # Detección de credenciales de servicio (clave privada)
            if firebase_config.get("private_key"):
                # Si se utiliza service account (preferido para Python Admin SDK)
                cred = credentials.Certificate(firebase_config)
            else:
                # Si se utiliza la configuración estándar (a veces necesaria)
                print("Usando credenciales estándar. Asegúrate de que las variables de entorno estén bien configuradas.")
                cred = credentials.ApplicationDefault()
            
            # Inicializar la app
            firebase_admin.initialize_app(cred, {'databaseURL': f'https://{firebase_config["projectId"]}.firebaseio.com'})
            db = firestore.client()
            print(f"Firestore inicializado con éxito para App ID: {app_id}")
            return True
        else:
            # Modo de Desarrollo Local (requiere un archivo serviceAccountKey.json o simulación)
            # Para simplificar el setup local: Si no hay credenciales, usamos un mock.
            # ADVERTENCIA: Para desarrollo local real, se requiere un archivo de credenciales.
            print("USANDO MODO MOCK/FALLBACK. Si no estás en la nube, la DB no funcionará.")
            # db permanece como None o requiere setup manual del usuario para probar localmente.
            return False
            
    except Exception as e:
        print(f"Error al inicializar Firestore: {e}")
        return False

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# --- Funciones de Utilidad para la Ruta de la Colección ---
def get_collection_ref(collection_name):
    """Retorna la referencia a la colección pública, asegurando persistencia."""
    global app_id
    # Estructura obligatoria para datos compartidos en el entorno:
    return db.collection('artifacts').document(app_id).collection('public').document('data').collection(collection_name)

# --- Funciones de Reintento para la API (Manejo de throttling) ---
def firestore_retry_wrapper(func, *args, max_retries=5):
    """Ejecuta una función de Firestore con reintentos y backoff exponencial."""
    for attempt in range(max_retries):
        try:
            return func(*args)
        except Exception as e:
            if attempt < max_retries - 1:
                delay = 2 ** attempt
                print(f"Firestore Error: {e}. Reintentando en {delay}s...")
                time.sleep(delay)
            else:
                print(f"Firestore Error: {e}. Falló después de {max_retries} intentos.")
                raise e

# --- USUARIOS ---
def authenticate_user(username, password):
    if db is None: return (False, None)
    hashed_pwd = hash_password(password)
    
    # 1. Intentar iniciar sesión
    try:
        users_ref = get_collection_ref('Usuarios')
        query = users_ref.where('username', '==', username).limit(1).stream()
        user_doc = next(query, None)
    except Exception:
        return (False, None) # Error de conexión

    if user_doc and user_doc.to_dict().get('password') == hashed_pwd:
        return (True, user_doc.to_dict().get('role'))
    
    # 2. Inicializar Admin por defecto si no hay usuarios (solo si se está ejecutando en el entorno)
    if not user_doc:
        try:
            if next(users_ref.limit(1).stream(), None) is None:
                if username == "admin" and password == "admin":
                    add_user("admin", "admin", "admin")
                    return (True, "admin")
        except:
            pass # No hacer nada si falla la creación

    return (False, None)

def get_users():
    if db is None: return []
    try:
        rows = get_collection_ref('Usuarios').stream()
        return [{**r.to_dict(), 'id': r.id} for r in rows]
    except: return []

def add_user(u, p, r):
    if db is None: return False
    try:
        # Verificar si el usuario ya existe
        query = get_collection_ref('Usuarios').where('username', '==', u).limit(1).stream()
        if next(query, None): return False
        
        get_collection_ref('Usuarios').add({'username': u, 'password': hash_password(p), 'role': r})
        return True
    except Exception as e:
        print(f"Error adding user: {e}")
        return False

def delete_user(uid):
    if db is None: return
    try:
        get_collection_ref('Usuarios').document(uid).delete()
    except: pass

# --- CICLOS ---
def get_ciclos():
    if db is None: return []
    try:
        rows = get_collection_ref('Ciclos').order_by('nombre', direction=firestore.Query.DESCENDING).stream()
        return [{**r.to_dict(), 'id': r.id} for r in rows]
    except: return []

def get_ciclo_activo():
    if db is None: return None
    try:
        rows = get_collection_ref('Ciclos').where('activo', '==', 1).limit(1).stream()
        row = next(rows, None)
        return {**row.to_dict(), 'id': row.id} if row else None
    except: return None

def add_ciclo(nombre):
    if db is None: return False
    try:
        # Desactivar todos los ciclos
        for c in get_collection_ref('Ciclos').where('activo', '==', 1).stream():
            c.reference.update({'activo': 0})
        
        # Insertar nuevo ciclo como activo
        get_collection_ref('Ciclos').add({'nombre': nombre, 'activo': 1})
        return True
    except: return False

def activar_ciclo(cid):
    if db is None: return
    try:
        # Desactivar todos los ciclos
        for c in get_collection_ref('Ciclos').where('activo', '==', 1).stream():
            c.reference.update({'activo': 0})
        
        # Activar el ciclo seleccionado
        get_collection_ref('Ciclos').document(cid).update({'activo': 1})
    except: pass

# --- CURSOS ---
def get_cursos():
    if db is None: return []
    try:
        ciclo = get_ciclo_activo()
        if not ciclo: return []
        
        rows = get_collection_ref('Cursos').where('ciclo_id', '==', ciclo['id']).order_by('nombre').stream()
        return [{**r.to_dict(), 'id': r.id} for r in rows]
    except: return []

def add_curso(nombre):
    if db is None: return False
    try:
        ciclo = get_ciclo_activo()
        if not ciclo: return False

        # Verificar si el curso ya existe en este ciclo
        query = get_collection_ref('Cursos').where('ciclo_id', '==', ciclo['id']).where('nombre', '==', nombre).limit(1).stream()
        if next(query, None): return False
        
        get_collection_ref('Cursos').add({'nombre': nombre, 'ciclo_id': ciclo['id']})
        return True
    except: return False

def delete_curso(cid):
    if db is None: return
    try:
        # Firestore no hace ON DELETE CASCADE automáticamente, eliminar subcolecciones manualmente
        # (Alumnos, Asistencia, Requisitos, etc., se quedan "huérfanas" pero no son necesarias
        # ya que la query de cursos activos depende del ciclo. Por ahora solo eliminamos el curso.)
        get_collection_ref('Cursos').document(cid).delete()
    except: pass

# --- ALUMNOS ---
def get_alumnos(curso_id):
    if db is None: return []
    try:
        rows = get_collection_ref('Alumnos').where('curso_id', '==', curso_id).order_by('nombre').stream()
        return [{**r.to_dict(), 'id': r.id} for r in rows]
    except: return []

def get_alumno_by_id(aid):
    if db is None: return None
    try:
        doc = get_collection_ref('Alumnos').document(aid).get()
        if not doc.exists: return None
        alumno = {**doc.to_dict(), 'id': doc.id}
        
        # Obtener nombre del curso
        curso_doc = get_collection_ref('Cursos').document(alumno['curso_id']).get()
        if curso_doc.exists: alumno['curso_nombre'] = curso_doc.to_dict().get('nombre', 'Curso Desconocido')

        return alumno
    except: return None

def add_alumno(curso_id, nombre, dni, obs, tutor_n, tutor_t):
    if db is None: return False
    try:
        # Verificar duplicado
        query = get_collection_ref('Alumnos').where('curso_id', '==', curso_id).where('nombre', '==', nombre).limit(1).stream()
        if next(query, None): return False
        
        data = {'curso_id': curso_id, 'nombre': nombre, 'dni': dni, 'observaciones': obs, 'tutor_nombre': tutor_n, 'tutor_telefono': tutor_t}
        get_collection_ref('Alumnos').add(data)
        return True
    except: return False

def update_alumno(aid, nombre, dni, obs, tutor_n, tutor_t):
    if db is None: return
    try:
        data = {'nombre': nombre, 'dni': dni, 'observaciones': obs, 'tutor_nombre': tutor_n, 'tutor_telefono': tutor_t}
        get_collection_ref('Alumnos').document(aid).update(data)
    except: pass

def delete_alumno(aid):
    if db is None: return
    try:
        # Eliminar alumno (la asistencia y requisitos cumplidos quedan huérfanos)
        get_collection_ref('Alumnos').document(aid).delete()
    except: pass

def search_students(term):
    """Búsqueda simple por nombre/DNI (Firestore no soporta busquedas LIKE nativas)."""
    if db is None: return []
    try:
        # Búsqueda manual:
        results = []
        alumnos_stream = get_collection_ref('Alumnos').stream()
        
        active_ciclo = get_ciclo_activo()
        if not active_ciclo: return []

        # Mapa de cursos del ciclo activo para filtrado
        active_curso_ids = {c['id']: c for c in get_cursos()}
        
        term_lower = term.lower()

        for a_doc in alumnos_stream:
            alumno = {**a_doc.to_dict(), 'id': a_doc.id}
            
            # Solo alumnos de cursos activos
            if alumno['curso_id'] not in active_curso_ids:
                continue
            
            nombre = alumno.get('nombre', '').lower()
            dni = alumno.get('dni', '').lower()

            if term_lower in nombre or term_lower in dni:
                alumno['curso_nombre'] = active_curso_ids[alumno['curso_id']]['nombre']
                alumno['ciclo_nombre'] = active_ciclo['nombre']
                results.append(alumno)
        
        # Ordenar por nombre después de filtrar
        return sorted(results, key=lambda x: x['nombre'])
    except Exception as e:
        print(f"Search error: {e}")
        return []

def get_report_data(curso_id, start, end):
    if db is None: return []
    try:
        alumnos = get_alumnos(curso_id)
        report_data = []

        for a in alumnos:
            # Obtener todas las asistencias del alumno en el rango de fechas
            rows = get_collection_ref('Asistencia').where('alumno_id', '==', a['id']).where('fecha', '>=', start).where('fecha', '<=', end).stream()
            counts = {'P': 0, 'T': 0, 'A': 0, 'J': 0, 'S': 0}
            
            for row in rows:
                status = row.to_dict().get('status', 'P')
                counts[status] = counts.get(status, 0) + 1
            
            p, t, aus, j, s = counts['P'], counts['T'], counts['A'], counts['J'], counts['S']
            
            # Cálculo de faltas: 1 A, 1 S, 0.25 T
            faltas = aus + s + (t * 0.25)
            total_dias = p + t + aus + j + s # Total de días con registro
            pct = (faltas / total_dias * 100) if total_dias > 0 else 0
            
            report_data.append({
                'nombre': a['nombre'], 
                'dni': a['dni'], 
                'tutor_nombre': a.get('tutor_nombre', '-'),
                'tutor_telefono': a.get('tutor_telefono', '-'),
                'p': p, 't': t, 'a': aus, 'j': j, 's': s, 
                'faltas': faltas, 
                'pct': round(pct, 1)
            })
        
        return report_data
    except Exception as e:
        print(f"Error getting report data: {e}")
        return []

# --- ASISTENCIA ---
def get_asistencia_diaria(curso_id, fecha):
    if db is None: return {}
    try:
        alumnos_ids = [a['id'] for a in get_alumnos(curso_id)]
        if not alumnos_ids: return {}

        results = {}
        # Consulta por ID y Fecha
        rows = get_collection_ref('Asistencia').where('fecha', '==', fecha).stream()
        
        for r in rows:
            data = r.to_dict()
            if data['alumno_id'] in alumnos_ids:
                 results[data['alumno_id']] = data['status']
        return results
    except: return {}

def register_asistencia(aid, fecha, status):
    if db is None: return
    try:
        # Crea un ID de documento único basado en alumno y fecha para asegurar unicidad
        doc_id = f"{aid}_{fecha}"
        
        data = {'alumno_id': aid, 'fecha': fecha, 'status': status}
        get_collection_ref('Asistencia').document(doc_id).set(data, merge=True)
    except: pass

# --- REQUISITOS ---
# Las funciones de requisitos (Requisitos, Requisitos_Cumplidos) se adaptan a Firestore de manera similar,
# usando los IDs del documento como claves foráneas.

def get_requisitos(curso_id):
    if db is None: return []
    try:
        rows = get_collection_ref('Requisitos').where('curso_id', '==', curso_id).stream()
        return [{**r.to_dict(), 'id': r.id} for r in rows]
    except: return []

def add_requisito(curso_id, desc):
    if db is None: return False
    try:
        get_collection_ref('Requisitos').add({'curso_id': curso_id, 'descripcion': desc})
        return True
    except: return False

def delete_requisito(rid):
    if db is None: return
    try:
        get_collection_ref('Requisitos').document(rid).delete()
    except: pass

def get_cumplimientos(rid):
    if db is None: return set()
    try:
        rows = get_collection_ref('Requisitos_Cumplidos').where('requisito_id', '==', rid).stream()
        return {r.to_dict()['alumno_id'] for r in rows}
    except: return set()

def toggle_cumplimiento(rid, aid, val):
    if db is None: return
    try:
        doc_id = f"{rid}_{aid}"
        doc_ref = get_collection_ref('Requisitos_Cumplidos').document(doc_id)
        
        if val:
            doc_ref.set({'requisito_id': rid, 'alumno_id': aid})
        else:
            doc_ref.delete()
    except: pass

def get_student_req_status(aid, cid):
    if db is None: return []
    try:
        reqs = get_requisitos(cid)
        done_ids = {r.to_dict()['requisito_id'] for r in get_collection_ref('Requisitos_Cumplidos').where('alumno_id', '==', aid).stream()}
        
        return [{'desc': r['descripcion'], 'ok': r['id'] in done_ids} for r in reqs]
    except: return []

# ======================================================================
# 2. INTERFAZ GRÁFICA WEB (Flet) - Lógica de Flet
# (Esta sección permanece casi igual, usando las nuevas funciones DB)
# ======================================================================

def main(page: ft.Page):
    page.title = "Sistema de Asistencia UNSAM"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 0
    COLOR_PRIMARY = ft.colors.BLUE_700
    COLOR_BG = ft.colors.BLUE_GREY_50
    COLOR_DANGER = ft.colors.RED_700
    COLOR_WARNING = ft.colors.ORANGE_500
    
    # Inicialización de la base de datos (Firestore)
    if not init_firestore():
        page.add(ft.Container(
            content=ft.Text("ERROR CRÍTICO: No se pudo conectar a Firestore. La persistencia fallará.", color=COLOR_DANGER, size=18, weight="bold"),
            alignment=ft.alignment.center, expand=True, padding=50
        ))
        page.update()
        return

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
        user = ft.TextField(label="Usuario", width=300, bgcolor="white", border_radius=10)
        pwd = ft.TextField(label="Contraseña", password=True, width=300, bgcolor="white", border_radius=10)
        
        def login_click(e):
            ok, role = authenticate_user(user.value, pwd.value)
            if ok:
                state["role"] = role
                state["username"] = user.value
                page.go("/dashboard")
            else:
                show_snack("Credenciales incorrectas", COLOR_DANGER)

        logo_widget = ft.Icon(ft.icons.SCHOOL, size=80, color=COLOR_PRIMARY)

        # Contenedor de la URL de la App ID para referencia del usuario
        app_id_info = ft.Text(f"App ID: {app_id}", size=10, color=ft.colors.GREY_600)

        return ft.View("/", [
            ft.Container(
                content=ft.Column([
                    logo_widget,
                    ft.Text("Sistema de Asistencia", size=24, weight=ft.FontWeight.BOLD),
                    ft.Text("UNSAM", size=16, color=ft.colors.GREY_600),
                    ft.Divider(height=20, color="transparent"),
                    user, pwd,
                    ft.ElevatedButton("INGRESAR", on_click=login_click, width=300, height=50, 
                                     bgcolor=COLOR_PRIMARY, color="white",
                                     style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10))),
                    ft.Container(height=10),
                    app_id_info # Mostrar el ID para fines de depuración/referencia
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
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

        cursos_col = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
        
        def load_cursos():
            cursos_col.controls.clear()
            for c in get_cursos():
                card = ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.icons.BOOK, color=COLOR_PRIMARY),
                        ft.Text(c['nombre'], weight=ft.FontWeight.BOLD, size=16, expand=True),
                        ft.IconButton(ft.icons.ARROW_FORWARD, on_click=lambda e, cid=c['id'], cn=c['nombre']: ir_curso(cid, cn)),
                        ft.IconButton(ft.icons.DELETE, icon_color=COLOR_DANGER, on_click=lambda e, cid=c['id']: del_c(cid)) if state["role"] == 'admin' else ft.Container()
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    padding=15,
                    bgcolor="white",
                    border_radius=10,
                    shadow=ft.BoxShadow(blur_radius=5, color=ft.colors.BLACK12)
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
            show_snack("Curso eliminado permanentemente.", COLOR_DANGER)

        def go_add_curso(e):
            if not ciclo_activo:
                show_snack("Debes crear/activar un ciclo lectivo primero.", COLOR_WARNING)
            else:
                page.go("/form_curso")

        load_cursos()

        admin_btn = ft.IconButton(ft.icons.SETTINGS, tooltip="Admin", icon_color="white", on_click=lambda _: page.go("/admin")) if state["role"] == 'admin' else ft.Container()

        return ft.View("/dashboard", [
            ft.AppBar(title=ft.Text("Panel Principal"), bgcolor=COLOR_PRIMARY, color="white", 
                      actions=[admin_btn, ft.IconButton(ft.icons.LOGOUT, icon_color="white", on_click=lambda _: page.go("/"))]),
            ft.Container(
                content=ft.Column([
                    ft.Container(content=ft.Text(f"Ciclo Lectivo: {nombre_ciclo}", weight=ft.FontWeight.BOLD, color=COLOR_PRIMARY), padding=5),
                    ft.Container(
                        content=ft.Row([search_field, ft.IconButton(ft.icons.SEARCH, on_click=go_search)]),
                        padding=10, bgcolor="white", border_radius=10
                    ),
                    ft.Row([
                        ft.Text("Mis Cursos", size=20, weight=ft.FontWeight.BOLD), 
                        ft.IconButton(ft.icons.ADD_CIRCLE, icon_color=ft.colors.GREEN_600, icon_size=30, on_click=go_add_curso)
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    cursos_col
                ]),
                padding=15, expand=True, bgcolor=COLOR_BG
            )
        ])

    def admin_view():
        return ft.View("/admin", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/dashboard")),
                      title=ft.Text("Administración"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(
                content=ft.Column([
                    ft.Text("Seleccione una opción:", size=18, weight=ft.FontWeight.BOLD),
                    ft.Divider(),
                    ft.ListTile(leading=ft.Icon(ft.icons.CALENDAR_MONTH, color=COLOR_PRIMARY), title=ft.Text("Ciclos Lectivos"), subtitle=ft.Text("Crear, cerrar y cambiar años escolares"), on_click=lambda _: page.go("/ciclos")),
                    ft.ListTile(leading=ft.Icon(ft.icons.PEOPLE, color=COLOR_PRIMARY), title=ft.Text("Usuarios"), subtitle=ft.Text("Gestionar preceptores y admins"), on_click=lambda _: page.go("/users"))
                ]),
                padding=20, bgcolor=COLOR_BG, expand=True
            )
        ])

    def ciclos_view():
        tf_new = ft.TextField(label="Año / Nombre Ciclo", expand=True, bgcolor="white", border_radius=10)
        list_col = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)

        def load():
            list_col.controls.clear()
            ciclos = get_ciclos()
            for c in ciclos:
                is_active = c.get('activo', 0) == 1
                icon = ft.icons.CHECK_CIRCLE if is_active else ft.icons.CIRCLE_OUTLINED
                color = ft.colors.GREEN_600 if is_active else ft.colors.GREY_600
                trailing = ft.Container()
                
                if not is_active:
                    trailing = ft.ElevatedButton("Activar", on_click=lambda e, cid=c['id']: activate(cid), bgcolor=COLOR_WARNING, color="white")
                else:
                    trailing = ft.Text("ACTIVO", color=ft.colors.GREEN_600, weight=ft.FontWeight.BOLD)

                card = ft.Container(
                    content=ft.ListTile(leading=ft.Icon(icon, color=color), title=ft.Text(c['nombre'], weight=ft.FontWeight.BOLD), trailing=trailing), 
                    bgcolor="white", border_radius=10, margin=ft.margin.only(bottom=5), shadow=ft.BoxShadow(blur_radius=2, color=ft.colors.BLACK12)
                )
                list_col.controls.append(card)
            page.update()

        def add(e):
            if tf_new.value:
                add_ciclo(tf_new.value); tf_new.value = ""; load(); show_snack("Ciclo creado y activado")
        
        def activate(cid):
            activar_ciclo(cid); load(); show_snack("Ciclo cambiado correctamente")

        load()
        return ft.View("/ciclos", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/admin")),
                      title=ft.Text("Gestión Ciclos Lectivos"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(
                content=ft.Column([
                    ft.Text("Crear Nuevo Ciclo (Cierra el actual)", weight=ft.FontWeight.BOLD), 
                    ft.Row([tf_new, ft.IconButton(ft.icons.ADD_CIRCLE, icon_color=ft.colors.GREEN_600, icon_size=40, on_click=add)]), 
                    ft.Divider(), 
                    ft.Text("Historial de Ciclos", weight=ft.FontWeight.BOLD), 
                    list_col
                ]), 
                padding=20, bgcolor=COLOR_BG, expand=True
            )
        ])

    def users_view():
        users_col = ft.Column()
        
        def load():
            users_col.controls.clear()
            for u in get_users():
                users_col.controls.append(ft.Container(
                    content=ft.ListTile(
                        leading=ft.Icon(ft.icons.PERSON), 
                        title=ft.Text(u['username']), 
                        subtitle=ft.Text(u['role']), 
                        trailing=ft.PopupMenuButton(items=[ft.PopupMenuItem(text="Eliminar", on_click=lambda e, uid=u['id']: rem(uid))]) if u['username'] != state['username'] else None
                    ), 
                    bgcolor="white", border_radius=10, margin=ft.margin.only(bottom=2), shadow=ft.BoxShadow(blur_radius=2, color=ft.colors.BLACK12)
                ))
            page.update()
            
        def add(e): 
            if add_user(u_tf.value, p_tf.value, r_dd.value): 
                u_tf.value=""; p_tf.value=""; 
                load()
                show_snack("Usuario creado con éxito")
            else: show_snack("Error: Usuario ya existe o campos vacíos", COLOR_DANGER)
            
        def rem(uid): 
            delete_user(uid)
            load()
            show_snack("Usuario eliminado", COLOR_DANGER)
            
        u_tf = ft.TextField(label="Usuario", expand=True, bgcolor="white", border_radius=10)
        p_tf = ft.TextField(label="Clave", password=True, expand=True, bgcolor="white", border_radius=10)
        r_dd = ft.Dropdown(options=[ft.dropdown.Option("preceptor"), ft.dropdown.Option("admin")], value="preceptor", width=120, bgcolor="white", border_radius=10)
        
        load()
        return ft.View("/users", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/admin")), title=ft.Text("Gestión Usuarios"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(
                content=ft.Column([
                    ft.Row([u_tf, p_tf, r_dd, ft.IconButton(ft.icons.ADD, on_click=add, icon_color=ft.colors.GREEN_600, icon_size=30)]), 
                    ft.Divider(), 
                    users_col
                ]), 
                padding=15, bgcolor=COLOR_BG, expand=True)
        ])

    def search_view():
        term = state["search_term"]
        results_col = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
        res = search_students(term)
        
        if not res: 
            results_col.controls.append(ft.Text("No se encontraron resultados."))
        else:
            for r in res:
                card = ft.Container(
                    content=ft.ListTile(
                        leading=ft.Icon(ft.icons.PERSON), 
                        title=ft.Text(r['nombre'], weight=ft.FontWeight.BOLD), 
                        subtitle=ft.Text(f"DNI: {r.get('dni', '-')} | Curso: {r['curso_nombre']} ({r['ciclo_nombre']})"), 
                        on_click=lambda e, s=r: go_detail(s)
                    ), 
                    bgcolor="white", border_radius=10, margin=ft.margin.only(bottom=5), shadow=ft.BoxShadow(blur_radius=3, color=ft.colors.BLACK12)
                )
                results_col.controls.append(card)
                
        def go_detail(s): 
            state["student_id_view"] = s['id']
            state["curso_id"] = s['curso_id']
            page.go("/student_detail")
            
        return ft.View("/search", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(f"Búsqueda: {term}"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(content=results_col, padding=15, expand=True, bgcolor=COLOR_BG)
        ])

    def student_detail_view():
        aid = state["student_id_view"]
        student = get_alumno_by_id(aid)
        
        if not student:
            return ft.View("/student_detail", [
                ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/search")), title=ft.Text("Error"), bgcolor=COLOR_PRIMARY, color="white"),
                ft.Container(content=ft.Text("Alumno no encontrado."), padding=20, expand=True, bgcolor=COLOR_BG)
            ])
            
        reqs = get_student_req_status(aid, student['curso_id'])
        req_col = ft.Column()
        for r in reqs: 
            req_col.controls.append(ft.Row([
                ft.Icon(ft.icons.CHECK_CIRCLE if r['ok'] else ft.icons.CANCEL, color=ft.colors.GREEN_600 if r['ok'] else COLOR_DANGER), 
                ft.Text(r['desc'])
            ]))
            
        card = ft.Container(
            content=ft.Column([
                ft.Text(student['nombre'], size=24, weight=ft.FontWeight.BOLD), 
                ft.Text(f"Curso: {student['curso_nombre']}", size=16), 
                ft.Text(f"DNI: {student.get('dni', 'No registrado')}", size=16), 
                ft.Divider(), 
                ft.Text("Datos de Tutor:", weight=ft.FontWeight.BOLD),
                ft.Text(f"Nombre: {student.get('tutor_nombre', '-')}", size=14),
                ft.Text(f"Teléfono: {student.get('tutor_telefono', '-')}", size=14),
                ft.Divider(),
                ft.Text("Observaciones:", weight=ft.FontWeight.BOLD), 
                ft.Text(student.get('observaciones', '-'), italic=True), 
                ft.Divider(), 
                ft.Text("Documentación:", weight=ft.FontWeight.BOLD), 
                req_col
            ]), 
            padding=20, bgcolor="white", border_radius=15, shadow=ft.BoxShadow(blur_radius=10, color=ft.colors.BLACK12)
        )
        
        return ft.View("/student_detail", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/search")), title=ft.Text("Ficha de Alumno"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(content=ft.Column([card], scroll=ft.ScrollMode.AUTO), padding=20, expand=True, bgcolor=COLOR_BG)
        ])

    def form_curso_view():
        tf = ft.TextField(label="Nombre del Curso", bgcolor="white", border_radius=10)
        def save(e): 
            if add_curso(tf.value): page.go("/dashboard")
            else: show_snack("Error: El curso ya existe o no hay ciclo activo.", COLOR_DANGER)
        return ft.View("/form_curso", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text("Nuevo Curso"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(content=ft.Column([tf, ft.ElevatedButton("Guardar", on_click=save, bgcolor=ft.colors.GREEN_600, color="white")]), padding=20, bgcolor=COLOR_BG, expand=True)
        ])

    def curso_view():
        alumnos_col = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
        
        def load_alumnos():
            alumnos_col.controls.clear()
            for a in get_alumnos(state["curso_id"]):
                alumnos_col.controls.append(ft.Container(
                    content=ft.ListTile(
                        leading=ft.Icon(ft.icons.PERSON), 
                        title=ft.Text(a['nombre']), 
                        subtitle=ft.Text(f"DNI: {a.get('dni', '-')}"), 
                        trailing=ft.PopupMenuButton(items=[
                            ft.PopupMenuItem(text="Editar", on_click=lambda e, aid=a['id']: go_edit(aid)), 
                            ft.PopupMenuItem(text="Eliminar", on_click=lambda e, aid=a['id']: del_s(aid), icon=ft.icons.DELETE, icon_color=COLOR_DANGER)
                        ])
                    ), 
                    bgcolor="white", border_radius=10, margin=ft.margin.only(bottom=2), shadow=ft.BoxShadow(blur_radius=2, color=ft.colors.BLACK12)
                ))
            page.update()
            
        def go_edit(aid): 
            state["student_id_edit"] = aid
            page.go("/form_student")
            
        def go_add(e): 
            state["student_id_edit"] = None
            page.go("/form_student")
            
        def del_s(aid): 
            delete_alumno(aid)
            load_alumnos()
            show_snack("Alumno eliminado.", COLOR_DANGER)
            
        load_alumnos()
        
        return ft.View("/curso", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(state["curso_nombre"]), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.ElevatedButton("Asistencia", icon=ft.icons.CHECK, on_click=lambda _: page.go("/asistencia"), expand=True, bgcolor=ft.colors.GREEN_600, color="white"), 
                        ft.ElevatedButton("Pedidos", icon=ft.icons.LIST, on_click=lambda _: page.go("/pedidos"), expand=True, bgcolor=ft.colors.ORANGE_600, color="white"), 
                        ft.ElevatedButton("Reportes", icon=ft.icons.BAR_CHART, on_click=lambda _: page.go("/reportes"), expand=True, bgcolor=ft.colors.CYAN_600, color="white")
                    ]), 
                    ft.Divider(), 
                    ft.Row([
                        ft.Text("Alumnos", size=18, weight=ft.FontWeight.BOLD), 
                        ft.IconButton(ft.icons.PERSON_ADD, icon_color=ft.colors.GREEN_600, on_click=go_add)
                    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN), 
                    alumnos_col
                ]), 
                padding=15, expand=True, bgcolor=COLOR_BG
            )
        ])

    def form_student_view():
        is_edit = state["student_id_edit"] is not None
        
        # Nuevos campos de tutor implementados (Punto 4)
        name = ft.TextField(label="Nombre", bgcolor="white", border_radius=10)
        dni = ft.TextField(label="DNI", bgcolor="white", border_radius=10)
        obs = ft.TextField(label="Observaciones", multiline=True, bgcolor="white", border_radius=10)
        tutor_n = ft.TextField(label="Nombre de Tutor", bgcolor="white", border_radius=10)
        tutor_t = ft.TextField(label="Teléfono de Tutor", bgcolor="white", border_radius=10)
        
        if is_edit:
            data = get_alumno_by_id(state["student_id_edit"])
            if data:
                name.value = data.get('nombre', '')
                dni.value = data.get('dni', '')
                obs.value = data.get('observaciones', '')
                tutor_n.value = data.get('tutor_nombre', '')
                tutor_t.value = data.get('tutor_telefono', '')
        
        def save(e):
            if name.value:
                if is_edit: 
                    update_alumno(state["student_id_edit"], name.value, dni.value, obs.value, tutor_n.value, tutor_t.value)
                else: 
                    add_alumno(state["curso_id"], name.value, dni.value, obs.value, tutor_n.value, tutor_t.value)
                page.go("/curso")
            else: 
                show_snack("El nombre del alumno es obligatorio.", COLOR_DANGER)
                
        return ft.View("/form_student", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Alumno"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(
                content=ft.Column([
                    ft.Text("Datos del Alumno:", weight=ft.FontWeight.BOLD), name, dni, obs, 
                    ft.Divider(),
                    ft.Text("Datos del Tutor:", weight=ft.FontWeight.BOLD), tutor_n, tutor_t,
                    ft.ElevatedButton("Guardar", on_click=save, bgcolor=ft.colors.GREEN_600, color="white")
                ], scroll=ft.ScrollMode.AUTO), 
                padding=20, bgcolor=COLOR_BG, expand=True)
        ])

    def asistencia_view():
        date_pk = ft.TextField(label="Fecha (AAAA-MM-DD)", value=date.today().isoformat(), bgcolor="white", border_radius=10)
        list_view = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
        status_vars = {} 
        
        def load_list(e=None):
            try:
                date.fromisoformat(date_pk.value)
            except ValueError:
                show_snack("Formato de fecha inválido. Use AAAA-MM-DD", COLOR_DANGER)
                return
            
            existing = get_asistencia_diaria(state["curso_id"], date_pk.value)
            list_view.controls.clear(); status_vars.clear()
            
            for a in get_alumnos(state["curso_id"]):
                dd = ft.Dropdown(options=[ft.dropdown.Option(x) for x in ["P", "T", "A", "J", "S", "N"]], 
                                 value=existing.get(a['id'], "P"), 
                                 width=80, bgcolor="white", border_radius=10)
                status_vars[a['id']] = dd
                
                list_view.controls.append(ft.Container(
                    content=ft.Row([ft.Text(a['nombre'], expand=True), dd]), 
                    padding=10, bgcolor="white", border_radius=5, margin=ft.margin.only(bottom=2), shadow=ft.BoxShadow(blur_radius=1, color=ft.colors.BLACK12)
                ))
            page.update()
        
        def save(e):
            try:
                d = date.fromisoformat(date_pk.value)
                if d > date.today(): show_snack("Fecha futura no permitida", COLOR_WARNING); return
            except ValueError:
                show_snack("Formato de fecha inválido. Use AAAA-MM-DD", COLOR_DANGER); return

            for aid, dd in status_vars.items(): 
                register_asistencia(aid, date_pk.value, dd.value)
            show_snack("Asistencia guardada en Firestore", ft.colors.BLUE_GREY_800); page.go("/curso")

        load_list()
        return ft.View("/asistencia", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Asistencia"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(
                content=ft.Column([
                    ft.Row([date_pk, ft.IconButton(ft.icons.REFRESH, on_click=load_list)]), 
                    ft.ElevatedButton("GUARDAR ASISTENCIA", on_click=save, bgcolor=ft.colors.GREEN_600, color="white", width=float("inf")), 
                    ft.Divider(), 
                    list_view
                ]), 
                padding=15, bgcolor=COLOR_BG, expand=True)
        ])

    def pedidos_view():
        req_dd = ft.Dropdown(label="Seleccionar Pedido", expand=True, bgcolor="white", border_radius=10, on_change=lambda e: load_checks())
        list_view = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
        req_map = {} 
        
        def load_reqs():
            reqs = get_requisitos(state["curso_id"])
            req_map.clear(); req_dd.options.clear()
            for r in reqs: 
                req_map[r['descripcion']] = r['id']
                req_dd.options.append(ft.dropdown.Option(r['descripcion']))
            
            if reqs: 
                req_dd.value = reqs[0]['descripcion']
            else: 
                req_dd.value = None
            
            page.update()
            load_checks()
            
        def load_checks():
            list_view.controls.clear()
            if not req_dd.value or req_dd.value not in req_map: 
                list_view.controls.append(ft.Text("Selecciona o agrega un requisito."))
                page.update()
                return
            
            rid = req_map[req_dd.value]
            done = get_cumplimientos(rid)
            
            for a in get_alumnos(state["curso_id"]):
                def on_change(e, aid=a['id'], rid=rid): 
                    toggle_cumplimiento(rid, aid, e.control.value)
                    
                list_view.controls.append(ft.Container(
                    content=ft.Checkbox(label=a['nombre'], value=(a['id'] in done), on_change=on_change), 
                    bgcolor="white", padding=10, border_radius=5, margin=ft.margin.only(bottom=2), shadow=ft.BoxShadow(blur_radius=1, color=ft.colors.BLACK12)
                ))
            page.update()
            
        def go_add(e): page.go("/form_requisito")
        
        def del_r(e): 
            if req_dd.value and req_dd.value in req_map: 
                delete_requisito(req_map[req_dd.value])
                load_reqs()
                show_snack("Pedido eliminado", COLOR_DANGER)
            
        load_reqs()
        
        return ft.View("/pedidos", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Pedidos (Documentación)"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(
                content=ft.Column([
                    ft.Row([
                        req_dd, 
                        ft.IconButton(ft.icons.ADD, on_click=go_add, icon_color=ft.colors.GREEN_600), 
                        ft.IconButton(ft.icons.DELETE, icon_color=COLOR_DANGER, on_click=del_r)
                    ]), 
                    ft.Divider(), 
                    list_view
                ]), 
                padding=15, bgcolor=COLOR_BG, expand=True)
        ])

    def form_requisito_view():
        tf = ft.TextField(label="Descripción del Requisito/Documento", bgcolor="white", border_radius=10)
        def save(e):
            if tf.value: add_requisito(state["curso_id"], tf.value); page.go("/pedidos")
        return ft.View("/form_requisito", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/pedidos")), title=ft.Text("Nuevo Pedido"), bgcolor=COLOR_PRIMARY, color="white"),
            ft.Container(content=ft.Column([tf, ft.ElevatedButton("Crear", on_click=save, bgcolor=ft.colors.GREEN_600, color="white")]), padding=20, bgcolor=COLOR_BG, expand=True)
        ])

    def reportes_view():
        # Punto 3: Periodos más amplios (inicio en el año actual)
        today = date.today()
        d1 = ft.TextField(label="Desde (AAAA-MM-DD)", value=today.replace(month=1, day=1).isoformat(), width=150, bgcolor="white", border_radius=10)
        d2 = ft.TextField(label="Hasta (AAAA-MM-DD)", value=today.isoformat(), width=150, bgcolor="white", border_radius=10)
        table_cont = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
        
        def gen(e):
            data = get_report_data(state["curso_id"], d1.value, d2.value)
            rows = []
            for d in data:
                # Punto 4: Alerta de 25 Faltas
                faltas = d['faltas']
                alert_color = COLOR_DANGER if faltas >= 25 else (COLOR_WARNING if faltas >= 15 else ft.colors.BLACK)
                
                rows.append(ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(d['nombre'], color=alert_color, weight=ft.FontWeight.BOLD if faltas >= 25 else ft.FontWeight.NORMAL)), 
                        ft.DataCell(ft.Text(str(d['p']), size=12)), 
                        ft.DataCell(ft.Text(str(d['t']), size=12)), 
                        ft.DataCell(ft.Text(str(d['a']), size=12)), 
                        ft.DataCell(ft.Text(str(d['j']), size=12)), 
                        ft.DataCell(ft.Text(str(d['s']), size=12)), 
                        ft.DataCell(ft.Text(f"{faltas:.2f}", color=alert_color, size=14, weight=ft.FontWeight.BOLD)), 
                        ft.DataCell(ft.Text(f"{d['pct']:.1f}%", color=alert_color, size=14, weight=ft.FontWeight.BOLD))
                    ]
                ))
            
            table = ft.DataTable(
                columns=[
                    ft.DataColumn(ft.Text("Alumno", weight=ft.FontWeight.BOLD)), 
                    ft.DataColumn(ft.Tooltip(message="Presentes", content=ft.Text("P"), height=25), numeric=True), 
                    ft.DataColumn(ft.Tooltip(message="Tardes (0.25F)", content=ft.Text("T"), height=25), numeric=True), 
                    ft.DataColumn(ft.Tooltip(message="Ausentes", content=ft.Text("A"), height=25), numeric=True), 
                    ft.DataColumn(ft.Tooltip(message="Justificadas", content=ft.Text("J"), height=25), numeric=True), 
                    ft.DataColumn(ft.Tooltip(message="Suspensiones", content=ft.Text("S"), height=25), numeric=True), 
                    ft.DataColumn(ft.Text("Faltas", color=COLOR_DANGER, weight=ft.FontWeight.BOLD), numeric=True), 
                    ft.DataColumn(ft.Text("% Aus.", color=COLOR_DANGER, weight=ft.FontWeight.BOLD), numeric=True)
                ], 
                rows=rows, 
                bgcolor="white", 
                border_radius=10, 
                column_spacing=10
            )
            table_cont.controls = [ft.Row([table], scroll=ft.ScrollMode.ALWAYS)]; page.update()
            
        def export(e):
            # Punto 1: La exportación de Excel necesita reescribirse para la web
            if not pd: 
                show_snack("Error: 'pandas' no instalado. No se puede exportar.", COLOR_DANGER)
                return
            
            try:
                data = get_report_data(state["curso_id"], d1.value, d2.value)
                if not data:
                    show_snack("No hay datos para exportar.", COLOR_WARNING)
                    return
                
                df = pd.DataFrame(data)
                df = df.rename(columns={'nombre': 'Alumno', 'dni': 'DNI', 'tutor_nombre': 'Tutor', 'tutor_telefono': 'Teléfono Tutor',
                                        'p': 'Presentes', 't': 'Tardes', 'a': 'Ausentes', 'j': 'Justificadas', 
                                        's': 'Suspensiones', 'faltas': 'Total Faltas', 'pct': '% Ausentismo'})
                
                # --- Lógica de Descarga Web (Usando io y page.launch_url) ---
                import io
                
                output = io.BytesIO()
                # Guardar el DataFrame en el buffer en formato Excel
                df.to_excel(output, index=False, engine='xlsxwriter')
                output.seek(0)
                
                # Convertir el buffer a base64 para incrustarlo en una URL de datos
                b64 = base64.b64encode(output.read()).decode()
                
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                fname = f"reporte_asistencia_{state['curso_nombre']}_{today.isoformat()}.xlsx"

                # Abrir una nueva ventana/pestaña con el archivo para forzar la descarga
                page.launch_url(f"data:{mime_type};base64,{b64}", web_window_name=fname)
                show_snack("¡Exportación lista! Revisa las descargas de tu navegador.", ft.colors.BLUE_GREY_800)
                
            except Exception as ex:
                show_snack(f"Error al exportar: {ex}", COLOR_DANGER)

        gen(None) # Generar al cargar
        
        return ft.View("/reportes", [
            ft.AppBar(leading=ft.IconButton(ft.icons.ARROW_BACK, icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Reportes"), bgcolor=COLOR_PRIMARY, color="white"), 
            ft.Container(
                content=ft.Column([
                    ft.Row([d1, d2, ft.ElevatedButton("Generar Reporte", on_click=gen, icon=ft.icons.CALCULATE, bgcolor=COLOR_PRIMARY, color="white")]), 
                    ft.ElevatedButton("Exportar Excel (Descargar)", icon=ft.icons.DOWNLOAD, on_click=export, bgcolor=ft.colors.ORANGE_600, color="white"), 
                    ft.Divider(), 
                    table_cont
                ]), 
                padding=15, bgcolor=COLOR_BG, expand=True)
        ])

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
    # La App ID se usa en get_collection_ref
    port = int(os.environ.get("PORT", 8000))
    # Importante: usar web_renderer="html" para asegurar compatibilidad en entornos serverless/Micro
    ft.app(target=main, view=ft.WEB_BROWSER, port=port, web_renderer="html")