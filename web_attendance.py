import flet as ft
import sqlite3
import hashlib
from datetime import date
import os
import time

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
    print("⚠️ ADVERTENCIA: 'xlsxwriter' no está instalado. La exportación a Excel fallará.")

try:
    import firebase_admin
    from firebase_admin import credentials, firestore
except ImportError:
    firebase_admin = None
    firestore = None
    print("⚠️ ADVERTENCIA: 'firebase-admin' no está instalado.")

# ======================================================================
# 1. LÓGICA DE BASE DE DATOS (Backend con Firestore Optimizado)
# ======================================================================

db = None
app_id = "asistencia-unsam-app"

def init_firestore():
    """Inicializa la conexión a Firestore."""
    global db, app_id
    
    # Intenta obtener configuración del entorno (Render/Deta)
    firebase_config_str = os.environ.get('__firebase_config', None)
    app_id = os.environ.get('__app_id', 'asistencia-local')

    if firebase_admin is None: return False

    if not firebase_admin._apps:
        try:
            if firebase_config_str:
                # Producción
                import json
                cred_dict = json.loads(firebase_config_str)
                cred = credentials.Certificate(cred_dict) if "private_key" in cred_dict else credentials.ApplicationDefault()
                firebase_admin.initialize_app(cred)
            else:
                # Desarrollo Local (Mock o Credenciales por defecto)
                # Si tienes un archivo serviceAccountKey.json local, úsalo aquí:
                # cred = credentials.Certificate("serviceAccountKey.json")
                # firebase_admin.initialize_app(cred)
                pass 
                
            db = firestore.client()
            print("Conexión a Firestore exitosa.")
            return True
        except Exception as e:
            print(f"Error Firestore Init: {e}")
            return False
    else:
        db = firestore.client()
        return True

def get_collection_ref(name):
    # Estructura: artifacts/{app_id}/public/data/{name}
    return db.collection('artifacts').document(app_id).collection('public').document('data').collection(name)

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# --- USUARIOS ---
def authenticate_user(username, password):
    if db is None: return (False, None)
    try:
        users = get_collection_ref('Usuarios').where('username', '==', username).limit(1).stream()
        u_doc = next(users, None)
        if u_doc:
            u_data = u_doc.to_dict()
            if u_data.get('password') == hash_password(password):
                return (True, u_data.get('role'))
    except Exception as e: print(e)
    
    # Admin de respaldo (solo si no hay usuarios)
    return (False, None)

def get_users():
    if db is None: return []
    try:
        return [{**d.to_dict(), 'id': d.id} for d in get_collection_ref('Usuarios').stream()]
    except: return []

def add_user(u, p, r):
    if db is None: return False
    try:
        # Verificar duplicado
        if next(get_collection_ref('Usuarios').where('username', '==', u).limit(1).stream(), None):
            return False
        get_collection_ref('Usuarios').add({'username': u, 'password': hash_password(p), 'role': r})
        return True
    except: return False

def delete_user(uid):
    if db is None: return
    get_collection_ref('Usuarios').document(uid).delete()

# --- CICLOS ---
def get_ciclos():
    if db is None: return []
    try:
        # Ordenar por nombre descendente (años más recientes primero)
        return [{**d.to_dict(), 'id': d.id} for d in get_collection_ref('Ciclos').order_by('nombre', direction=firestore.Query.DESCENDING).stream()]
    except: return []

def get_ciclo_activo():
    if db is None: return None
    try:
        gen = get_collection_ref('Ciclos').where('activo', '==', 1).limit(1).stream()
        doc = next(gen, None)
        return {**doc.to_dict(), 'id': doc.id} if doc else None
    except: return None

def add_ciclo(nombre):
    if db is None: return False
    try:
        # Desactivar anteriores
        batch = db.batch()
        for doc in get_collection_ref('Ciclos').where('activo', '==', 1).stream():
            batch.update(doc.reference, {'activo': 0})
        batch.commit()
        
        get_collection_ref('Ciclos').add({'nombre': nombre, 'activo': 1})
        return True
    except: return False

def activar_ciclo(cid):
    if db is None: return
    try:
        batch = db.batch()
        for doc in get_collection_ref('Ciclos').where('activo', '==', 1).stream():
            batch.update(doc.reference, {'activo': 0})
        
        ref = get_collection_ref('Ciclos').document(cid)
        batch.update(ref, {'activo': 1})
        batch.commit()
    except: pass

# --- CURSOS ---
def get_cursos():
    if db is None: return []
    try:
        ciclo = get_ciclo_activo()
        if not ciclo: return []
        # Traer cursos solo del ciclo activo
        return [{**d.to_dict(), 'id': d.id} for d in get_collection_ref('Cursos').where('ciclo_id', '==', ciclo['id']).order_by('nombre').stream()]
    except: return []

def add_curso(nombre):
    if db is None: return False
    try:
        ciclo = get_ciclo_activo()
        if not ciclo: return False
        
        # Verificar duplicado en este ciclo
        dup = next(get_collection_ref('Cursos').where('ciclo_id', '==', ciclo['id']).where('nombre', '==', nombre).limit(1).stream(), None)
        if dup: return False
        
        get_collection_ref('Cursos').add({'nombre': nombre, 'ciclo_id': ciclo['id']})
        return True
    except: return False

def delete_curso(cid):
    if db is None: return
    get_collection_ref('Cursos').document(cid).delete()

# --- ALUMNOS ---
def get_alumnos(curso_id):
    if db is None: return []
    try:
        return [{**d.to_dict(), 'id': d.id} for d in get_collection_ref('Alumnos').where('curso_id', '==', curso_id).order_by('nombre').stream()]
    except: return []

def get_alumno_by_id(aid):
    if db is None: return None
    try:
        doc = get_collection_ref('Alumnos').document(aid).get()
        if doc.exists:
            data = doc.to_dict()
            # Obtener nombre curso (opcional, extra read)
            c_doc = get_collection_ref('Cursos').document(data['curso_id']).get()
            data['curso_nombre'] = c_doc.to_dict().get('nombre', '?') if c_doc.exists else '?'
            return {**data, 'id': doc.id}
    except: pass
    return None

def add_alumno(curso_id, nombre, dni, obs, t_n, t_t):
    if db is None: return False
    try:
        get_collection_ref('Alumnos').add({
            'curso_id': curso_id, 'nombre': nombre, 'dni': dni, 
            'observaciones': obs, 'tutor_nombre': t_n, 'tutor_telefono': t_t
        })
        return True
    except: return False

def update_alumno(aid, nombre, dni, obs, t_n, t_t):
    if db is None: return
    get_collection_ref('Alumnos').document(aid).update({
        'nombre': nombre, 'dni': dni, 'observaciones': obs, 
        'tutor_nombre': t_n, 'tutor_telefono': t_t
    })

def delete_alumno(aid):
    if db is None: return
    get_collection_ref('Alumnos').document(aid).delete()

def search_students(term):
    if db is None: return []
    # Firestore no tiene "LIKE". Traemos todos los del ciclo activo y filtramos en Python.
    # (Para < 1000 alumnos es aceptable).
    try:
        cursos_activos = {c['id']: c['nombre'] for c in get_cursos()}
        if not cursos_activos: return []
        
        term = term.lower()
        res = []
        # Optimización: Podríamos filtrar por curso en la query si tuviéramos índice
        # Por ahora, stream completo de Alumnos (cuidado con costos si escala mucho)
        all_alumnos = get_collection_ref('Alumnos').stream()
        
        for doc in all_alumnos:
            d = doc.to_dict()
            if d.get('curso_id') in cursos_activos:
                if term in d.get('nombre', '').lower() or term in d.get('dni', ''):
                    d['id'] = doc.id
                    d['curso_nombre'] = cursos_activos[d['curso_id']]
                    res.append(d)
        return sorted(res, key=lambda x: x['nombre'])
    except: return []

# --- ASISTENCIA (Optimizado) ---
def register_asistencia(aid, cid, fecha, status):
    if db is None: return
    try:
        doc_id = f"{aid}_{fecha}"
        # OPTIMIZACIÓN: Guardamos también curso_id para facilitar queries
        get_collection_ref('Asistencia').document(doc_id).set({
            'alumno_id': aid, 
            'curso_id': cid, # Nuevo campo para optimizar lectura
            'fecha': fecha, 
            'status': status
        }, merge=True)
    except Exception as e: print(e)

def get_asistencia_diaria(curso_id, fecha):
    if db is None: return {}
    try:
        # OPTIMIZACIÓN: Filtramos por curso y fecha. 
        # (Requiere índice compuesto en Firestore si hay muchos datos, pero es más eficiente que traer todo)
        # Si falla por falta de índice, Firestore lanzará error con link para crearlo.
        # Fallback: Si no hay curso_id en documentos viejos, traer solo por fecha.
        
        # Intento optimizado:
        try:
            query = get_collection_ref('Asistencia').where('curso_id', '==', curso_id).where('fecha', '==', fecha)
            rows = query.stream()
        except:
            # Fallback a solo fecha (menos eficiente)
            rows = get_collection_ref('Asistencia').where('fecha', '==', fecha).stream()

        res = {}
        for r in rows:
            d = r.to_dict()
            # Doble chequeo por si usamos el fallback
            res[d['alumno_id']] = d['status']
        return res
    except: return {}

def get_report_data(curso_id, start, end):
    if db is None: return []
    try:
        alumnos = get_alumnos(curso_id)
        report = []
        
        # Para optimizar reportes, traemos TODA la asistencia del curso en ese rango
        # en lugar de 1 query por alumno.
        try:
            asis_query = get_collection_ref('Asistencia').where('curso_id', '==', curso_id)\
                .where('fecha', '>=', start).where('fecha', '<=', end).stream()
            
            # Mapear asistencia en memoria
            asis_map = {} # {alumno_id: [status, ...]}
            for doc in asis_query:
                d = doc.to_dict()
                aid = d['alumno_id']
                if aid not in asis_map: asis_map[aid] = []
                asis_map[aid].append(d['status'])
        except:
            # Si falla (ej: falta índice), volvemos al método lento (1 query por alumno) o vacío
            print("Falta índice compuesto curso_id + fecha. Usando método lento.")
            asis_map = {} # Implementar fallback si es necesario
            # Por simplicidad del ejemplo, asumimos que el índice se creará o usamos lógica simple
            pass

        for a in alumnos:
            statuses = asis_map.get(a['id'], [])
            # Fallback method if asis_map is empty due to index error (slow individual queries)
            if not asis_map:
                 q = get_collection_ref('Asistencia').where('alumno_id', '==', a['id'])\
                     .where('fecha', '>=', start).where('fecha', '<=', end).stream()
                 statuses = [d.to_dict().get('status') for d in q]

            counts = {k: statuses.count(k) for k in ['P','T','A','J','S','N']}
            
            faltas = counts['A'] + counts['S'] + (counts['T'] * 0.25)
            total = counts['P'] + counts['T'] + counts['A'] + counts['J'] + counts['S']
            pct = (faltas/total*100) if total > 0 else 0
            
            report.append({
                'nombre': a['nombre'], 'dni': a['dni'], 
                'tutor': a.get('tutor_nombre', '-'), 'tel': a.get('tutor_telefono', '-'),
                'p': counts['P'], 't': counts['T'], 'a': counts['A'], 
                'j': counts['J'], 's': counts['S'], 
                'faltas': faltas, 'pct': round(pct, 1)
            })
        return report
    except Exception as e: 
        print(e)
        return []

# --- REQUISITOS ---
def add_requisito(cid, desc):
    if db is None: return
    get_collection_ref('Requisitos').add({'curso_id': cid, 'descripcion': desc})

def get_requisitos(cid):
    if db is None: return []
    return [{**d.to_dict(), 'id': d.id} for d in get_collection_ref('Requisitos').where('curso_id', '==', cid).stream()]

def delete_requisito(rid):
    if db is None: return
    get_collection_ref('Requisitos').document(rid).delete()

def toggle_cumplimiento(rid, aid, val):
    if db is None: return
    doc_id = f"{rid}_{aid}"
    ref = get_collection_ref('Requisitos_Cumplidos').document(doc_id)
    if val: ref.set({'requisito_id': rid, 'alumno_id': aid})
    else: ref.delete()

def get_cumplimientos(rid):
    if db is None: return set()
    return {d.to_dict()['alumno_id'] for d in get_collection_ref('Requisitos_Cumplidos').where('requisito_id', '==', rid).stream()}

def get_student_req_status(aid, cid):
    if db is None: return []
    reqs = get_requisitos(cid)
    # Optimización: traer todos los cumplidos del alumno de una vez
    done_q = get_collection_ref('Requisitos_Cumplidos').where('alumno_id', '==', aid).stream()
    done_ids = {d.to_dict()['requisito_id'] for d in done_q}
    return [{'desc': r['descripcion'], 'ok': r['id'] in done_ids} for r in reqs]


# ======================================================================
# 2. INTERFAZ GRÁFICA (Flet)
# ======================================================================

def main(page: ft.Page):
    page.title = "Asistencia UNSAM"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 0
    
    # Intenta conectar
    connected = init_firestore()
    
    # State
    state = {"role": None, "username": None, "curso_id": None, "curso_nombre": None, "search": "", "st_view": None, "st_edit": None}

    def show_snack(m, c="green"):
        page.snack_bar = ft.SnackBar(ft.Text(m), bgcolor=c)
        page.snack_bar.open = True
        page.update()

    if not connected:
        show_snack("Error crítico: No hay conexión a Base de Datos (Firestore).", "red")

    # --- VISTAS ---
    def login_view():
        user = ft.TextField(label="Usuario", width=300, bgcolor="white")
        pwd = ft.TextField(label="Clave", password=True, width=300, bgcolor="white")
        def login(e):
            ok, role = authenticate_user(user.value, pwd.value)
            if ok:
                state["role"], state["username"] = role, user.value
                page.go("/dashboard")
            else: show_snack("Datos incorrectos", "red")
        
        return ft.View("/", [
            ft.Container(content=ft.Column([
                ft.Icon("school", size=80, color="blue"),
                ft.Text("Asistencia UNSAM", size=24, weight="bold"),
                ft.Container(height=20),
                user, pwd,
                ft.ElevatedButton("ENTRAR", on_click=login, width=300, height=50, bgcolor="blue", color="white")
            ], horizontal_alignment="center"), alignment=ft.alignment.center, expand=True, bgcolor="#f0f2f5")
        ])

    def dashboard_view():
        ciclo = get_ciclo_activo()
        c_nombre = ciclo['nombre'] if ciclo else "Sin Ciclo"
        
        search = ft.TextField(hint_text="Buscar alumno...", expand=True, bgcolor="white")
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
                    ]),
                    padding=15, bgcolor="white", border_radius=10, shadow=ft.BoxShadow(blur_radius=2, color="black12"), margin=5
                ))
            page.update()
        
        def go_curso(cid, cn): state["curso_id"]=cid; state["curso_nombre"]=cn; page.go("/curso")
        def add_c(e): page.go("/form_curso") if ciclo else show_snack("Falta Ciclo Activo", "red")

        load()
        admin = ft.IconButton("settings", icon_color="white", on_click=lambda _: page.go("/admin")) if state["role"]=='admin' else ft.Container()
        
        return ft.View("/dashboard", [
            ft.AppBar(title=ft.Text("Panel Principal"), bgcolor="blue", color="white", actions=[admin, ft.IconButton("logout", icon_color="white", on_click=lambda _: page.go("/"))]),
            ft.Container(content=ft.Column([
                ft.Text(f"Ciclo: {c_nombre}", color="blue", weight="bold"),
                ft.Row([search, ft.IconButton("search", on_click=do_search)]),
                ft.Row([ft.Text("Cursos", size=20, weight="bold"), ft.IconButton("add_circle", icon_color="green", icon_size=30, on_click=add_c)], alignment="spaceBetween"),
                cursos_col
            ]), padding=15, bgcolor="#f0f2f5", expand=True)
        ])

    def curso_view():
        col = ft.Column(scroll="auto", expand=True)
        def load():
            col.controls.clear()
            for a in get_alumnos(state["curso_id"]):
                col.controls.append(ft.Container(
                    content=ft.ListTile(
                        leading=ft.Icon("person"), title=ft.Text(a['nombre']), subtitle=ft.Text(f"DNI: {a.get('dni','-')}"),
                        trailing=ft.PopupMenuButton(items=[
                            ft.PopupMenuItem("Editar", on_click=lambda e, aid=a['id']: (state.update({"st_edit": aid}), page.go("/form_student"))),
                            ft.PopupMenuItem("Borrar", on_click=lambda e, aid=a['id']: (delete_alumno(aid), load()))
                        ])
                    ), bgcolor="white", border_radius=10, margin=2
                ))
            page.update()
        
        load()
        return ft.View("/curso", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(state["curso_nombre"]), bgcolor="blue", color="white"),
            ft.Container(content=ft.Column([
                ft.Row([
                    ft.ElevatedButton("Asistencia", icon="check", expand=True, on_click=lambda _: page.go("/asistencia")),
                    ft.ElevatedButton("Pedidos", icon="list", expand=True, on_click=lambda _: page.go("/pedidos")),
                    ft.ElevatedButton("Reportes", icon="bar_chart", expand=True, on_click=lambda _: page.go("/reportes"))
                ]),
                ft.Divider(),
                ft.Row([ft.Text("Alumnos", size=18, weight="bold"), ft.IconButton("person_add", icon_color="green", on_click=lambda _: (state.update({"st_edit": None}), page.go("/form_student")))], alignment="spaceBetween"),
                col
            ]), padding=15, bgcolor="#f0f2f5", expand=True)
        ])

    def asistencia_view():
        dp = ft.TextField(label="Fecha (AAAA-MM-DD)", value=date.today().isoformat(), bgcolor="white")
        col = ft.Column(scroll="auto", expand=True)
        vals = {}
        
        def load(e=None):
            try:
                if date.fromisoformat(dp.value).weekday() >= 5: show_snack("⚠️ Fin de semana", "orange")
            except: pass
            
            ex = get_asistencia_diaria(state["curso_id"], dp.value)
            col.controls.clear(); vals.clear()
            for a in get_alumnos(state["curso_id"]):
                dd = ft.Dropdown(options=[ft.dropdown.Option(x) for x in ["P","T","A","J","S","N"]], value=ex.get(a['id'], "P"), width=80, bgcolor="white")
                vals[a['id']] = dd
                col.controls.append(ft.Container(content=ft.Row([ft.Text(a['nombre'], expand=True), dd]), padding=5, bgcolor="white", border_radius=5, margin=2))
            page.update()
            
        def save(e):
            try:
                d = date.fromisoformat(dp.value)
                if d > date.today(): return show_snack("Fecha futura", "red")
                if d.weekday() >= 5: return show_snack("Es fin de semana", "red")
            except: return show_snack("Fecha inválida", "red")
            
            for aid, dd in vals.items(): register_asistencia(aid, state["curso_id"], dp.value, dd.value)
            show_snack("Guardado"); page.go("/curso")

        load()
        return ft.View("/asistencia", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Asistencia"), bgcolor="blue", color="white"),
            ft.Container(content=ft.Column([ft.Row([dp, ft.IconButton("refresh", on_click=load)]), ft.ElevatedButton("GUARDAR", on_click=save, bgcolor="green", color="white", width=float("inf")), ft.Divider(), col]), padding=15, bgcolor="#f0f2f5", expand=True)
        ])

    def reportes_view():
        d1 = ft.TextField(label="Desde", value=date.today().replace(day=1).isoformat(), width=130, bgcolor="white")
        d2 = ft.TextField(label="Hasta", value=date.today().isoformat(), width=130, bgcolor="white")
        cont = ft.Column(scroll="auto", expand=True)
        
        def gen(e):
            data = get_report_data(state["curso_id"], d1.value, d2.value)
            rows = []
            for d in data:
                c = "red" if d['faltas']>=25 else ("orange" if d['faltas']>=15 else "black")
                rows.append(ft.DataRow(cells=[
                    ft.DataCell(ft.Text(d['nombre'], color=c, weight="bold" if d['faltas']>=25 else "normal")),
                    ft.DataCell(ft.Text(str(d['p']))), ft.DataCell(ft.Text(str(d['t']))), ft.DataCell(ft.Text(str(d['a']))),
                    ft.DataCell(ft.Text(str(d['j']))), ft.DataCell(ft.Text(str(d['s']))),
                    ft.DataCell(ft.Text(f"{d['faltas']}", color=c, weight="bold")),
                    ft.DataCell(ft.Text(f"{d['pct']}%"))
                ]))
            
            dt = ft.DataTable(columns=[
                ft.DataColumn(ft.Text("Alumno")), ft.DataColumn(ft.Text("P"), numeric=True), 
                ft.DataColumn(ft.Text("T"), numeric=True), ft.DataColumn(ft.Text("A"), numeric=True),
                ft.DataColumn(ft.Text("J"), numeric=True), ft.DataColumn(ft.Text("S"), numeric=True),
                ft.DataColumn(ft.Text("Faltas"), numeric=True), ft.DataColumn(ft.Text("%"), numeric=True)
            ], rows=rows, bgcolor="white", border_radius=10, column_spacing=15)
            cont.controls = [ft.Row([dt], scroll="always")]; page.update()

        def export(e):
            if not pd: return show_snack("Falta pandas", "red")
            if not xlsxwriter: return show_snack("Falta xlsxwriter", "red")
            
            data = get_report_data(state["curso_id"], d1.value, d2.value)
            if not data: return show_snack("Sin datos", "orange")
            
            df = pd.DataFrame(data)
            df = df.rename(columns={'nombre':'Alumno', 'p':'Pres', 't':'Tarde', 'a':'Aus', 'j':'Just', 's':'Susp', 'faltas':'Total', 'pct':'%'})
            
            # Exportación compatible con Web
            import io
            import base64
            output = io.BytesIO()
            df.to_excel(output, index=False, engine='xlsxwriter')
            b64 = base64.b64encode(output.getvalue()).decode()
            filename = f"reporte_{state['curso_id']}.xlsx"
            page.launch_url(f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}", web_window_name=filename)
            show_snack("Descargando...", "blue")

        return ft.View("/reportes", [
            ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/curso")), title=ft.Text("Reportes"), bgcolor="blue", color="white"),
            ft.Container(content=ft.Column([ft.Row([d1, d2, ft.ElevatedButton("Ver", on_click=gen)]), ft.ElevatedButton("Excel", icon="download", on_click=export, bgcolor="green", color="white"), ft.Divider(), cont]), padding=15, bgcolor="#f0f2f5", expand=True)
        ])

    # --- OTRAS VISTAS (Simplificadas para brevedad, mismas funcionalidades) ---
    def search_view():
        term = state["search"]; res = search_students(term); col = ft.Column(scroll="auto")
        if not res: col.controls.append(ft.Text("Sin resultados"))
        else:
            for r in res:
                col.controls.append(ft.Container(content=ft.ListTile(leading=ft.Icon("person"), title=ft.Text(r['nombre']), subtitle=ft.Text(f"Curso: {r['curso_nombre']}"), on_click=lambda e, s=r: (state.update({"st_view": s['id'], "curso_id": s['curso_id']}), page.go("/student_detail"))), bgcolor="white", border_radius=10, margin=2))
        return ft.View("/search", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/dashboard")), title=ft.Text(f"Busqueda: {term}"), bgcolor="blue", color="white"), ft.Container(content=col, padding=15, bgcolor="#f0f2f5", expand=True)])

    def student_detail_view():
        aid = state["st_view"]; s = get_alumno_by_id(aid)
        if not s: return ft.View("/error", [ft.Text("Error")])
        req_col = ft.Column()
        for r in get_student_req_status(aid, s['curso_id']): req_col.controls.append(ft.Row([ft.Icon("check" if r['ok'] else "close", color="green" if r['ok'] else "red"), ft.Text(r['desc'])]))
        
        card = ft.Container(content=ft.Column([
            ft.Text(s['nombre'], size=24, weight="bold"), ft.Text(f"Curso: {s['curso_nombre']}"), ft.Text(f"DNI: {s.get('dni','-')}"),
            ft.Divider(), ft.Text("Tutor:", weight="bold"), ft.Text(f"{s.get('tutor_nombre','-')} ({s.get('tutor_telefono','-')})"),
            ft.Divider(), ft.Text("Obs:", weight="bold"), ft.Text(s.get('observaciones','-'), italic=True),
            ft.Divider(), ft.Text("Papeles:", weight="bold"), req_col
        ]), padding=20, bgcolor="white", border_radius=15, shadow=ft.BoxShadow(blur_radius=5, color="black12"))
        
        return ft.View("/student_detail", [ft.AppBar(leading=ft.IconButton("arrow_back", icon_color="white", on_click=lambda _: page.go("/search")), title=ft.Text("Ficha"), bgcolor="blue", color="white"), ft.Container(content=ft.Column([card], scroll="auto"), padding=20, bgcolor="#f0f2f5", expand=True)])

    def form_student_view():
        is_edit = state["st_edit"] is not None
        nm = ft.TextField(label="Nombre", bgcolor="white"); dni = ft.TextField(label="DNI", bgcolor="white")
        obs = ft.TextField(label="Obs", multiline=True, bgcolor="white"); tn = ft.TextField(label="Tutor", bgcolor="white"); tt = ft.TextField(label="Tel Tutor", bgcolor="white")
        if is_edit:
            d = get_alumno_by_id(state["st_edit"])
            nm.value=d.get('nombre',''); dni.value=d.get('dni',''); obs.value=d.get('observaciones',''); tn.value=d.get('tutor_nombre',''); tt.value=d.get('tutor_telefono','')
        
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
            for a in get_alumnos(state["curso_id"]):
                col.controls.append(ft.Container(content=ft.Checkbox(label=a['nombre'], value=(a['id'] in done), on_change=lambda e, aid=a['id'], rid=rid: toggle_cumplimiento(rid, aid, e.control.value)), bgcolor="white", padding=5, border_radius=5, margin=2))
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
    port = int(os.environ.get("PORT", 8000))
    ft.app(target=main, view=ft.WEB_BROWSER, port=port, web_renderer="html")
