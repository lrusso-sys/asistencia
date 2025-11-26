import sqlite3
from datetime import date

# ----------------------------------------------------------------------
# Configuración de la Base de Datos y Utilidades
# ----------------------------------------------------------------------

DB_NAME = 'asistencia_alumnos.db'

def get_db_connection():
    """Establece la conexión a la base de datos SQLite."""
    conn = sqlite3.connect(DB_NAME)
    # Permite acceder a las columnas por nombre (útil para retornar diccionarios/objetos)
    conn.row_factory = sqlite3.Row 
    return conn

def init_db():
    """
    Inicializa y crea las tablas de la base de datos si no existen.
    Se añade 'ON DELETE CASCADE' para mantener la integridad:
    Si se borra un curso, se borran sus alumnos.
    Si se borra un alumno, se borran sus registros de asistencia.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Tabla de Cursos (Grupos)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Cursos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE
        )
    """)

    # 2. Tabla de Alumnos (con borrado en cascada)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Alumnos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            curso_id INTEGER NOT NULL,
            nombre TEXT NOT NULL,
            UNIQUE(curso_id, nombre), -- No puede haber dos alumnos con el mismo nombre en el mismo curso
            FOREIGN KEY (curso_id) REFERENCES Cursos(id) ON DELETE CASCADE
        )
    """)

    # 3. Tabla de Asistencia (con borrado en cascada)
    # status: P=Presente, A=Ausente, J=Ausente Justificado, S=Suspensión, T=Tarde, N=No Corresponde
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Asistencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            alumno_id INTEGER NOT NULL,
            fecha TEXT NOT NULL, -- Formato YYYY-MM-DD
            status TEXT NOT NULL,
            UNIQUE(alumno_id, fecha),
            FOREIGN KEY (alumno_id) REFERENCES Alumnos(id) ON DELETE CASCADE
        )
    """)

    # 4. Tabla de Feriados
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Feriados (
            fecha TEXT PRIMARY KEY, -- Formato YYYY-MM-DD
            descripcion TEXT
        )
    """)
    
    conn.commit()
    conn.close()
    print("Base de datos inicializada correctamente.")

# ----------------------------------------------------------------------
# CRUD: Cursos
# ----------------------------------------------------------------------

def add_curso(nombre):
    """C (Crear): Agrega un nuevo curso."""
    conn = get_db_connection()
    try:
        conn.execute("INSERT INTO Cursos (nombre) VALUES (?)", (nombre,))
        conn.commit()
        return True, f"Curso '{nombre}' agregado exitosamente."
    except sqlite3.IntegrityError:
        return False, f"Error: El curso '{nombre}' ya existe."
    finally:
        conn.close()

def get_cursos():
    """R (Leer): Retorna todos los cursos."""
    conn = get_db_connection()
    cursos = conn.execute("SELECT * FROM Cursos ORDER BY nombre").fetchall()
    conn.close()
    return [dict(curso) for curso in cursos]

def update_curso(curso_id, new_nombre):
    """U (Actualizar): Cambia el nombre de un curso."""
    conn = get_db_connection()
    try:
        conn.execute("UPDATE Cursos SET nombre = ? WHERE id = ?", (new_nombre, curso_id))
        conn.commit()
        if conn.rowcount > 0:
            return True, "Curso actualizado."
        return False, "Error: ID de curso no encontrado."
    except sqlite3.IntegrityError:
        return False, f"Error: Ya existe un curso llamado '{new_nombre}'."
    finally:
        conn.close()

def delete_curso(curso_id):
    """D (Borrar): Elimina un curso y sus alumnos/asistencias asociadas (CASCADE)."""
    conn = get_db_connection()
    try:
        # La integridad CASCADE se encarga de borrar alumnos y asistencias
        conn.execute("DELETE FROM Cursos WHERE id = ?", (curso_id,))
        conn.commit()
        if conn.rowcount > 0:
            return True, "Curso y datos asociados eliminados correctamente."
        return False, "Error: ID de curso no encontrado."
    except Exception as e:
        return False, f"Error al eliminar curso: {e}"
    finally:
        conn.close()

# ----------------------------------------------------------------------
# CRUD: Alumnos
# ----------------------------------------------------------------------

def add_alumno(curso_id, alumno_nombre):
    """C (Crear): Agrega un alumno a un curso."""
    conn = get_db_connection()
    try:
        conn.execute("INSERT INTO Alumnos (curso_id, nombre) VALUES (?, ?)", (curso_id, alumno_nombre))
        conn.commit()
        return True, f"Alumno '{alumno_nombre}' agregado."
    except sqlite3.IntegrityError:
        return False, f"Error: Ya existe un alumno llamado '{alumno_nombre}' en este curso."
    finally:
        conn.close()

def get_alumnos_by_curso(curso_id):
    """R (Leer): Retorna todos los alumnos de un curso específico."""
    conn = get_db_connection()
    alumnos = conn.execute("SELECT * FROM Alumnos WHERE curso_id = ? ORDER BY nombre", (curso_id,)).fetchall()
    conn.close()
    return [dict(alumno) for alumno in alumnos]

def get_alumno_by_id(alumno_id):
    """R (Leer): Retorna un único alumno por su ID."""
    conn = get_db_connection()
    alumno = conn.execute("SELECT * FROM Alumnos WHERE id = ?", (alumno_id,)).fetchone()
    conn.close()
    return dict(alumno) if alumno else None

def update_alumno(alumno_id, new_nombre=None, new_curso_id=None):
    """U (Actualizar): Actualiza el nombre o el curso de un alumno."""
    conn = get_db_connection()
    try:
        if new_nombre and new_curso_id:
            conn.execute("UPDATE Alumnos SET nombre = ?, curso_id = ? WHERE id = ?", (new_nombre, new_curso_id, alumno_id))
        elif new_nombre:
            conn.execute("UPDATE Alumnos SET nombre = ? WHERE id = ?", (new_nombre, alumno_id))
        elif new_curso_id:
            conn.execute("UPDATE Alumnos SET curso_id = ? WHERE id = ?", (new_curso_id, alumno_id))
        else:
            return False, "No se proporcionaron datos para actualizar."

        conn.commit()
        if conn.rowcount > 0:
            return True, "Alumno actualizado."
        return False, "Error: ID de alumno no encontrado o no hubo cambios."
    except sqlite3.IntegrityError:
        return False, "Error de integridad: Este nombre ya existe en el curso de destino."
    finally:
        conn.close()

def delete_alumno(alumno_id):
    """D (Borrar): Elimina un alumno y sus asistencias asociadas (CASCADE)."""
    conn = get_db_connection()
    try:
        # La integridad CASCADE se encarga de borrar las asistencias
        conn.execute("DELETE FROM Alumnos WHERE id = ?", (alumno_id,))
        conn.commit()
        if conn.rowcount > 0:
            return True, "Alumno y asistencias eliminados correctamente."
        return False, "Error: ID de alumno no encontrado."
    except Exception as e:
        return False, f"Error al eliminar alumno: {e}"
    finally:
        conn.close()

# ----------------------------------------------------------------------
# Gestión de Asistencia y Feriados
# ----------------------------------------------------------------------

def add_feriado(fecha_str, descripcion):
    """Agrega o reemplaza un día feriado."""
    conn = get_db_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO Feriados (fecha, descripcion) VALUES (?, ?)", (fecha_str, descripcion))
        conn.commit()
        return True, f"Feriado '{descripcion}' ({fecha_str}) agregado/actualizado."
    except Exception as e:
        return False, f"Error al agregar feriado: {e}"
    finally:
        conn.close()

def get_feriados():
    """Retorna la lista de feriados."""
    conn = get_db_connection()
    feriados = conn.execute("SELECT * FROM Feriados ORDER BY fecha").fetchall()
    conn.close()
    return [dict(f) for f in feriados]

def register_asistencia(alumno_id, fecha_str, status):
    """Registra o actualiza la asistencia de un alumno para una fecha específica."""
    valid_statuses = ['P', 'A', 'J', 'S', 'T', 'N']
    if status not in valid_statuses:
        return False, f"Error: Estado de asistencia '{status}' no válido. Use {valid_statuses}"

    conn = get_db_connection()
    try:
        # INSERT OR REPLACE actualiza si ya existe un registro para esa fecha/alumno
        conn.execute("INSERT OR REPLACE INTO Asistencia (alumno_id, fecha, status) VALUES (?, ?, ?)", 
                     (alumno_id, fecha_str, status))
        conn.commit()
        return True, f"Asistencia de alumno {alumno_id} el {fecha_str} registrada/actualizada como '{status}'."
    except Exception as e:
        return False, f"Error al registrar asistencia: {e}"
    finally:
        conn.close()

def get_asistencia_by_alumno(alumno_id):
    """Retorna todos los registros de asistencia de un alumno."""
    conn = get_db_connection()
    asistencias = conn.execute("SELECT fecha, status FROM Asistencia WHERE alumno_id = ? ORDER BY fecha", (alumno_id,)).fetchall()
    conn.close()
    return [dict(a) for a in asistencias]


# ----------------------------------------------------------------------
# Función de Cálculo Central y Reportes
# ----------------------------------------------------------------------

def calculate_absences(alumno_id):
    """
    Calcula el total de faltas para un alumno,
    donde T (Tarde) = 0.25 falta.
    A (Ausente) y S (Suspensión) = 1 falta.
    """
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Obtener la información del alumno
    cursor.execute("""
        SELECT a.nombre as alumno_nombre, c.nombre as curso_nombre
        FROM Alumnos a
        JOIN Cursos c ON a.curso_id = c.id
        WHERE a.id = ?
    """, (alumno_id,))
    alumno_info = cursor.fetchone()

    if not alumno_info:
        conn.close()
        return {"error": "Alumno no encontrado."}

    # 2. Obtener todos los registros de asistencia del alumno
    asistencias = get_asistencia_by_alumno(alumno_id)

    # 3. Obtener la lista de feriados
    feriados_data = get_feriados()
    feriados = {f['fecha'] for f in feriados_data}

    conn.close()
    
    total_faltas = 0.0
    detalle = {
        'T': 0, # Tardanzas (0.25 c/u)
        'A': 0, # Ausentes (1 c/u)
        'S': 0  # Suspensiones (1 c/u)
    }
    
    # Calcular faltas
    for asistencia in asistencias:
        fecha = asistencia['fecha']
        status = asistencia['status']
        
        # Si la fecha es un feriado, no cuenta como falta
        if fecha in feriados:
            continue
            
        if status == 'T':
            total_faltas += 0.25
            detalle['T'] += 1
        elif status == 'A':
            total_faltas += 1.0
            detalle['A'] += 1
        elif status == 'S':
            total_faltas += 1.0
            detalle['S'] += 1
        
    reporte = {
        'alumno_nombre': alumno_info['alumno_nombre'],
        'curso_nombre': alumno_info['curso_nombre'],
        'total_faltas': total_faltas,
        'tardanzas_count': detalle['T'],
        'ausentes_count': detalle['A'],
        'suspensiones_count': detalle['S']
    }
    
    return reporte

# ----------------------------------------------------------------------
# Ejemplo de Uso / Pruebas de CRUD
# ----------------------------------------------------------------------

if __name__ == '__main__':
    # --- 1. Inicializar la base de datos ---
    init_db()

    # --- 2. Pruebas de CRUD de Cursos ---
    print("\n--- CRUD Cursos ---")
    add_curso("4to A - Matutino")
    add_curso("6to B - Vespertino")
    add_curso("5to C - Tarde")
    print("Cursos Iniciales:", get_cursos())
    
    curso_to_update = get_cursos()[0]['id'] # Tomamos el ID del primer curso
    update_curso(curso_to_update, "4to A - Mañana")
    print("Cursos Actualizados:", get_cursos())

    # --- 3. Pruebas de CRUD de Alumnos ---
    print("\n--- CRUD Alumnos ---")
    curso_id_4to = get_cursos()[0]['id'] # ID de "4to A - Mañana"
    curso_id_6to = get_cursos()[1]['id'] # ID de "6to B - Vespertino"
    
    # Crear Alumnos
    add_alumno(curso_id_4to, "Ana Torres") # Alumno ID 1
    add_alumno(curso_id_4to, "Juan Pérez") # Alumno ID 2
    add_alumno(curso_id_6to, "Sofía Gómez") # Alumno ID 3

    print("Alumnos en 4to A:", get_alumnos_by_curso(curso_id_4to))
    
    # Actualizar Alumno (cambiar nombre)
    alumno_to_update = get_alumnos_by_curso(curso_id_4to)[0]['id']
    update_alumno(alumno_to_update, new_nombre="Ana García")
    print("Alumno Actualizado:", get_alumno_by_id(alumno_to_update))
    
    # Actualizar Alumno (cambiar de curso)
    alumno_to_move = get_alumnos_by_curso(curso_id_4to)[1]['id'] # Juan Pérez
    update_alumno(alumno_to_move, new_curso_id=curso_id_6to)
    print("Alumnos en 4to A después de mover:", get_alumnos_by_curso(curso_id_4to))
    print("Alumnos en 6to B después de mover:", get_alumnos_by_curso(curso_id_6to))

    # --- 4. Pruebas de Asistencia y Reporte (usando IDs después del movimiento) ---
    print("\n--- Asistencia y Reporte ---")
    ana_id = get_alumno_by_id(alumno_to_update)['id']
    juan_id = get_alumno_by_id(alumno_to_move)['id']
    
    add_feriado(date(2025, 11, 26).isoformat(), "Feriado Nacional")

    # Asistencia de Ana
    register_asistencia(ana_id, date(2025, 11, 25).isoformat(), 'P') # Presente
    register_asistencia(ana_id, date(2025, 11, 28).isoformat(), 'A') # Ausente (1 falta)
    register_asistencia(ana_id, date(2025, 11, 29).isoformat(), 'T') # Tarde (0.25 falta)
    register_asistencia(ana_id, date(2025, 11, 26).isoformat(), 'A') # Día feriado (0 faltas)

    # Asistencia de Juan (ahora en 6to B)
    register_asistencia(juan_id, date(2025, 11, 28).isoformat(), 'T') # Tarde (0.25 falta)
    register_asistencia(juan_id, date(2025, 11, 29).isoformat(), 'T') # Tarde (0.25 falta)
    
    reporte_ana = calculate_absences(ana_id)
    print(f"\nReporte de {reporte_ana['alumno_nombre']}: Total Faltas = {reporte_ana['total_faltas']}") # Esperado: 1.25
    
    reporte_juan = calculate_absences(juan_id)
    print(f"Reporte de {reporte_juan['alumno_nombre']}: Total Faltas = {reporte_juan['total_faltas']}") # Esperado: 0.5

    # --- 5. Pruebas de Borrado ---
    print("\n--- Borrado de Datos ---")
    # Borrar un alumno: debería borrar sus asistencias
    delete_alumno(ana_id) 
    print(f"¿Ana existe aún? {get_alumno_by_id(ana_id)}") # Debe ser None
    print(f"Asistencias de Ana: {get_asistencia_by_alumno(ana_id)}") # Debe ser lista vacía

    # Borrar un curso: debería borrar el curso y sus alumnos (Juan y Sofía)
    delete_curso(curso_id_6to)
    print(f"Alumnos en 6to B después de borrar curso: {get_alumnos_by_curso(curso_id_6to)}") # Debe ser lista vacía