from flask import Flask, render_template, request, session, redirect, url_for
import sqlite3
import pandas as pd
import openpyxl
import unicodedata
import urllib.request
import json
import time
import os
import math
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'llave_secreta_cintac_2026_prod')

# Soporte automático para ejecución local y despliegue Serverless en Vercel (/tmp)
if os.environ.get('VERCEL') or not os.access(os.path.dirname(os.path.abspath(__file__)), os.W_OK):
    DB_PATH = '/tmp/cintac_comex.db'
else:
    DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cintac_comex.db')

# Límites operativos para Pruebas de Borde (Boundary Testing)
PESO_MIN_TON = 0.01
PESO_MAX_TON = 10000.00
VALOR_CARGA_MAX_USD = 100000000.00

# Caché en memoria para la API de Dólar Observado (evita latencia en pruebas de estrés)
_CACHE_DOLAR = {
    'valor': 965.0,
    'fuente': 'Valor referencial Banco Central',
    'timestamp': 0
}

def obtener_tasa_dolar_clp():
    ahora = time.time()
    if ahora - _CACHE_DOLAR['timestamp'] < 3600 and _CACHE_DOLAR['timestamp'] > 0:
        return _CACHE_DOLAR['valor'], _CACHE_DOLAR['fuente']

    try:
        req = urllib.request.Request(
            'https://mindicador.cl/api/dolar',
            headers={'User-Agent': 'Mozilla/5.0 CINTAC-Cotizador/1.0'}
        )
        with urllib.request.urlopen(req, timeout=2.0) as respuesta:
            datos = json.loads(respuesta.read().decode('utf-8'))
            valor_actual = float(datos['serie'][0]['valor'])
            if valor_actual > 0:
                _CACHE_DOLAR['valor'] = valor_actual
                _CACHE_DOLAR['fuente'] = 'Dólar Observado en línea (mindicador.cl)'
                _CACHE_DOLAR['timestamp'] = ahora
                return valor_actual, _CACHE_DOLAR['fuente']
    except Exception:
        pass

    _CACHE_DOLAR['timestamp'] = ahora
    return _CACHE_DOLAR['valor'], 'Tasa referencial de respaldo (Modo sin conexión)'

@app.after_request
def add_security_headers(response):
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    response.headers["Vary"] = "Cookie"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    return response

# ==========================================
# 1. BASE DE DATOS NORMALIZADA (3FN - 4 TABLAS)
# ==========================================
def poblar_rutas_oficiales_cintac(cursor):
    rutas_seed = [
        ('Importación', 'Shanghai', 'China', 0, 'San Antonio', 'Chile', 1, 25, 32, 1600.0, 2800.0, 2300.0, 3800.0),
        ('Importación', 'Shanghai', 'China', 0, 'Valparaíso', 'Chile', 1, 26, 34, 1650.0, 2850.0, 2350.0, 3850.0),
        ('Importación', 'Ningbo', 'China', 0, 'San Antonio', 'Chile', 1, 26, 32, 1600.0, 2800.0, 2300.0, 3800.0),
        ('Importación', 'Ningbo', 'China', 0, 'Valparaíso', 'Chile', 1, 27, 34, 1650.0, 2850.0, 2350.0, 3850.0),
        ('Importación', 'Qingdao', 'China', 0, 'San Antonio', 'Chile', 1, 28, 34, 1700.0, 2900.0, 2400.0, 3900.0),
        ('Importación', 'Qingdao', 'China', 0, 'Valparaíso', 'Chile', 1, 29, 36, 1750.0, 2950.0, 2450.0, 3950.0),
        ('Importación', 'Shenzhen / Yantian', 'China', 0, 'San Antonio', 'Chile', 1, 27, 33, 1650.0, 2850.0, 2350.0, 3850.0),
        ('Importación', 'Shenzhen / Yantian', 'China', 0, 'Valparaíso', 'Chile', 1, 28, 35, 1700.0, 2900.0, 2400.0, 3900.0),
        ('Importación', 'Tianjin', 'China', 0, 'San Antonio', 'Chile', 1, 30, 36, 1800.0, 3000.0, 2500.0, 4000.0),
        ('Importación', 'Tianjin', 'China', 0, 'Valparaíso', 'Chile', 1, 31, 38, 1850.0, 3050.0, 2550.0, 4050.0),
        ('Importación', 'Busan', 'Corea del Sur', 0, 'San Antonio', 'Chile', 1, 30, 38, 1700.0, 2900.0, 2400.0, 3900.0),
        ('Importación', 'Busan', 'Corea del Sur', 0, 'Valparaíso', 'Chile', 1, 31, 40, 1750.0, 2950.0, 2450.0, 3950.0),
        ('Importación', 'Tokio / Yokohama', 'Japón', 0, 'San Antonio', 'Chile', 1, 32, 40, 1900.0, 3100.0, 2600.0, 4100.0),
        ('Importación', 'Tokio / Yokohama', 'Japón', 0, 'Valparaíso', 'Chile', 1, 33, 42, 1950.0, 3150.0, 2650.0, 4150.0),
        ('Importación', 'Ho Chi Minh', 'Vietnam', 0, 'San Antonio', 'Chile', 1, 35, 42, 1900.0, 3200.0, 2700.0, 4300.0),
        ('Importación', 'Ho Chi Minh', 'Vietnam', 0, 'Valparaíso', 'Chile', 1, 36, 44, 1950.0, 3250.0, 2750.0, 4350.0),
        ('Exportación', 'San Antonio', 'Chile', 1, 'Callao', 'Perú', 0, 3, 5, 900.0, 900.0, 970.0, 970.0),
        ('Exportación', 'Valparaíso', 'Chile', 1, 'Callao', 'Perú', 0, 3, 5, 900.0, 900.0, 970.0, 970.0)
    ]

    for op, p_or, pais_or, nac_or, p_des, pais_des, nac_des, tmin, tmax, t20min, t20max, t40min, t40max in rutas_seed:
        cursor.execute('INSERT OR IGNORE INTO puertos (nombre_puerto, pais, es_nacional) VALUES (?, ?, ?)', (p_or, pais_or, nac_or))
        cursor.execute('INSERT OR IGNORE INTO puertos (nombre_puerto, pais, es_nacional) VALUES (?, ?, ?)', (p_des, pais_des, nac_des))

    cursor.execute('SELECT id_puerto, nombre_puerto FROM puertos')
    mapa_p = {r['nombre_puerto']: r['id_puerto'] for r in cursor.fetchall()}

    for op, p_or, _, _, p_des, _, _, tmin, tmax, t20min, t20max, t40min, t40max in rutas_seed:
        cursor.execute('''
            INSERT OR REPLACE INTO tarifas_contenedores
            (tipo_operacion, id_origen, id_destino, tipo_transporte, transito_min, transito_max, tarifa_20_min, tarifa_20, tarifa_40_min, tarifa_40)
            VALUES (?, ?, ?, 'Maritimo', ?, ?, ?, ?, ?, ?)
        ''', (op, mapa_p[p_or], mapa_p[p_des], tmin, tmax, t20min, t20max, t40min, t40max))

def _inicializar_esquema_si_falta(conn):
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(tarifas_contenedores)")
    columnas_actuales = [col['name'] for col in cursor.fetchall()]
    if 'tarifa_20_min' not in columnas_actuales:
        cursor.execute("DROP TABLE IF EXISTS tarifas")
        cursor.execute("DROP TABLE IF EXISTS tarifas_contenedores")
        cursor.execute("DROP TABLE IF EXISTS usuarios")
        cursor.execute("DROP TABLE IF EXISTS puertos")
        cursor.execute("DROP TABLE IF EXISTS roles")

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS roles (
                id_rol INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre_rol TEXT UNIQUE NOT NULL CHECK (nombre_rol IN ('admin', 'empleado')),
                descripcion TEXT NOT NULL
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS usuarios (
                id_usuario INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                correo TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                id_rol INTEGER NOT NULL,
                fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (id_rol) REFERENCES roles(id_rol) ON DELETE RESTRICT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS puertos (
                id_puerto INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre_puerto TEXT UNIQUE NOT NULL,
                pais TEXT NOT NULL,
                es_nacional INTEGER NOT NULL DEFAULT 0 CHECK (es_nacional IN (0, 1))
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS tarifas_contenedores (
                id_tarifa INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo_operacion TEXT NOT NULL CHECK (tipo_operacion IN ('Importación', 'Exportación')),
                id_origen INTEGER NOT NULL,
                id_destino INTEGER NOT NULL,
                tipo_transporte TEXT NOT NULL DEFAULT 'Maritimo',
                transito_min INTEGER NOT NULL CHECK (transito_min > 0),
                transito_max INTEGER NOT NULL CHECK (transito_max >= transito_min),
                tarifa_20_min REAL NOT NULL CHECK (tarifa_20_min > 0),
                tarifa_20 REAL NOT NULL CHECK (tarifa_20 >= tarifa_20_min),
                tarifa_40_min REAL NOT NULL CHECK (tarifa_40_min > 0),
                tarifa_40 REAL NOT NULL CHECK (tarifa_40 >= tarifa_40_min),
                fecha_actualizacion DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (id_origen) REFERENCES puertos(id_puerto) ON DELETE RESTRICT,
                FOREIGN KEY (id_destino) REFERENCES puertos(id_puerto) ON DELETE RESTRICT,
                UNIQUE(tipo_operacion, id_origen, id_destino)
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_usuarios_correo ON usuarios(correo);')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_puertos_tipo ON puertos(es_nacional, pais);')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_tarifas_ruta ON tarifas_contenedores(tipo_operacion, id_origen, id_destino);')

        cursor.executemany(
            'INSERT INTO roles (nombre_rol, descripcion) VALUES (?, ?)',
            [('admin', 'Jefatura / Administrador COMEX'), ('empleado', 'Analista de Operaciones COMEX')]
        )
        cursor.execute("SELECT id_rol, nombre_rol FROM roles")
        mapa_roles = {r['nombre_rol']: r['id_rol'] for r in cursor.fetchall()}
        cursor.executemany('''
            INSERT INTO usuarios (nombre, correo, password_hash, id_rol)
            VALUES (?, ?, ?, ?)
        ''', [
            ('Administrador COMEX', 'admin@cintac.cl', generate_password_hash('admin123'), mapa_roles['admin']),
            ('Analista COMEX 1', 'empleado1@cintac.cl', generate_password_hash('emp123'), mapa_roles['empleado']),
            ('Analista COMEX 2', 'empleado2@cintac.cl', generate_password_hash('emp123'), mapa_roles['empleado'])
        ])
        poblar_rutas_oficiales_cintac(cursor)
        conn.commit()

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=15.0)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON;')
    _inicializar_esquema_si_falta(conn)
    return conn

def init_db():
    conn = get_db_connection()
    conn.close()

def obtener_catalogo_rutas():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT tc.id_tarifa, tc.tipo_operacion, po.nombre_puerto AS puerto_origen, po.pais AS pais_origen,
               pd.nombre_puerto AS puerto_destino, pd.pais AS pais_destino, tc.tipo_transporte,
               tc.transito_min, tc.transito_max, tc.tarifa_20_min, tc.tarifa_20, tc.tarifa_40_min, tc.tarifa_40
        FROM tarifas_contenedores tc
        JOIN puertos po ON tc.id_origen = po.id_puerto
        JOIN puertos pd ON tc.id_destino = pd.id_puerto
        ORDER BY tc.tipo_operacion, po.pais, po.nombre_puerto, pd.nombre_puerto
    ''')
    filas = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return filas

def obtener_usuarios_bd():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT u.id_usuario, u.nombre, u.correo, r.nombre_rol AS rol, r.descripcion AS rol_desc
        FROM usuarios u JOIN roles r ON u.id_rol = r.id_rol ORDER BY u.id_usuario ASC
    ''')
    us = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return us

# ==========================================
# 2. MOTOR DE VALIDACIÓN Y CONSOLIDACIÓN EXCEL
# ==========================================
def normalizar_texto(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto).strip().lower()) if unicodedata.category(c) != 'Mn')

def estandarizar_nombre_puerto(nombre):
    s = str(nombre).strip()
    norm = normalizar_texto(s)
    mapa_canonico = {
        'valparaiso': 'Valparaíso',
        'san antonio': 'San Antonio',
        'callao': 'Callao',
        'shanghai': 'Shanghai',
        'ningbo': 'Ningbo',
        'qingdao': 'Qingdao',
        'shenzhen / yantian': 'Shenzhen / Yantian',
        'shenzhen': 'Shenzhen',
        'tianjin': 'Tianjin',
        'busan': 'Busan',
        'tokio / yokohama': 'Tokio / Yokohama',
        'ho chi minh': 'Ho Chi Minh'
    }
    return mapa_canonico.get(norm, s)

def estandarizar_pais(pais):
    s = str(pais).strip()
    norm = normalizar_texto(s)
    mapa_pais = {
        'chile': 'Chile',
        'peru': 'Perú',
        'japon': 'Japón',
        'china': 'China',
        'corea del sur': 'Corea del Sur',
        'vietnam': 'Vietnam',
        'espana': 'España',
        'ee.uu.': 'EE.UU.',
        'singapur': 'Singapur'
    }
    return mapa_pais.get(norm, s)

def limpiar_numero(val):
    if isinstance(val, (int, float)):
        return float(val)
    val = (
        str(val).upper()
        .replace('DÍAS', '').replace('DIAS', '').replace('DÍA', '').replace('DIA', '')
        .replace('USD', '').replace('US$', '').replace('$', '').strip()
    )
    if ',' in val and '.' in val:
        val = val.replace('.', '').replace(',', '.') if val.rfind(',') > val.rfind('.') else val.replace(',', '')
    elif ',' in val:
        val = val.replace(',', '.') if len(val.split(',')[-1]) == 2 else val.replace(',', '')
    return float(val)

def extraer_puerto_y_pais(texto, pais_defecto='Perú'):
    s = str(texto).strip()
    if '(' in s and ')' in s:
        p = s.split('(')[0].strip()
        pais = s.split('(')[1].replace(')', '').strip()
        return estandarizar_nombre_puerto(p), estandarizar_pais(pais)
    return estandarizar_nombre_puerto(s), estandarizar_pais(pais_defecto)

def validar_excel_tarifas(archivo_stream):
    caracteres_peligrosos = ('=', '+', '@', '-', '\t', '\r')
    try:
        wb = openpyxl.load_workbook(archivo_stream, data_only=False)
        if 'Tarifas Referencia' in wb.sheetnames:
            ws = wb['Tarifas Referencia']
        elif 'Importaciones' in wb.sheetnames:
            ws = wb['Importaciones']
        else:
            ws = wb.active
        datos = list(ws.values)
    except Exception:
        return False, "No se pudo abrir el archivo. Verifique que sea un documento Excel (.xlsx) válido y no esté dañado.", 0

    if not datos:
        return False, "El archivo Excel seleccionado se encuentra vacío.", 0

    cabeceras_raw = datos[0]
    for idx_col, valor_cab in enumerate(cabeceras_raw):
        if valor_cab is not None:
            val_str = str(valor_cab).strip()
            if val_str.startswith(caracteres_peligrosos):
                letra_col = chr(65 + idx_col) if idx_col < 26 else f"Col {idx_col+1}"
                return False, f"Por seguridad se rechazó el archivo: se detectó una fórmula no permitida ('{val_str}') en la celda {letra_col}1.", 0

    # Mapear columnas dinámicamente por nombre (soporta plantillas de 10, 12 y 13 columnas)
    mapa_cols = {}
    for idx_col, valor_cab in enumerate(cabeceras_raw):
        if valor_cab is not None and str(valor_cab).strip() != "":
            mapa_cols[normalizar_texto(valor_cab)] = idx_col

    def buscar_col(posibles):
        for p in posibles:
            if p in mapa_cols:
                return mapa_cols[p]
        return None

    idx_p_orig = buscar_col(['puerto origen'])
    idx_pais_orig = buscar_col(['pais origen', 'pais'])
    idx_p_dest = buscar_col(['puerto destino'])
    idx_pais_dest = buscar_col(['pais destino'])
    idx_modalidad = buscar_col(['modalidad', 'tipo operacion'])
    idx_t20_min = buscar_col(["tarifa 20' min (us$)", "tarifa 20 min"])
    idx_t20_max = buscar_col(["tarifa 20' max (us$)", "tarifa 20 max"])
    idx_t40_min = buscar_col(["tarifa 40' min (us$)", "tarifa 40 min"])
    idx_t40_max = buscar_col(["tarifa 40' max (us$)", "tarifa 40 max"])
    idx_d_min = buscar_col(['transito min (dias)', 'transito min'])
    idx_d_max = buscar_col(['transito max (dias)', 'transito max'])

    indices_obligatorios = [idx_p_orig, idx_pais_orig, idx_p_dest, idx_t20_min, idx_t20_max, idx_t40_min, idx_t40_max, idx_d_min, idx_d_max]
    if any(i is None for i in indices_obligatorios):
        return False, "El archivo no contiene todas las columnas obligatorias de la plantilla oficial (Puerto Origen, País Origen, Puerto Destino, Tarifas 20'/40' y Días de Tránsito).", 0

    registros_unicos = {}
    filas_referencia_consolidadas = 0

    for idx_fila, fila in enumerate(datos[1:], start=2):
        v_porig = fila[idx_p_orig] if idx_p_orig < len(fila) else None
        v_paisorig = fila[idx_pais_orig] if idx_pais_orig < len(fila) else None
        v_pdest = fila[idx_p_dest] if idx_p_dest < len(fila) else None

        if (v_porig is None or str(v_porig).strip() == "") and (v_pdest is None or str(v_pdest).strip() == ""):
            continue

        for val_c in fila:
            if val_c is not None and str(val_c).strip() != "":
                val_s = str(val_c).strip()
                if val_s.startswith(caracteres_peligrosos):
                    return False, f"Por seguridad se canceló la carga: la fila {idx_fila} contiene una fórmula o carácter inválido ('{val_s}').", 0

        p_or_raw = str(v_porig).strip() if v_porig is not None else ""
        pais_or_raw = str(v_paisorig).strip() if v_paisorig is not None else ""
        p_des_raw = str(v_pdest).strip() if v_pdest is not None else ""

        if not p_or_raw or not pais_or_raw or not p_des_raw:
            return False, f"Error en la fila {idx_fila}: Los campos Puerto Origen, País Origen y Puerto Destino no pueden estar vacíos.", 0

        p_or, pais_or = extraer_puerto_y_pais(p_or_raw, pais_or_raw)

        if idx_modalidad is not None and idx_modalidad < len(fila) and fila[idx_modalidad] is not None:
            es_exp = ('export' in normalizar_texto(fila[idx_modalidad]))
        else:
            es_exp = (normalizar_texto(pais_or) == 'chile')

        if idx_pais_dest is not None and idx_pais_dest < len(fila) and fila[idx_pais_dest] is not None and str(fila[idx_pais_dest]).strip() != "":
            pais_defecto_dest = str(fila[idx_pais_dest]).strip()
        else:
            pais_defecto_dest = 'Perú' if es_exp else 'Chile'

        p_des, pais_des = extraer_puerto_y_pais(p_des_raw, pais_defecto_dest)

        try:
            t20_1 = limpiar_numero(fila[idx_t20_min])
            t20_2 = limpiar_numero(fila[idx_t20_max])
            t40_1 = limpiar_numero(fila[idx_t40_min])
            t40_2 = limpiar_numero(fila[idx_t40_max])
            d1 = int(round(limpiar_numero(fila[idx_d_min])))
            d2 = int(round(limpiar_numero(fila[idx_d_max])))

            t20_min, t20_max = min(t20_1, t20_2), max(t20_1, t20_2)
            t40_min, t40_max = min(t40_1, t40_2), max(t40_1, t40_2)
            tmin, tmax = min(d1, d2), max(d1, d2)
        except Exception:
            return False, f"Error en la fila {idx_fila} ({p_or} ➔ {p_des}): Verifique que las tarifas y días de tránsito sean valores numéricos válidos.", 0

        if tmin <= 0 or t20_min <= 0 or t40_min <= 0:
            return False, f"Error en la fila {idx_fila} ({p_or} ➔ {p_des}): Las tarifas y los días de tránsito deben ser mayores a 0.", 0

        tipo_op = 'Exportación' if es_exp else 'Importación'
        clave_ruta = (tipo_op, p_or.lower(), p_des.lower())

        if clave_ruta in registros_unicos:
            filas_referencia_consolidadas += 1
            continue

        registros_unicos[clave_ruta] = {
            'tipo_operacion': tipo_op,
            'puerto_origen': p_or,
            'pais_origen': pais_or,
            'es_nac_origen': 1 if es_exp else 0,
            'puerto_destino': p_des,
            'pais_destino': pais_des,
            'es_nac_destino': 0 if es_exp else 1,
            'transito_min': tmin,
            'transito_max': tmax,
            'tarifa_20_min': t20_min,
            'tarifa_20': t20_max,
            'tarifa_40_min': t40_min,
            'tarifa_40': t40_max
        }

    if not registros_unicos:
        return False, "La planilla no contiene filas de tarifas válidas.", 0

    return True, list(registros_unicos.values()), filas_referencia_consolidadas

init_db()

# ==========================================
# 3. RUTAS WEB Y LÓGICA DE NEGOCIO
# ==========================================
def renderizar_panel(tab_activa='cotizador-view', **kwargs):
    es_admin = ('rol' in session and session['rol'] == 'admin')
    tasa_clp, fuente_dolar = obtener_tasa_dolar_clp()
    return render_template(
        'index.html',
        rutas=obtener_catalogo_rutas(),
        usuarios=obtener_usuarios_bd() if es_admin else [],
        es_admin=es_admin,
        correo_usuario=session.get('correo'),
        nombre_usuario=session.get('nombre'),
        tasa_dolar=f"{tasa_clp:,.2f}",
        fuente_dolar=fuente_dolar,
        tab_activa=tab_activa,
        **kwargs
    )

@app.route('/', methods=['GET', 'POST'])
@app.route('/panel_interno', methods=['GET', 'POST'])
def vista_principal():
    if request.path == '/panel_interno' and 'rol' not in session:
        return redirect(url_for('login'))

    es_admin = ('rol' in session and session['rol'] == 'admin')
    if request.method == 'POST':
        pestaña_destino = 'cotizador-view'
    elif es_admin and request.path == '/panel_interno':
        pestaña_destino = 'admin-view'
    else:
        pestaña_destino = 'cotizador-view'

    res, err = None, None
    if request.method == 'POST':
        op = request.form.get('tipoOperacion', 'Importación').strip()
        orig = request.form.get('puertoOrigen', '').strip()
        dest = request.form.get('puertoDestino', '').strip()
        eq = request.form.get('tipoEquipo', '40 HC').strip()
        unidad = request.form.get('unidadPeso', 'TON').strip()
        peso_raw = request.form.get('pesoCarga', '').strip()
        val_raw = request.form.get('valorCarga', '').strip()

        try:
            peso_num = float(peso_raw)
            if math.isnan(peso_num) or math.isinf(peso_num):
                raise ValueError()
            peso_ton = peso_num / 1000.0 if unidad == 'KG' else peso_num
            if peso_ton < PESO_MIN_TON:
                err = f"El peso ingresado ({peso_ton:g} Ton) es menor al mínimo permitido para operar ({PESO_MIN_TON} Ton / 10 Kg)."
            elif peso_ton > PESO_MAX_TON:
                err = f"El peso ingresado ({peso_ton:,.2f} Ton) supera el límite máximo por cotización ({PESO_MAX_TON:,.0f} Ton)."
        except Exception:
            err = "Por favor ingrese un peso numérico válido mayor a cero."

        val_usd = 0.0
        if not err and val_raw != '':
            try:
                val_usd = float(val_raw)
                if math.isnan(val_usd) or math.isinf(val_usd):
                    raise ValueError()
                if val_usd < 0:
                    err = "El valor declarado de la carga en dólares no puede ser un número negativo."
                elif val_usd > VALOR_CARGA_MAX_USD:
                    err = f"El valor declarado de la carga excede el monto máximo permitido (${VALOR_CARGA_MAX_USD:,.0f} USD)."
            except Exception:
                err = "El valor declarado de la carga debe ser un monto numérico válido."

        if not err:
            if not orig or not dest or orig == dest:
                err = "Seleccione un puerto de origen y un puerto de destino válidos."
            elif eq not in ('20 HC', '40 HC'):
                err = "Seleccione un tipo de contenedor válido (20' HC o 40' HC)."
            else:
                p_orig = orig.split(' (')[0].strip()
                p_dest = dest.split(' (')[0].strip()
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT tc.* FROM tarifas_contenedores tc
                    JOIN puertos po ON tc.id_origen = po.id_puerto
                    JOIN puertos pd ON tc.id_destino = pd.id_puerto
                    WHERE tc.tipo_operacion = ? AND po.nombre_puerto = ? AND pd.nombre_puerto = ?
                ''', (op, p_orig, p_dest))
                fila = cursor.fetchone()
                conn.close()

                if not fila:
                    err = "La ruta seleccionada no se encuentra disponible en el tarifario vigente."
                else:
                    tope = 23.0 if eq == '20 HC' else 25.0
                    nombre_eq = "Contenedor 20' High Cube" if eq == '20 HC' else "Contenedor 40' High Cube"
                    cant = math.ceil(peso_ton / tope)

                    tarifa_min = fila['tarifa_20_min'] if eq == '20 HC' else fila['tarifa_40_min']
                    tarifa_max = fila['tarifa_20'] if eq == '20 HC' else fila['tarifa_40']

                    flete_min_usd = cant * tarifa_min
                    flete_max_usd = cant * tarifa_max
                    total_min_usd = flete_min_usd + val_usd
                    total_max_usd = flete_max_usd + val_usd

                    tasa_clp, fuente_dolar = obtener_tasa_dolar_clp()
                    total_min_clp = round(total_min_usd * tasa_clp)
                    total_max_clp = round(total_max_usd * tasa_clp)

                    tmin, tmax = fila['transito_min'], fila['transito_max']
                    tiene_rango_tarifa = (tarifa_min != tarifa_max)

                    res = {
                        'tipo_operacion': op,
                        'ruta_texto': f"{orig} ➔ {dest}",
                        'nombre_equipo': nombre_eq,
                        'tope_romana': f"{tope:g}",
                        'cantidad_contenedores': cant,
                        'tiene_rango_tarifa': tiene_rango_tarifa,
                        'tarifa_unitaria_texto': f"${tarifa_min:,.0f} a ${tarifa_max:,.0f} USD" if tiene_rango_tarifa else f"${tarifa_max:,.2f} USD",
                        'costo_flete_texto': f"${flete_min_usd:,.2f} – ${flete_max_usd:,.2f} USD" if tiene_rango_tarifa else f"${flete_max_usd:,.2f} USD",
                        'tiene_valor_carga': (val_usd > 0),
                        'valor_carga': f"{val_usd:,.2f}",
                        'costo_total': f"{total_max_usd:,.2f}",
                        'costo_total_rango_usd': f"${total_min_usd:,.2f} a ${total_max_usd:,.2f} USD" if tiene_rango_tarifa else f"${total_max_usd:,.2f} USD",
                        'costo_total_clp': f"${total_max_clp:,.0f} CLP".replace(',', '.'),
                        'costo_rango_clp': (
                            f"${total_min_clp:,.0f} a ${total_max_clp:,.0f} CLP".replace(',', '.')
                            if tiene_rango_tarifa else f"${total_max_clp:,.0f} CLP".replace(',', '.')
                        ),
                        'tasa_dolar': f"${tasa_clp:,.2f} CLP",
                        'fuente_dolar': fuente_dolar,
                        'texto_transito': f"{tmin} días" if tmin == tmax else f"{tmin} a {tmax} días",
                        'capacidad_maxima': f"{(cant * tope):g}",
                        'peso_ingresado': f"{peso_ton:,.2f}",
                        'alerta_romana_activa': (cant > 1)
                    }

    return renderizar_panel(tab_activa=pestaña_destino, resultado=res, error_cotizacion=err)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'rol' in session:
        return redirect(url_for('vista_principal'))
    error = None
    if request.method == 'POST':
        c = request.form.get('correo', '').strip().lower()
        p = request.form.get('password', '')
        conn = get_db_connection()
        usr = conn.execute(
            "SELECT u.*, r.nombre_rol AS rol FROM usuarios u JOIN roles r ON u.id_rol=r.id_rol WHERE LOWER(u.correo)=?",
            (c,)
        ).fetchone()
        conn.close()
        if usr and check_password_hash(usr['password_hash'], p):
            session.clear()
            session.update({
                'id_usuario': usr['id_usuario'],
                'nombre': usr['nombre'],
                'correo': usr['correo'],
                'rol': usr['rol']
            })
            return redirect('/panel_interno')
        error = 'Correo o contraseña incorrectos. Intente nuevamente.'
    return render_template('login.html', error=error)

@app.route('/crear_usuario', methods=['POST'])
def crear_usuario():
    if session.get('rol') != 'admin':
        return redirect(url_for('login'))

    n = request.form.get('nombre_nuevo', '').strip()
    c = request.form.get('correo_nuevo', '').strip().lower()
    p = request.form.get('password_nuevo', '')
    r = request.form.get('rol_nuevo', 'empleado').strip()

    if not n or not c or not p:
        return renderizar_panel('admin-view', error_usuario="Por favor complete el nombre, correo y contraseña.")
    if '@' not in c or ' ' in c or any(ch in c for ch in ("'", '"', ';')):
        return renderizar_panel('admin-view', error_usuario="Ingrese una dirección de correo corporativo válida.")
    if len(p) < 6 or len(p) > 64:
        return renderizar_panel('admin-view', error_usuario="La contraseña debe tener entre 6 y 64 caracteres.")
    if r not in ('admin', 'empleado'):
        return renderizar_panel('admin-view', error_usuario="Seleccione un perfil de acceso válido.")

    conn = get_db_connection()
    try:
        rol_id = conn.execute("SELECT id_rol FROM roles WHERE nombre_rol=?", (r,)).fetchone()['id_rol']
        conn.execute(
            "INSERT INTO usuarios (nombre, correo, password_hash, id_rol) VALUES (?, ?, ?, ?)",
            (n, c, generate_password_hash(p), rol_id)
        )
        conn.commit()
    except sqlite3.IntegrityError:
        conn.rollback()
        conn.close()
        return renderizar_panel('admin-view', error_usuario=f"El correo '{c}' ya se encuentra registrado.")
    except Exception:
        conn.rollback()
        conn.close()
        return renderizar_panel('admin-view', error_usuario="No se pudo registrar el usuario. Verifique los datos.")

    conn.close()
    nombre_perfil = "Administrador" if r == 'admin' else "Analista COMEX"
    return renderizar_panel(
        'admin-view',
        mensaje_usuario=f"Se creó la cuenta para {n} ({c}) con perfil de {nombre_perfil}."
    )

@app.route('/eliminar_usuario/<int:id_usuario>', methods=['POST'])
def eliminar_usuario(id_usuario):
    if session.get('rol') != 'admin':
        return redirect(url_for('login'))

    # Protección: impedir que el administrador elimine su propia sesión activa o la cuenta raíz
    if id_usuario == session.get('id_usuario'):
        return renderizar_panel('admin-view', error_usuario="Por seguridad no puede eliminar su propia cuenta con sesión activa.")

    conn = get_db_connection()
    try:
        usr = conn.execute("SELECT correo FROM usuarios WHERE id_usuario = ?", (id_usuario,)).fetchone()
        if not usr:
            conn.close()
            return renderizar_panel('admin-view', error_usuario="El usuario seleccionado ya no existe.")
        if usr['correo'].lower() == 'admin@cintac.cl':
            conn.close()
            return renderizar_panel('admin-view', error_usuario="La cuenta principal (admin@cintac.cl) está protegida y no puede eliminarse.")

        conn.execute("DELETE FROM usuarios WHERE id_usuario = ?", (id_usuario,))
        conn.commit()
        correo_borrado = usr['correo']
    except Exception:
        conn.rollback()
        conn.close()
        return renderizar_panel('admin-view', error_usuario="No se pudo eliminar la cuenta seleccionada.")

    conn.close()
    return renderizar_panel('admin-view', mensaje_usuario=f"La cuenta '{correo_borrado}' fue eliminada correctamente.")

@app.route('/cargar_matriz', methods=['POST'])
def cargar_matriz():
    if session.get('rol') != 'admin':
        return redirect(url_for('login'))

    f = request.files.get('archivo_matriz')
    if not f or not f.filename:
        return renderizar_panel('admin-view', error_matriz="Seleccione un archivo antes de subir.")
    if not f.filename.lower().endswith('.xlsx'):
        return renderizar_panel('admin-view', error_matriz="Formato no compatible. Suba únicamente un archivo Excel (.xlsx).")

    ok, res, consolidadas = validar_excel_tarifas(f)
    if not ok:
        return renderizar_panel('admin-view', error_matriz=res)

    conn = get_db_connection()
    try:
        for r in res:
            conn.execute(
                'INSERT OR IGNORE INTO puertos (nombre_puerto, pais, es_nacional) VALUES (?, ?, ?)',
                (r['puerto_origen'], r['pais_origen'], r['es_nac_origen'])
            )
            conn.execute(
                'INSERT OR IGNORE INTO puertos (nombre_puerto, pais, es_nacional) VALUES (?, ?, ?)',
                (r['puerto_destino'], r['pais_destino'], r['es_nac_destino'])
            )
        mp = {fila['nombre_puerto']: fila['id_puerto'] for fila in conn.execute('SELECT id_puerto, nombre_puerto FROM puertos').fetchall()}

        conn.execute('DELETE FROM tarifas_contenedores')
        for r in res:
            conn.execute('''
                INSERT INTO tarifas_contenedores
                (tipo_operacion, id_origen, id_destino, tipo_transporte, transito_min, transito_max, tarifa_20_min, tarifa_20, tarifa_40_min, tarifa_40)
                VALUES (?, ?, ?, 'Maritimo', ?, ?, ?, ?, ?, ?)
            ''', (
                r['tipo_operacion'], mp[r['puerto_origen']], mp[r['puerto_destino']],
                r['transito_min'], r['transito_max'], r['tarifa_20_min'], r['tarifa_20'], r['tarifa_40_min'], r['tarifa_40']
            ))
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        return renderizar_panel('admin-view', error_matriz=f"No se pudo actualizar el tarifario: {str(e)}")

    conn.close()
    detalle_extra = f" (se consolidaron {consolidadas} filas de referencia adicionales)" if consolidadas > 0 else ""
    return renderizar_panel(
        'admin-view',
        mensaje_matriz=f"El tarifario se actualizó correctamente con {len(res)} rutas marítimas oficiales desde '{f.filename}'{detalle_extra}."
    )

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)