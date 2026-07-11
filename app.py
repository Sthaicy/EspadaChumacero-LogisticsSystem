"""
ESPADACHUMACERO S.R.L. — Control de Inventarios v3.0
Compatible con ejecución directa (python app.py) y como .exe (PyInstaller)

CAMBIOS v3.0:
  - Salida de almacén: producto desde lista (solo comprados), unidad y cod autofill
  - Cantidades/precios con step=1 (suben de uno en uno)
  - Validación duplicados: código/nombre de proyecto, cod_prod en compras, herramientas, consumibles
  - Historial servicios y ventas: exportar Excel
  - Historial: precio total del proyecto en cabecera de cada bloque
  - Historial: egresos de almacén con costo y total egresos por proyecto
  - Todo conectado a listas reales (selects desde BD)
"""

import sys, os, socket, threading, webbrowser, tkinter as tk, time
from tkinter import messagebox
import pystray
from waitress import serve
from PIL import Image, ImageDraw
import sqlite3, zipfile, io
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── RUTAS DINÁMICAS ────────────────────────────────────────────────────────────
def get_base_dir():
    if getattr(sys, 'frozen', False): return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def get_runtime_dir():
    if getattr(sys, 'frozen', False): return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR    = get_base_dir()
RUNTIME_DIR = get_runtime_dir()
DB_PATH     = os.path.join(RUNTIME_DIR, 'chumacero.db')

# ── BACKUP AUTOMÁTICO ──────────────────────────────────────────────────────────
# Carpeta de backups: ./backups junto al ejecutable (separada de la carpeta de la BD)
BACKUP_DIR  = os.path.join(RUNTIME_DIR, 'backups')
BACKUP_LOG  = os.path.join(BACKUP_DIR, 'backup_log.txt')
MAX_BACKUPS = 30

def _log_backup(msg):
    """Escribe una línea en el log de backup."""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(BACKUP_LOG, 'a', encoding='utf-8') as f:
            f.write(f'[{ts}] {msg}\n')
    except Exception as e:
        print(f'[BACKUP LOG ERROR] {e}')

def hacer_backup():
    """Crea un backup comprimido de la BD y limpia los más antiguos."""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = f'backup_chumacero_{ts}.zip'
        dest = os.path.join(BACKUP_DIR, name)

        if not os.path.exists(DB_PATH):
            _log_backup(f'ERROR — BD no encontrada en: {DB_PATH}')
            return

        # Crear ZIP con la BD
        with zipfile.ZipFile(dest, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(DB_PATH, arcname='chumacero.db')

        size_kb = os.path.getsize(dest) // 1024
        _log_backup(f'OK    — {name}  ({size_kb} KB)')

        # Mantener solo los últimos MAX_BACKUPS
        archivos = sorted(
            [f for f in os.listdir(BACKUP_DIR)
             if f.startswith('backup_chumacero_') and f.endswith('.zip')],
            reverse=True
        )
        for viejo in archivos[MAX_BACKUPS:]:
            os.remove(os.path.join(BACKUP_DIR, viejo))
            _log_backup(f'BORRADO — {viejo} (límite {MAX_BACKUPS} backups)')

    except Exception as e:
        _log_backup(f'ERROR — {e}')

def _scheduler_backup():
    """Hilo que espera hasta las 12:00 y ejecuta el backup cada 24 h."""
    _log_backup('Scheduler iniciado')
    while True:
        ahora   = datetime.now()
        # Próxima ejecución: hoy a las 12:00, o mañana si ya pasó
        proxima = ahora.replace(hour=12, minute=0, second=0, microsecond=0)
        if ahora >= proxima:
            from datetime import timedelta
            proxima += timedelta(days=1)
        segundos = (proxima - ahora).total_seconds()
        time.sleep(segundos)
        hacer_backup()
        time.sleep(60)   # pequeña pausa antes de recalcular el ciclo siguiente

def iniciar_scheduler_backup():
    t = threading.Thread(target=_scheduler_backup, daemon=True, name='BackupScheduler')
    t.start()

from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static'))
app.secret_key = 'chumacero_espada_2024_secret'

# ── BASE DE DATOS ──────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db(); c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS Clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, nombre TEXT, identificacion TEXT,
        celular TEXT, direccion TEXT, correo TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Proveedores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, nombre TEXT, identificacion TEXT,
        celular TEXT, direccion TEXT, correo TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Trabajadores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, nombre TEXT, identificacion TEXT,
        celular TEXT, direccion TEXT, correo TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Servicios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, codigo_proyecto TEXT, nombre_proyecto TEXT,
        descripcion TEXT, cliente_id INTEGER, tipo_servicio TEXT,
        precio_total REAL DEFAULT 0, saldo_cobrar REAL DEFAULT 0,
        FOREIGN KEY(cliente_id) REFERENCES Clientes(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS Ingresos_Servicios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, codigo_proyecto TEXT, nombre_proyecto TEXT,
        descripcion TEXT, cliente_id INTEGER, tipo_servicio TEXT,
        monto_facturado REAL DEFAULT 0, nro_factura TEXT,
        monto REAL DEFAULT 0, tipo_pago TEXT, observaciones TEXT,
        servicio_id INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Ventas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, codigo_proyecto TEXT, nombre_proyecto TEXT,
        descripcion TEXT, cliente_id INTEGER,
        precio_total REAL DEFAULT 0, saldo_cobrar REAL DEFAULT 0,
        FOREIGN KEY(cliente_id) REFERENCES Clientes(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS Ingresos_Ventas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, codigo_proyecto TEXT, nombre_proyecto TEXT,
        descripcion TEXT, cliente_id INTEGER,
        monto_facturado REAL DEFAULT 0, nro_factura TEXT,
        monto REAL DEFAULT 0, tipo_pago TEXT, observaciones TEXT,
        venta_id INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Almacen (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_prod TEXT, insumo TEXT, entradas REAL DEFAULT 0,
        salidas REAL DEFAULT 0, stock REAL DEFAULT 0,
        unidad TEXT, descripcion_cliente TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Salidas_Almacen (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_prod TEXT, fecha TEXT, insumo TEXT,
        cantidad REAL DEFAULT 0, unidad TEXT,
        proyecto TEXT, codigo_proyecto TEXT,
        sobrantes REAL DEFAULT 0, costo REAL DEFAULT 0)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Compras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, proveedor_id INTEGER, cod_prod TEXT,
        insumo TEXT, cantidad REAL DEFAULT 0, unidad TEXT,
        precio_unitario REAL DEFAULT 0, precio_total REAL DEFAULT 0,
        tipo_documento TEXT, nro_documento TEXT,
        tipo_pago TEXT, detalle TEXT,
        FOREIGN KEY(proveedor_id) REFERENCES Proveedores(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS Herramientas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_ingreso TEXT, nombre TEXT, descripcion TEXT,
        marca TEXT, costo REAL DEFAULT 0, observaciones TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Asignacion_Herramientas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, herramienta_id INTEGER, descripcion TEXT,
        cantidad INTEGER DEFAULT 1, responsable_id INTEGER,
        observaciones TEXT, devuelto INTEGER DEFAULT 0, fecha_devolucion TEXT,
        FOREIGN KEY(herramienta_id) REFERENCES Herramientas(id),
        FOREIGN KEY(responsable_id) REFERENCES Trabajadores(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS Consumibles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT, descripcion TEXT, observaciones TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Entrega_Consumibles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT, consumible_id INTEGER, descripcion TEXT,
        cantidad REAL DEFAULT 0, responsable_id INTEGER,
        proyecto TEXT, observaciones TEXT,
        FOREIGN KEY(consumible_id) REFERENCES Consumibles(id),
        FOREIGN KEY(responsable_id) REFERENCES Trabajadores(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS Proyectos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT UNIQUE NOT NULL)''')

    # ── CATÁLOGO DE INSUMOS ──
    # Cada insumo es único y tiene su propio código de producto fijo.
    # Ese código funciona como PREFIJO para el correlativo del Almacén de Sobrantes
    # (ej: cod_prod "10MT_INX" → piezas de sobrante "10MT_INX_01", "10MT_INX_02"...).
    # El espesor es opcional: hay insumos que no se diferencian por espesor.
    c.execute('''CREATE TABLE IF NOT EXISTS Insumos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_prod TEXT UNIQUE NOT NULL,
        nombre TEXT NOT NULL,
        unidad TEXT NOT NULL,
        espesor TEXT,
        detalle TEXT)''')

    # Contador correlativo persistente por cod_prod de insumo — nunca decrece ni se reutiliza
    c.execute('''CREATE TABLE IF NOT EXISTS Sobrantes_Contadores (
        cod_insumo TEXT PRIMARY KEY,
        ultimo INTEGER DEFAULT 0)''')

    # Contador correlativo persistente para codigo_proyecto en Órdenes (SERV-) y Ventas (VENT-)
    c.execute('''CREATE TABLE IF NOT EXISTS Codigos_Proyecto_Contadores (
        prefijo TEXT PRIMARY KEY,
        ultimo INTEGER DEFAULT 0)''')

    # ── CATÁLOGOS CONFIGURABLES PARA COMPRAS ──
    c.execute('''CREATE TABLE IF NOT EXISTS Tipos_Pago (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        valor TEXT UNIQUE NOT NULL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Tipos_Documento (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        valor TEXT UNIQUE NOT NULL)''')

    # ── ALMACÉN DE SOBRANTES ── (cada fila = una pieza/lote único de sobrante)
    c.execute('''CREATE TABLE IF NOT EXISTS Almacen_Sobrantes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_prod TEXT UNIQUE, fecha_entrada TEXT, insumo TEXT, espesor TEXT, tamano TEXT,
        unidad TEXT, entradas REAL DEFAULT 1, salidas REAL DEFAULT 0, stock REAL DEFAULT 0,
        proyecto_origen TEXT, descripcion TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Entradas_Almacen_Sobrantes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_prod TEXT, fecha TEXT, insumo TEXT, espesor TEXT, tamano TEXT,
        unidad TEXT, cantidad REAL DEFAULT 0,
        proyecto_origen TEXT, descripcion TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS Salidas_Almacen_Sobrantes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cod_prod TEXT, fecha TEXT, insumo TEXT,
        cantidad REAL DEFAULT 0, unidad TEXT,
        proyecto_origen TEXT, proyecto_destino TEXT)''')

    # Migración defensiva: agrega proyecto_origen si la tabla ya existía sin esa columna
    cols = [r['name'] for r in c.execute("PRAGMA table_info(Salidas_Almacen_Sobrantes)").fetchall()]
    if 'proyecto_origen' not in cols:
        c.execute('ALTER TABLE Salidas_Almacen_Sobrantes ADD COLUMN proyecto_origen TEXT')

    # Siembra inicial de Tipos_Pago / Tipos_Documento (solo si están vacíos)
    if not c.execute('SELECT 1 FROM Tipos_Pago').fetchone():
        for v in ['Caja Chica','Unión','Ganadero','BCP','Otros']:
            c.execute('INSERT OR IGNORE INTO Tipos_Pago(valor) VALUES(?)', (v,))
    if not c.execute('SELECT 1 FROM Tipos_Documento').fetchone():
        for v in ['Factura','Recibo','Nota','Otro']:
            c.execute('INSERT OR IGNORE INTO Tipos_Documento(valor) VALUES(?)', (v,))

    conn.commit(); conn.close()

TIPOS_SERVICIO = ['Corte CNC','Corte y Plegado','Tornería','Fresado','Soldadura','Taladrado']

# ── BANDEJA DEL SISTEMA ────────────────────────────────────────────────────────
def obtener_ip_red():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80)); ip = s.getsockname()[0]; s.close(); return ip
    except: return "No disponible"

def mostrar_info_ip():
    ventana = tk.Tk()
    ventana.title("🌐 Acceder a EspadaChumacero")
    ventana.geometry("500x350"); ventana.resizable(False,False)
    ventana.configure(bg='#f0f0f0')
    ventana.protocol("WM_DELETE_WINDOW", ventana.destroy)
    tk.Label(ventana, text="ESPADACHUMACERO - Sistema de Inventarios",
             font=("Arial",14,"bold"), bg='#f0f0f0', fg='#2E86AB').pack(pady=10)
    frame = tk.Frame(ventana, bg='white', relief=tk.RIDGE, bd=2)
    frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
    tk.Label(frame, text="📌 ACCEDER DESDE ESTA PC:",
             font=("Arial",10,"bold"), bg='white').pack(anchor='w', padx=10, pady=5)
    tk.Label(frame, text="http://localhost:5000",
             fg="blue", bg='white', font=("Arial",10)).pack(anchor='w', padx=30)
    tk.Label(frame, text="http://127.0.0.1:5000",
             fg="blue", bg='white', font=("Arial",10)).pack(anchor='w', padx=30)
    tk.Label(frame, text="\n📡 ACCEDER DESDE OTROS DISPOSITIVOS (misma red):",
             font=("Arial",10,"bold"), bg='white').pack(anchor='w', padx=10, pady=5)
    ip_red = obtener_ip_red()
    if ip_red != "No disponible":
        tk.Label(frame, text=f"http://{ip_red}:5000",
                 fg="green", bg='white', font=("Arial",10,"bold")).pack(anchor='w', padx=30)
    else:
        tk.Label(frame, text="⚠️ No se detectó IP de red",
                 fg="red", bg='white').pack(anchor='w', padx=30)
    tk.Label(frame, text="\n🔧 CONFIGURAR FIREWALL (si no acceden):",
             font=("Arial",9,"bold"), bg='white').pack(anchor='w', padx=10, pady=5)
    cmd = 'netsh advfirewall firewall add rule name="Chumacero" dir=in action=allow protocol=tcp localport=5000'
    tk.Label(frame, text=cmd, font=("Consolas",7), bg='#f5f5f5', fg='#555',
             wraplength=420).pack(anchor='w', padx=35, pady=5)
    frame_botones = tk.Frame(ventana, bg='#f0f0f0')
    frame_botones.pack(pady=10)
    def abrir_local():
        webbrowser.open("http://localhost:5000"); ventana.destroy()
    def abrir_red():
        if ip_red != "No disponible":
            webbrowser.open(f"http://{ip_red}:5000"); ventana.destroy()
        else: messagebox.showerror("Error", "IP de red no disponible")
    tk.Button(frame_botones, text="🌐 Abrir Localhost", command=abrir_local,
              bg='#2E86AB', fg='white', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botones, text="📡 Abrir IP Red", command=abrir_red,
              bg='#28B463', fg='white', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botones, text="📋 Cerrar", command=ventana.destroy,
              bg='#E74C3C', fg='white', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
    ventana.attributes('-topmost', True)
    ventana.mainloop()

def detener_servidor(icono, item): icono.stop(); os._exit(0)
def abrir_navegador_bandeja(icono, item): webbrowser.open("http://127.0.0.1:5000")
def mostrar_ip(icono, item): mostrar_info_ip()
def backup_manual_bandeja(icono, item):
    threading.Thread(target=hacer_backup, daemon=True).start()

def crear_icono_bandeja():
    try:
        base_path = sys._MEIPASS if hasattr(sys,'_MEIPASS') else os.path.dirname(os.path.abspath(__file__))
        icono_path = os.path.join(base_path, "logo.ico")
        if os.path.exists(icono_path): image = Image.open(icono_path).resize((64,64))
        else:
            image = Image.new('RGB',(64,64),color='#2E86AB')
            draw = ImageDraw.Draw(image)
            draw.rectangle([10,10,54,54], outline='white', width=3)
            draw.text((22,18),"E", fill='white', size=30)
    except: image = Image.new('RGB',(64,64),color='#2E86AB')
    ip_red = obtener_ip_red()
    tooltip = "EspadaChumacero\nhttp://localhost:5000"
    if ip_red != "No disponible": tooltip += f"\nRed: http://{ip_red}:5000"
    menu = pystray.Menu(
        pystray.MenuItem("🌐 Abrir navegador", abrir_navegador_bandeja),
        pystray.MenuItem("📡 Mostrar IP de red", mostrar_ip),
        pystray.MenuItem("💾 Hacer backup ahora", backup_manual_bandeja),
        pystray.MenuItem("🟢 EspadaChumacero - Activo", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("❌ Salir", detener_servidor))
    icon = pystray.Icon("espadachumacero", image, tooltip, menu)
    icon.run()

# ══════════════════════════════════════════════════════════════════════════════
#  API HELPERS — usados por los selects del frontend
# ══════════════════════════════════════════════════════════════════════════════

@app.route('/api/almacen/<cod>')
def api_almacen_cod(cod):
    conn = get_db()
    row = conn.execute('SELECT * FROM Almacen WHERE cod_prod=?',(cod,)).fetchone(); conn.close()
    return jsonify(dict(row) if row else {})

@app.route('/api/productos_comprados')
def api_productos_comprados():
    """Lista de productos únicos del almacén (que entraron via compras)"""
    conn = get_db()
    rows = conn.execute('''SELECT DISTINCT cod_prod, insumo, unidad, stock
        FROM Almacen WHERE cod_prod IS NOT NULL AND cod_prod != ''
        ORDER BY insumo''').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/check_duplicado', methods=['POST'])
def api_check_duplicado():
    """Verifica si ya existe un código o nombre en una tabla dada"""
    d = request.json
    tabla  = d.get('tabla')
    campo  = d.get('campo')
    valor  = d.get('valor','').strip()
    exc_id = d.get('excluir_id')  # para ediciones, excluir el propio registro

    tablas_permitidas = {
        'Servicios':  ['codigo_proyecto','nombre_proyecto'],
        'Ventas':     ['codigo_proyecto','nombre_proyecto'],
        'Compras':    ['cod_prod'],
        'Herramientas': ['nombre'],
        'Consumibles':  ['nombre'],
        'Clientes':     ['nombre','identificacion'],
        'Proveedores':  ['nombre','identificacion'],
        'Trabajadores': ['nombre','identificacion'],
    }
    if tabla not in tablas_permitidas or campo not in tablas_permitidas[tabla]:
        return jsonify({'existe': False})

    conn = get_db()
    if exc_id:
        row = conn.execute(f'SELECT id FROM {tabla} WHERE {campo}=? AND id!=?',(valor,exc_id)).fetchone()
    else:
        row = conn.execute(f'SELECT id FROM {tabla} WHERE {campo}=?',(valor,)).fetchone()
    conn.close()
    return jsonify({'existe': row is not None})

@app.route('/api/proyectos')
def api_proyectos():
    """Todos los proyectos (servicios + ventas) para selects"""
    conn = get_db()
    serv = conn.execute('SELECT codigo_proyecto,nombre_proyecto FROM Servicios ORDER BY nombre_proyecto').fetchall()
    vent = conn.execute('SELECT codigo_proyecto,nombre_proyecto FROM Ventas ORDER BY nombre_proyecto').fetchall()
    conn.close()
    vistos = set()
    result = []
    for r in list(serv)+list(vent):
        key = r['codigo_proyecto']
        if key and key not in vistos:
            vistos.add(key)
            result.append(dict(r))
    return jsonify(result)

# ══════════════════════════════════════════════════════════════════════════════
#  RUTAS PRINCIPALES
# ══════════════════════════════════════════════════════════════════════════════

@app.route('/')
def index(): return render_template('index.html')

# ── CLIENTES ──
@app.route('/clientes')
def clientes():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Clientes ORDER BY id DESC').fetchall(); conn.close()
    return render_template('clientes.html', rows=rows, today=date.today().isoformat())

@app.route('/clientes/add', methods=['POST'])
def clientes_add():
    d = request.form; conn = get_db()
    conn.execute('INSERT INTO Clientes(fecha,nombre,identificacion,celular,direccion,correo) VALUES(?,?,?,?,?,?)',
        (d['fecha'],d['nombre'],d['identificacion'],d['celular'],d['direccion'],d['correo']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/clientes/edit/<int:rid>', methods=['POST'])
def clientes_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('UPDATE Clientes SET fecha=?,nombre=?,identificacion=?,celular=?,direccion=?,correo=? WHERE id=?',
        (d['fecha'],d['nombre'],d['identificacion'],d['celular'],d['direccion'],d['correo'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/clientes/delete/<int:rid>', methods=['POST'])
def clientes_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Clientes WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/clientes/export')
def clientes_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Clientes').fetchall(); conn.close()
    return export_excel(rows,['ID','Fecha','Nombre','ID/RUC','Celular','Dirección','Correo'],'Clientes')

# ── PROVEEDORES ──
@app.route('/proveedores')
def proveedores():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Proveedores ORDER BY id DESC').fetchall(); conn.close()
    return render_template('proveedores.html', rows=rows, today=date.today().isoformat())

@app.route('/proveedores/add', methods=['POST'])
def proveedores_add():
    d = request.form; conn = get_db()
    conn.execute('INSERT INTO Proveedores(fecha,nombre,identificacion,celular,direccion,correo) VALUES(?,?,?,?,?,?)',
        (d['fecha'],d['nombre'],d['identificacion'],d['celular'],d['direccion'],d['correo']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/proveedores/edit/<int:rid>', methods=['POST'])
def proveedores_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('UPDATE Proveedores SET fecha=?,nombre=?,identificacion=?,celular=?,direccion=?,correo=? WHERE id=?',
        (d['fecha'],d['nombre'],d['identificacion'],d['celular'],d['direccion'],d['correo'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/proveedores/delete/<int:rid>', methods=['POST'])
def proveedores_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Proveedores WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/proveedores/export')
def proveedores_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Proveedores').fetchall(); conn.close()
    return export_excel(rows,['ID','Fecha','Nombre','ID/NIT','Celular','Dirección','Correo'],'Proveedores')

# ── TRABAJADORES ──
@app.route('/trabajadores')
def trabajadores():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Trabajadores ORDER BY id DESC').fetchall(); conn.close()
    return render_template('trabajadores.html', rows=rows, today=date.today().isoformat())

@app.route('/trabajadores/add', methods=['POST'])
def trabajadores_add():
    d = request.form; conn = get_db()
    conn.execute('INSERT INTO Trabajadores(fecha,nombre,identificacion,celular,direccion,correo) VALUES(?,?,?,?,?,?)',
        (d['fecha'],d['nombre'],d['identificacion'],d['celular'],d['direccion'],d['correo']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/trabajadores/edit/<int:rid>', methods=['POST'])
def trabajadores_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('UPDATE Trabajadores SET fecha=?,nombre=?,identificacion=?,celular=?,direccion=?,correo=? WHERE id=?',
        (d['fecha'],d['nombre'],d['identificacion'],d['celular'],d['direccion'],d['correo'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/trabajadores/delete/<int:rid>', methods=['POST'])
def trabajadores_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Trabajadores WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/trabajadores/export')
def trabajadores_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Trabajadores').fetchall(); conn.close()
    return export_excel(rows,['ID','Fecha','Nombre','CI','Celular','Dirección','Correo'],'Trabajadores')

# ── SERVICIOS ──
@app.route('/ordenes')
def ordenes():
    conn = get_db()
    proyectos = conn.execute('''SELECT * FROM Servicios
        WHERE saldo_cobrar > 0 ORDER BY id DESC''').fetchall()
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    conn.close()
    return render_template('ordenes.html', proyectos=proyectos,
        today=date.today().isoformat(), tipos_servicio=TIPOS_SERVICIO, tipos_pago=tipos_pago)

@app.route('/ordenes/registrar', methods=['POST'])
def ordenes_registrar():
    d = request.form; precio = float(d.get('precio_total') or 0); conn = get_db()
    nom = d.get('nombre_proyecto','').strip()
    if conn.execute('SELECT id FROM Servicios WHERE nombre_proyecto=?',(nom,)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe un proyecto con el nombre "{nom}"'})
    cod = _siguiente_codigo_proyecto(conn, 'SERV')
    conn.execute('''INSERT INTO Servicios(fecha,codigo_proyecto,nombre_proyecto,descripcion,
        cliente_id,tipo_servicio,precio_total,saldo_cobrar) VALUES(?,?,?,?,?,?,?,?)''',
        (d['fecha'],cod,nom,d['descripcion'],d.get('cliente_id') or None,
         d.get('tipo_servicio',''),precio,precio))
    conn.commit(); conn.close(); return jsonify({'ok':True,'codigo_proyecto':cod})

@app.route('/ordenes/pago/<int:sid>', methods=['POST'])
def ordenes_pago(sid):
    d = request.form
    mf = float(d.get('monto_facturado') or 0)
    mo = float(d.get('monto') or 0)
    conn = get_db()
    s = conn.execute('SELECT * FROM Servicios WHERE id=?',(sid,)).fetchone()
    if not s: conn.close(); return jsonify({'ok':False,'msg':'Proyecto no encontrado'})
    conn.execute('''INSERT INTO Ingresos_Servicios(fecha,codigo_proyecto,nombre_proyecto,
        descripcion,cliente_id,tipo_servicio,monto_facturado,nro_factura,monto,
        tipo_pago,observaciones,servicio_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
        (d['fecha'],s['codigo_proyecto'],s['nombre_proyecto'],s['descripcion'],
         s['cliente_id'],s['tipo_servicio'],mf,d.get('nro_factura',''),
         mo,d.get('tipo_pago',''),d.get('observaciones',''),sid))
    conn.execute('UPDATE Servicios SET saldo_cobrar=? WHERE id=?',(max(0,s['saldo_cobrar']-mf-mo),sid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/edit/<int:sid>', methods=['POST'])
def ordenes_edit(sid):
    d = request.form; conn = get_db()
    conn.execute('''UPDATE Servicios SET fecha=?,nombre_proyecto=?,descripcion=?,
        cliente_id=?,tipo_servicio=?,precio_total=? WHERE id=?''',
        (d['fecha'],d['nombre_proyecto'],d['descripcion'],d.get('cliente_id') or None,
         d.get('tipo_servicio',''),float(d.get('precio_total') or 0),sid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/delete/<int:sid>', methods=['POST'])
def ordenes_delete(sid):
    conn = get_db(); conn.execute('DELETE FROM Servicios WHERE id=?',(sid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/ingresos')
def ordenes_ingresos():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Ingresos_Servicios ORDER BY id DESC').fetchall()
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    conn.close()
    return render_template('ordenes_ingresos.html', rows=rows, tipos_pago=tipos_pago)

@app.route('/ordenes/ingresos/edit/<int:rid>', methods=['POST'])
def ordenes_ingresos_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('''UPDATE Ingresos_Servicios SET fecha=?,monto_facturado=?,
        nro_factura=?,monto=?,tipo_pago=?,observaciones=? WHERE id=?''',
        (d['fecha'],float(d.get('monto_facturado') or 0),d['nro_factura'],
         float(d.get('monto') or 0),d['tipo_pago'],d['observaciones'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/ingresos/delete/<int:rid>', methods=['POST'])
def ordenes_ingresos_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Ingresos_Servicios WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/historial')
def ordenes_historial():
    conn = get_db()
    ingresos = conn.execute('SELECT * FROM Ingresos_Servicios ORDER BY codigo_proyecto,id').fetchall()
    salidas  = conn.execute('''SELECT sa.*,
        COALESCE((SELECT AVG(precio_unitario) FROM Compras WHERE cod_prod=sa.cod_prod),0) as costo_unitario
        FROM Salidas_Almacen sa ORDER BY sa.codigo_proyecto,sa.id''').fetchall()
    # Precio total de cada proyecto (para mostrar en cabecera)
    proyectos_info = {}
    for s in conn.execute('SELECT codigo_proyecto,nombre_proyecto,precio_total FROM Servicios').fetchall():
        proyectos_info[s['codigo_proyecto']] = dict(s)
    pi = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in ingresos]))
    ps = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in salidas]))
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    conn.close()
    return render_template('ordenes_historial.html', ingresos=ingresos, salidas=salidas,
        proyectos=list(dict.fromkeys(pi+ps)), proyectos_info=proyectos_info, tipos_pago=tipos_pago)

@app.route('/ordenes/historial/edit/<int:rid>', methods=['POST'])
def ordenes_historial_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('''UPDATE Ingresos_Servicios SET fecha=?,monto_facturado=?,
        nro_factura=?,monto=?,tipo_pago=?,observaciones=? WHERE id=?''',
        (d['fecha'],float(d.get('monto_facturado') or 0),d['nro_factura'],
         float(d.get('monto') or 0),d['tipo_pago'],d['observaciones'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/historial/delete/<int:rid>', methods=['POST'])
def ordenes_historial_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Ingresos_Servicios WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ordenes/historial/export')
def ordenes_historial_export():
    conn = get_db()
    rows = conn.execute('SELECT fecha,codigo_proyecto,nombre_proyecto,monto_facturado,nro_factura,monto,tipo_pago,observaciones FROM Ingresos_Servicios ORDER BY codigo_proyecto,id').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Historial_Servicios')

@app.route('/ordenes/export')
def ordenes_export():
    conn = get_db()
    rows = conn.execute('SELECT codigo_proyecto,nombre_proyecto,tipo_servicio,precio_total,saldo_cobrar FROM Servicios WHERE saldo_cobrar>0').fetchall(); conn.close()
    return export_excel(rows,['Cód.','Proyecto','Tipo Servicio','Precio Total','Saldo'],'Servicios')

@app.route('/ordenes/ingresos/export')
def ordenes_ingresos_export():
    conn = get_db()
    rows = conn.execute('SELECT fecha,codigo_proyecto,nombre_proyecto,monto_facturado,nro_factura,monto,tipo_pago,observaciones FROM Ingresos_Servicios').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Ingresos_Servicios')

# ── VENTAS ──
@app.route('/ventas')
def ventas():
    conn = get_db()
    proyectos = conn.execute('''SELECT * FROM Ventas
        WHERE saldo_cobrar>0 ORDER BY id DESC''').fetchall()
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    conn.close()
    return render_template('ventas.html', proyectos=proyectos,
        today=date.today().isoformat(), tipos_pago=tipos_pago)

@app.route('/ventas/registrar', methods=['POST'])
def ventas_registrar():
    d = request.form; precio=float(d.get('precio_total') or 0); conn=get_db()
    nom = d.get('nombre_proyecto','').strip()
    if conn.execute('SELECT id FROM Ventas WHERE nombre_proyecto=?',(nom,)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe una venta con el nombre "{nom}"'})
    cod = _siguiente_codigo_proyecto(conn, 'VENT')
    conn.execute('''INSERT INTO Ventas(fecha,codigo_proyecto,nombre_proyecto,descripcion,
        cliente_id,precio_total,saldo_cobrar) VALUES(?,?,?,?,?,?,?)''',
        (d['fecha'],cod,nom,d['descripcion'],d.get('cliente_id') or None,precio,precio))
    conn.commit(); conn.close(); return jsonify({'ok':True,'codigo_proyecto':cod})

@app.route('/ventas/pago/<int:vid>', methods=['POST'])
def ventas_pago(vid):
    d = request.form
    mf = float(d.get('monto_facturado') or 0)
    mo = float(d.get('monto') or 0)
    conn = get_db()
    v = conn.execute('SELECT * FROM Ventas WHERE id=?',(vid,)).fetchone()
    if not v: conn.close(); return jsonify({'ok':False,'msg':'Proyecto no encontrado'})
    conn.execute('''INSERT INTO Ingresos_Ventas(fecha,codigo_proyecto,nombre_proyecto,
        descripcion,cliente_id,monto_facturado,nro_factura,monto,tipo_pago,observaciones,venta_id)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
        (d['fecha'],v['codigo_proyecto'],v['nombre_proyecto'],v['descripcion'],
         v['cliente_id'],mf,d.get('nro_factura',''),mo,d.get('tipo_pago',''),d.get('observaciones',''),vid))
    conn.execute('UPDATE Ventas SET saldo_cobrar=? WHERE id=?',(max(0,v['saldo_cobrar']-mf-mo),vid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/edit/<int:vid>', methods=['POST'])
def ventas_edit(vid):
    d = request.form; conn = get_db()
    conn.execute('''UPDATE Ventas SET fecha=?,nombre_proyecto=?,descripcion=?,
        cliente_id=?,precio_total=? WHERE id=?''',
        (d['fecha'],d['nombre_proyecto'],d['descripcion'],d.get('cliente_id') or None,
         float(d.get('precio_total') or 0),vid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/delete/<int:vid>', methods=['POST'])
def ventas_delete(vid):
    conn = get_db(); conn.execute('DELETE FROM Ventas WHERE id=?',(vid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/ingresos')
def ventas_ingresos():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Ingresos_Ventas ORDER BY id DESC').fetchall()
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    conn.close()
    return render_template('ventas_ingresos.html', rows=rows, tipos_pago=tipos_pago)

@app.route('/ventas/ingresos/edit/<int:rid>', methods=['POST'])
def ventas_ingresos_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('''UPDATE Ingresos_Ventas SET fecha=?,monto_facturado=?,
        nro_factura=?,monto=?,tipo_pago=?,observaciones=? WHERE id=?''',
        (d['fecha'],float(d.get('monto_facturado') or 0),d['nro_factura'],
         float(d.get('monto') or 0),d['tipo_pago'],d['observaciones'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/ingresos/delete/<int:rid>', methods=['POST'])
def ventas_ingresos_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Ingresos_Ventas WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/historial')
def ventas_historial():
    conn = get_db()
    ingresos = conn.execute('SELECT * FROM Ingresos_Ventas ORDER BY codigo_proyecto,id').fetchall()
    salidas  = conn.execute('''SELECT sa.*,
        COALESCE((SELECT AVG(precio_unitario) FROM Compras WHERE cod_prod=sa.cod_prod),0) as costo_unitario
        FROM Salidas_Almacen sa ORDER BY sa.codigo_proyecto,sa.id''').fetchall()
    proyectos_info = {}
    for v in conn.execute('SELECT codigo_proyecto,nombre_proyecto,precio_total FROM Ventas').fetchall():
        proyectos_info[v['codigo_proyecto']] = dict(v)
    pi = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in ingresos]))
    ps = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in salidas]))
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    conn.close()
    return render_template('ventas_historial.html', ingresos=ingresos, salidas=salidas,
        proyectos=list(dict.fromkeys(pi+ps)), proyectos_info=proyectos_info, tipos_pago=tipos_pago)

@app.route('/ventas/historial/edit/<int:rid>', methods=['POST'])
def ventas_historial_edit(rid):
    d = request.form; conn = get_db()
    conn.execute('''UPDATE Ingresos_Ventas SET fecha=?,monto_facturado=?,
        nro_factura=?,monto=?,tipo_pago=?,observaciones=? WHERE id=?''',
        (d['fecha'],float(d.get('monto_facturado') or 0),d['nro_factura'],
         float(d.get('monto') or 0),d['tipo_pago'],d['observaciones'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/historial/delete/<int:rid>', methods=['POST'])
def ventas_historial_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Ingresos_Ventas WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/ventas/historial/export')
def ventas_historial_export():
    conn = get_db()
    rows = conn.execute('SELECT fecha,codigo_proyecto,nombre_proyecto,monto_facturado,nro_factura,monto,tipo_pago,observaciones FROM Ingresos_Ventas ORDER BY codigo_proyecto,id').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Historial_Ventas')

@app.route('/ventas/export')
def ventas_export():
    conn = get_db()
    rows = conn.execute('SELECT codigo_proyecto,nombre_proyecto,precio_total,saldo_cobrar FROM Ventas WHERE saldo_cobrar>0').fetchall(); conn.close()
    return export_excel(rows,['Cód.','Proyecto','Precio Total','Saldo'],'Ventas')

@app.route('/ventas/ingresos/export')
def ventas_ingresos_export():
    conn = get_db()
    rows = conn.execute('SELECT fecha,codigo_proyecto,nombre_proyecto,monto_facturado,nro_factura,monto,tipo_pago,observaciones FROM Ingresos_Ventas').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Ingresos_Ventas')

# ── PROYECTOS ──
@app.route('/proyectos')
def proyectos_page():
    conn = get_db()
    proy = conn.execute('SELECT * FROM Proyectos ORDER BY nombre').fetchall()
    salidas = conn.execute('''
        SELECT sa.proyecto, sa.cantidad,
               COALESCE((SELECT AVG(precio_unitario) FROM Compras WHERE cod_prod=sa.cod_prod),0) as costo_unit
        FROM Salidas_Almacen sa WHERE sa.proyecto IS NOT NULL AND sa.proyecto != ""
    ''').fetchall()
    gastos = {}
    for s in salidas:
        nom = s['proyecto']
        gastos[nom] = gastos.get(nom, 0) + (s['cantidad'] or 0) * (s['costo_unit'] or 0)
    total_gasto = sum(gastos.values())
    conn.close()
    return render_template('proyectos.html', proyectos=proy, gastos=gastos, total_gasto=total_gasto)

@app.route('/proyectos/check', methods=['POST'])
def proyectos_check():
    data = request.get_json()
    nombre = (data.get('nombre') or '').strip()
    conn = get_db()
    existe = conn.execute('SELECT id FROM Proyectos WHERE LOWER(nombre)=LOWER(?)', (nombre,)).fetchone()
    conn.close()
    return jsonify({'existe': existe is not None})

@app.route('/proyectos/add', methods=['POST'])
def proyectos_add():
    nombre = (request.form.get('nombre') or '').strip()
    if not nombre:
        return jsonify({'ok': False, 'msg': 'El nombre es obligatorio'})
    conn = get_db()
    if conn.execute('SELECT id FROM Proyectos WHERE LOWER(nombre)=LOWER(?)', (nombre,)).fetchone():
        conn.close()
        return jsonify({'ok': False, 'msg': 'Ya existe un proyecto con ese nombre'})
    conn.execute('INSERT INTO Proyectos(nombre) VALUES(?)', (nombre,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/proyectos/delete/<int:rid>', methods=['POST'])
def proyectos_delete(rid):
    conn = get_db()
    conn.execute('DELETE FROM Proyectos WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

# ── INSUMOS ── (catálogo único: cada insumo tiene su propio Cód. Prod. fijo,
# que funciona como prefijo para el correlativo del Almacén de Sobrantes)
@app.route('/api/insumos')
def api_insumos():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Insumos ORDER BY nombre').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/insumos')
def insumos():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Insumos ORDER BY nombre').fetchall()
    conn.close()
    return render_template('insumos.html', rows=rows)

@app.route('/insumos/add', methods=['POST'])
def insumos_add():
    cod_prod = (request.form.get('cod_prod') or '').strip().upper()
    nombre   = (request.form.get('nombre') or '').strip()
    unidad   = (request.form.get('unidad') or '').strip()
    espesor  = (request.form.get('espesor') or '').strip()
    detalle  = (request.form.get('detalle') or '').strip()
    if not cod_prod or not nombre or not unidad:
        return jsonify({'ok': False, 'msg': 'Cód. Prod., Insumo y Unidad son obligatorios'})
    conn = get_db()
    if conn.execute('SELECT id FROM Insumos WHERE LOWER(cod_prod)=LOWER(?)', (cod_prod,)).fetchone():
        conn.close()
        return jsonify({'ok': False, 'msg': 'Ese Cód. Prod. ya está en uso por otro insumo'})
    if conn.execute('SELECT id FROM Insumos WHERE LOWER(nombre)=LOWER(?)', (nombre,)).fetchone():
        conn.close()
        return jsonify({'ok': False, 'msg': 'Ese insumo ya existe'})
    conn.execute('INSERT INTO Insumos(cod_prod,nombre,unidad,espesor,detalle) VALUES(?,?,?,?,?)',
        (cod_prod, nombre, unidad, espesor, detalle))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/insumos/edit/<int:rid>', methods=['POST'])
def insumos_edit(rid):
    d = request.form; conn = get_db()
    nombre = (d.get('nombre') or '').strip()
    unidad = (d.get('unidad') or '').strip()
    espesor = (d.get('espesor') or '').strip()
    detalle = (d.get('detalle') or '').strip()
    if not nombre or not unidad:
        conn.close(); return jsonify({'ok': False, 'msg': 'Insumo y Unidad son obligatorios'})
    dup = conn.execute('SELECT id FROM Insumos WHERE LOWER(nombre)=LOWER(?) AND id!=?', (nombre, rid)).fetchone()
    if dup:
        conn.close(); return jsonify({'ok': False, 'msg': 'Ya existe otro insumo con ese nombre'})
    conn.execute('UPDATE Insumos SET nombre=?,unidad=?,espesor=?,detalle=? WHERE id=?',
        (nombre, unidad, espesor, detalle, rid))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/insumos/delete/<int:rid>', methods=['POST'])
def insumos_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Insumos WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/insumos/export')
def insumos_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Insumos ORDER BY nombre').fetchall(); conn.close()
    return export_excel(rows, ['ID','Cód. Prod.','Insumo','Unidad','Espesor','Detalle'], 'Insumos')

# ── TIPOS DE PAGO / TIPOS DE DOCUMENTO (configurables desde Configuraciones) ──
@app.route('/configuraciones/tipos_pago/add', methods=['POST'])
def tipos_pago_add():
    valor = (request.form.get('valor') or '').strip()
    if not valor:
        return jsonify({'ok': False, 'msg': 'El valor es obligatorio'})
    conn = get_db()
    if conn.execute('SELECT id FROM Tipos_Pago WHERE LOWER(valor)=LOWER(?)', (valor,)).fetchone():
        conn.close()
        return jsonify({'ok': False, 'msg': 'Ese tipo de pago ya existe'})
    conn.execute('INSERT INTO Tipos_Pago(valor) VALUES(?)', (valor,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/configuraciones/tipos_pago/delete/<int:rid>', methods=['POST'])
def tipos_pago_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Tipos_Pago WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/configuraciones/tipos_documento/add', methods=['POST'])
def tipos_documento_add():
    valor = (request.form.get('valor') or '').strip()
    if not valor:
        return jsonify({'ok': False, 'msg': 'El valor es obligatorio'})
    conn = get_db()
    if conn.execute('SELECT id FROM Tipos_Documento WHERE LOWER(valor)=LOWER(?)', (valor,)).fetchone():
        conn.close()
        return jsonify({'ok': False, 'msg': 'Ese tipo de documento ya existe'})
    conn.execute('INSERT INTO Tipos_Documento(valor) VALUES(?)', (valor,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/configuraciones/tipos_documento/delete/<int:rid>', methods=['POST'])
def tipos_documento_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Tipos_Documento WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

# ── ALMACÉN DE SOBRANTES ──
# Cada fila es una PIEZA ÚNICA de sobrante (un corte = un código de producto propio).
# El código se genera como {PREFIJO}_{NN}, con NN correlativo y SIEMPRE creciente
# por prefijo — nunca se reutiliza aunque se borren registros.
def _siguiente_cod_sobrante(conn, cod_insumo):
    conn.execute(
        "INSERT INTO Sobrantes_Contadores(cod_insumo,ultimo) VALUES(?,1) "
        "ON CONFLICT(cod_insumo) DO UPDATE SET ultimo=ultimo+1", (cod_insumo,))
    row = conn.execute('SELECT ultimo FROM Sobrantes_Contadores WHERE cod_insumo=?', (cod_insumo,)).fetchone()
    return f"{cod_insumo}_{row['ultimo']:02d}"

def _siguiente_codigo_proyecto(conn, prefijo):
    """Genera el siguiente código de proyecto para Órdenes (SERV-) o Ventas (VENT-).
    Correlativo persistente — nunca decrece ni se reutiliza, igual que en Sobrantes."""
    conn.execute(
        "INSERT INTO Codigos_Proyecto_Contadores(prefijo,ultimo) VALUES(?,1) "
        "ON CONFLICT(prefijo) DO UPDATE SET ultimo=ultimo+1", (prefijo,))
    row = conn.execute('SELECT ultimo FROM Codigos_Proyecto_Contadores WHERE prefijo=?', (prefijo,)).fetchone()
    return f"{prefijo}-{row['ultimo']:04d}"

@app.route('/api/sobrantes/siguiente_cod/<insumo>')
def api_sobrantes_siguiente_cod(insumo):
    conn = get_db()
    ins = conn.execute('SELECT * FROM Insumos WHERE nombre=?', (insumo,)).fetchone()
    if not ins:
        conn.close()
        return jsonify({'cod_prod': '', 'error': 'Insumo no encontrado en catálogo'})
    # Vista previa sin consumir el contador: muestra cuál sería el próximo número
    row = conn.execute('SELECT ultimo FROM Sobrantes_Contadores WHERE cod_insumo=?', (ins['cod_prod'],)).fetchone()
    siguiente = (row['ultimo'] + 1) if row else 1
    conn.close()
    return jsonify({'cod_prod': f"{ins['cod_prod']}_{siguiente:02d}", 'unidad': ins['unidad'], 'espesor': ins['espesor'] or ''})

@app.route('/sobrantes')
def sobrantes():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Almacen_Sobrantes ORDER BY insumo, cod_prod').fetchall()
    insumos = conn.execute('SELECT * FROM Insumos ORDER BY nombre').fetchall()
    proyectos = conn.execute('SELECT id, nombre FROM Proyectos ORDER BY nombre').fetchall()
    conn.close()
    return render_template('sobrantes.html', rows=rows, insumos=insumos,
        proyectos=proyectos, today=date.today().isoformat())

@app.route('/sobrantes/entrada', methods=['POST'])
def sobrantes_entrada():
    d = request.form; conn = get_db()
    cant = float(d.get('cantidad') or 0)
    insumo = d['insumo']; tamano = d.get('tamano') or ''
    proyecto = d.get('proyecto_origen') or ''
    descripcion = d.get('descripcion') or ''

    ins = conn.execute('SELECT * FROM Insumos WHERE nombre=?', (insumo,)).fetchone()
    if not ins:
        conn.close(); return jsonify({'ok': False, 'msg': 'Insumo no encontrado en el catálogo'})

    # Espesor y unidad vienen siempre del catálogo de Insumos — no se capturan a mano.
    espesor = ins['espesor'] or ''
    unidad = ins['unidad']

    # El cód. prod. siempre se genera nuevo — cada pieza de sobrante es única,
    # aunque coincida insumo/espesor/tamaño con una entrada anterior.
    cod_prod = _siguiente_cod_sobrante(conn, ins['cod_prod'])

    conn.execute('''INSERT INTO Entradas_Almacen_Sobrantes
        (cod_prod,fecha,insumo,espesor,tamano,unidad,cantidad,proyecto_origen,descripcion)
        VALUES(?,?,?,?,?,?,?,?,?)''',
        (cod_prod, d['fecha'], insumo, espesor, tamano, unidad, cant, proyecto, descripcion))

    conn.execute('''INSERT INTO Almacen_Sobrantes
        (cod_prod,fecha_entrada,insumo,espesor,tamano,unidad,entradas,salidas,stock,proyecto_origen,descripcion)
        VALUES(?,?,?,?,?,?,?,0,?,?,?)''',
        (cod_prod, d['fecha'], insumo, espesor, tamano, unidad, cant, cant, proyecto, descripcion))
    conn.commit(); conn.close()
    return jsonify({'ok': True, 'cod_prod': cod_prod})

@app.route('/sobrantes/salida', methods=['POST'])
def sobrantes_salida():
    d = request.form; conn = get_db()
    cant = float(d.get('cantidad') or 0)
    cod_prod = d['cod_prod']

    row = conn.execute('SELECT * FROM Almacen_Sobrantes WHERE cod_prod=?', (cod_prod,)).fetchone()
    if not row:
        conn.close(); return jsonify({'ok': False, 'msg': 'Producto no encontrado'})
    if cant > (row['stock'] or 0):
        conn.close(); return jsonify({'ok': False, 'msg': 'Stock insuficiente'})

    conn.execute('''INSERT INTO Salidas_Almacen_Sobrantes
        (cod_prod,fecha,insumo,cantidad,unidad,proyecto_origen,proyecto_destino) VALUES(?,?,?,?,?,?,?)''',
        (cod_prod, d['fecha'], d['insumo'], cant, d['unidad'], row['proyecto_origen'] or '', d.get('proyecto_destino') or ''))

    nuevo_stock = (row['stock'] or 0) - cant
    if nuevo_stock <= 0:
        # Stock agotado: la pieza desaparece del Almacén de Sobrantes
        conn.execute('DELETE FROM Almacen_Sobrantes WHERE cod_prod=?', (cod_prod,))
    else:
        conn.execute('UPDATE Almacen_Sobrantes SET salidas=salidas+?,stock=? WHERE cod_prod=?',
            (cant, nuevo_stock, cod_prod))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/edit/<int:rid>', methods=['POST'])
def sobrantes_edit(rid):
    d = request.form; conn = get_db()
    # Insumo/Espesor/Unidad quedan fijos: ya determinaron el cod_prod de esta pieza.
    conn.execute('''UPDATE Almacen_Sobrantes SET tamano=?,
        proyecto_origen=?,descripcion=? WHERE id=?''',
        (d.get('tamano') or '', d.get('proyecto_origen') or '', d.get('descripcion') or '', rid))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/delete/<int:rid>', methods=['POST'])
def sobrantes_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Almacen_Sobrantes WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/export')
def sobrantes_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Almacen_Sobrantes ORDER BY insumo, cod_prod').fetchall(); conn.close()
    return export_excel(rows, ['ID','Cód. Prod.','Fecha Entrada','Insumo','Espesor','Tamaño','Unidad',
        'Entradas','Salidas','Stock','Proyecto Origen','Descripción'], 'Almacen_Sobrantes')

@app.route('/sobrantes/entradas')
def sobrantes_entradas():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Entradas_Almacen_Sobrantes ORDER BY id DESC').fetchall()
    insumos = conn.execute('SELECT * FROM Insumos ORDER BY nombre').fetchall()
    proyectos = conn.execute('SELECT id, nombre FROM Proyectos ORDER BY nombre').fetchall()
    conn.close()
    return render_template('sobrantes_entradas.html', rows=rows, insumos=insumos, proyectos=proyectos)

@app.route('/sobrantes/entradas/edit/<int:rid>', methods=['POST'])
def sobrantes_entradas_edit(rid):
    d = request.form; conn = get_db()
    old = conn.execute('SELECT * FROM Entradas_Almacen_Sobrantes WHERE id=?', (rid,)).fetchone()
    nueva_cant = float(d.get('cantidad') or 0)
    diff = nueva_cant - float(old['cantidad'] or 0)
    tamano = d.get('tamano') or ''
    proyecto_origen = d.get('proyecto_origen') or ''
    descripcion = d.get('descripcion') or ''
    # Insumo/Espesor/Unidad quedan fijos: ya determinaron el cod_prod de esta pieza
    # al momento de la entrada original, y cambiarlos dejaría el código incoherente.
    # Solo se ajustan fecha, tamaño, cantidad, proyecto origen y descripción.
    conn.execute('''UPDATE Entradas_Almacen_Sobrantes SET fecha=?,
        tamano=?,cantidad=?,proyecto_origen=?,descripcion=? WHERE id=?''',
        (d['fecha'], tamano, nueva_cant, proyecto_origen, descripcion, rid))
    conn.execute('UPDATE Almacen_Sobrantes SET tamano=?,proyecto_origen=?,descripcion=? WHERE cod_prod=?',
        (tamano, proyecto_origen, descripcion, old['cod_prod']))
    if diff != 0:
        conn.execute('UPDATE Almacen_Sobrantes SET entradas=entradas+?,stock=stock+? WHERE cod_prod=?',
            (diff, diff, old['cod_prod']))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/entradas/delete/<int:rid>', methods=['POST'])
def sobrantes_entradas_delete(rid):
    conn = get_db()
    row = conn.execute('SELECT * FROM Entradas_Almacen_Sobrantes WHERE id=?', (rid,)).fetchone()
    if row:
        conn.execute('UPDATE Almacen_Sobrantes SET entradas=entradas-?,stock=stock-? WHERE cod_prod=?',
            (row['cantidad'], row['cantidad'], row['cod_prod']))
        conn.execute('DELETE FROM Almacen_Sobrantes WHERE cod_prod=? AND stock<=0', (row['cod_prod'],))
    conn.execute('DELETE FROM Entradas_Almacen_Sobrantes WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/entradas/export')
def sobrantes_entradas_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Entradas_Almacen_Sobrantes ORDER BY id DESC').fetchall()
    conn.close()
    return export_excel(rows, ['ID','Cód.','Fecha','Insumo','Espesor','Tamaño','Unidad',
        'Cantidad','Proyecto Origen','Descripción'], 'Entradas_Almacen_Sobrantes')

@app.route('/sobrantes/salidas')
def sobrantes_salidas():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Salidas_Almacen_Sobrantes ORDER BY id DESC').fetchall()
    proyectos = conn.execute('SELECT id, nombre FROM Proyectos ORDER BY nombre').fetchall()
    conn.close()
    return render_template('sobrantes_salidas.html', rows=rows, proyectos=proyectos)

@app.route('/sobrantes/salidas/edit/<int:rid>', methods=['POST'])
def sobrantes_salidas_edit(rid):
    d = request.form; conn = get_db()
    old = conn.execute('SELECT * FROM Salidas_Almacen_Sobrantes WHERE id=?', (rid,)).fetchone()
    nueva_cant = float(d.get('cantidad') or 0)
    diff = nueva_cant - float(old['cantidad'] or 0)  # diff>0 = sale más stock
    # Insumo/Unidad pertenecen a la pieza original (cod_prod) y no se reasignan aquí —
    # solo se ajustan fecha, cantidad y proyecto destino.
    conn.execute('''UPDATE Salidas_Almacen_Sobrantes SET fecha=?,
        cantidad=?,proyecto_destino=? WHERE id=?''',
        (d['fecha'], nueva_cant, d.get('proyecto_destino') or '', rid))
    if diff != 0:
        conn.execute('UPDATE Almacen_Sobrantes SET salidas=salidas+?,stock=stock-? WHERE cod_prod=?',
            (diff, diff, old['cod_prod']))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/salidas/delete/<int:rid>', methods=['POST'])
def sobrantes_salidas_delete(rid):
    conn = get_db()
    row = conn.execute('SELECT * FROM Salidas_Almacen_Sobrantes WHERE id=?', (rid,)).fetchone()
    if row:
        existe = conn.execute('SELECT id FROM Almacen_Sobrantes WHERE cod_prod=?', (row['cod_prod'],)).fetchone()
        if existe:
            conn.execute('UPDATE Almacen_Sobrantes SET salidas=salidas-?,stock=stock+? WHERE cod_prod=?',
                (row['cantidad'], row['cantidad'], row['cod_prod']))
        # Si la pieza ya no existe (stock se agotó y se borró), no se puede revivir
        # automáticamente sin perder trazabilidad.
    conn.execute('DELETE FROM Salidas_Almacen_Sobrantes WHERE id=?', (rid,))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/sobrantes/salidas/export')
def sobrantes_salidas_export():
    conn = get_db()
    rows = conn.execute('''SELECT id,cod_prod,fecha,insumo,cantidad,unidad,
        proyecto_origen,proyecto_destino FROM Salidas_Almacen_Sobrantes ORDER BY id DESC''').fetchall()
    conn.close()
    return export_excel(rows, ['ID','Cód.','Fecha','Insumo','Cantidad','Unidad',
        'Proyecto Origen','Proyecto Destino'], 'Salidas_Almacen_Sobrantes')

# ── ALMACÉN ──
@app.route('/almacen')
def almacen():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Almacen ORDER BY cod_prod').fetchall()
    proyectos = conn.execute('SELECT id, nombre FROM Proyectos ORDER BY nombre').fetchall()
    productos = conn.execute('''SELECT cod_prod,insumo,unidad,stock FROM Almacen
        WHERE cod_prod IS NOT NULL AND cod_prod!='' ORDER BY insumo''').fetchall()
    conn.close()
    return render_template('almacen.html', rows=rows, proyectos=proyectos,
        productos=productos, today=date.today().isoformat())

@app.route('/almacen/salida', methods=['POST'])
def almacen_salida():
    d = request.form; cant=float(d.get('cantidad') or 0); conn=get_db()
    conn.execute('''INSERT INTO Salidas_Almacen(cod_prod,fecha,insumo,cantidad,unidad,
        proyecto) VALUES(?,?,?,?,?,?)''',
        (d['cod_prod'],d['fecha'],d['insumo'],cant,d['unidad'],d['proyecto']))
    if conn.execute('SELECT id FROM Almacen WHERE cod_prod=?',(d['cod_prod'],)).fetchone():
        conn.execute('UPDATE Almacen SET salidas=salidas+?,stock=stock-? WHERE cod_prod=?',(cant,cant,d['cod_prod']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/almacen/excedente', methods=['POST'])
def almacen_excedente():
    """Devuelve al stock del Almacén normal material que se sacó pero no se usó
    por completo (ej: se sacaron 5 electrodos, se usaron 3, los 2 sobrantes
    vuelven a sumarse al stock). No tiene relación con el Almacén de Sobrantes
    (que es para retazos/cortes de planchas, perfiles, tubos, etc.).

    Se registra como una salida en NEGATIVO en Salidas_Almacen, ligada al
    Proyecto Origen indicado. Esto deja rastro visible en la pantalla de
    Salidas de Almacén y, automáticamente, reduce el "Material Gastado" de
    ese proyecto en la pantalla Proyectos (que calcula el gasto sumando
    cantidad × precio_unitario_promedio de Compras sobre todas las salidas
    de ese proyecto)."""
    d = request.form; cant = float(d.get('cantidad') or 0); conn = get_db()
    cod_prod = (d.get('cod_prod') or '').strip()
    insumo = (d.get('insumo') or '').strip()
    unidad = (d.get('unidad') or '').strip()
    proyecto_origen = (d.get('proyecto_origen') or '').strip()
    if not cod_prod or not insumo:
        conn.close(); return jsonify({'ok': False, 'msg': 'Selecciona un producto válido del almacén'})
    if not proyecto_origen:
        conn.close(); return jsonify({'ok': False, 'msg': 'Selecciona el proyecto origen del excedente'})
    if cant <= 0:
        conn.close(); return jsonify({'ok': False, 'msg': 'Ingresa una cantidad mayor a 0'})

    fila = conn.execute('SELECT id FROM Almacen WHERE cod_prod=?', (cod_prod,)).fetchone()
    if not fila:
        conn.close(); return jsonify({'ok': False, 'msg': 'Producto no encontrado en el almacén'})

    fecha = d.get('fecha') or date.today().isoformat()
    conn.execute('''INSERT INTO Salidas_Almacen(cod_prod,fecha,insumo,cantidad,unidad,proyecto)
        VALUES(?,?,?,?,?,?)''', (cod_prod, fecha, insumo, -cant, unidad, proyecto_origen))
    conn.execute('UPDATE Almacen SET entradas=entradas+?,stock=stock+? WHERE cod_prod=?', (cant, cant, cod_prod))
    conn.commit(); conn.close()
    return jsonify({'ok': True})

@app.route('/almacen/salida/edit/<int:rid>', methods=['POST'])
def almacen_salida_edit(rid):
    d = request.form; conn=get_db()
    old = conn.execute('SELECT * FROM Salidas_Almacen WHERE id=?',(rid,)).fetchone()
    nueva_cant = float(d.get('cantidad') or 0)
    diff = nueva_cant - float(old['cantidad'])
    # El producto (cod_prod/insumo/unidad) de una salida ya registrada NO se reasigna:
    # solo se ajusta la cantidad/proyecto. El stock se corrige sobre el producto
    # ORIGINAL de la salida (old['cod_prod']), nunca sobre un valor que pudiera
    # haber llegado distinto desde el formulario.
    conn.execute('''UPDATE Salidas_Almacen SET fecha=?,cantidad=?,
        proyecto=?,codigo_proyecto=? WHERE id=?''',
        (d['fecha'],nueva_cant,d['proyecto'],d.get('codigo_proyecto') or '',rid))
    conn.execute('UPDATE Almacen SET salidas=salidas+?,stock=stock-? WHERE cod_prod=?',
        (diff,diff,old['cod_prod']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/almacen/salida/delete/<int:rid>', methods=['POST'])
def almacen_salida_delete(rid):
    conn = get_db()
    row = conn.execute('SELECT * FROM Salidas_Almacen WHERE id=?',(rid,)).fetchone()
    conn.execute('UPDATE Almacen SET salidas=salidas-?,stock=stock+? WHERE cod_prod=?',
        (row['cantidad'],row['cantidad'],row['cod_prod']))
    conn.execute('DELETE FROM Salidas_Almacen WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/almacen/salidas')
def almacen_salidas():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Salidas_Almacen ORDER BY id DESC').fetchall()
    proyectos = conn.execute('SELECT id, nombre FROM Proyectos ORDER BY nombre').fetchall()
    conn.close()
    return render_template('almacen_salidas.html', rows=rows, proyectos=proyectos)

@app.route('/almacen/entradas')
def almacen_entradas():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Compras ORDER BY id DESC').fetchall(); conn.close()
    return render_template('almacen_entradas.html', rows=rows)

@app.route('/almacen/edit/<int:rid>', methods=['POST'])
def almacen_edit(rid):
    d = request.form; conn=get_db()
    conn.execute('UPDATE Almacen SET cod_prod=?,insumo=?,unidad=?,descripcion_cliente=? WHERE id=?',
        (d['cod_prod'],d['insumo'],d['unidad'],d['descripcion_cliente'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/almacen/delete/<int:rid>', methods=['POST'])
def almacen_delete(rid):
    conn = get_db(); conn.execute('DELETE FROM Almacen WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/almacen/export')
def almacen_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Almacen').fetchall(); conn.close()
    return export_excel(rows,['ID','Cód.','Insumo','Entradas','Salidas','Stock','Unidad','Descripción'],'Almacen')

@app.route('/almacen/salidas/export')
def almacen_salidas_export():
    conn = get_db()
    rows = conn.execute('''SELECT id,cod_prod,fecha,insumo,cantidad,unidad,proyecto,
        costo FROM Salidas_Almacen''').fetchall(); conn.close()
    return export_excel(rows,['ID','Cód.','Fecha','Insumo','Cantidad','Unidad','Proyecto','Costo'],'Salidas_Almacen')

# ── COMPRAS ──
@app.route('/compras')
def compras():
    conn = get_db()
    proveedores = conn.execute('SELECT id,nombre FROM Proveedores ORDER BY nombre').fetchall()
    rows = conn.execute('''SELECT c.*,p.nombre as proveedor_nombre FROM Compras c
        LEFT JOIN Proveedores p ON c.proveedor_id=p.id ORDER BY c.id DESC''').fetchall()
    tipos_pago = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Pago ORDER BY valor').fetchall()]
    tipos_documento = [r['valor'] for r in conn.execute('SELECT valor FROM Tipos_Documento ORDER BY valor').fetchall()]
    insumos = conn.execute('SELECT cod_prod,nombre,unidad FROM Insumos ORDER BY nombre').fetchall()
    conn.close()
    return render_template('compras.html', proveedores=proveedores, rows=rows, today=date.today().isoformat(),
        tipos_pago=tipos_pago, tipos_documento=tipos_documento, insumos=insumos)

@app.route('/compras/add', methods=['POST'])
def compras_add():
    d = request.form; cant=float(d.get('cantidad') or 0); pu=float(d.get('precio_unitario') or 0)
    total=cant*pu; conn=get_db()
    cod_prod = (d.get('cod_prod') or '').strip()
    insumo   = (d.get('insumo') or '').strip()
    if not cod_prod or not insumo:
        conn.close(); return jsonify({'ok': False, 'msg': 'Selecciona un insumo válido del catálogo'})
    conn.execute('''INSERT INTO Compras(fecha,proveedor_id,cod_prod,insumo,cantidad,unidad,
        precio_unitario,precio_total,tipo_documento,nro_documento,tipo_pago,detalle) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
        (d['fecha'],d.get('proveedor_id') or None,cod_prod,insumo,cant,d['unidad'],
         pu,total,d['tipo_documento'],d['nro_documento'],d['tipo_pago'],d['detalle']))
    if conn.execute('SELECT id FROM Almacen WHERE cod_prod=?',(cod_prod,)).fetchone():
        conn.execute('UPDATE Almacen SET entradas=entradas+?,stock=stock+? WHERE cod_prod=?',(cant,cant,cod_prod))
    else:
        conn.execute('INSERT INTO Almacen(cod_prod,insumo,entradas,salidas,stock,unidad) VALUES(?,?,?,0,?,?)',
            (cod_prod,insumo,cant,cant,d['unidad']))
    conn.commit(); conn.close(); return jsonify({'ok':True,'total':total})

@app.route('/compras/edit/<int:rid>', methods=['POST'])
def compras_edit(rid):
    d = request.form; cant=float(d.get('cantidad') or 0); pu=float(d.get('precio_unitario') or 0); conn=get_db()
    old = conn.execute('SELECT * FROM Compras WHERE id=?',(rid,)).fetchone()
    diff = cant - float(old['cantidad'])

    # El producto (cod_prod/insumo/unidad) de una compra ya registrada viene del
    # catálogo de Insumos y no se reasigna aquí — solo se ajustan fecha, proveedor,
    # cantidad, precio, documento y detalle. El stock se corrige sobre el producto
    # ORIGINAL de la compra (old['cod_prod']).
    conn.execute('''UPDATE Compras SET fecha=?,proveedor_id=?,cantidad=?,
        unidad=?,precio_unitario=?,precio_total=?,tipo_documento=?,nro_documento=?,tipo_pago=?,detalle=? WHERE id=?''',
        (d['fecha'],d.get('proveedor_id') or None,cant,d['unidad'],
         pu,cant*pu,d['tipo_documento'],d['nro_documento'],d['tipo_pago'],d['detalle'],rid))

    if diff != 0 and old['cod_prod']:
        conn.execute('UPDATE Almacen SET entradas=entradas+?,stock=stock+? WHERE cod_prod=?',
            (diff,diff,old['cod_prod']))

    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/compras/delete/<int:rid>', methods=['POST'])
def compras_delete(rid):
    conn = get_db(); row=conn.execute('SELECT * FROM Compras WHERE id=?',(rid,)).fetchone()
    conn.execute('UPDATE Almacen SET entradas=entradas-?,stock=stock-? WHERE cod_prod=?',
        (row['cantidad'],row['cantidad'],row['cod_prod']))
    conn.execute('DELETE FROM Compras WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/compras/export')
def compras_export():
    conn = get_db()
    rows = conn.execute('''SELECT c.fecha,p.nombre,c.cod_prod,c.insumo,c.cantidad,c.unidad,
        c.precio_unitario,c.precio_total,c.tipo_documento,c.nro_documento,c.tipo_pago,c.detalle
        FROM Compras c LEFT JOIN Proveedores p ON c.proveedor_id=p.id''').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Proveedor','Cód.','Insumo','Cant.','Unidad',
        'P.Unit.','P.Total','Tipo Doc.','N° Doc.','Pago','Detalle'],'Compras')

# ── HERRAMIENTAS ──
@app.route('/herramientas')
def herramientas():
    conn = get_db()
    lista=conn.execute('SELECT * FROM Herramientas ORDER BY id DESC').fetchall()
    trabajadores=conn.execute('SELECT id,nombre FROM Trabajadores ORDER BY nombre').fetchall()
    ts=date.today().isoformat()
    asignaciones=conn.execute('''SELECT a.*,h.nombre as herramienta_nombre,t.nombre as responsable_nombre
        FROM Asignacion_Herramientas a LEFT JOIN Herramientas h ON a.herramienta_id=h.id
        LEFT JOIN Trabajadores t ON a.responsable_id=t.id WHERE a.devuelto=0 AND a.fecha=?''',(ts,)).fetchall()
    conn.close()
    return render_template('herramientas.html',lista=lista,trabajadores=trabajadores,asignaciones=asignaciones,today=ts)

@app.route('/herramientas/add', methods=['POST'])
def herramientas_add():
    d=request.form; conn=get_db()
    # Verificar nombre duplicado
    if conn.execute('SELECT id FROM Herramientas WHERE nombre=?',(d['nombre'],)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe una herramienta con el nombre "{d["nombre"]}"'})
    conn.execute('INSERT INTO Herramientas(fecha_ingreso,nombre,descripcion,marca,costo,observaciones) VALUES(?,?,?,?,?,?)',
        (d['fecha_ingreso'],d['nombre'],d['descripcion'],d['marca'],float(d.get('costo') or 0),d['observaciones']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/herramientas/edit/<int:rid>', methods=['POST'])
def herramientas_edit(rid):
    d=request.form; conn=get_db()
    conn.execute('UPDATE Herramientas SET fecha_ingreso=?,nombre=?,descripcion=?,marca=?,costo=?,observaciones=? WHERE id=?',
        (d['fecha_ingreso'],d['nombre'],d['descripcion'],d['marca'],float(d.get('costo') or 0),d['observaciones'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/herramientas/delete/<int:rid>', methods=['POST'])
def herramientas_delete(rid):
    conn=get_db(); conn.execute('DELETE FROM Herramientas WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/herramientas/asignar', methods=['POST'])
def herramientas_asignar():
    d=request.form; conn=get_db()
    conn.execute('''INSERT INTO Asignacion_Herramientas(fecha,herramienta_id,descripcion,
        cantidad,responsable_id,observaciones) VALUES(?,?,?,?,?,?)''',
        (d['fecha'],int(d['herramienta_id']),d['descripcion'],int(d['cantidad']),int(d['responsable_id']),d['observaciones']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/herramientas/devolver/<int:rid>', methods=['POST'])
def herramientas_devolver(rid):
    conn=get_db()
    conn.execute('UPDATE Asignacion_Herramientas SET devuelto=1,fecha_devolucion=? WHERE id=?',(date.today().isoformat(),rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/herramientas/historial')
def herramientas_historial():
    conn=get_db()
    rows=conn.execute('''SELECT a.*,h.nombre as herramienta_nombre,t.nombre as responsable_nombre
        FROM Asignacion_Herramientas a LEFT JOIN Herramientas h ON a.herramienta_id=h.id
        LEFT JOIN Trabajadores t ON a.responsable_id=t.id ORDER BY a.id DESC''').fetchall(); conn.close()
    return render_template('herramientas_historial.html',rows=rows)

@app.route('/herramientas/historial/export')
def herramientas_historial_export():
    conn=get_db()
    rows=conn.execute('''SELECT a.id,a.fecha,h.nombre,a.descripcion,a.cantidad,t.nombre,
        a.observaciones,CASE WHEN a.devuelto=1 THEN 'Sí' ELSE 'No' END,a.fecha_devolucion
        FROM Asignacion_Herramientas a LEFT JOIN Herramientas h ON a.herramienta_id=h.id
        LEFT JOIN Trabajadores t ON a.responsable_id=t.id''').fetchall(); conn.close()
    return export_excel(rows,['ID','Fecha','Herramienta','Descripción','Cant.','Responsable','Obs.','Devuelto','F.Dev.'],'Historial_Herramientas')

# ── CONSUMIBLES ──
@app.route('/consumibles')
def consumibles():
    conn=get_db()
    lista=conn.execute('SELECT * FROM Consumibles ORDER BY nombre').fetchall()
    trabajadores=conn.execute('SELECT id,nombre FROM Trabajadores ORDER BY nombre').fetchall()
    ts=date.today().isoformat()
    eh=conn.execute('''SELECT e.*,c.nombre as consumible_nombre,t.nombre as responsable_nombre
        FROM Entrega_Consumibles e LEFT JOIN Consumibles c ON e.consumible_id=c.id
        LEFT JOIN Trabajadores t ON e.responsable_id=t.id WHERE e.fecha=? ORDER BY e.id DESC''',(ts,)).fetchall()
    ps=conn.execute('SELECT codigo_proyecto,nombre_proyecto FROM Servicios ORDER BY nombre_proyecto').fetchall()
    pv=conn.execute('SELECT codigo_proyecto,nombre_proyecto FROM Ventas ORDER BY nombre_proyecto').fetchall()
    proyectos=list({p['codigo_proyecto']:p for p in list(ps)+list(pv)}.values())
    conn.close()
    return render_template('consumibles.html',lista=lista,trabajadores=trabajadores,
        entregas_hoy=eh,today=ts,proyectos=proyectos)

@app.route('/consumibles/add', methods=['POST'])
def consumibles_add():
    d=request.form; conn=get_db()
    if conn.execute('SELECT id FROM Consumibles WHERE nombre=?',(d['nombre'],)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe un consumible con el nombre "{d["nombre"]}"'})
    conn.execute('INSERT INTO Consumibles(nombre,descripcion,observaciones) VALUES(?,?,?)',
        (d['nombre'],d['descripcion'],d['observaciones']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/consumibles/edit/<int:rid>', methods=['POST'])
def consumibles_edit(rid):
    d=request.form; conn=get_db()
    conn.execute('UPDATE Consumibles SET nombre=?,descripcion=?,observaciones=? WHERE id=?',
        (d['nombre'],d['descripcion'],d['observaciones'],rid))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/consumibles/delete/<int:rid>', methods=['POST'])
def consumibles_delete(rid):
    conn=get_db(); conn.execute('DELETE FROM Consumibles WHERE id=?',(rid,))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/consumibles/entregar', methods=['POST'])
def consumibles_entregar():
    d=request.form; conn=get_db()
    conn.execute('''INSERT INTO Entrega_Consumibles(fecha,consumible_id,descripcion,
        cantidad,responsable_id,proyecto,observaciones) VALUES(?,?,?,?,?,?,?)''',
        (d['fecha'],int(d['consumible_id']),d['descripcion'],float(d.get('cantidad') or 0),
         int(d['responsable_id']),d['proyecto'],d['observaciones']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/consumibles/historial')
def consumibles_historial():
    conn=get_db()
    rows=conn.execute('''SELECT e.*,c.nombre as consumible_nombre,t.nombre as responsable_nombre
        FROM Entrega_Consumibles e LEFT JOIN Consumibles c ON e.consumible_id=c.id
        LEFT JOIN Trabajadores t ON e.responsable_id=t.id ORDER BY e.consumible_id,e.id DESC''').fetchall(); conn.close()
    return render_template('consumibles_historial.html',rows=rows)

@app.route('/consumibles/historial/export')
def consumibles_historial_export():
    conn=get_db()
    rows=conn.execute('''SELECT e.fecha,c.nombre,e.descripcion,e.cantidad,t.nombre,e.proyecto,e.observaciones
        FROM Entrega_Consumibles e LEFT JOIN Consumibles c ON e.consumible_id=c.id
        LEFT JOIN Trabajadores t ON e.responsable_id=t.id''').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Consumible','Descripción','Cant.','Responsable','Proyecto','Obs.'],'Historial_Consumibles')

# ── CONFIGURACIONES ──
@app.route('/configuraciones')
def configuraciones():
    conn=get_db()
    tipos_pago = conn.execute('SELECT * FROM Tipos_Pago ORDER BY valor').fetchall()
    tipos_documento = conn.execute('SELECT * FROM Tipos_Documento ORDER BY valor').fetchall()
    conn.close()
    return render_template('configuraciones.html',compras_prov=[],tipos_pago=tipos_pago,tipos_documento=tipos_documento)

@app.route('/configuraciones/backup')
def backup():
    fecha=datetime.now().strftime('%Y%m%d_%H%M%S'); buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as zf:
        zf.write(DB_PATH,f'chumacero_{fecha}.db')
    buf.seek(0)
    return send_file(buf,mimetype='application/zip',as_attachment=True,
        download_name=f'backup_chumacero_{fecha}.zip')

@app.route('/configuraciones/importar', methods=['POST'])
def importar_backup():
    f=request.files.get('archivo')
    if not f: return jsonify({'ok':False,'msg':'No se envió archivo'})
    try:
        buf=io.BytesIO(f.read())
        with zipfile.ZipFile(buf,'r') as zf:
            names=[n for n in zf.namelist() if n.endswith('.db')]
            if not names: return jsonify({'ok':False,'msg':'No se encontró .db en el ZIP'})
            with zf.open(names[0]) as db_file:
                with open(DB_PATH,'wb') as out: out.write(db_file.read())
        return jsonify({'ok':True,'msg':'Base de datos restaurada correctamente'})
    except Exception as e: return jsonify({'ok':False,'msg':str(e)})

@app.route('/configuraciones/eliminar', methods=['POST'])
def eliminar_tabla():
    tabla=request.form.get('tabla')
    permitidas=['Clientes','Proveedores','Trabajadores','Servicios','Ingresos_Servicios',
        'Ventas','Ingresos_Ventas','Almacen','Salidas_Almacen','Compras',
        'Herramientas','Asignacion_Herramientas','Consumibles','Entrega_Consumibles',
        'Insumos','Almacen_Sobrantes','Entradas_Almacen_Sobrantes','Salidas_Almacen_Sobrantes']
    if tabla in permitidas:
        conn=get_db(); conn.execute(f'DELETE FROM {tabla}')
        conn.commit(); conn.close(); return jsonify({'ok':True,'msg':f'Tabla {tabla} limpiada'})
    return jsonify({'ok':False,'msg':'Tabla no permitida'})

# ── EXCEL HELPER ───────────────────────────────────────────────────────────────
def export_excel(rows, headers, sheet_name):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title=sheet_name
    hf=PatternFill('solid',fgColor='1A3A5C')
    hfont=Font(bold=True,color='FFFFFF',name='Calibri',size=11)
    af=PatternFill('solid',fgColor='EBF3FB')
    thin=Border(left=Side(style='thin',color='C0C0C0'),right=Side(style='thin',color='C0C0C0'),
        top=Side(style='thin',color='C0C0C0'),bottom=Side(style='thin',color='C0C0C0'))
    for col,h in enumerate(headers,1):
        cell=ws.cell(row=1,column=col,value=h)
        cell.fill=hf; cell.font=hfont
        cell.alignment=Alignment(horizontal='center',vertical='center'); cell.border=thin
    ws.auto_filter.ref=f'A1:{ws.cell(1,len(headers)).column_letter}1'
    for ri,row in enumerate(rows,2):
        fill=af if ri%2==0 else PatternFill('solid',fgColor='FFFFFF')
        for ci,val in enumerate(row,1):
            cell=ws.cell(row=ri,column=ci,value=val)
            cell.fill=fill; cell.border=thin
            cell.font=Font(name='Calibri',size=10); cell.alignment=Alignment(vertical='center')
    for col in ws.columns:
        mx=max((len(str(c.value or '')) for c in col),default=10)
        ws.column_dimensions[col[0].column_letter].width=min(mx+4,45)
    ws.row_dimensions[1].height=22
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,download_name=f'{sheet_name}_{datetime.now().strftime("%Y%m%d")}.xlsx')

# ── ARRANQUE ───────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    init_db()
    iniciar_scheduler_backup()
    threading.Thread(target=lambda: serve(app,host='0.0.0.0',port=5000,threads=8), daemon=True).start()
    time.sleep(2)
    webbrowser.open("http://127.0.0.1:5000")
    crear_icono_bandeja()
