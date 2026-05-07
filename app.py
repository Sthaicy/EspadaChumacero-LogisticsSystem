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
    conn.commit(); conn.close()

TIPOS_SERVICIO = ['Corte CNC','Corte y Plegado','Tornería','Fresado','Soldadura','Taladrado']
TIPOS_PAGO     = ['Caja Chica','Unión','Ganadero','BCP','Otros']

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
    clientes  = conn.execute('SELECT id,nombre FROM Clientes ORDER BY nombre').fetchall()
    proyectos = conn.execute('''SELECT s.*,c.nombre as cliente_nombre
        FROM Servicios s LEFT JOIN Clientes c ON s.cliente_id=c.id
        WHERE s.saldo_cobrar > 0 ORDER BY s.id DESC''').fetchall()
    conn.close()
    return render_template('ordenes.html', clientes=clientes, proyectos=proyectos,
        today=date.today().isoformat(), tipos_servicio=TIPOS_SERVICIO, tipos_pago=TIPOS_PAGO)

@app.route('/ordenes/registrar', methods=['POST'])
def ordenes_registrar():
    d = request.form; precio = float(d.get('precio_total') or 0); conn = get_db()
    # Verificar duplicados
    cod = d.get('codigo_proyecto','').strip()
    nom = d.get('nombre_proyecto','').strip()
    if conn.execute('SELECT id FROM Servicios WHERE codigo_proyecto=?',(cod,)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe un proyecto con el código "{cod}"'})
    if conn.execute('SELECT id FROM Servicios WHERE nombre_proyecto=?',(nom,)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe un proyecto con el nombre "{nom}"'})
    conn.execute('''INSERT INTO Servicios(fecha,codigo_proyecto,nombre_proyecto,descripcion,
        cliente_id,tipo_servicio,precio_total,saldo_cobrar) VALUES(?,?,?,?,?,?,?,?)''',
        (d['fecha'],cod,nom,d['descripcion'],d.get('cliente_id') or None,
         d.get('tipo_servicio',''),precio,precio))
    conn.commit(); conn.close(); return jsonify({'ok':True})

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
    rows = conn.execute('''SELECT i.*,c.nombre as cliente_nombre
        FROM Ingresos_Servicios i LEFT JOIN Clientes c ON i.cliente_id=c.id
        ORDER BY i.id DESC''').fetchall(); conn.close()
    return render_template('ordenes_ingresos.html', rows=rows, tipos_pago=TIPOS_PAGO)

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
    ingresos = conn.execute('''SELECT i.*,c.nombre as cliente_nombre FROM Ingresos_Servicios i
        LEFT JOIN Clientes c ON i.cliente_id=c.id ORDER BY i.codigo_proyecto,i.id''').fetchall()
    salidas  = conn.execute('''SELECT sa.*,
        COALESCE((SELECT AVG(precio_unitario) FROM Compras WHERE cod_prod=sa.cod_prod),0) as costo_unitario
        FROM Salidas_Almacen sa ORDER BY sa.codigo_proyecto,sa.id''').fetchall()
    # Precio total de cada proyecto (para mostrar en cabecera)
    proyectos_info = {}
    for s in conn.execute('SELECT codigo_proyecto,nombre_proyecto,precio_total FROM Servicios').fetchall():
        proyectos_info[s['codigo_proyecto']] = dict(s)
    pi = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in ingresos]))
    ps = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in salidas]))
    conn.close()
    return render_template('ordenes_historial.html', ingresos=ingresos, salidas=salidas,
        proyectos=list(dict.fromkeys(pi+ps)), proyectos_info=proyectos_info, tipos_pago=TIPOS_PAGO)

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
    rows = conn.execute('''SELECT i.fecha,i.codigo_proyecto,i.nombre_proyecto,
        c.nombre,i.monto_facturado,i.nro_factura,i.monto,i.tipo_pago,i.observaciones
        FROM Ingresos_Servicios i LEFT JOIN Clientes c ON i.cliente_id=c.id
        ORDER BY i.codigo_proyecto,i.id''').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','Cliente','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Historial_Servicios')

@app.route('/ordenes/export')
def ordenes_export():
    conn = get_db()
    rows = conn.execute('''SELECT s.codigo_proyecto,s.nombre_proyecto,s.tipo_servicio,
        c.nombre,s.precio_total,s.saldo_cobrar FROM Servicios s
        LEFT JOIN Clientes c ON s.cliente_id=c.id WHERE s.saldo_cobrar>0''').fetchall(); conn.close()
    return export_excel(rows,['Cód.','Proyecto','Tipo Servicio','Cliente','Precio Total','Saldo'],'Servicios')

@app.route('/ordenes/ingresos/export')
def ordenes_ingresos_export():
    conn = get_db()
    rows = conn.execute('''SELECT i.fecha,i.codigo_proyecto,i.nombre_proyecto,c.nombre,
        i.monto_facturado,i.nro_factura,i.monto,i.tipo_pago,i.observaciones
        FROM Ingresos_Servicios i LEFT JOIN Clientes c ON i.cliente_id=c.id''').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','Cliente','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Ingresos_Servicios')

# ── VENTAS ──
@app.route('/ventas')
def ventas():
    conn = get_db()
    clientes  = conn.execute('SELECT id,nombre FROM Clientes ORDER BY nombre').fetchall()
    proyectos = conn.execute('''SELECT v.*,c.nombre as cliente_nombre FROM Ventas v
        LEFT JOIN Clientes c ON v.cliente_id=c.id WHERE v.saldo_cobrar>0 ORDER BY v.id DESC''').fetchall()
    conn.close()
    return render_template('ventas.html', clientes=clientes, proyectos=proyectos,
        today=date.today().isoformat(), tipos_pago=TIPOS_PAGO)

@app.route('/ventas/registrar', methods=['POST'])
def ventas_registrar():
    d = request.form; precio=float(d.get('precio_total') or 0); conn=get_db()
    cod = d.get('codigo_proyecto','').strip()
    nom = d.get('nombre_proyecto','').strip()
    if conn.execute('SELECT id FROM Ventas WHERE codigo_proyecto=?',(cod,)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe una venta con el código "{cod}"'})
    if conn.execute('SELECT id FROM Ventas WHERE nombre_proyecto=?',(nom,)).fetchone():
        conn.close(); return jsonify({'ok':False,'msg':f'Ya existe una venta con el nombre "{nom}"'})
    conn.execute('''INSERT INTO Ventas(fecha,codigo_proyecto,nombre_proyecto,descripcion,
        cliente_id,precio_total,saldo_cobrar) VALUES(?,?,?,?,?,?,?)''',
        (d['fecha'],cod,nom,d['descripcion'],d.get('cliente_id') or None,precio,precio))
    conn.commit(); conn.close(); return jsonify({'ok':True})

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
    rows = conn.execute('''SELECT i.*,c.nombre as cliente_nombre FROM Ingresos_Ventas i
        LEFT JOIN Clientes c ON i.cliente_id=c.id ORDER BY i.id DESC''').fetchall(); conn.close()
    return render_template('ventas_ingresos.html', rows=rows, tipos_pago=TIPOS_PAGO)

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
    ingresos = conn.execute('''SELECT i.*,c.nombre as cliente_nombre FROM Ingresos_Ventas i
        LEFT JOIN Clientes c ON i.cliente_id=c.id ORDER BY i.codigo_proyecto,i.id''').fetchall()
    salidas  = conn.execute('''SELECT sa.*,
        COALESCE((SELECT AVG(precio_unitario) FROM Compras WHERE cod_prod=sa.cod_prod),0) as costo_unitario
        FROM Salidas_Almacen sa ORDER BY sa.codigo_proyecto,sa.id''').fetchall()
    proyectos_info = {}
    for v in conn.execute('SELECT codigo_proyecto,nombre_proyecto,precio_total FROM Ventas').fetchall():
        proyectos_info[v['codigo_proyecto']] = dict(v)
    pi = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in ingresos]))
    ps = list(dict.fromkeys([r['codigo_proyecto'] or '(sin código)' for r in salidas]))
    conn.close()
    return render_template('ventas_historial.html', ingresos=ingresos, salidas=salidas,
        proyectos=list(dict.fromkeys(pi+ps)), proyectos_info=proyectos_info, tipos_pago=TIPOS_PAGO)

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
    rows = conn.execute('''SELECT i.fecha,i.codigo_proyecto,i.nombre_proyecto,
        c.nombre,i.monto_facturado,i.nro_factura,i.monto,i.tipo_pago,i.observaciones
        FROM Ingresos_Ventas i LEFT JOIN Clientes c ON i.cliente_id=c.id
        ORDER BY i.codigo_proyecto,i.id''').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','Cliente','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Historial_Ventas')

@app.route('/ventas/export')
def ventas_export():
    conn = get_db()
    rows = conn.execute('''SELECT v.codigo_proyecto,v.nombre_proyecto,c.nombre,
        v.precio_total,v.saldo_cobrar FROM Ventas v LEFT JOIN Clientes c ON v.cliente_id=c.id
        WHERE v.saldo_cobrar>0''').fetchall(); conn.close()
    return export_excel(rows,['Cód.','Proyecto','Cliente','Precio Total','Saldo'],'Ventas')

@app.route('/ventas/ingresos/export')
def ventas_ingresos_export():
    conn = get_db()
    rows = conn.execute('''SELECT i.fecha,i.codigo_proyecto,i.nombre_proyecto,c.nombre,
        i.monto_facturado,i.nro_factura,i.monto,i.tipo_pago,i.observaciones
        FROM Ingresos_Ventas i LEFT JOIN Clientes c ON i.cliente_id=c.id''').fetchall(); conn.close()
    return export_excel(rows,['Fecha','Cód.','Proyecto','Cliente','M.Fact.','N°Fact.','Monto','Pago','Obs.'],'Ingresos_Ventas')

# ── ALMACÉN ──
@app.route('/almacen')
def almacen():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Almacen ORDER BY cod_prod').fetchall()
    proyectos = conn.execute('''
        SELECT codigo_proyecto,nombre_proyecto FROM Servicios WHERE codigo_proyecto IS NOT NULL
        UNION
        SELECT codigo_proyecto,nombre_proyecto FROM Ventas WHERE codigo_proyecto IS NOT NULL
        ORDER BY nombre_proyecto''').fetchall()
    productos = conn.execute('''SELECT cod_prod,insumo,unidad,stock FROM Almacen
        WHERE cod_prod IS NOT NULL AND cod_prod!='' ORDER BY insumo''').fetchall()
    conn.close()
    return render_template('almacen.html', rows=rows, proyectos=proyectos,
        productos=productos, today=date.today().isoformat())

@app.route('/almacen/salida', methods=['POST'])
def almacen_salida():
    d = request.form; cant=float(d.get('cantidad') or 0); conn=get_db()
    conn.execute('''INSERT INTO Salidas_Almacen(cod_prod,fecha,insumo,cantidad,unidad,
        proyecto,codigo_proyecto,sobrantes) VALUES(?,?,?,?,?,?,?,?)''',
        (d['cod_prod'],d['fecha'],d['insumo'],cant,d['unidad'],
         d['proyecto'],d['codigo_proyecto'],float(d.get('sobrantes') or 0)))
    if conn.execute('SELECT id FROM Almacen WHERE cod_prod=?',(d['cod_prod'],)).fetchone():
        conn.execute('UPDATE Almacen SET salidas=salidas+?,stock=stock-? WHERE cod_prod=?',(cant,cant,d['cod_prod']))
    conn.commit(); conn.close(); return jsonify({'ok':True})

@app.route('/almacen/salida/edit/<int:rid>', methods=['POST'])
def almacen_salida_edit(rid):
    d = request.form; conn=get_db()
    old = conn.execute('SELECT * FROM Salidas_Almacen WHERE id=?',(rid,)).fetchone()
    diff = float(d.get('cantidad') or 0)-float(old['cantidad'])
    conn.execute('''UPDATE Salidas_Almacen SET cod_prod=?,fecha=?,insumo=?,cantidad=?,
        unidad=?,proyecto=?,codigo_proyecto=?,sobrantes=? WHERE id=?''',
        (d['cod_prod'],d['fecha'],d['insumo'],float(d.get('cantidad') or 0),
         d['unidad'],d['proyecto'],d['codigo_proyecto'],float(d.get('sobrantes') or 0),rid))
    conn.execute('UPDATE Almacen SET salidas=salidas+?,stock=stock-? WHERE cod_prod=?',(diff,diff,d['cod_prod']))
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
    rows = conn.execute('SELECT * FROM Salidas_Almacen ORDER BY id DESC').fetchall(); conn.close()
    return render_template('almacen_salidas.html', rows=rows)

@app.route('/almacen/entradas')
def almacen_entradas():
    conn = get_db()
    rows = conn.execute('''SELECT c.*,p.nombre as proveedor_nombre FROM Compras c
        LEFT JOIN Proveedores p ON c.proveedor_id=p.id ORDER BY c.id DESC''').fetchall(); conn.close()
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
    return export_excel(rows,['ID','Cód.','Insumo','Entradas','Salidas','Stock','Unidad','Descripción/Cliente'],'Almacen')

@app.route('/almacen/salidas/export')
def almacen_salidas_export():
    conn = get_db()
    rows = conn.execute('SELECT * FROM Salidas_Almacen').fetchall(); conn.close()
    return export_excel(rows,['ID','Cód.','Fecha','Insumo','Cantidad','Unidad','Proyecto','Cód.Proy.','Sobrantes','Costo'],'Salidas_Almacen')

# ── COMPRAS ──
@app.route('/compras')
def compras():
    conn = get_db()
    proveedores = conn.execute('SELECT id,nombre FROM Proveedores ORDER BY nombre').fetchall()
    rows = conn.execute('''SELECT c.*,p.nombre as proveedor_nombre FROM Compras c
        LEFT JOIN Proveedores p ON c.proveedor_id=p.id ORDER BY c.id DESC''').fetchall(); conn.close()
    return render_template('compras.html', proveedores=proveedores, rows=rows, today=date.today().isoformat())

@app.route('/compras/add', methods=['POST'])
def compras_add():
    d = request.form; cant=float(d.get('cantidad') or 0); pu=float(d.get('precio_unitario') or 0)
    total=cant*pu; conn=get_db()
    conn.execute('''INSERT INTO Compras(fecha,proveedor_id,cod_prod,insumo,cantidad,unidad,
        precio_unitario,precio_total,tipo_documento,nro_documento,tipo_pago,detalle) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
        (d['fecha'],d.get('proveedor_id') or None,d['cod_prod'],d['insumo'],cant,d['unidad'],
         pu,total,d['tipo_documento'],d['nro_documento'],d['tipo_pago'],d['detalle']))
    if conn.execute('SELECT id FROM Almacen WHERE cod_prod=?',(d['cod_prod'],)).fetchone():
        conn.execute('UPDATE Almacen SET entradas=entradas+?,stock=stock+? WHERE cod_prod=?',(cant,cant,d['cod_prod']))
    else:
        conn.execute('INSERT INTO Almacen(cod_prod,insumo,entradas,salidas,stock,unidad) VALUES(?,?,?,0,?,?)',
            (d['cod_prod'],d['insumo'],cant,cant,d['unidad']))
    conn.commit(); conn.close(); return jsonify({'ok':True,'total':total})

@app.route('/compras/edit/<int:rid>', methods=['POST'])
def compras_edit(rid):
    d = request.form; cant=float(d.get('cantidad') or 0); pu=float(d.get('precio_unitario') or 0); conn=get_db()
    old = conn.execute('SELECT * FROM Compras WHERE id=?',(rid,)).fetchone()
    old_cod  = old['cod_prod'] or ''
    new_cod  = (d.get('cod_prod') or '').strip()
    diff = cant - float(old['cantidad'])

    conn.execute('''UPDATE Compras SET fecha=?,proveedor_id=?,cod_prod=?,insumo=?,cantidad=?,
        unidad=?,precio_unitario=?,precio_total=?,tipo_documento=?,nro_documento=?,tipo_pago=?,detalle=? WHERE id=?''',
        (d['fecha'],d.get('proveedor_id') or None,new_cod,d['insumo'],cant,d['unidad'],
         pu,cant*pu,d['tipo_documento'],d['nro_documento'],d['tipo_pago'],d['detalle'],rid))

    if new_cod:
        existing = conn.execute('SELECT id FROM Almacen WHERE cod_prod=?',(new_cod,)).fetchone()
        if existing:
            # Ya existe en almacén: ajustar diferencia de cantidad
            conn.execute('UPDATE Almacen SET entradas=entradas+?,stock=stock+? WHERE cod_prod=?',(diff,diff,new_cod))
        else:
            # cod_prod nuevo (antes estaba vacío): crear fila en Almacén
            conn.execute('INSERT INTO Almacen(cod_prod,insumo,entradas,salidas,stock,unidad) VALUES(?,?,?,0,?,?)',
                (new_cod,d['insumo'],cant,cant,d['unidad']))
        # Si cambió el código, revertir el stock del código viejo
        if old_cod and old_cod != new_cod:
            conn.execute('UPDATE Almacen SET entradas=entradas-?,stock=stock-? WHERE cod_prod=?',
                (float(old['cantidad']),float(old['cantidad']),old_cod))

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
    cp=conn.execute('''SELECT p.nombre,COUNT(*) as cant,SUM(c.precio_total) as total
        FROM Compras c LEFT JOIN Proveedores p ON c.proveedor_id=p.id
        GROUP BY c.proveedor_id ORDER BY total DESC''').fetchall(); conn.close()
    return render_template('configuraciones.html',compras_prov=cp)

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
        'Herramientas','Asignacion_Herramientas','Consumibles','Entrega_Consumibles']
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
    threading.Thread(target=lambda: serve(app,host='0.0.0.0',port=5000,threads=8), daemon=True).start()
    time.sleep(2)
    webbrowser.open("http://127.0.0.1:5000")
    crear_icono_bandeja()
