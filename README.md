# EspadaChumacero LogisticsSystem

Sistema web de control de inventarios y gestión logística desarrollado para EspadaChumacero S.R.L., empresa de estructuras metálicas y maquinaria industrial (Cochabamba, Bolivia). Centraliza compras, stock de almacén, retazos reutilizables, consumo de material por proyecto, herramientas y personal, y se distribuye como un ejecutable de Windows que levanta su propio servidor local.

## Resumen

El sistema reemplaza el control manual en hojas de cálculo por una aplicación con base de datos relacional, trazabilidad de movimientos y cálculo automático del costo de material por proyecto. Está pensado para operar en una red local: un equipo ejecuta la aplicación y el resto accede desde el navegador.

## Funcionalidades

**Almacén**
- Stock actual calculado a partir de compras (entradas) y salidas por proyecto.
- Registro de salidas con búsqueda de producto y proyecto, y validación de stock disponible.
- Devolución de material no utilizado ("Excedente"): suma el stock de vuelta y registra una salida negativa ligada al proyecto de origen, de modo que el gasto del proyecto se reduce automáticamente.
- Historial de entradas y salidas con edición, eliminación y exportación a Excel.

**Almacén de Sobrantes**
- Gestión de retazos de planchas, perfiles y tubos, donde cada pieza es única (tamaño, espesor y proyecto de origen propios).
- Código de producto generado automáticamente a partir del código del insumo, con un correlativo persistente que nunca se reutiliza, incluso si se eliminan registros.
- Salidas por pieza hacia otros proyectos; la pieza se retira del inventario cuando su stock llega a cero.
- Historial de entradas y salidas con recálculo coherente del stock al editar o eliminar movimientos.

**Catálogos y compras**
- Catálogo de insumos con código fijo, unidad, espesor opcional y detalle.
- Compras con proveedor, tipo de documento y tipo de pago; los dos últimos son configurables desde la aplicación.
- Autocompletado del código y la unidad a partir del insumo elegido.

**Proyectos**
- Registro de proyectos únicos por nombre, con validación de duplicados en tiempo real.
- Material gastado por proyecto, calculado como cantidad por precio unitario promedio de compra, incluyendo las devoluciones.

**Recursos y registros**
- Asignación y devolución de herramientas por trabajador.
- Entrega de consumibles por proyecto.
- Registro de trabajadores y proveedores.

**Administración**
- Copia de seguridad manual en ZIP e importación para restaurar la base de datos.
- Copia de seguridad automática diaria a las 12:00, en una carpeta separada de la base de datos, con rotación de las últimas 30 copias y registro de actividad.
- Icono en la bandeja del sistema con acceso rápido al navegador, a la dirección IP de red y a la copia de seguridad manual.
- Tablas con filtros por columna tipo Excel, búsqueda global y exportación a Excel en cada pantalla.

## Decisiones de diseño relevantes

- **Integridad de datos al editar.** El ajuste de stock tras editar una compra o una salida siempre se aplica sobre el producto original del registro, tomado de la base de datos y no del formulario, para evitar inconsistencias entre movimientos y existencias.
- **Campos ligados a catálogos.** Los campos que provienen de otra tabla (insumo, proyecto, proveedor, tipos de pago y documento) solo se pueden seleccionar desde su catálogo, tanto al registrar como al editar.
- **Correlativos persistentes.** Los códigos de sobrantes y de proyectos se generan con contadores almacenados en tablas propias, no con el máximo de las filas existentes, por lo que no se repiten tras eliminaciones.
- **Búsqueda sin dependencias pesadas.** Los selectores con autocompletado usan datos embebidos desde Jinja2 y JavaScript sin bibliotecas externas, lo que resultó más estable que un componente de terceros.
- **Operación sin conexión a internet.** Toda la aplicación corre localmente sobre SQLite.

## Tecnologías

| Capa | Herramientas |
|---|---|
| Backend | Python, Flask, SQLite |
| Servidor | Waitress |
| Frontend | Jinja2, HTML, CSS, JavaScript |
| Exportación | openpyxl |
| Escritorio | pystray, Pillow, Tkinter |
| Distribución | PyInstaller |

## Estructura del proyecto

```
LogisticsSystem/
├── app.py                # Aplicación Flask, esquema de base de datos y rutas
├── app.spec              # Configuración de PyInstaller
├── logo.ico              # Icono del ejecutable
├── INICIAR_SERVIDOR.bat  # Arranque en Windows
├── iniciar_servidor.sh   # Arranque en Linux/macOS
├── static/
│   ├── css/style.css
│   ├── js/main.js
│   └── img/
└── templates/            # Plantillas Jinja2 de cada pantalla
```

La base de datos (`chumacero.db`) y la carpeta `backups/` se crean automáticamente en la primera ejecución.

## Instalación y ejecución

Requisitos: Python 3.10 o superior.

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows (PowerShell)
pip install flask waitress pystray pillow openpyxl
python app.py
```

La aplicación abre el navegador en `http://127.0.0.1:5000`. Desde otros equipos de la misma red local se accede con `http://IP-DEL-SERVIDOR:5000`.

## Generar el ejecutable de Windows

```bash
pip install pyinstaller
pyinstaller app.spec
```

El resultado queda en `dist/EspadaChumacero.exe`.


## Nota

El repositorio no incluye datos reales de la empresa; la base de datos se genera vacía al iniciar la aplicación.
