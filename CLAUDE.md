# CLAUDE.md

Documentación del sistema para quien (persona o IA) trabaje sobre este repositorio.

## Qué es esto

**Costa Taller Mecánico** (Taller Costa) — sistema interno de gestión para un taller de reparación de maquinaria pesada (grúas, autoelevadores, camiones, motores industriales). Cubre todo el ciclo: presupuesto → trabajo en taller → cobro, más cuentas corrientes, gastos e informes de resultados.

Es una app monolítica de un solo taller: no hay multi-tenant ni usuarios diferenciados, solo una contraseña de acceso compartida.

## Stack

- **Backend**: Python 3.13, Flask, Flask-SQLAlchemy
- **Base de datos**: el código soporta SQLite (`sqlite:///taller.db`, default) o Postgres vía la variable de entorno `DATABASE_URL` (tiene `psycopg2-binary` instalado). **En producción (Render) corre sobre Postgres real** — confirmado el 2026-09-14 por un error de `psycopg2` en un `ALTER TABLE`. Esa instancia Postgres **no aparece** en `list_postgres_instances` del workspace de Render conectado, así que probablemente sea un Postgres externo (Neon/Supabase/ElephantSQL/etc.) apuntado vía `DATABASE_URL` en las variables de entorno del servicio — no confirmado cuál. En local, por default, se usa SQLite (`instance/taller.db`).
- **PDF**: ReportLab, para el presupuesto imprimible (`/presupuestos/<id>/pdf`)
- **Excel**: openpyxl, exportación completa de todas las tablas para PowerBI (`/exportar-excel`)
- **Servidor**: Gunicorn (`gunicorn app:app`) en producción; servidor de desarrollo de Flask (`python app.py`, puerto 8080) en local. `gunicorn.conf.py` existe pero está vacío (sin configuración custom).
- **Frontend**: Jinja2 server-rendered, CSS propio en `templates/base.html` (sin framework JS ni bundler), íconos Tabler (`ti ti-*`)
- **Hosting**: Render, plan **free** — servicio web `TallerCosta` (`srv-d8vn66sm0tmc73d17m60`), autodeploy en cada push a `main`, URL pública `https://tallercosta.onrender.com`

### Cómo correr localmente

```
cd C:\TallerCosta
venv\Scripts\activate
python app.py
```

Sirve en `http://0.0.0.0:8080`. Contraseña de acceso (login único, hardcodeada): `costa2026` (variable `PASSWORD` en `app.py`).

### Otros archivos de arranque

- `Iniciar_Taller.vbs`: lanzador para Windows sin ventana de consola, pensado para los dueños del taller (Lucas y Leonardo Costa). Levanta `python app.py` oculto y abre el navegador en `http://192.168.0.12:8080` (IP fija de red local del taller).
- `config.yml` + `Localtonet Installer.exe`: rastros de un túnel Localtonet para exponer el puerto 8080 a internet como alternativa/backup al hosting en Render. No se confirmó que esté en uso activo.

## Estructura de rutas

Todas las rutas (salvo `/login` y estáticos) están protegidas por un interceptor `@app.before_request` que exige `session['autenticado']`. La sesión dura 1 hora (`permanent_session_lifetime`).

| Módulo | Rutas |
|---|---|
| **Auth** | `/login` (GET/POST), `/logout` |
| **Inicio** | `/` — dashboard con KPIs por rango de fechas + trabajos activos + presupuestos pendientes |
| **Clientes** | `/clientes` (listado + búsqueda), `/clientes/nuevo`, `/clientes/<id>/editar` |
| **Proveedores** | `/proveedores` (listado + búsqueda), `/proveedores/nuevo`, `/proveedores/<id>/editar` — calcado de Clientes |
| **Presupuestos** | `/presupuestos` (listado por tabs: activos/aceptados/rechazados), `/presupuestos/nuevo`, `/presupuestos/<id>/editar`, `/presupuestos/<id>/convertir` (pasa a Trabajo), `/presupuestos/<id>/estado` (POST, cambia estado), `/presupuestos/<id>/pdf` |
| **Trabajos** | `/trabajos` (listado + filtros), `/trabajos/<id>` (detalle), `/trabajos/<id>/estado` (POST), `/trabajos/<id>/gasto` (POST, alta), `/trabajos/gasto/<id>/eliminar` (POST), `/trabajos/<id>/cobro` (POST, alta), `/trabajos/cobro/<id>/eliminar` (POST), `/trabajos/<id>/notas` (POST) |
| **Cuenta corriente** | `/cuenta-corriente?cliente_id=` |
| **Historial** | `/historial` — búsqueda técnica global por cliente/máquina, incluye trabajos pasados y anulados |
| **Gastos generales** | `/gastos-generales` (tabla con filtros por categoría/estado/texto + alta vía panel lateral), `/gastos-generales/<id>/anular` (POST, soft-delete), `/gastos-generales/<id>/pagar` (POST, salda un pendiente) |
| **Resultados** | `/resultados?mes=&anio=` — estado de resultados financiero/económico |
| **Configuración** | `/configuracion` (listas desplegables), `/configuracion/agregar` (POST), `/configuracion/<id>/eliminar` (POST) |
| **Excel** | `/exportar-excel` — descarga un `.xlsx` con todas las tablas |

## Modelo de datos

Todos los modelos están en `app.py` (sin `models.py` separado).

- **Cliente**: `empresa*`, `contacto`, `telefono`, `email`, `direccion`, `cuit`, `notas`, `creado`. `1—N` con `Presupuesto` y `Trabajo` (backref `cliente`).
- **Proveedor**: mismos campos que Cliente pero con `nombre*` en vez de `empresa`. **Sin relación FK** con nada — ver "Proveedor como texto libre" más abajo.
- **Presupuesto**: `numero` (`PRES-0001...`, autogenerado), `cliente_id`, datos de la máquina (`tipo_equipo`, `identificador`, `marca`, `modelo`, `tipo_trabajo`), `observaciones`, `estado`, `total`, `creado`. `1—N` con `ItemPresupuesto` (cascade delete). `1—1` con `Trabajo` (backref `presupuesto`, `uselist=False`).
  - **Estados**: `borrador` → `enviado` → `aceptado` | `rechazado`. Un presupuesto `aceptado` se convierte en `Trabajo` vía `/presupuestos/<id>/convertir`.
- **ItemPresupuesto**: línea de detalle de un presupuesto (`descripcion`, `cantidad`, `precio_unitario`, `descuento`, `subtotal`).
- **Trabajo**: `numero` (`REP-0001...`, autogenerado), `cliente_id`, `presupuesto_id` (nullable — puede crearse sin presupuesto previo, aunque hoy el único flujo de alta es `convertir_presupuesto`), datos de la máquina, `observaciones`, `estado`, `presupuestado`, `fecha_ingreso`, `fecha_entrega`, `creado`. `1—N` con `GastoTrabajo` y `Cobro` (cascade delete). Propiedades calculadas: `total_gastos`, `total_cobrado`, `saldo`, `ganancia`.
  - **Estados**: `en_curso` → `finalizado` → `entregado`, o `anulado` en cualquier momento. **`anulado` debe excluirse de todos los totales financieros** (ver más abajo).
- **GastoTrabajo**: gasto imputado a un trabajo puntual (repuestos, insumos). `proveedor` es texto libre.
- **Cobro**: pago recibido por un trabajo.
- **GastoGeneral**: gasto del taller no ligado a un trabajo (alquiler, servicios, etc.). Tiene `anulado` (bool, soft-delete) — igual que `Trabajo.anulado`, **debe excluirse de los totales** cuando está anulado. También tiene `pagado` (bool, default `True`) y `fecha_pago` (nullable): al cargar un gasto se elige pagarlo en el momento o dejarlo como saldo pendiente; `/gastos-generales/<id>/pagar` salda un pendiente después. Nota: `pagado`/`fecha_pago` **no** se tienen en cuenta todavía en los cálculos de `/resultados` ni del KPI de Inicio (siguen sumando el gasto esté pagado o pendiente, mientras no esté anulado) — si se quiere que el estado de resultados sea "caja" (solo lo efectivamente pagado) en vez de "devengado", falta esa lógica.
- **OpcionLista**: tabla genérica `(tipo, valor, orden)` que alimenta los desplegables configurables desde `/configuracion`: `tipo_equipo`, `tipo_trabajo`, `categoria_gasto`, `forma_pago`. Sembrada por `seed_opciones()` al arrancar si está vacía.

### Por qué Cliente/Proveedor son tablas propias y no `OpcionLista`

`OpcionLista` alcanza para listas simples de un solo valor (categorías, formas de pago). Cliente y Proveedor necesitan más campos (contacto, teléfono, CUIT, notas) y, en el caso de Cliente, relaciones reales con Presupuesto/Trabajo — por eso tienen modelo y CRUD propio, calcado uno del otro.

### Proveedor como texto libre en los gastos

`GastoTrabajo.proveedor` y `GastoGeneral.proveedor` siguen siendo `db.Column(db.String)`, **no** `ForeignKey(Proveedor.id)`. La decisión fue que el formulario elija el nombre de un `<select>` poblado con `Proveedor.query`, pero sin migrar el esquema ni los datos históricos ya cargados como texto libre. Si en algún momento se necesita reportar gastos por proveedor de forma confiable, ahí sí conviene pasar a FK.

## Decisiones y cambios importantes (historial reciente)

1. Excel: se agregaron columnas "N° Presupuesto" y "Observaciones" a la hoja "Trabajos"; nueva hoja "Proveedores".
2. `editar_presupuesto`: bloquea la edición si el trabajo asociado ya está `finalizado` o `entregado` (redirige a `/presupuestos`).
3. Sesión con expiración de 1 hora (`permanent_session_lifetime` + `session.permanent = True` en el login).
4. Se agregó la ruta `/presupuestos/<id>/convertir` — **faltaba por completo** (el botón "Convertir a trabajo" del listado apuntaba a una ruta inexistente, 404). Crea el `Trabajo` a partir del `Presupuesto` aceptado.
5. **Bug corregido**: trabajos y gastos generales `anulado` no se excluían de los KPIs de Inicio, Resultados y Cuenta Corriente (inflaban ventas/saldos/gastos con datos cancelados). Corregido en las tres pantallas — si se agrega un nuevo cálculo de totales en el futuro, recordar excluir `estado == 'anulado'` / `anulado == True`.
6. Módulo completo de Proveedores (modelo, CRUD, desplegable en los dos formularios de carga de gastos, hoja en el Excel).
7. Se vació una vez la base de datos de producción (clientes/proveedores/presupuestos/trabajos/gastos/cobros) manteniendo las listas de Configuración, para arrancar con datos limpios. Se hizo con una ruta temporal `/admin/reset-datos` que se agregó, se usó una vez y se eliminó del código en el commit siguiente — **no debería volver a existir en el repo**; si hace falta repetir el proceso, replicar el mismo patrón (ruta temporal + confirmación escrita + eliminarla después) en vez de dejar un endpoint de borrado permanente.
8. Rediseño de Gastos Generales (tabla + filtros + alta por panel lateral) agregó columnas nuevas a `GastoGeneral` (`pagado`, `fecha_pago`). Esto **rompió producción** (`/gastos-generales` y `/` daban 500) porque `db.create_all()` no altera tablas existentes y la tabla `gasto_general` de producción ya existía sin esas columnas — confirmó en la práctica el riesgo del punto 3 de "Pendientes". Se arregló con el mismo patrón de ruta temporal (`/admin/migrar-schema`, `ALTER TABLE ... ADD COLUMN`, sin tocar datos), usada una vez y eliminada. Ojo: la sintaxis DDL tiene que ser compatible con Postgres, no solo SQLite — `DEFAULT 1` en una columna `BOOLEAN` falla en Postgres (`DatatypeMismatch`), hay que usar `DEFAULT TRUE`. **Lección para el futuro**: cualquier columna nueva en un modelo existente necesita este mismo tipo de migración manual antes (o en el mismo deploy) de que el código que la usa llegue a producción — nunca asumir que `db.create_all()` alcanza.

## Pendientes / riesgos conocidos

1. **🔴 CRÍTICO — sin sistema de migraciones (Alembic/Flask-Migrate)**: `db.create_all()` (al final de `app.py`) solo crea tablas que no existen, nunca altera columnas de tablas existentes. Confirmado en producción el 2026-09-14: agregar `pagado`/`fecha_pago` a `GastoGeneral` rompió `/gastos-generales` y `/` (500) hasta hacer un `ALTER TABLE` manual — ver punto 8 de "Decisiones". **Regla a partir de ahora**: cualquier columna/tabla nueva en un modelo existente necesita una migración manual (ruta temporal con `ALTER TABLE`, usada una vez y eliminada, como se hizo dos veces esta sesión) antes de que el código que la usa se despliegue — nunca asumir que alcanza con cambiar el modelo en Python. Instalar Flask-Migrate eliminaría este riesgo de raíz.
2. **Persistencia de datos en producción (Render, plan free)**: confirmado el 2026-09-14 que la base de producción es **Postgres real** (no SQLite) y que los datos persisten entre deploys — la tabla `gasto_general` sobrevivió varios pushes con sus filas intactas. El proveedor exacto de ese Postgres no se identificó (no aparece en `list_postgres_instances` del workspace de Render conectado). Sigue siendo buena idea confirmar dónde vive y que tenga backups, pero el riesgo de "cada push borra todo" que se sospechaba antes **no aplica**.
3. **`instance/taller.db` local desactualizado**: le falta la columna `gasto_general.anulado` (y ahora también `pagado`/`fecha_pago`) — rompe `/` si se corre la app en local contra ese archivo. Falta migrar o recrear la base local (borrarla y dejar que `db.create_all()` la regenere es la forma más simple, ya que es solo de prueba).
4. **`__pycache__/` versionado en git**: no hay `.gitignore` en el repo; el bytecode compilado (`app.cpython-313.pyc`) queda trackeado y genera diffs de ruido en cada commit.
5. **Login único sin usuarios**: una sola contraseña compartida (`costa2026`) para toda la app, sin distinguir quién hizo cada acción. Suficiente para el tamaño actual del equipo, pero sin auditoría.
6. **`app.secret_key` hardcodeado** en el código fuente (`'taller_costa_2025'`), no en variable de entorno. Cualquiera con acceso al repo puede firmar cookies de sesión válidas.
7. **Excel — asimetría menor**: la hoja "Gastos Generales" no exporta el campo `notas` (sí lo hace "Cobranzas" con un campo análogo en `Cobro`).
8. **Sin suite de tests en el repo**: las pruebas de las últimas features (rutas nuevas, exclusión de anulados, proveedores, reset de datos) se hicieron con scripts ad-hoc contra una base aislada durante el desarrollo, pero no quedaron como tests versionados en el proyecto.
