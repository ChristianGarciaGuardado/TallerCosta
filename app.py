from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
import os

app = Flask(__name__)

DATABASE_URL = os.environ.get('DATABASE_URL', 'sqlite:///taller.db')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'taller_costa_2025'

db = SQLAlchemy(app)

# ═══════════════════════════════════════════════════════
# MODELOS
# ═══════════════════════════════════════════════════════

class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    empresa = db.Column(db.String(100), nullable=False)
    contacto = db.Column(db.String(100))
    telefono = db.Column(db.String(30))
    email = db.Column(db.String(100))
    direccion = db.Column(db.String(200))
    cuit = db.Column(db.String(20))
    notas = db.Column(db.Text)
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    presupuestos = db.relationship('Presupuesto', backref='cliente', lazy=True)
    trabajos = db.relationship('Trabajo', backref='cliente', lazy=True)

class Presupuesto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(20), unique=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey('cliente.id'), nullable=False)
    tipo_equipo = db.Column(db.String(50))
    identificador = db.Column(db.String(50))
    marca = db.Column(db.String(50))
    modelo = db.Column(db.String(50))
    tipo_trabajo = db.Column(db.String(50))
    observaciones = db.Column(db.Text)
    estado = db.Column(db.String(20), default='borrador')  # borrador, enviado, aceptado, rechazado
    total = db.Column(db.Float, default=0)
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    items = db.relationship('ItemPresupuesto', backref='presupuesto', lazy=True, cascade='all, delete-orphan')
    trabajo = db.relationship('Trabajo', backref='presupuesto', lazy=True, uselist=False)

class ItemPresupuesto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    presupuesto_id = db.Column(db.Integer, db.ForeignKey('presupuesto.id'), nullable=False)
    descripcion = db.Column(db.String(200), nullable=False)
    cantidad = db.Column(db.Float, default=1)
    precio_unitario = db.Column(db.Float, default=0)
    subtotal = db.Column(db.Float, default=0)

class Trabajo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(20), unique=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey('cliente.id'), nullable=False)
    presupuesto_id = db.Column(db.Integer, db.ForeignKey('presupuesto.id'), nullable=True)
    tipo_equipo = db.Column(db.String(50))
    identificador = db.Column(db.String(50))
    marca = db.Column(db.String(50))
    modelo = db.Column(db.String(50))
    tipo_trabajo = db.Column(db.String(50))
    observaciones = db.Column(db.Text)
    estado = db.Column(db.String(20), default='en_curso')  # en_curso, finalizado, entregado
    presupuestado = db.Column(db.Float, default=0)
    fecha_ingreso = db.Column(db.Date, default=date.today)
    fecha_entrega = db.Column(db.Date, nullable=True)
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    gastos = db.relationship('GastoTrabajo', backref='trabajo', lazy=True, cascade='all, delete-orphan')
    cobros = db.relationship('Cobro', backref='trabajo', lazy=True, cascade='all, delete-orphan')

    @property
    def total_gastos(self):
        return sum(g.monto for g in self.gastos)

    @property
    def total_cobrado(self):
        return sum(c.monto for c in self.cobros)

    @property
    def saldo(self):
        return self.presupuestado - self.total_cobrado

    @property
    def ganancia(self):
        return self.total_cobrado - self.total_gastos

class GastoTrabajo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    trabajo_id = db.Column(db.Integer, db.ForeignKey('trabajo.id'), nullable=False)
    fecha = db.Column(db.Date, default=date.today)
    categoria = db.Column(db.String(50))
    proveedor = db.Column(db.String(100))
    descripcion = db.Column(db.String(200))
    monto = db.Column(db.Float, default=0)
    forma_pago = db.Column(db.String(30))

class Cobro(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    trabajo_id = db.Column(db.Integer, db.ForeignKey('trabajo.id'), nullable=False)
    fecha = db.Column(db.Date, default=date.today)
    monto = db.Column(db.Float, default=0)
    forma_pago = db.Column(db.String(30))
    comprobante = db.Column(db.String(50))
    notas = db.Column(db.Text)

class GastoGeneral(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, default=date.today)
    categoria = db.Column(db.String(50))
    proveedor = db.Column(db.String(100))
    monto = db.Column(db.Float, default=0)
    forma_pago = db.Column(db.String(30))
    comprobante = db.Column(db.String(50))
    notas = db.Column(db.Text)

# ═══════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════

def gen_numero_presupuesto():
    ultimo = Presupuesto.query.order_by(Presupuesto.id.desc()).first()
    n = (ultimo.id + 1) if ultimo else 1
    return f"PRES-{n:04d}"

def gen_numero_trabajo():
    ultimo = Trabajo.query.order_by(Trabajo.id.desc()).first()
    n = (ultimo.id + 1) if ultimo else 1
    return f"REP-{n:04d}"

def fmt_moneda(valor):
    if valor is None:
        return "$0"
    return f"${valor:,.0f}".replace(",", ".")

app.jinja_env.filters['moneda'] = fmt_moneda

# ═══════════════════════════════════════════════════════
# RUTAS — INICIO
# ═══════════════════════════════════════════════════════

@app.route('/')
def inicio():
    trabajos_activos = Trabajo.query.filter(Trabajo.estado != 'entregado').order_by(Trabajo.creado.desc()).all()
    presupuestos_pendientes = Presupuesto.query.filter(
        Presupuesto.estado.in_(['borrador', 'enviado'])
    ).count()
    total_cobrado_mes = 0
    total_saldo = 0
    mes_actual = date.today().month
    anio_actual = date.today().year
    cobros_mes = Cobro.query.filter(
        db.extract('month', Cobro.fecha) == mes_actual,
        db.extract('year', Cobro.fecha) == anio_actual
    ).all()
    total_cobrado_mes = sum(c.monto for c in cobros_mes)
    for t in Trabajo.query.filter(Trabajo.estado != 'entregado').all():
        total_saldo += t.saldo
    return render_template('inicio.html',
        trabajos=trabajos_activos,
        presupuestos_pendientes=presupuestos_pendientes,
        total_cobrado_mes=total_cobrado_mes,
        total_saldo=total_saldo)

# ═══════════════════════════════════════════════════════
# RUTAS — CLIENTES
# ═══════════════════════════════════════════════════════

@app.route('/clientes')
def clientes():
    q = request.args.get('q', '')
    if q:
        lista = Cliente.query.filter(Cliente.empresa.ilike(f'%{q}%')).order_by(Cliente.empresa).all()
    else:
        lista = Cliente.query.order_by(Cliente.empresa).all()
    return render_template('clientes.html', clientes=lista, q=q)

@app.route('/clientes/nuevo', methods=['GET', 'POST'])
def nuevo_cliente():
    if request.method == 'POST':
        c = Cliente(
            empresa=request.form['empresa'],
            contacto=request.form.get('contacto'),
            telefono=request.form.get('telefono'),
            email=request.form.get('email'),
            direccion=request.form.get('direccion'),
            cuit=request.form.get('cuit'),
            notas=request.form.get('notas')
        )
        db.session.add(c)
        db.session.commit()
        return redirect(url_for('clientes'))
    return render_template('cliente_form.html', cliente=None)

@app.route('/clientes/<int:id>/editar', methods=['GET', 'POST'])
def editar_cliente(id):
    c = Cliente.query.get_or_404(id)
    if request.method == 'POST':
        c.empresa = request.form['empresa']
        c.contacto = request.form.get('contacto')
        c.telefono = request.form.get('telefono')
        c.email = request.form.get('email')
        c.direccion = request.form.get('direccion')
        c.cuit = request.form.get('cuit')
        c.notas = request.form.get('notas')
        db.session.commit()
        return redirect(url_for('clientes'))
    return render_template('cliente_form.html', cliente=c)

# ═══════════════════════════════════════════════════════
# RUTAS — PRESUPUESTOS
# ═══════════════════════════════════════════════════════

@app.route('/presupuestos')
def presupuestos():
    estado = request.args.get('estado', 'activos')
    if estado == 'activos':
        lista = Presupuesto.query.filter(
            Presupuesto.estado.in_(['borrador', 'enviado'])
        ).order_by(Presupuesto.creado.desc()).all()
    elif estado == 'aceptados':
        lista = Presupuesto.query.filter_by(estado='aceptado').order_by(Presupuesto.creado.desc()).all()
    else:
        lista = Presupuesto.query.filter_by(estado='rechazado').order_by(Presupuesto.creado.desc()).all()
    return render_template('presupuestos.html', lista=lista, estado=estado)

@app.route('/presupuestos/nuevo', methods=['GET', 'POST'])
def nuevo_presupuesto():
    clientes = Cliente.query.order_by(Cliente.empresa).all()
    if request.method == 'POST':
        p = Presupuesto(
            numero=gen_numero_presupuesto(),
            cliente_id=request.form['cliente_id'],
            tipo_equipo=request.form.get('tipo_equipo'),
            identificador=request.form.get('identificador'),
            marca=request.form.get('marca'),
            modelo=request.form.get('modelo'),
            tipo_trabajo=request.form.get('tipo_trabajo'),
            observaciones=request.form.get('observaciones'),
            estado=request.form.get('estado', 'borrador')
        )
        db.session.add(p)
        db.session.flush()
        descripciones = request.form.getlist('descripcion[]')
        cantidades = request.form.getlist('cantidad[]')
        precios = request.form.getlist('precio[]')
        total = 0
        for desc, cant, precio in zip(descripciones, cantidades, precios):
            if desc.strip():
                cant_f = float(cant or 1)
                precio_f = float(precio or 0)
                sub = cant_f * precio_f
                total += sub
                item = ItemPresupuesto(
                    presupuesto_id=p.id,
                    descripcion=desc,
                    cantidad=cant_f,
                    precio_unitario=precio_f,
                    subtotal=sub
                )
                db.session.add(item)
        p.total = total
        db.session.commit()
        return redirect(url_for('presupuestos'))
    return render_template('presupuesto_form.html', presupuesto=None, clientes=clientes)

@app.route('/presupuestos/<int:id>/editar', methods=['GET', 'POST'])
def editar_presupuesto(id):
    p = Presupuesto.query.get_or_404(id)
    clientes = Cliente.query.order_by(Cliente.empresa).all()
    if request.method == 'POST':
        p.cliente_id = request.form['cliente_id']
        p.tipo_equipo = request.form.get('tipo_equipo')
        p.identificador = request.form.get('identificador')
        p.marca = request.form.get('marca')
        p.modelo = request.form.get('modelo')
        p.tipo_trabajo = request.form.get('tipo_trabajo')
        p.observaciones = request.form.get('observaciones')
        p.estado = request.form.get('estado', p.estado)
        for item in p.items:
            db.session.delete(item)
        descripciones = request.form.getlist('descripcion[]')
        cantidades = request.form.getlist('cantidad[]')
        precios = request.form.getlist('precio[]')
        total = 0
        for desc, cant, precio in zip(descripciones, cantidades, precios):
            if desc.strip():
                cant_f = float(cant or 1)
                precio_f = float(precio or 0)
                sub = cant_f * precio_f
                total += sub
                item = ItemPresupuesto(
                    presupuesto_id=p.id,
                    descripcion=desc,
                    cantidad=cant_f,
                    precio_unitario=precio_f,
                    subtotal=sub
                )
                db.session.add(item)
        p.total = total
        db.session.commit()
        return redirect(url_for('presupuestos'))
    return render_template('presupuesto_form.html', presupuesto=p, clientes=clientes)

@app.route('/presupuestos/<int:id>/estado', methods=['POST'])
def cambiar_estado_presupuesto(id):
    p = Presupuesto.query.get_or_404(id)
    p.estado = request.form['estado']
    db.session.commit()
    return redirect(url_for('presupuestos'))

@app.route('/presupuestos/<int:id>/convertir', methods=['GET', 'POST'])
def convertir_presupuesto(id):
    p = Presupuesto.query.get_or_404(id)
    if request.method == 'POST':
        t = Trabajo(
            numero=gen_numero_trabajo(),
            cliente_id=p.cliente_id,
            presupuesto_id=p.id,
            tipo_equipo=p.tipo_equipo,
            identificador=p.identificador,
            marca=p.marca,
            modelo=p.modelo,
            tipo_trabajo=p.tipo_trabajo,
            observaciones=request.form.get('observaciones', p.observaciones),
            presupuestado=p.total,
            fecha_ingreso=datetime.strptime(request.form['fecha_ingreso'], '%Y-%m-%d').date()
        )
        db.session.add(t)
        p.estado = 'aceptado'
        db.session.commit()
        return redirect(url_for('detalle_trabajo', id=t.id))
    return render_template('convertir.html', presupuesto=p)

# ═══════════════════════════════════════════════════════
# RUTAS — TRABAJOS
# ═══════════════════════════════════════════════════════

@app.route('/trabajos')
def trabajos():
    cliente_id = request.args.get('cliente_id', '')
    estado = request.args.get('estado', '')
    maquina = request.args.get('maquina', '')
    q = Trabajo.query
    if cliente_id:
        q = q.filter_by(cliente_id=cliente_id)
    if estado:
        q = q.filter_by(estado=estado)
    if maquina:
        q = q.filter(Trabajo.identificador.ilike(f'%{maquina}%'))
    lista = q.order_by(Trabajo.creado.desc()).all()
    clientes = Cliente.query.order_by(Cliente.empresa).all()
    return render_template('trabajos.html', trabajos=lista, clientes=clientes,
                           cliente_id=cliente_id, estado=estado, maquina=maquina)

@app.route('/trabajos/<int:id>')
def detalle_trabajo(id):
    t = Trabajo.query.get_or_404(id)
    return render_template('detalle_trabajo.html', trabajo=t)

@app.route('/trabajos/<int:id>/estado', methods=['POST'])
def cambiar_estado_trabajo(id):
    t = Trabajo.query.get_or_404(id)
    t.estado = request.form['estado']
    if t.estado == 'entregado':
        t.fecha_entrega = date.today()
    db.session.commit()
    return redirect(url_for('detalle_trabajo', id=id))

@app.route('/trabajos/<int:id>/gasto', methods=['POST'])
def agregar_gasto_trabajo(id):
    t = Trabajo.query.get_or_404(id)
    g = GastoTrabajo(
        trabajo_id=id,
        fecha=datetime.strptime(request.form['fecha'], '%Y-%m-%d').date(),
        categoria=request.form['categoria'],
        proveedor=request.form.get('proveedor'),
        descripcion=request.form.get('descripcion'),
        monto=float(request.form['monto'] or 0),
        forma_pago=request.form.get('forma_pago')
    )
    db.session.add(g)
    db.session.commit()
    return redirect(url_for('detalle_trabajo', id=id))

@app.route('/trabajos/gasto/<int:id>/eliminar', methods=['POST'])
def eliminar_gasto_trabajo(id):
    g = GastoTrabajo.query.get_or_404(id)
    trabajo_id = g.trabajo_id
    db.session.delete(g)
    db.session.commit()
    return redirect(url_for('detalle_trabajo', id=trabajo_id))

@app.route('/trabajos/<int:id>/cobro', methods=['POST'])
def agregar_cobro(id):
    t = Trabajo.query.get_or_404(id)
    c = Cobro(
        trabajo_id=id,
        fecha=datetime.strptime(request.form['fecha'], '%Y-%m-%d').date(),
        monto=float(request.form['monto'] or 0),
        forma_pago=request.form['forma_pago'],
        comprobante=request.form.get('comprobante'),
        notas=request.form.get('notas')
    )
    db.session.add(c)
    db.session.commit()
    return redirect(url_for('cuenta_corriente') + f'?cliente_id={t.cliente_id}')

@app.route('/trabajos/cobro/<int:id>/eliminar', methods=['POST'])
def eliminar_cobro(id):
    c = Cobro.query.get_or_404(id)
    trabajo_id = c.trabajo_id
    t = Trabajo.query.get(trabajo_id)
    db.session.delete(c)
    db.session.commit()
    return redirect(url_for('cuenta_corriente') + f'?cliente_id={t.cliente_id}')

# ═══════════════════════════════════════════════════════
# RUTAS — CUENTA CORRIENTE
# ═══════════════════════════════════════════════════════

@app.route('/cuenta-corriente')
def cuenta_corriente():
    cliente_id = request.args.get('cliente_id', '')
    clientes = Cliente.query.order_by(Cliente.empresa).all()
    cliente_sel = None
    trabajos = []
    if cliente_id:
        cliente_sel = Cliente.query.get(cliente_id)
        trabajos = Trabajo.query.filter_by(
            cliente_id=cliente_id
        ).order_by(Trabajo.creado.desc()).all()
    return render_template('cuenta_corriente.html',
        clientes=clientes,
        cliente_sel=cliente_sel,
        trabajos=trabajos,
        cliente_id=cliente_id)

# ═══════════════════════════════════════════════════════
# RUTAS — HISTORIAL
# ═══════════════════════════════════════════════════════

@app.route('/historial')
def historial():
    cliente_q = request.args.get('cliente', '')
    maquina_q = request.args.get('maquina', '')
    q = Trabajo.query.filter_by(estado='entregado')
    if cliente_q:
        q = q.join(Cliente).filter(Cliente.empresa.ilike(f'%{cliente_q}%'))
    if maquina_q:
        q = q.filter(Trabajo.identificador.ilike(f'%{maquina_q}%'))
    lista = q.order_by(Trabajo.fecha_entrega.desc()).all()
    return render_template('historial.html', trabajos=lista,
                           cliente_q=cliente_q, maquina_q=maquina_q)

# ═══════════════════════════════════════════════════════
# RUTAS — GASTOS GENERALES
# ═══════════════════════════════════════════════════════

@app.route('/gastos-generales', methods=['GET', 'POST'])
def gastos_generales():
    if request.method == 'POST':
        g = GastoGeneral(
            fecha=datetime.strptime(request.form['fecha'], '%Y-%m-%d').date(),
            categoria=request.form['categoria'],
            proveedor=request.form.get('proveedor'),
            monto=float(request.form['monto'] or 0),
            forma_pago=request.form.get('forma_pago'),
            comprobante=request.form.get('comprobante'),
            notas=request.form.get('notas')
        )
        db.session.add(g)
        db.session.commit()
        return redirect(url_for('gastos_generales'))
    lista = GastoGeneral.query.order_by(GastoGeneral.fecha.desc()).limit(50).all()
    return render_template('gastos_generales.html', gastos=lista)

@app.route('/gastos-generales/<int:id>/eliminar', methods=['POST'])
def eliminar_gasto_general(id):
    g = GastoGeneral.query.get_or_404(id)
    db.session.delete(g)
    db.session.commit()
    return redirect(url_for('gastos_generales'))

# ═══════════════════════════════════════════════════════
# RUTAS — RESULTADOS
# ═══════════════════════════════════════════════════════

@app.route('/resultados')
def resultados():
    mes = int(request.args.get('mes', date.today().month))
    anio = int(request.args.get('anio', date.today().year))

    # Ventas (presupuestado de trabajos ingresados ese mes)
    trabajos_mes = Trabajo.query.filter(
        db.extract('month', Trabajo.fecha_ingreso) == mes,
        db.extract('year', Trabajo.fecha_ingreso) == anio
    ).all()
    ventas = sum(t.presupuestado for t in trabajos_mes)

    # Cobrado ese mes
    cobros = Cobro.query.filter(
        db.extract('month', Cobro.fecha) == mes,
        db.extract('year', Cobro.fecha) == anio
    ).all()
    cobrado = sum(c.monto for c in cobros)

    # Saldo pendiente (todos los trabajos activos)
    saldo_total = sum(t.saldo for t in Trabajo.query.filter(Trabajo.estado != 'entregado').all())

    # Gastos de trabajos ese mes
    gastos_trabajos = GastoTrabajo.query.filter(
        db.extract('month', GastoTrabajo.fecha) == mes,
        db.extract('year', GastoTrabajo.fecha) == anio
    ).all()

    # Gastos generales ese mes
    gastos_gen = GastoGeneral.query.filter(
        db.extract('month', GastoGeneral.fecha) == mes,
        db.extract('year', GastoGeneral.fecha) == anio
    ).all()

    # Desglose categorias
    cats = {}
    for g in gastos_trabajos + gastos_gen:
        cats[g.categoria] = cats.get(g.categoria, 0) + g.monto

    total_gastos = sum(cats.values())
    resultado_financiero = cobrado - total_gastos
    resultado_economico = ventas - total_gastos

    meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
             'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

    return render_template('resultados.html',
        mes=mes, anio=anio, meses=meses,
        ventas=ventas, cobrado=cobrado, saldo_total=saldo_total,
        total_gastos=total_gastos, cats=cats,
        resultado_financiero=resultado_financiero,
        resultado_economico=resultado_economico)

# ═══════════════════════════════════════════════════════
# INICIO
# ═══════════════════════════════════════════════════════

# Crear tablas al iniciar — funciona tanto local como en Render
with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8080, threaded=True)
