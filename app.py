from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import os
import io

app = Flask(__name__, static_folder='static', static_url_path='/static')

# ─── Base de datos ────────────────────────────────────────────────────────────
# Usa PostgreSQL en Render (variable DATABASE_URL) o SQLite localmente
DATABASE_URL = os.environ.get('DATABASE_URL', 'sqlite:///taller.db')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'taller_costa_2025'

db = SQLAlchemy(app)

# ═══════════════════════════════════════════════════════
# MODELOS — Tablas de la base de datos
# ═══════════════════════════════════════════════════════

class Cliente(db.Model):
    """Clientes del taller"""
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
    """Presupuestos generados para clientes"""
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
    """Ítems de cada presupuesto (descripcion, cantidad, precio)"""
    id = db.Column(db.Integer, primary_key=True)
    presupuesto_id = db.Column(db.Integer, db.ForeignKey('presupuesto.id'), nullable=False)
    descripcion = db.Column(db.String(200), nullable=False)
    cantidad = db.Column(db.Float, default=1)
    precio_unitario = db.Column(db.Float, default=0)
    descuento = db.Column(db.Float, default=0)  # porcentaje 0-100
    subtotal = db.Column(db.Float, default=0)

class Trabajo(db.Model):
    """Trabajos activos e históricos del taller"""
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
    """Gastos asociados a un trabajo específico (repuestos, insumos, etc.)"""
    id = db.Column(db.Integer, primary_key=True)
    trabajo_id = db.Column(db.Integer, db.ForeignKey('trabajo.id'), nullable=False)
    fecha = db.Column(db.Date, default=date.today)
    categoria = db.Column(db.String(50))
    proveedor = db.Column(db.String(100))
    descripcion = db.Column(db.String(200))
    monto = db.Column(db.Float, default=0)
    forma_pago = db.Column(db.String(30))

class Cobro(db.Model):
    """Pagos recibidos por un trabajo"""
    id = db.Column(db.Integer, primary_key=True)
    trabajo_id = db.Column(db.Integer, db.ForeignKey('trabajo.id'), nullable=False)
    fecha = db.Column(db.Date, default=date.today)
    monto = db.Column(db.Float, default=0)
    forma_pago = db.Column(db.String(30))
    comprobante = db.Column(db.String(50))
    notas = db.Column(db.Text)

class GastoGeneral(db.Model):
    """Gastos del taller no asociados a un trabajo (alquiler, servicios, etc.)"""
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.Date, default=date.today)
    categoria = db.Column(db.String(50))
    proveedor = db.Column(db.String(100))
    monto = db.Column(db.Float, default=0)
    forma_pago = db.Column(db.String(30))
    comprobante = db.Column(db.String(50))
    notas = db.Column(db.Text)

class OpcionLista(db.Model):
    """Opciones configurables de los desplegables (tipos de equipo, trabajo, etc.)"""
    id = db.Column(db.Integer, primary_key=True)
    tipo = db.Column(db.String(50), nullable=False)   # tipo_equipo, tipo_trabajo, categoria_gasto, forma_pago
    valor = db.Column(db.String(100), nullable=False)
    orden = db.Column(db.Integer, default=0)          # para ordenar los items en el desplegable

    def __repr__(self):
        return f'<OpcionLista {self.tipo}: {self.valor}>'

# ═══════════════════════════════════════════════════════
# HELPERS — Funciones auxiliares
# ═══════════════════════════════════════════════════════

def gen_numero_presupuesto():
    """Genera el próximo número de presupuesto (PRES-0001, PRES-0002, etc.)"""
    ultimo = Presupuesto.query.order_by(Presupuesto.id.desc()).first()
    n = (ultimo.id + 1) if ultimo else 1
    return f"PRES-{n:04d}"

def gen_numero_trabajo():
    """Genera el próximo número de reparación (REP-0001, REP-0002, etc.)"""
    ultimo = Trabajo.query.order_by(Trabajo.id.desc()).first()
    n = (ultimo.id + 1) if ultimo else 1
    return f"REP-{n:04d}"

def fmt_moneda(valor):
    """Formatea un número como moneda argentina ($1.234.567)"""
    if valor is None:
        return "$0"
    return f"${valor:,.0f}".replace(",", ".")

def get_opciones(tipo):
    """Devuelve la lista de opciones para un tipo de desplegable"""
    return [o.valor for o in OpcionLista.query.filter_by(tipo=tipo).order_by(OpcionLista.orden, OpcionLista.valor).all()]

def seed_opciones():
    """Carga las opciones por defecto si la tabla está vacía"""
    if OpcionLista.query.count() == 0:
        defaults = {
            'tipo_equipo': ['Camión', 'Pluma / Grúa', 'Autoelevador / Clark',
                            'Excavadora', 'Retroexcavadora', 'Compactadora',
                            'Generador', 'Compresor', 'Otro'],
            'tipo_trabajo': ['Mecánica general', 'Motor', 'Transmisión',
                             'Hidráulica', 'Eléctrico', 'Neumático',
                             'Chapa y pintura', 'Restauración', 'Preventivo', 'Otro'],
            'categoria_gasto': ['Repuestos', 'Insumos', 'Mano de Obra Externa',
                                'Herramientas', 'Alquiler', 'Servicios (luz/agua/gas)',
                                'Seguros', 'Impuestos', 'Administración', 'Fletes', 'Otros'],
            'forma_pago': ['Efectivo', 'Transferencia', 'Cheque',
                           'Tarjeta Débito', 'Tarjeta Crédito', 'Cuenta Corriente'],
        }
        for tipo, valores in defaults.items():
            for i, valor in enumerate(valores):
                db.session.add(OpcionLista(tipo=tipo, valor=valor, orden=i))
        db.session.commit()

app.jinja_env.filters['moneda'] = fmt_moneda

# ═══════════════════════════════════════════════════════
# RUTAS — INICIO
# ═══════════════════════════════════════════════════════

@app.route('/')
def inicio():
    """Pantalla principal con KPIs y trabajos activos"""
    trabajos_activos = Trabajo.query.filter(Trabajo.estado != 'entregado').order_by(Trabajo.creado.desc()).all()
    presupuestos_pendientes = Presupuesto.query.filter(
        Presupuesto.estado.in_(['borrador', 'enviado'])
    ).count()
    mes_actual = date.today().month
    anio_actual = date.today().year
    cobros_mes = Cobro.query.filter(
        db.extract('month', Cobro.fecha) == mes_actual,
        db.extract('year', Cobro.fecha) == anio_actual
    ).all()
    total_cobrado_mes = sum(c.monto for c in cobros_mes)
    total_saldo = sum(t.saldo for t in Trabajo.query.filter(Trabajo.estado != 'entregado').all())
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
        lista = Presupuesto.query.filter(Presupuesto.estado.in_(['borrador', 'enviado'])).order_by(Presupuesto.creado.desc()).all()
    elif estado == 'aceptados':
        lista = Presupuesto.query.filter_by(estado='aceptado').order_by(Presupuesto.creado.desc()).all()
    else:
        lista = Presupuesto.query.filter_by(estado='rechazado').order_by(Presupuesto.creado.desc()).all()
    return render_template('presupuestos.html', lista=lista, estado=estado)

@app.route('/presupuestos/nuevo', methods=['GET', 'POST'])
def nuevo_presupuesto():
    clientes = Cliente.query.order_by(Cliente.empresa).all()
    tipos_equipo = get_opciones('tipo_equipo')
    tipos_trabajo = get_opciones('tipo_trabajo')
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
        descuentos = request.form.getlist('descuento[]')
        total = 0
        for desc, cant, precio, desc_pct in zip(descripciones, cantidades, precios, descuentos):
            if desc.strip():
                cant_f = float(cant or 1)
                precio_f = float(precio or 0)
                desc_f = float(desc_pct or 0)
                sub = cant_f * precio_f * (1 - desc_f / 100)
                total += sub
                item = ItemPresupuesto(
                    presupuesto_id=p.id,
                    descripcion=desc,
                    cantidad=cant_f,
                    precio_unitario=precio_f,
                    descuento=desc_f,
                    subtotal=sub
                )
                db.session.add(item)
        p.total = total
        db.session.commit()
        return redirect(url_for('presupuestos'))
    return render_template('presupuesto_form.html', presupuesto=None,
                           clientes=clientes, tipos_equipo=tipos_equipo, tipos_trabajo=tipos_trabajo)

@app.route('/presupuestos/<int:id>/editar', methods=['GET', 'POST'])
def editar_presupuesto(id):
    p = Presupuesto.query.get_or_404(id)
    clientes = Cliente.query.order_by(Cliente.empresa).all()
    tipos_equipo = get_opciones('tipo_equipo')
    tipos_trabajo = get_opciones('tipo_trabajo')
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
        descuentos = request.form.getlist('descuento[]')
        total = 0
        for desc, cant, precio, desc_pct in zip(descripciones, cantidades, precios, descuentos):
            if desc.strip():
                cant_f = float(cant or 1)
                precio_f = float(precio or 0)
                desc_f = float(desc_pct or 0)
                sub = cant_f * precio_f * (1 - desc_f / 100)
                total += sub
                item = ItemPresupuesto(
                    presupuesto_id=p.id,
                    descripcion=desc,
                    cantidad=cant_f,
                    precio_unitario=precio_f,
                    descuento=desc_f,
                    subtotal=sub
                )
                db.session.add(item)
        p.total = total
        db.session.commit()
        return redirect(url_for('presupuestos'))
    return render_template('presupuesto_form.html', presupuesto=p,
                           clientes=clientes, tipos_equipo=tipos_equipo, tipos_trabajo=tipos_trabajo)

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

@app.route('/presupuestos/<int:id>/pdf')
def pdf_presupuesto(id):
    """Genera y descarga el PDF del presupuesto"""
    p = Presupuesto.query.get_or_404(id)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=1.2*cm, leftMargin=1.2*cm,
                            topMargin=1*cm, bottomMargin=1.5*cm)
    elements = []
    styles = getSampleStyleSheet()

    # ── Colores ──────────────────────────────────────────
    navy = colors.HexColor('#1F3864')
    gray_light = colors.HexColor('#F2F2F2')
    border_color = colors.HexColor('#CCCCCC')

    # ── Header: Logo + datos taller ───────────────────────
    logo_path = os.path.join(app.root_path, 'static', 'Logo.png')
    print(f"Buscando logo en: {logo_path}", flush=True)
    print(f"Existe: {os.path.exists(logo_path)}", flush=True)
    header_data = []
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=6.5*cm, height=2.5*cm)
        servicios = Paragraph(
            "Reparación integral de grúas y autoelevadores<br/>"
            "Sistemas hidráulicos<br/>"
            "Motores industriales<br/>"
            "Transmisiones automáticas",
            ParagraphStyle('serv', fontSize=8, leading=12, textColor=colors.black)
        )
        header_data = [[logo, servicios]]
    else:
        header_data = [[
            Paragraph("<b>COSTA MECÁNICA INTEGRAL</b>",
                      ParagraphStyle('t', fontSize=14, textColor=navy)),
            Paragraph("Reparación integral de maquinaria pesada",
                      ParagraphStyle('s', fontSize=9))
        ]]

    header_table = Table(header_data, colWidths=[7*cm, 10*cm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
    ]))
    elements.append(header_table)

    # ── Línea + dirección + número ─────────────────────────
    elements.append(Spacer(1, 0.3*cm))
    dir_pres_data = [[
        Paragraph("<b>📍 Carlos F. Melo 488, CABA</b><br/><b>📞 1173662508 / 1161415101</b>",
                  ParagraphStyle('dir', fontSize=8, textColor=colors.black)),
        Paragraph("<b>PRESUPUESTO</b>",
                  ParagraphStyle('ptitle', fontSize=14, textColor=navy, alignment=TA_CENTER)),
        Paragraph(f"<b>N° {p.numero.replace('PRES-','')}</b>",
                  ParagraphStyle('pnum', fontSize=14, textColor=colors.white,
                                 backColor=navy, alignment=TA_CENTER))
    ]]
    dir_table = Table(dir_pres_data, colWidths=[7*cm, 5*cm, 5*cm])
    dir_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (2,0), (2,0), navy),
        ('LEFTPADDING', (2,0), (2,0), 6),
        ('RIGHTPADDING', (2,0), (2,0), 6),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(dir_table)
    elements.append(Spacer(1, 0.5*cm))

    # ── Información del cliente y presupuesto ─────────────
    valido_hasta = p.creado.date() + timedelta(days=15)
    info_title = Table([
        [Paragraph("<b>INFORMACIÓN DEL CLIENTE Y PRESUPUESTO</b>",
                   ParagraphStyle('it', fontSize=9, textColor=colors.white, alignment=TA_CENTER))]
    ], colWidths=[17*cm])
    info_title.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), navy),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(info_title)

    lbl = ParagraphStyle('lbl', fontSize=8, textColor=colors.black)
    val = ParagraphStyle('val', fontSize=8)

    info_data = [
        [Paragraph('<b>N° Presupuesto:</b>', lbl), Paragraph(str(p.id), val),
         Paragraph('<b>Fecha:</b>', lbl), Paragraph(p.creado.strftime('%d/%m/%Y'), val),
         Paragraph('<b>Empresa:</b>', lbl), Paragraph(p.cliente.empresa or '', val)],
        [Paragraph('<b>N° Reparación:</b>', lbl), Paragraph('', val),
         Paragraph('<b>Válido hasta:</b>', lbl), Paragraph(valido_hasta.strftime('%d/%m/%Y'), val),
         Paragraph('<b>Contacto:</b>', lbl), Paragraph(p.cliente.contacto or '', val)],
    ]
    info_table = Table(info_data, colWidths=[3*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 4*cm])
    info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('BACKGROUND', (0,0), (-1,-1), gray_light),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('FONTSIZE', (0,0), (-1,-1), 8),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 0.3*cm))

    # ── Datos del equipo ──────────────────────────────────
    equipo_title = Table([
        [Paragraph("<b>DATOS DEL EQUIPO / MÁQUINA</b>",
                   ParagraphStyle('et', fontSize=9, textColor=colors.white, alignment=TA_CENTER))]
    ], colWidths=[17*cm])
    equipo_title.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), navy),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(equipo_title)

    equipo_data = [
        [Paragraph('<b>Tipo de Equipo:</b>', lbl), Paragraph(p.tipo_equipo or '', val),
         Paragraph('<b>Marca:</b>', lbl), Paragraph(p.marca or '', val),
         Paragraph('<b>Modelo:</b>', lbl), Paragraph(p.modelo or '', val)],
        [Paragraph('<b>N° Serie / Interno:</b>', lbl), Paragraph(p.identificador or '', val),
         Paragraph('<b>Tipo de Trabajo:</b>', lbl), Paragraph(p.tipo_trabajo or '', val),
         Paragraph('', lbl), Paragraph('', val)],
    ]
    equipo_table = Table(equipo_data, colWidths=[3*cm, 3*cm, 3*cm, 3*cm, 2*cm, 3*cm])
    equipo_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('BACKGROUND', (0,0), (-1,-1), gray_light),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('SPAN',(3,1), (5,1)),
    ]))
    elements.append(equipo_table)
    elements.append(Spacer(1, 0.3*cm))

    # ── Tabla de ítems ────────────────────────────────────
    hdr_style = ParagraphStyle('hdr', fontSize=8, textColor=colors.white, alignment=TA_CENTER)
    items_header = [
        Paragraph('<b>#</b>', hdr_style),
        Paragraph('<b>Descripción del Trabajo / Repuesto</b>', hdr_style),
        Paragraph('<b>Cantidad</b>', hdr_style),
        Paragraph('<b>P. Unit. ($)</b>', hdr_style),
        Paragraph('<b>Desc. (%)</b>', hdr_style),
        Paragraph('<b>Subtotal ($)</b>', hdr_style),
    ]
    items_data = [items_header]
    cell_style = ParagraphStyle('cell', fontSize=8)
    cell_right = ParagraphStyle('cellr', fontSize=8, alignment=TA_RIGHT)

    for i, item in enumerate(p.items, 1):
        items_data.append([
            Paragraph(str(i), cell_style),
            Paragraph(item.descripcion, cell_style),
            Paragraph(fmt_numero(item.cantidad), cell_right),
            Paragraph(fmt_moneda(item.precio_unitario), cell_right),
            Paragraph(f"{item.descuento:.0f}%" if item.descuento else '', cell_right),
            Paragraph(fmt_moneda(item.subtotal), cell_right),
        ])

    # Filas vacías hasta 15
    for i in range(len(p.items) + 1, 13):
        items_data.append([Paragraph(str(i), cell_style), '', '', '', '', ''])

    items_table = Table(items_data, colWidths=[1*cm, 7.5*cm, 2*cm, 2.5*cm, 2*cm, 2*cm])
    items_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), navy),
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, gray_light]),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,1), (1,-1), 'LEFT'),
    ]))
    elements.append(items_table)

    # ── Total final ───────────────────────────────────────
    total_data = [
        ['', '', '', '', Paragraph('<b>TOTAL FINAL:</b>',
                                    ParagraphStyle('tf', fontSize=9, textColor=colors.white, alignment=TA_RIGHT)),
         Paragraph(f'<b>{fmt_moneda(p.total)}</b>',
                   ParagraphStyle('tv', fontSize=9, textColor=colors.white, alignment=TA_RIGHT))]
    ]
    total_table = Table(total_data, colWidths=[1*cm, 7.5*cm, 1*cm, 1*cm, 3.25*cm, 3.25*cm])
    total_table.setStyle(TableStyle([
        ('BACKGROUND', (4,0), (-1,-1), navy),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('GRID', (4,0), (-1,-1), 0.5, border_color),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(total_table)
    elements.append(Spacer(1, 0.5*cm))

    # ── Condiciones ───────────────────────────────────────
    cond_title = Table([
        [Paragraph("<b>CONDICIONES Y NOTAS</b>",
                   ParagraphStyle('ct', fontSize=9, textColor=colors.white, alignment=TA_CENTER))]
    ], colWidths=[17*cm])
    cond_title.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), navy),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(cond_title)

    condiciones = p.observaciones or (
        "• El presupuesto tiene validez de 15 días desde la fecha de emisión.\n"
        "• Los repuestos no incluidos en este presupuesto serán cotizados por separado.\n"
        "• El tiempo estimado de entrega se acordará al confirmar el trabajo."
    )
    cond_table = Table([
        [Paragraph(condiciones.replace('\n', '<br/>'),
                   ParagraphStyle('cond', fontSize=8, leading=12))]
    ], colWidths=[17*cm])
    cond_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
    ]))
    elements.append(cond_table)
    elements.append(Spacer(1, 2.5*cm))

    # ── Firmas ────────────────────────────────────────────
    firma_data = [
        ['_' * 30, '', '_' * 30],
        [Paragraph('Firma taller', ParagraphStyle('f', fontSize=8, alignment=TA_CENTER)),
         '',
         Paragraph('Firma cliente', ParagraphStyle('f', fontSize=8, alignment=TA_CENTER))],
    ]
    firma_table = Table(firma_data, colWidths=[6*cm, 5*cm, 6*cm])
    firma_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(firma_table)

    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True,
                     download_name=f'Presupuesto_{p.numero}.pdf',
                     mimetype='application/pdf')

def fmt_numero(n):
    """Formatea cantidad: sin decimales si es entero"""
    if n == int(n):
        return str(int(n))
    return f"{n:.2f}"

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
    categorias = get_opciones('categoria_gasto')
    formas_pago = get_opciones('forma_pago')
    return render_template('detalle_trabajo.html', trabajo=t,
                           categorias=categorias, formas_pago=formas_pago)

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
    Trabajo.query.get_or_404(id)
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
    t = Trabajo.query.get(c.trabajo_id)
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
    formas_pago = get_opciones('forma_pago')
    cliente_sel = None
    trabajos = []
    if cliente_id:
        cliente_sel = Cliente.query.get(cliente_id)
        trabajos = Trabajo.query.filter_by(cliente_id=cliente_id).order_by(Trabajo.creado.desc()).all()
    return render_template('cuenta_corriente.html',
        clientes=clientes, cliente_sel=cliente_sel,
        trabajos=trabajos, cliente_id=cliente_id, formas_pago=formas_pago)

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
    categorias = get_opciones('categoria_gasto')
    formas_pago = get_opciones('forma_pago')
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
    return render_template('gastos_generales.html', gastos=lista,
                           categorias=categorias, formas_pago=formas_pago)

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
    trabajos_mes = Trabajo.query.filter(
        db.extract('month', Trabajo.fecha_ingreso) == mes,
        db.extract('year', Trabajo.fecha_ingreso) == anio
    ).all()
    ventas = sum(t.presupuestado for t in trabajos_mes)
    cobros = Cobro.query.filter(
        db.extract('month', Cobro.fecha) == mes,
        db.extract('year', Cobro.fecha) == anio
    ).all()
    cobrado = sum(c.monto for c in cobros)
    saldo_total = sum(t.saldo for t in Trabajo.query.filter(Trabajo.estado != 'entregado').all())
    gastos_trabajos = GastoTrabajo.query.filter(
        db.extract('month', GastoTrabajo.fecha) == mes,
        db.extract('year', GastoTrabajo.fecha) == anio
    ).all()
    gastos_gen = GastoGeneral.query.filter(
        db.extract('month', GastoGeneral.fecha) == mes,
        db.extract('year', GastoGeneral.fecha) == anio
    ).all()
    cats = {}
    for g in gastos_trabajos + gastos_gen:
        cats[g.categoria] = cats.get(g.categoria, 0) + g.monto
    total_gastos = sum(cats.values())
    meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
             'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    return render_template('resultados.html',
        mes=mes, anio=anio, meses=meses,
        ventas=ventas, cobrado=cobrado, saldo_total=saldo_total,
        total_gastos=total_gastos, cats=cats,
        resultado_financiero=cobrado - total_gastos,
        resultado_economico=ventas - total_gastos)

# ═══════════════════════════════════════════════════════
# RUTAS — CONFIGURACIÓN
# ═══════════════════════════════════════════════════════

@app.route('/configuracion')
def configuracion():
    """Pantalla para gestionar las opciones de los desplegables"""
    tipos = {
        'tipo_equipo': 'Tipos de equipo',
        'tipo_trabajo': 'Tipos de trabajo',
        'categoria_gasto': 'Categorías de gasto',
        'forma_pago': 'Formas de pago',
    }
    opciones = {}
    for tipo in tipos:
        opciones[tipo] = OpcionLista.query.filter_by(tipo=tipo).order_by(OpcionLista.orden, OpcionLista.valor).all()
    return render_template('configuracion.html', tipos=tipos, opciones=opciones)

@app.route('/configuracion/agregar', methods=['POST'])
def agregar_opcion():
    tipo = request.form['tipo']
    valor = request.form['valor'].strip()
    if valor:
        # Verificar que no exista ya
        existe = OpcionLista.query.filter_by(tipo=tipo, valor=valor).first()
        if not existe:
            db.session.add(OpcionLista(tipo=tipo, valor=valor))
            db.session.commit()
    return redirect(url_for('configuracion'))

@app.route('/configuracion/<int:id>/eliminar', methods=['POST'])
def eliminar_opcion(id):
    o = OpcionLista.query.get_or_404(id)
    db.session.delete(o)
    db.session.commit()
    return redirect(url_for('configuracion'))

# ═══════════════════════════════════════════════════════
# RUTAS — EXPORTAR EXCEL
# ═══════════════════════════════════════════════════════

@app.route('/exportar-excel')
def exportar_excel():
    """Genera un Excel con todas las tablas para usar en PowerBI"""
    wb = openpyxl.Workbook()

    navy = "1F3864"
    white = "FFFFFF"

    def hdr_style(cell, text):
        cell.value = text
        cell.font = Font(bold=True, color=white, name='Arial', size=10)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # ── Hoja Clientes ──────────────────────────────────
    ws = wb.active
    ws.title = "Clientes"
    headers = ['ID', 'Empresa', 'Contacto', 'Teléfono', 'Email', 'Dirección', 'CUIT']
    for ci, h in enumerate(headers, 1):
        hdr_style(ws.cell(1, ci), h)
    for c in Cliente.query.order_by(Cliente.empresa).all():
        ws.append([c.id, c.empresa, c.contacto, c.telefono, c.email, c.direccion, c.cuit])

    # ── Hoja Presupuestos ──────────────────────────────
    ws2 = wb.create_sheet("Presupuestos")
    headers2 = ['ID', 'Número', 'Cliente', 'Tipo Equipo', 'Identificador', 'Marca',
                'Modelo', 'Tipo Trabajo', 'Estado', 'Total', 'Fecha']
    for ci, h in enumerate(headers2, 1):
        hdr_style(ws2.cell(1, ci), h)
    for p in Presupuesto.query.order_by(Presupuesto.creado.desc()).all():
        ws2.append([p.id, p.numero, p.cliente.empresa, p.tipo_equipo, p.identificador,
                    p.marca, p.modelo, p.tipo_trabajo, p.estado, p.total,
                    p.creado.strftime('%d/%m/%Y')])

    # ── Hoja Trabajos ──────────────────────────────────
    ws3 = wb.create_sheet("Trabajos")
    headers3 = ['ID', 'Número', 'Cliente', 'Tipo Equipo', 'Identificador', 'Marca',
                'Modelo', 'Tipo Trabajo', 'Estado', 'Presupuestado', 'Cobrado',
                'Saldo', 'Gastos', 'Ganancia', 'F. Ingreso', 'F. Entrega']
    for ci, h in enumerate(headers3, 1):
        hdr_style(ws3.cell(1, ci), h)
    for t in Trabajo.query.order_by(Trabajo.creado.desc()).all():
        ws3.append([t.id, t.numero, t.cliente.empresa, t.tipo_equipo, t.identificador,
                    t.marca, t.modelo, t.tipo_trabajo, t.estado, t.presupuestado,
                    t.total_cobrado, t.saldo, t.total_gastos, t.ganancia,
                    t.fecha_ingreso.strftime('%d/%m/%Y') if t.fecha_ingreso else '',
                    t.fecha_entrega.strftime('%d/%m/%Y') if t.fecha_entrega else ''])

    # ── Hoja Gastos ────────────────────────────────────
    ws4 = wb.create_sheet("Gastos")
    headers4 = ['ID', 'Trabajo', 'Cliente', 'Fecha', 'Categoría', 'Proveedor', 'Descripción', 'Monto', 'Forma Pago']
    for ci, h in enumerate(headers4, 1):
        hdr_style(ws4.cell(1, ci), h)
    for g in GastoTrabajo.query.order_by(GastoTrabajo.fecha.desc()).all():
        ws4.append([g.id, g.trabajo.numero, g.trabajo.cliente.empresa,
                    g.fecha.strftime('%d/%m/%Y'), g.categoria, g.proveedor,
                    g.descripcion, g.monto, g.forma_pago])

    # ── Hoja Gastos Generales ──────────────────────────
    ws5 = wb.create_sheet("Gastos Generales")
    headers5 = ['ID', 'Fecha', 'Categoría', 'Proveedor', 'Monto', 'Forma Pago', 'Comprobante']
    for ci, h in enumerate(headers5, 1):
        hdr_style(ws5.cell(1, ci), h)
    for g in GastoGeneral.query.order_by(GastoGeneral.fecha.desc()).all():
        ws5.append([g.id, g.fecha.strftime('%d/%m/%Y'), g.categoria,
                    g.proveedor, g.monto, g.forma_pago, g.comprobante])

    # ── Hoja Cobranzas ─────────────────────────────────
    ws6 = wb.create_sheet("Cobranzas")
    headers6 = ['ID', 'Trabajo', 'Cliente', 'Fecha', 'Monto', 'Forma Pago', 'Comprobante', 'Notas']
    for ci, h in enumerate(headers6, 1):
        hdr_style(ws6.cell(1, ci), h)
    for c in Cobro.query.order_by(Cobro.fecha.desc()).all():
        ws6.append([c.id, c.trabajo.numero, c.trabajo.cliente.empresa,
                    c.fecha.strftime('%d/%m/%Y'), c.monto, c.forma_pago,
                    c.comprobante, c.notas])

    # Ajustar anchos de columna automáticamente
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max((len(str(cell.value or '')) for cell in col), default=10)
            sheet.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True,
                     download_name=f'TallerCosta_{date.today().strftime("%Y%m%d")}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ═══════════════════════════════════════════════════════
# INICIO — Crear tablas y cargar datos por defecto
# ═══════════════════════════════════════════════════════

with app.app_context():
    db.create_all()       # Crea las tablas si no existen
    seed_opciones()       # Carga opciones por defecto si la tabla está vacía

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=8080, threaded=True)
