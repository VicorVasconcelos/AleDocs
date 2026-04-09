import os
import io
import zipfile
from datetime import datetime
import sqlite3
from flask import (
    Flask, render_template, request, redirect,
    url_for, jsonify, send_file, flash
)
from docxtpl import DocxTemplate

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'aledocs-dev-key-2024-change-in-production')

DATABASE = os.path.join('instance', 'aledocs.db')
TEMPLATES_DIR = 'docx_templates'

MESES_PT = {
    1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
    5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
    9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
}

DOCUMENTOS_LISTA = [
    {
        'id': 'termo_entrega',
        'nome': 'Termo de Entrega de Documentos',
        'descricao': 'Registro formal dos documentos entregues pelo cliente para o serviço.',
        'icon': 'bi-file-earmark-check-fill',
    },
    {
        'id': 'checklist_averbacao',
        'nome': 'Checklist de Averbação de Georreferenciamento',
        'descricao': 'Lista de verificação de todas as etapas e documentos para a averbação.',
        'icon': 'bi-list-check',
    },
    {
        'id': 'contrato_georreferenciamento',
        'nome': 'Contrato de Georreferenciamento',
        'descricao': 'Contrato de prestação de serviços técnicos de georreferenciamento de imóvel rural.',
        'icon': 'bi-geo-alt-fill',
    },
    {
        'id': 'memorial_descritivo',
        'nome': 'Memorial Descritivo',
        'descricao': 'Memorial descritivo com planilha de coordenadas e confrontações do imóvel.',
        'icon': 'bi-map-fill',
    },
    {
        'id': 'recibo_servico',
        'nome': 'Recibo de Prestação de Serviço',
        'descricao': 'Recibo de pagamento pelos serviços técnicos e/ou advocatícios prestados.',
        'icon': 'bi-receipt-cutoff',
    },
    {
        'id': 'contrato_dacao',
        'nome': 'Contrato de Serviços Advocatícios e Técnicos (Dação em Pagamento)',
        'descricao': 'Contrato com cláusula de dação em pagamento de área do imóvel.',
        'icon': 'bi-briefcase-fill',
    },
    {
        'id': 'procuracao',
        'nome': 'Procuração Particular',
        'descricao': 'Procuração outorgando poderes ao responsável técnico/advogado para representar o cliente junto a cartórios, INCRA e órgãos públicos.',
        'icon': 'bi-pen-fill',
    },
]


# ─── Banco de Dados ────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs('instance', exist_ok=True)
    conn = get_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cpf_cnpj TEXT,
            rg TEXT,
            orgao_expedidor TEXT,
            profissao TEXT,
            estado_civil TEXT,
            nome_conjuge TEXT,
            cpf_conjuge TEXT,
            endereco TEXT,
            complemento TEXT,
            cidade TEXT,
            estado TEXT,
            cep TEXT,
            telefone TEXT,
            email TEXT,
            nome_propriedade TEXT,
            municipio_imovel TEXT,
            uf_imovel TEXT,
            matricula_imovel TEXT,
            comarca TEXT,
            codigo_incra TEXT,
            numero_ccir TEXT,
            numero_car TEXT,
            area_imovel TEXT,
            area_registrada TEXT,
            perimetro_imovel TEXT,
            fuso_sirgas TEXT,
            numero_sigef TEXT,
            confrontante_norte TEXT,
            confrontante_sul TEXT,
            confrontante_leste TEXT,
            confrontante_oeste TEXT,
            valor_servico TEXT,
            forma_pagamento TEXT,
            observacoes TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')
    # Migração: adiciona colunas novas em bancos existentes
    novas_colunas = [
        ('rg', 'TEXT'), ('orgao_expedidor', 'TEXT'),
        ('nome_conjuge', 'TEXT'), ('cpf_conjuge', 'TEXT'),
        ('municipio_imovel', 'TEXT'), ('uf_imovel', 'TEXT'),
        ('codigo_incra', 'TEXT'), ('numero_ccir', 'TEXT'), ('numero_car', 'TEXT'),
        ('area_registrada', 'TEXT'), ('perimetro_imovel', 'TEXT'),
        ('fuso_sirgas', 'TEXT'), ('numero_sigef', 'TEXT'),
        ('confrontante_norte', 'TEXT'), ('confrontante_sul', 'TEXT'),
        ('confrontante_leste', 'TEXT'), ('confrontante_oeste', 'TEXT'),
        ('valor_servico', 'TEXT'), ('forma_pagamento', 'TEXT'),
    ]
    for col, tipo in novas_colunas:
        try:
            conn.execute(f'ALTER TABLE clientes ADD COLUMN {col} {tipo}')
        except Exception:
            pass  # coluna já existe

    conn.execute('''
        CREATE TABLE IF NOT EXISTS emissores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            documento TEXT,
            nacionalidade TEXT,
            estado_civil TEXT,
            profissao TEXT,
            registro_tipo TEXT,
            registro_numero TEXT,
            telefone TEXT,
            email TEXT,
            endereco TEXT,
            cidade TEXT,
            estado TEXT,
            observacoes TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    ''')

    # Migração de emissores já existentes.
    novas_colunas_emissor = [
        ('nacionalidade', 'TEXT'),
        ('estado_civil', 'TEXT'),
    ]
    for col, tipo in novas_colunas_emissor:
        try:
            conn.execute(f'ALTER TABLE emissores ADD COLUMN {col} {tipo}')
        except Exception:
            pass

    # Emissor inicial para manter compatibilidade com os modelos já existentes.
    total_emissores = conn.execute('SELECT COUNT(1) FROM emissores').fetchone()[0]
    if total_emissores == 0:
        conn.execute('''
            INSERT INTO emissores
                (nome, nacionalidade, estado_civil, profissao, registro_tipo, registro_numero, cidade, estado)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            'Alessandra Barbara',
            'Brasileira',
            'Solteira',
            'Agrimensora',
            'Registro',
            '',
            '',
            '',
        ))

    conn.commit()
    conn.close()


# ─── Helpers ──────────────────────────────────────────────────────────────────

def formatar_data(dt=None):
    if dt is None:
        dt = datetime.now()
    return f"{dt.day} de {MESES_PT[dt.month]} de {dt.year}"


def montar_contexto_cliente(cliente_row):
    c = dict(cliente_row)

    # Alias descritivo para os templates
    c['nome_cliente'] = c.get('nome', '')

    now = datetime.now()
    c['data_atual'] = formatar_data(now)
    c['data_curta'] = now.strftime('%d/%m/%Y')
    c['ano_atual'] = str(now.year)

    # Endereço completo concatenado
    parts = [c.get('endereco', '')]
    if c.get('complemento'):
        parts.append(c['complemento'])
    if c.get('cidade'):
        parts.append(c['cidade'])
    if c.get('estado'):
        parts.append(c['estado'])
    if c.get('cep'):
        parts.append(f"CEP {c['cep']}")
    c['endereco_completo'] = ', '.join(p for p in parts if p)

    # Município/UF do imóvel: usa dados do imóvel ou cai para cidade do cliente
    if not c.get('municipio_imovel'):
        c['municipio_imovel'] = c.get('cidade', '')
    if not c.get('uf_imovel'):
        c['uf_imovel'] = c.get('estado', '')
    c['municipio_uf_imovel'] = ', '.join(
        p for p in [c.get('municipio_imovel', ''), c.get('uf_imovel', '')] if p
    )

    # Substitui campos vazios por linha para permitir preenchimento manual posterior.
    for key in list(c.keys()):
        val = c[key]
        if val is None or (isinstance(val, str) and not val.strip()):
            c[key] = '______________________________'

    return c


def montar_contexto_emissor(emissor_row):
    e = dict(emissor_row) if emissor_row else {}

    registro_numero = (e.get('registro_numero') or '').strip()
    registro_emissor = f'Registro {registro_numero}' if registro_numero else ''

    cidade_uf = ', '.join(p for p in [e.get('cidade', ''), e.get('estado', '')] if p)
    linha_contato = '  |  '.join(
        p for p in [registro_emissor, e.get('telefone', ''), e.get('email', ''), cidade_uf] if p
    )

    ctx = {
        'emissor_id': str(e.get('id', '')),
        'nome_emissor': e.get('nome', ''),
        'documento_emissor': e.get('documento', ''),
        'nacionalidade_emissor': e.get('nacionalidade', ''),
        'estado_civil_emissor': e.get('estado_civil', ''),
        'profissao_emissor': e.get('profissao', ''),
        'registro_tipo_emissor': (e.get('registro_tipo') or '').strip(),
        'registro_numero_emissor': registro_numero,
        'registro_emissor': registro_emissor,
        'telefone_emissor': e.get('telefone', ''),
        'email_emissor': e.get('email', ''),
        'endereco_emissor': e.get('endereco', ''),
        'cidade_emissor': e.get('cidade', ''),
        'estado_emissor': e.get('estado', ''),
        'cidade_uf_emissor': cidade_uf,
        'linha_emissor': linha_contato,
        'observacoes_emissor': e.get('observacoes', ''),
    }

    for key in list(ctx.keys()):
        val = ctx[key]
        if val is None or (isinstance(val, str) and not val.strip()):
            ctx[key] = '______________________________'

    # Se não houver número de registro, mantém esse campo em branco por solicitação do usuário.
    if not registro_numero:
        ctx['registro_emissor'] = ''
        ctx['registro_numero_emissor'] = ''

    return ctx


def validar_dados_emissor(dados):
    campos_obrigatorios = [
        ('nome', 'Nome do emissor'),
        ('documento', 'CPF/CNPJ do emissor'),
    ]
    faltantes = [rotulo for chave, rotulo in campos_obrigatorios if not (dados.get(chave) or '').strip()]
    return faltantes


def verificar_templates():
    """Retorna lista de templates faltando."""
    faltando = []
    for doc in DOCUMENTOS_LISTA:
        path = os.path.join(TEMPLATES_DIR, f'{doc["id"]}.docx')
        if not os.path.exists(path):
            faltando.append(doc['nome'])
    return faltando


# ─── Rotas ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    conn = get_db()
    clientes = conn.execute(
        'SELECT id, nome, cpf_cnpj FROM clientes ORDER BY nome'
    ).fetchall()
    conn.close()

    faltando = verificar_templates()
    if faltando:
        flash(
            f'⚠️ Modelos .docx não encontrados. Execute "python create_templates.py" primeiro.',
            'warning'
        )

    return render_template('index.html', clientes=clientes)


@app.route('/clientes')
def listar_clientes():
    q = request.args.get('q', '').strip()
    conn = get_db()
    if q:
        clientes = conn.execute(
            'SELECT * FROM clientes WHERE nome LIKE ? OR cpf_cnpj LIKE ? ORDER BY nome',
            (f'%{q}%', f'%{q}%')
        ).fetchall()
    else:
        clientes = conn.execute('SELECT * FROM clientes ORDER BY nome').fetchall()
    conn.close()
    return render_template('listar_clientes.html', clientes=clientes, q=q)


@app.route('/clientes/novo', methods=['GET', 'POST'])
def novo_cliente():
    if request.method == 'POST':
        nome = request.form.get('nome', '').strip()
        if not nome:
            flash('O nome do cliente é obrigatório.', 'danger')
            return render_template('cliente_form.html', dados=request.form, cliente=None)

        conn = get_db()
        try:
            conn.execute('''
                INSERT INTO clientes
                    (nome, cpf_cnpj, rg, orgao_expedidor, profissao, estado_civil,
                     nome_conjuge, cpf_conjuge,
                     endereco, complemento, cidade, estado, cep, telefone, email,
                     nome_propriedade, municipio_imovel, uf_imovel,
                     matricula_imovel, comarca, codigo_incra, numero_ccir, numero_car,
                     area_imovel, area_registrada, perimetro_imovel, fuso_sirgas, numero_sigef,
                     confrontante_norte, confrontante_sul, confrontante_leste, confrontante_oeste,
                     valor_servico, forma_pagamento, observacoes)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', (
                nome,
                request.form.get('cpf_cnpj', '').strip(),
                request.form.get('rg', '').strip(),
                request.form.get('orgao_expedidor', '').strip(),
                request.form.get('profissao', '').strip(),
                request.form.get('estado_civil', '').strip(),
                request.form.get('nome_conjuge', '').strip(),
                request.form.get('cpf_conjuge', '').strip(),
                request.form.get('endereco', '').strip(),
                request.form.get('complemento', '').strip(),
                request.form.get('cidade', '').strip(),
                request.form.get('estado', '').strip(),
                request.form.get('cep', '').strip(),
                request.form.get('telefone', '').strip(),
                request.form.get('email', '').strip(),
                request.form.get('nome_propriedade', '').strip(),
                request.form.get('municipio_imovel', '').strip(),
                request.form.get('uf_imovel', '').strip(),
                request.form.get('matricula_imovel', '').strip(),
                request.form.get('comarca', '').strip(),
                request.form.get('codigo_incra', '').strip(),
                request.form.get('numero_ccir', '').strip(),
                request.form.get('numero_car', '').strip(),
                request.form.get('area_imovel', '').strip(),
                request.form.get('area_registrada', '').strip(),
                request.form.get('perimetro_imovel', '').strip(),
                request.form.get('fuso_sirgas', '').strip(),
                request.form.get('numero_sigef', '').strip(),
                request.form.get('confrontante_norte', '').strip(),
                request.form.get('confrontante_sul', '').strip(),
                request.form.get('confrontante_leste', '').strip(),
                request.form.get('confrontante_oeste', '').strip(),
                request.form.get('valor_servico', '').strip(),
                request.form.get('forma_pagamento', '').strip(),
                request.form.get('observacoes', '').strip(),
            ))
            conn.commit()
            cliente_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
            flash(f'Cliente "{nome}" cadastrado com sucesso!', 'success')
            return redirect(url_for('gerar_docs', cliente_id=cliente_id))
        finally:
            conn.close()

    return render_template('cliente_form.html', dados={}, cliente=None)


@app.route('/clientes/<int:cliente_id>/editar', methods=['GET', 'POST'])
def editar_cliente(cliente_id):
    conn = get_db()
    cliente = conn.execute('SELECT * FROM clientes WHERE id=?', (cliente_id,)).fetchone()

    if not cliente:
        conn.close()
        flash('Cliente não encontrado.', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        nome = request.form.get('nome', '').strip()
        if not nome:
            flash('O nome do cliente é obrigatório.', 'danger')
            conn.close()
            return render_template('cliente_form.html', dados=request.form, cliente=cliente)

        try:
            conn.execute('''
                UPDATE clientes SET
                    nome=?, cpf_cnpj=?, rg=?, orgao_expedidor=?, profissao=?, estado_civil=?,
                    nome_conjuge=?, cpf_conjuge=?,
                    endereco=?, complemento=?, cidade=?, estado=?, cep=?, telefone=?, email=?,
                    nome_propriedade=?, municipio_imovel=?, uf_imovel=?,
                    matricula_imovel=?, comarca=?, codigo_incra=?, numero_ccir=?, numero_car=?,
                    area_imovel=?, area_registrada=?, perimetro_imovel=?, fuso_sirgas=?, numero_sigef=?,
                    confrontante_norte=?, confrontante_sul=?, confrontante_leste=?, confrontante_oeste=?,
                    valor_servico=?, forma_pagamento=?, observacoes=?
                WHERE id=?
            ''', (
                nome,
                request.form.get('cpf_cnpj', '').strip(),
                request.form.get('rg', '').strip(),
                request.form.get('orgao_expedidor', '').strip(),
                request.form.get('profissao', '').strip(),
                request.form.get('estado_civil', '').strip(),
                request.form.get('nome_conjuge', '').strip(),
                request.form.get('cpf_conjuge', '').strip(),
                request.form.get('endereco', '').strip(),
                request.form.get('complemento', '').strip(),
                request.form.get('cidade', '').strip(),
                request.form.get('estado', '').strip(),
                request.form.get('cep', '').strip(),
                request.form.get('telefone', '').strip(),
                request.form.get('email', '').strip(),
                request.form.get('nome_propriedade', '').strip(),
                request.form.get('municipio_imovel', '').strip(),
                request.form.get('uf_imovel', '').strip(),
                request.form.get('matricula_imovel', '').strip(),
                request.form.get('comarca', '').strip(),
                request.form.get('codigo_incra', '').strip(),
                request.form.get('numero_ccir', '').strip(),
                request.form.get('numero_car', '').strip(),
                request.form.get('area_imovel', '').strip(),
                request.form.get('area_registrada', '').strip(),
                request.form.get('perimetro_imovel', '').strip(),
                request.form.get('fuso_sirgas', '').strip(),
                request.form.get('numero_sigef', '').strip(),
                request.form.get('confrontante_norte', '').strip(),
                request.form.get('confrontante_sul', '').strip(),
                request.form.get('confrontante_leste', '').strip(),
                request.form.get('confrontante_oeste', '').strip(),
                request.form.get('valor_servico', '').strip(),
                request.form.get('forma_pagamento', '').strip(),
                request.form.get('observacoes', '').strip(),
                cliente_id,
            ))
            conn.commit()
            flash(f'Dados de "{nome}" atualizados com sucesso!', 'success')
            return redirect(url_for('gerar_docs', cliente_id=cliente_id))
        finally:
            conn.close()

    conn.close()
    return render_template('cliente_form.html', dados=dict(cliente), cliente=cliente)


@app.route('/clientes/<int:cliente_id>/excluir', methods=['POST'])
def excluir_cliente(cliente_id):
    conn = get_db()
    cliente = conn.execute('SELECT nome FROM clientes WHERE id=?', (cliente_id,)).fetchone()
    if cliente:
        conn.execute('DELETE FROM clientes WHERE id=?', (cliente_id,))
        conn.commit()
        flash(f'Cliente "{cliente["nome"]}" excluído com sucesso.', 'success')
    else:
        flash('Cliente não encontrado.', 'danger')
    conn.close()
    return redirect(url_for('listar_clientes'))


@app.route('/emissores')
def listar_emissores():
    q = request.args.get('q', '').strip()
    conn = get_db()
    if q:
        rows = conn.execute(
            'SELECT * FROM emissores WHERE nome LIKE ? OR registro_numero LIKE ? ORDER BY nome',
            (f'%{q}%', f'%{q}%')
        ).fetchall()
    else:
        rows = conn.execute('SELECT * FROM emissores ORDER BY nome').fetchall()
    conn.close()

    emissores = []
    for row in rows:
        emissor = dict(row)
        faltantes = validar_dados_emissor(emissor)
        emissor['cadastro_completo'] = len(faltantes) == 0
        emissor['faltantes'] = faltantes
        emissores.append(emissor)

    return render_template('listar_emissores.html', emissores=emissores, q=q)


@app.route('/emissores/novo', methods=['GET', 'POST'])
def novo_emissor():
    if request.method == 'POST':
        nome = request.form.get('nome', '').strip()
        if not nome:
            flash('O nome do emissor é obrigatório.', 'danger')
            return render_template('emissor_form.html', dados=request.form, emissor=None)

        faltantes = validar_dados_emissor(request.form)
        if faltantes:
            flash('Preencha todos os campos obrigatórios do emissor: ' + ', '.join(faltantes) + '.', 'warning')
            return render_template('emissor_form.html', dados=request.form, emissor=None)

        conn = get_db()
        try:
            conn.execute('''
                INSERT INTO emissores
                    (nome, documento, nacionalidade, estado_civil, profissao, registro_tipo, registro_numero,
                     telefone, email, endereco, cidade, estado, observacoes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                nome,
                request.form.get('documento', '').strip(),
                request.form.get('nacionalidade', '').strip(),
                request.form.get('estado_civil', '').strip(),
                request.form.get('profissao', '').strip(),
                request.form.get('registro_tipo', '').strip(),
                request.form.get('registro_numero', '').strip(),
                request.form.get('telefone', '').strip(),
                request.form.get('email', '').strip(),
                request.form.get('endereco', '').strip(),
                request.form.get('cidade', '').strip(),
                request.form.get('estado', '').strip(),
                request.form.get('observacoes', '').strip(),
            ))
            conn.commit()
            flash(f'Emissor "{nome}" cadastrado com sucesso!', 'success')
            return redirect(url_for('listar_emissores'))
        finally:
            conn.close()

    return render_template('emissor_form.html', dados={}, emissor=None)


@app.route('/emissores/<int:emissor_id>/editar', methods=['GET', 'POST'])
def editar_emissor(emissor_id):
    conn = get_db()
    emissor = conn.execute('SELECT * FROM emissores WHERE id=?', (emissor_id,)).fetchone()

    if not emissor:
        conn.close()
        flash('Emissor não encontrado.', 'danger')
        return redirect(url_for('listar_emissores'))

    if request.method == 'POST':
        nome = request.form.get('nome', '').strip()
        if not nome:
            flash('O nome do emissor é obrigatório.', 'danger')
            conn.close()
            return render_template('emissor_form.html', dados=request.form, emissor=emissor)

        faltantes = validar_dados_emissor(request.form)
        if faltantes:
            flash('Preencha todos os campos obrigatórios do emissor: ' + ', '.join(faltantes) + '.', 'warning')
            conn.close()
            return render_template('emissor_form.html', dados=request.form, emissor=emissor)

        try:
            conn.execute('''
                UPDATE emissores SET
                    nome=?, documento=?, nacionalidade=?, estado_civil=?, profissao=?, registro_tipo=?, registro_numero=?,
                    telefone=?, email=?, endereco=?, cidade=?, estado=?, observacoes=?
                WHERE id=?
            ''', (
                nome,
                request.form.get('documento', '').strip(),
                request.form.get('nacionalidade', '').strip(),
                request.form.get('estado_civil', '').strip(),
                request.form.get('profissao', '').strip(),
                request.form.get('registro_tipo', '').strip(),
                request.form.get('registro_numero', '').strip(),
                request.form.get('telefone', '').strip(),
                request.form.get('email', '').strip(),
                request.form.get('endereco', '').strip(),
                request.form.get('cidade', '').strip(),
                request.form.get('estado', '').strip(),
                request.form.get('observacoes', '').strip(),
                emissor_id,
            ))
            conn.commit()
            flash(f'Emissor "{nome}" atualizado com sucesso!', 'success')
            return redirect(url_for('listar_emissores'))
        finally:
            conn.close()

    conn.close()
    return render_template('emissor_form.html', dados=dict(emissor), emissor=emissor)


@app.route('/emissores/<int:emissor_id>/excluir', methods=['POST'])
def excluir_emissor(emissor_id):
    conn = get_db()
    total = conn.execute('SELECT COUNT(1) FROM emissores').fetchone()[0]
    emissor = conn.execute('SELECT nome FROM emissores WHERE id=?', (emissor_id,)).fetchone()

    if not emissor:
        flash('Emissor não encontrado.', 'danger')
    elif total <= 1:
        flash('É necessário manter ao menos um emissor cadastrado.', 'warning')
    else:
        conn.execute('DELETE FROM emissores WHERE id=?', (emissor_id,))
        conn.commit()
        flash(f'Emissor "{emissor["nome"]}" excluído com sucesso.', 'success')

    conn.close()
    return redirect(url_for('listar_emissores'))


@app.route('/gerar/<int:cliente_id>')
def gerar_docs(cliente_id):
    conn = get_db()
    cliente = conn.execute('SELECT * FROM clientes WHERE id=?', (cliente_id,)).fetchone()
    emissores = conn.execute('SELECT id, nome, registro_tipo, registro_numero FROM emissores ORDER BY nome').fetchall()
    conn.close()

    if not cliente:
        flash('Cliente não encontrado.', 'danger')
        return redirect(url_for('index'))

    if not emissores:
        flash('Cadastre pelo menos um emissor antes de gerar documentos.', 'warning')

    return render_template(
        'gerar_docs.html',
        cliente=cliente,
        documentos=DOCUMENTOS_LISTA,
        emissores=emissores,
    )


@app.route('/gerar/documentos', methods=['POST'])
def processar_geracao():
    cliente_id = request.form.get('cliente_id')
    emissor_id = request.form.get('emissor_id')
    documentos_selecionados = request.form.getlist('documentos')

    if not cliente_id:
        flash('Cliente não identificado.', 'danger')
        return redirect(url_for('index'))

    if not documentos_selecionados:
        flash('Selecione pelo menos um documento para gerar.', 'warning')
        return redirect(url_for('gerar_docs', cliente_id=cliente_id))

    if not emissor_id:
        flash('Selecione um emissor para gerar os documentos.', 'warning')
        return redirect(url_for('gerar_docs', cliente_id=cliente_id))

    conn = get_db()
    cliente = conn.execute('SELECT * FROM clientes WHERE id=?', (cliente_id,)).fetchone()
    emissor = conn.execute('SELECT * FROM emissores WHERE id=?', (emissor_id,)).fetchone()
    conn.close()

    if not cliente:
        flash('Cliente não encontrado.', 'danger')
        return redirect(url_for('index'))

    if not emissor:
        flash('Emissor não encontrado.', 'danger')
        return redirect(url_for('gerar_docs', cliente_id=cliente_id))

    faltantes_emissor = validar_dados_emissor(dict(emissor))
    if faltantes_emissor:
        flash(
            'O emissor selecionado está com cadastro incompleto. '
            'Complete os campos: ' + ', '.join(faltantes_emissor) + '.',
            'warning'
        )
        return redirect(url_for('editar_emissor', emissor_id=emissor['id']))

    contexto = montar_contexto_cliente(cliente)
    contexto.update(montar_contexto_emissor(emissor))

    # Valida IDs contra lista conhecida (evita path traversal)
    ids_validos = {d['id'] for d in DOCUMENTOS_LISTA}
    docs_validos = [d for d in documentos_selecionados if d in ids_validos]

    if not docs_validos:
        flash('Nenhum documento válido selecionado.', 'warning')
        return redirect(url_for('gerar_docs', cliente_id=cliente_id))

    zip_buffer = io.BytesIO()
    erros = []

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for doc_id in docs_validos:
            template_path = os.path.join(TEMPLATES_DIR, f'{doc_id}.docx')
            if not os.path.exists(template_path):
                erros.append(f'Modelo "{doc_id}.docx" não encontrado. Execute create_templates.py.')
                continue
            try:
                tpl = DocxTemplate(template_path)
                tpl.render(contexto)
                buf = io.BytesIO()
                tpl.save(buf)
                buf.seek(0)
                # Nome seguro para o arquivo
                nome_safe = ''.join(
                    c for c in contexto['nome'] if c.isalnum() or c in ' _-'
                ).strip().replace(' ', '_')
                zf.writestr(f'{doc_id}_{nome_safe}.docx', buf.getvalue())
            except Exception as exc:
                erros.append(f'Erro ao gerar "{doc_id}": {exc}')

    if erros:
        for erro in erros:
            flash(erro, 'warning')

    zip_buffer.seek(0)
    nome_safe = ''.join(
        c for c in contexto['nome'] if c.isalnum() or c in ' _-'
    ).strip().replace(' ', '_')
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_zip = f'documentos_{nome_safe}_{timestamp}.zip'

    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name=nome_zip,
    )


# ─── API JSON ─────────────────────────────────────────────────────────────────

@app.route('/api/clientes')
def api_clientes():
    q = request.args.get('q', '').strip()
    conn = get_db()
    if q:
        rows = conn.execute(
            '''SELECT id, nome, cpf_cnpj FROM clientes
               WHERE nome LIKE ? OR cpf_cnpj LIKE ?
               ORDER BY nome LIMIT 40''',
            (f'%{q}%', f'%{q}%'),
        ).fetchall()
    else:
        rows = conn.execute(
            'SELECT id, nome, cpf_cnpj FROM clientes ORDER BY nome LIMIT 40'
        ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/clientes/<int:cliente_id>')
def api_cliente(cliente_id):
    conn = get_db()
    row = conn.execute('SELECT * FROM clientes WHERE id=?', (cliente_id,)).fetchone()
    conn.close()
    if not row:
        return jsonify({'error': 'Cliente não encontrado'}), 404
    return jsonify(dict(row))


# ─── Inicialização ────────────────────────────────────────────────────────────

if __name__ == '__main__':
    init_db()
    app.run(debug=True, port=5000)
