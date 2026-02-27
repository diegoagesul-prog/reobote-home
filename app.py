from flask import Flask, render_template_string, request, redirect, session, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from datetime import datetime
import os
from io import BytesIO
import json
import psycopg2
import psycopg2.extras

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "reobote_home_secret")


# =============================
# POSTGRES CONNECTION
# =============================
def get_db():
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL não configurada (Render → Web Service → Environment).")

    # compatibilidade: alguns ambientes usam postgres://
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)

    return psycopg2.connect(db_url)


def safe_json_load(x):
    if x is None:
        return {}
    if isinstance(x, (dict, list)):
        return x
    try:
        return json.loads(x)
    except Exception:
        return {}


# =============================
# INIT DB (Postgres)
# =============================
def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id SERIAL PRIMARY KEY,
        nome TEXT,
        email TEXT UNIQUE,
        senha TEXT,
        tipo TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS obras (
        id SERIAL PRIMARY KEY,
        nome TEXT UNIQUE
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS servicos (
        id SERIAL PRIMARY KEY,
        nome TEXT UNIQUE,
        unidade TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS insumos (
        id SERIAL PRIMARY KEY,
        nome TEXT UNIQUE,
        unidade TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS funcoes (
        id SERIAL PRIMARY KEY,
        nome TEXT UNIQUE
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS produtividade (
        id SERIAL PRIMARY KEY,
        obra TEXT,
        servico TEXT,
        servico_unidade TEXT,
        quantidade DOUBLE PRECISION,
        horas DOUBLE PRECISION,
        funcoes JSONB,
        insumos JSONB,
        observacao TEXT,
        data TEXT,
        user_id INTEGER
    );
    """)

    # Admin padrão
    senha_hash = generate_password_hash("123456")
    cur.execute("""
      INSERT INTO usuarios (nome, email, senha, tipo)
      VALUES (%s, %s, %s, %s)
      ON CONFLICT (email) DO NOTHING;
    """, ("Administrador", "admin@reobote.com", senha_hash, "admin"))

    conn.commit()
    cur.close()
    conn.close()


# roda init_db ao iniciar
init_db()


# =============================
# TEMPLATE BASE
# =============================
base_html = """
<!DOCTYPE html>
<html>
<head>
  <title>Reobote Home</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

  <!-- Tom Select (autocomplete) -->
  <link href="https://cdn.jsdelivr.net/npm/tom-select@2.3.1/dist/css/tom-select.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/tom-select@2.3.1/dist/js/tom-select.complete.min.js"></script>

  <style>
    body { background: #f8f9fa; }
    .card { margin-bottom: 20px; }
    h2 { color: #0d6efd; }
    .mini { font-size: .92rem; color: #6c757d; }
    .badge-item { margin-right: 6px; margin-bottom: 6px; display:inline-block; }
    .line { border-bottom: 1px solid #eee; padding-bottom: 10px; margin-bottom: 10px; }
    .ts-control { padding: .375rem .75rem; border-radius: .375rem; }
  </style>
</head>
<body>
<div class="container mt-4">
  <div class="text-center mb-4">
    <img src="/static/logo.png" width="120"><br>
    <h2 class="mt-2">REOBOTE HOME</h2>
    <hr>
  </div>
  {{ conteudo|safe }}
</div>

<script>
let tsFunc = null;
let tsIns = null;

function setServicoUnidade(){
  const sel = document.getElementById("servico_select");
  const uni = document.getElementById("servico_unidade");
  if(!sel || !uni) return;
  const opt = sel.selectedOptions && sel.selectedOptions[0];
  uni.value = (opt && opt.dataset && opt.dataset.unidade) ? opt.dataset.unidade : "";
}

function getInsumoUnidadeByValue(val){
  const sel = document.getElementById("ins_select");
  if(!sel) return "un";
  const opt = Array.from(sel.options).find(o => o.value === val);
  if(!opt) return "un";
  return (opt.dataset && opt.dataset.unidade) ? opt.dataset.unidade : "un";
}

function updateInsumoUnidFromSelected(){
  const sel = document.getElementById("ins_select");
  const unidEl = document.getElementById("ins_unid");
  if(!sel || !unidEl) return;
  unidEl.value = getInsumoUnidadeByValue(sel.value);
}

function addFunc(){
  const sel = document.getElementById("func_select");
  const nome = sel ? sel.value : "";
  const qtdEl = document.getElementById("func_qtd");
  const qtd = qtdEl ? (qtdEl.value || "0") : "0";
  if(!nome) return;

  const wrap = document.createElement("div");
  wrap.className = "badge bg-light text-dark badge-item p-2";
  wrap.innerHTML = `
    <strong>${nome}</strong> - ${qtd}
    <button type="button" class="btn btn-sm btn-danger ms-2" onclick="this.parentElement.remove()">x</button>
    <input type="hidden" name="funcoes_nome[]" value="${nome}">
    <input type="hidden" name="funcoes_qtd[]" value="${qtd}">
  `;
  document.getElementById("func_list").appendChild(wrap);

  if(tsFunc) tsFunc.clear(true);
  if(qtdEl) qtdEl.value = "1";
}

function addIns(){
  const sel = document.getElementById("ins_select");
  const nome = sel ? sel.value : "";
  const qtdEl = document.getElementById("ins_qtd");
  const qtd = qtdEl ? (qtdEl.value || "0") : "0";
  const unidEl = document.getElementById("ins_unid");
  const unid = unidEl ? (unidEl.value || "un") : "un";
  if(!nome) return;

  const wrap = document.createElement("div");
  wrap.className = "badge bg-light text-dark badge-item p-2";
  wrap.innerHTML = `
    <strong>${nome}</strong> - ${qtd} ${unid}
    <button type="button" class="btn btn-sm btn-danger ms-2" onclick="this.parentElement.remove()">x</button>
    <input type="hidden" name="insumos_nome[]" value="${nome}">
    <input type="hidden" name="insumos_qtd[]" value="${qtd}">
    <input type="hidden" name="insumos_unid[]" value="${unid}">
  `;
  document.getElementById("ins_list").appendChild(wrap);

  if(tsIns) tsIns.clear(true);
  if(qtdEl) qtdEl.value = "1";
  if(unidEl) unidEl.value = "un";
}

document.addEventListener("DOMContentLoaded", () => {
  setServicoUnidade();

  const funcSelect = document.getElementById("func_select");
  const insSelect = document.getElementById("ins_select");

  if(funcSelect && !tsFunc){
    tsFunc = new TomSelect("#func_select", {
      create: false,
      placeholder: "Digite para buscar função...",
      allowEmptyOption: true,
      maxItems: 1,
      sortField: { field: "text", direction: "asc" }
    });
  }

  if(insSelect && !tsIns){
    tsIns = new TomSelect("#ins_select", {
      create: false,
      placeholder: "Digite para buscar insumo...",
      allowEmptyOption: true,
      maxItems: 1,
      sortField: { field: "text", direction: "asc" },
      onChange: function(){ updateInsumoUnidFromSelected(); }
    });
  }

  updateInsumoUnidFromSelected();
});
</script>

</body>
</html>
"""


# =============================
# LOGIN
# =============================
@app.route("/", methods=["GET", "POST"])
def login():
    msg = ""
    if request.method == "POST":
        email = request.form["email"].strip()
        senha = request.form["senha"].strip()

        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT id, nome, email, senha, tipo FROM usuarios WHERE email = %s", (email,))
        user = cur.fetchone()
        cur.close()
        conn.close()

        if user and check_password_hash(user[3], senha):
            session["user_id"] = user[0]
            session["nome"] = user[1]
            session["tipo"] = user[4]
            return redirect("/dashboard")
        msg = "⚠️ Login inválido!"

    conteudo = f"""
    <div class="card p-4">
      <h4>Login</h4>
      <form method="POST">
        <input class="form-control mb-2" name="email" placeholder="Email" required>
        <input class="form-control mb-2" name="senha" type="password" placeholder="Senha" required>
        <button class="btn btn-primary w-100">Entrar</button>
      </form>
      <p class="text-danger mt-2">{msg}</p>
    </div>
    """
    return render_template_string(base_html, conteudo=conteudo)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


# =============================
# DASHBOARD
# =============================
@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if "user_id" not in session:
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT nome FROM obras ORDER BY nome")
    obras = [o[0] for o in cur.fetchall()]

    cur.execute("SELECT nome, COALESCE(unidade,'un') FROM servicos ORDER BY nome")
    servicos = cur.fetchall()

    cur.execute("SELECT nome FROM funcoes ORDER BY nome")
    funcoes = [f[0] for f in cur.fetchall()]

    cur.execute("SELECT nome, COALESCE(unidade,'un') FROM insumos ORDER BY nome")
    insumos = cur.fetchall()

    if request.method == "POST":
        obra = request.form.get("obra", "").strip()
        servico = request.form.get("servico", "").strip()
        servico_unidade = request.form.get("servico_unidade", "").strip()
        observacao = request.form.get("observacao", "").strip()

        try:
            quantidade = float(request.form.get("quantidade", "0"))
        except ValueError:
            quantidade = 0.0

        try:
            horas = float(request.form.get("horas", "0"))
        except ValueError:
            horas = 0.0

        # Funções dict: {"Pedreiro": 2, ...}
        funcoes_nomes = request.form.getlist("funcoes_nome[]")
        funcoes_qtds = request.form.getlist("funcoes_qtd[]")
        funcoes_usadas = {}
        for nome, qtd in zip(funcoes_nomes, funcoes_qtds):
            nome = (nome or "").strip()
            try:
                q = int(qtd)
            except ValueError:
                q = 0
            if nome and q > 0:
                funcoes_usadas[nome] = funcoes_usadas.get(nome, 0) + q

        # Insumos dict: {"Cimento":{"quantidade":2,"unidade":"kg"}}
        insumos_nomes = request.form.getlist("insumos_nome[]")
        insumos_qtds = request.form.getlist("insumos_qtd[]")
        insumos_unids = request.form.getlist("insumos_unid[]")
        insumos_usados = {}
        for nome, qtd, unid in zip(insumos_nomes, insumos_qtds, insumos_unids):
            nome = (nome or "").strip()
            unid = (unid or "").strip() or "un"
            try:
                q = float(qtd)
            except ValueError:
                q = 0.0
            if nome and q > 0:
                if nome in insumos_usados:
                    insumos_usados[nome]["quantidade"] = float(insumos_usados[nome]["quantidade"]) + q
                    insumos_usados[nome]["unidade"] = unid
                else:
                    insumos_usados[nome] = {"quantidade": q, "unidade": unid}

        if obra and servico and servico_unidade and quantidade > 0 and horas > 0:
            data = datetime.now().strftime("%d/%m/%Y")
            cur.execute(
                """INSERT INTO produtividade
                   (obra, servico, servico_unidade, quantidade, horas, funcoes, insumos, observacao, data, user_id)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (
                    obra, servico, servico_unidade, quantidade, horas,
                    psycopg2.extras.Json(funcoes_usadas),
                    psycopg2.extras.Json(insumos_usados),
                    observacao, data, session["user_id"]
                )
            )
            conn.commit()

    cur.close()
    conn.close()

    obras_opt = "".join([f'<option value="{o}">{o}</option>' for o in obras])

    servicos_opt = ""
    for s, un in servicos:
        servicos_opt += f'<option value="{s}" data-unidade="{un}">{s} ({un})</option>'

    funcoes_opt = "".join([f'<option value="{f}">{f}</option>' for f in funcoes])

    insumos_opt = ""
    for nome, unid in insumos:
        insumos_opt += f'<option value="{nome}" data-unidade="{unid}">{nome} ({unid})</option>'

    export_btn = ""
    admin_links = ""
    if session.get("tipo") == "admin":
        export_btn = '<a href="/exportar" class="btn btn-secondary mb-2">Exportar Excel (Power BI)</a>'
        admin_links = """
        <hr>
        <a href="/criar_usuario" class="btn btn-warning mb-2">Criar Usuário</a>
        <a href="/cadastros" class="btn btn-info mb-2">Cadastros</a>
        <a href="/usuarios" class="btn btn-danger mb-2">Gerenciar Usuários</a>
        """

    conteudo = f"""
    <p><strong>Logado como:</strong> {session['nome']} ({session['tipo']}) | <a href="/logout">Sair</a></p>

    <div class="card p-4">
      <h4>Registrar Produtividade</h4>
      <form method="POST">

        <select class="form-control mb-2" name="obra" required>
          <option value="">Selecione obra</option>
          {obras_opt}
        </select>

        <select class="form-control mb-2" name="servico" id="servico_select" required onchange="setServicoUnidade()">
          <option value="">Selecione serviço</option>
          {servicos_opt}
        </select>

        <input class="form-control mb-2" name="servico_unidade" id="servico_unidade"
               placeholder="Unidade do serviço" readonly required>

        <input class="form-control mb-2" name="quantidade" type="number" step="0.01"
               placeholder="Quantidade executada" required>

        <input class="form-control mb-3" name="horas" type="number" step="0.1"
               placeholder="Horas trabalhadas" required>

        <div class="line">
          <h5>Funções (autocomplete)</h5>
          <div class="row g-2">
            <div class="col-md-6">
              <select class="form-control" id="func_select">
                <option value="">Selecione / digite...</option>
                {funcoes_opt}
              </select>
            </div>
            <div class="col-md-4">
              <input class="form-control" type="number" id="func_qtd" min="1" step="1" placeholder="Qtd" value="1">
            </div>
            <div class="col-md-2">
              <button type="button" class="btn btn-primary w-100" onclick="addFunc()">Adicionar</button>
            </div>
          </div>
          <div class="mt-2 mini">Itens adicionados:</div>
          <div id="func_list" class="mt-1"></div>
        </div>

        <div class="line">
          <h5>Insumos (autocomplete)</h5>
          <div class="row g-2">
            <div class="col-md-6">
              <select class="form-control" id="ins_select">
                <option value="">Selecione / digite...</option>
                {insumos_opt}
              </select>
            </div>
            <div class="col-md-3">
              <input class="form-control" type="number" id="ins_qtd" min="0" step="0.01" placeholder="Qtd" value="1">
            </div>
            <div class="col-md-1">
              <input class="form-control" type="text" id="ins_unid" placeholder="un" readonly>
            </div>
            <div class="col-md-2">
              <button type="button" class="btn btn-primary w-100" onclick="addIns()">Adicionar</button>
            </div>
          </div>
          <div class="mt-2 mini">Itens adicionados:</div>
          <div id="ins_list" class="mt-1"></div>
        </div>

        <textarea class="form-control mb-2" name="observacao" placeholder="Observações"></textarea>
        <button class="btn btn-success w-100">Salvar</button>
      </form>
    </div>

    {export_btn}
    {admin_links}
    """
    return render_template_string(base_html, conteudo=conteudo)


# =============================
# ADMIN - CRIAR USUÁRIO
# =============================
@app.route("/criar_usuario", methods=["GET", "POST"])
def criar_usuario():
    if "user_id" not in session or session.get("tipo") != "admin":
        return redirect("/")

    msg = ""
    if request.method == "POST":
        nome = request.form["nome"].strip()
        email = request.form["email"].strip()
        senha = request.form["senha"].strip()
        tipo = request.form["tipo"]
        senha_hash = generate_password_hash(senha)

        conn = get_db()
        cur = conn.cursor()
        try:
            cur.execute(
                "INSERT INTO usuarios (nome,email,senha,tipo) VALUES (%s,%s,%s,%s)",
                (nome, email, senha_hash, tipo)
            )
            conn.commit()
            msg = "✅ Usuário criado!"
        except Exception:
            conn.rollback()
            msg = "⚠️ Email já cadastrado!"
        cur.close()
        conn.close()

    conteudo = f"""
    <p><a href="/dashboard">← Voltar</a></p>
    <div class="card p-4">
      <h4>Criar Usuário</h4>
      <form method="POST">
        <input class="form-control mb-2" name="nome" placeholder="Nome" required>
        <input class="form-control mb-2" name="email" placeholder="Email" type="email" required>
        <input class="form-control mb-2" name="senha" placeholder="Senha" type="password" required>
        <select class="form-control mb-2" name="tipo">
          <option value="usuario">Usuário</option>
          <option value="admin">Administrador</option>
        </select>
        <button class="btn btn-warning w-100">Criar</button>
      </form>
      <p class="text-success mt-2">{msg}</p>
    </div>
    """
    return render_template_string(base_html, conteudo=conteudo)


# =============================
# ADMIN - CADASTROS
# =============================
@app.route("/cadastros", methods=["GET", "POST"])
def cadastros():
    if "user_id" not in session or session.get("tipo") != "admin":
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    msg = ""

    if request.method == "POST":
        if request.form.get("obra", "").strip():
            try:
                cur.execute("INSERT INTO obras (nome) VALUES (%s)", (request.form["obra"].strip(),))
                conn.commit()
                msg = "✅ Obra cadastrada!"
            except Exception:
                conn.rollback()
                msg = "⚠️ Obra já existe."

        if request.form.get("servico", "").strip() and request.form.get("unidade", "").strip():
            try:
                cur.execute("INSERT INTO servicos (nome,unidade) VALUES (%s,%s)",
                            (request.form["servico"].strip(), request.form["unidade"].strip()))
                conn.commit()
                msg = "✅ Serviço cadastrado!"
            except Exception:
                conn.rollback()
                msg = "⚠️ Serviço já existe."

        if request.form.get("insumo", "").strip() and request.form.get("unidade_insumo", "").strip():
            try:
                cur.execute("INSERT INTO insumos (nome,unidade) VALUES (%s,%s)",
                            (request.form["insumo"].strip(), request.form["unidade_insumo"].strip()))
                conn.commit()
                msg = "✅ Insumo cadastrado!"
            except Exception:
                conn.rollback()
                msg = "⚠️ Insumo já existe."

        if request.form.get("funcao", "").strip():
            try:
                cur.execute("INSERT INTO funcoes (nome) VALUES (%s)", (request.form["funcao"].strip(),))
                conn.commit()
                msg = "✅ Função cadastrada!"
            except Exception:
                conn.rollback()
                msg = "⚠️ Função já existe."

    cur.execute("SELECT nome FROM obras ORDER BY nome")
    obras = [o[0] for o in cur.fetchall()]
    cur.execute("SELECT nome, COALESCE(unidade,'un') FROM servicos ORDER BY nome")
    servicos = cur.fetchall()
    cur.execute("SELECT nome, COALESCE(unidade,'un') FROM insumos ORDER BY nome")
    insumos = cur.fetchall()
    cur.execute("SELECT nome FROM funcoes ORDER BY nome")
    funcoes = [f[0] for f in cur.fetchall()]

    cur.close()
    conn.close()

    conteudo = f"""
    <p><a href="/dashboard">← Voltar</a></p>

    <div class="card p-4">
      <h4>Obras</h4>
      <form method="POST" class="mb-2">
        <input class="form-control mb-2" name="obra" placeholder="Nome da obra" required>
        <button class="btn btn-info w-100">Adicionar</button>
      </form>
      <ul>{"".join([f"<li>{o}</li>" for o in obras])}</ul>
    </div>

    <div class="card p-4">
      <h4>Serviços</h4>
      <form method="POST" class="mb-2">
        <input class="form-control mb-2" name="servico" placeholder="Nome do serviço" required>
        <input class="form-control mb-2" name="unidade" placeholder="Unidade (ex: m², m³, un, kg)" required>
        <button class="btn btn-info w-100">Adicionar</button>
      </form>
      <ul>{"".join([f"<li>{s[0]} ({s[1]})</li>" for s in servicos])}</ul>
    </div>

    <div class="card p-4">
      <h4>Insumos</h4>
      <form method="POST" class="mb-2">
        <input class="form-control mb-2" name="insumo" placeholder="Nome do insumo" required>
        <input class="form-control mb-2" name="unidade_insumo" placeholder="Unidade (ex: kg, m, un, saco)" required>
        <button class="btn btn-info w-100">Adicionar</button>
      </form>
      <ul>{"".join([f"<li>{i[0]} ({i[1]})</li>" for i in insumos])}</ul>
    </div>

    <div class="card p-4">
      <h4>Funções</h4>
      <form method="POST" class="mb-2">
        <input class="form-control mb-2" name="funcao" placeholder="Nome da função (ex: Armador)" required>
        <button class="btn btn-info w-100">Adicionar</button>
      </form>
      <ul>{"".join([f"<li>{f}</li>" for f in funcoes])}</ul>
    </div>

    <p class="text-success">{msg}</p>
    """
    return render_template_string(base_html, conteudo=conteudo)


# =============================
# ADMIN - USUÁRIOS
# =============================
@app.route("/usuarios", methods=["GET"])
def usuarios():
    if "user_id" not in session or session.get("tipo") != "admin":
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, nome, email, tipo FROM usuarios ORDER BY id")
    users = cur.fetchall()
    cur.close()
    conn.close()

    rows = ""
    for uid, nome, email, tipo in users:
        disabled = "disabled" if uid == session["user_id"] else ""
        rows += f"""
        <tr>
          <td>{uid}</td>
          <td>{nome}</td>
          <td>{email}</td>
          <td>{tipo}</td>
          <td>
            <form method="POST" action="/usuarios/excluir" style="display:inline;">
              <input type="hidden" name="id" value="{uid}">
              <button class="btn btn-sm btn-danger" {disabled}
                      onclick="return confirm('Excluir este usuário?');">Excluir</button>
            </form>
          </td>
        </tr>
        """

    conteudo = f"""
    <p><a href="/dashboard">← Voltar</a></p>
    <div class="card p-4">
      <h4>Gerenciar Usuários</h4>
      <p class="mini">Você pode excluir usuários (não pode excluir a si mesmo).</p>
      <table class="table table-striped">
        <thead><tr><th>ID</th><th>Nome</th><th>Email</th><th>Tipo</th><th>Ações</th></tr></thead>
        <tbody>{rows}</tbody>
      </table>
    </div>
    """
    return render_template_string(base_html, conteudo=conteudo)


@app.route("/usuarios/excluir", methods=["POST"])
def usuarios_excluir():
    if "user_id" not in session or session.get("tipo") != "admin":
        return redirect("/")

    try:
        uid = int(request.form.get("id", "0"))
    except ValueError:
        uid = 0

    if uid <= 0 or uid == session["user_id"]:
        return redirect("/usuarios")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM usuarios WHERE id = %s", (uid,))
    conn.commit()
    cur.close()
    conn.close()
    return redirect("/usuarios")


# =============================
# EXPORTAR (1 ABA, POWER BI)
# =============================
@app.route("/exportar")
def exportar():
    if "user_id" not in session or session.get("tipo") != "admin":
        return redirect("/dashboard")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
      SELECT p.id, p.obra, p.servico, p.servico_unidade, p.quantidade, p.horas,
             p.funcoes, p.insumos, p.observacao, p.data, p.user_id,
             u.nome
      FROM produtividade p
      LEFT JOIN usuarios u ON u.id = p.user_id
      ORDER BY p.id DESC
    """)
    dados = cur.fetchall()
    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"

    ws.append([
        "registro_id",
        "data",
        "obra",
        "servico",
        "servico_unidade",
        "quantidade_servico",
        "horas",
        "horas_por_unidade_servico",
        "observacao",
        "user_id",
        "usuario_nome",
        "tipo_item",
        "item_nome",
        "item_qtd",
        "item_unidade",
        "homem_hora"
    ])

    for row in dados:
        (rid, obra, servico, serv_unid, qtd_serv, horas,
         funcoes_raw, insumos_raw, obs, data, user_id, user_nome) = row

        try:
            horas_por_unid = (float(horas) / float(qtd_serv)) if float(qtd_serv) != 0 else 0
        except Exception:
            horas_por_unid = 0

        funcoes = safe_json_load(funcoes_raw) or {}
        insumos_data = safe_json_load(insumos_raw) or {}

        wrote_any = False

        if isinstance(funcoes, dict) and len(funcoes) > 0:
            for funcao, qtd in funcoes.items():
                try:
                    qtd_int = int(qtd)
                except Exception:
                    qtd_int = 0
                if not funcao or qtd_int <= 0:
                    continue
                try:
                    hh = float(horas) * float(qtd_int)
                except Exception:
                    hh = 0

                ws.append([
                    rid, data, obra, servico, serv_unid, qtd_serv, horas, horas_por_unid, obs,
                    user_id, user_nome or "",
                    "FUNCAO",
                    str(funcao),
                    qtd_int,
                    "",
                    hh
                ])
                wrote_any = True

        if isinstance(insumos_data, dict) and len(insumos_data) > 0:
            for insumo_nome, val in insumos_data.items():
                if not insumo_nome:
                    continue

                if isinstance(val, dict):
                    qtd = val.get("quantidade", "")
                    unid = val.get("unidade", "")
                else:
                    s = str(val).strip()
                    parts = s.split()
                    qtd = parts[0] if parts else ""
                    unid = " ".join(parts[1:]) if len(parts) > 1 else ""

                try:
                    qtd_num = float(qtd)
                except Exception:
                    qtd_num = 0.0
                if qtd_num <= 0:
                    continue

                ws.append([
                    rid, data, obra, servico, serv_unid, qtd_serv, horas, horas_por_unid, obs,
                    user_id, user_nome or "",
                    "INSUMO",
                    str(insumo_nome),
                    qtd_num,
                    str(unid),
                    ""
                ])
                wrote_any = True

        if not wrote_any:
            ws.append([
                rid, data, obra, servico, serv_unid, qtd_serv, horas, horas_por_unid, obs,
                user_id, user_nome or "",
                "REGISTRO",
                "",
                "",
                "",
                ""
            ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="produtividade_powerbi.xlsx", as_attachment=True)


# =============================
# RUN
# =============================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)


