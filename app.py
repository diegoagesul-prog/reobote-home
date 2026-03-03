# app.py — Reobote Home (Produtividade + Permissões por Obra + Custo de Efetivo (Mês) + PDF)
from flask import (
    Flask,
    render_template_string,
    request,
    redirect,
    session,
    send_file,
    abort,
    Response,
)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from datetime import datetime
import os
from io import BytesIO
import json
import secrets
import psycopg2
import psycopg2.extras
from psycopg2.pool import SimpleConnectionPool

# PDF (server-side)
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "reobote_home_secret")

DB_POOL = None  # pool global


# -------------------------
# DB / Pool
# -------------------------
def _normalize_db_url(db_url: str) -> str:
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
    return db_url


def get_db():
    """Pega uma conexão do pool (reutiliza conexões)."""
    global DB_POOL
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL nao configurada.")
    db_url = _normalize_db_url(db_url)

    if DB_POOL is None:
        DB_POOL = SimpleConnectionPool(1, 10, db_url)

    return DB_POOL.getconn()


def put_db(conn):
    """Devolve conexão pro pool."""
    global DB_POOL
    if DB_POOL and conn:
        DB_POOL.putconn(conn)


# -------------------------
# Helpers
# -------------------------
def safe_json_load(x):
    if x is None:
        return {}
    if isinstance(x, (dict, list)):
        return x
    try:
        return json.loads(x)
    except:
        return {}


def csrf_token():
    tok = session.get("_csrf")
    if not tok:
        tok = secrets.token_urlsafe(32)
        session["_csrf"] = tok
    return tok


def require_csrf():
    if request.form.get("_csrf", "") != session.get("_csrf", ""):
        abort(400)


def is_admin():
    return session.get("tipo") == "admin"


def require_login():
    return "user_id" in session


def password_policy_ok(pw):
    pw = pw or ""
    if len(pw) < 8:
        return False, "Minimo 8 caracteres."
    if not any(c.isalpha() for c in pw) or not any(c.isdigit() for c in pw):
        return False, "Pelo menos 1 letra e 1 numero."
    return True, ""


def fmt_num(x, dec=2):
    """Formata numero removendo zeros à direita. Se for None -> ''. Se for 0 -> ''. """
    if x is None:
        return ""
    try:
        v = float(x)
    except:
        return ""
    if v == 0:
        return ""
    s = f"{v:.{dec}f}".rstrip("0").rstrip(".")
    return s


def parse_month(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    if len(s) == 7 and s[4] == "-" and s[:4].isdigit() and s[5:].isdigit():
        y = int(s[:4])
        m = int(s[5:])
        if 1 <= m <= 12 and 2000 <= y <= 2100:
            return f"{y:04d}-{m:02d}"
        return ""
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            m = int(parts[0])
            y = int(parts[1])
            if 1 <= m <= 12 and 2000 <= y <= 2100:
                return f"{y:04d}-{m:02d}"
    return ""


# -------------------------
# Permissões por Obra
# -------------------------
def get_all_obras():
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT id,nome FROM obras ORDER BY nome")
        return cur.fetchall()
    finally:
        cur.close()
        put_db(conn)


def get_allowed_obras(user_id: int):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT o.id,o.nome FROM obras o "
            "JOIN usuario_obras uo ON uo.obra_id=o.id "
            "WHERE uo.user_id=%s ORDER BY o.nome",
            (user_id,),
        )
        return cur.fetchall()
    finally:
        cur.close()
        put_db(conn)


def user_can_use_obra_name(user_id: int, obra_nome: str) -> bool:
    obra_nome = (obra_nome or "").strip()
    if not obra_nome:
        return False
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT 1 FROM usuario_obras uo "
            "JOIN obras o ON o.id=uo.obra_id "
            "WHERE uo.user_id=%s AND o.nome=%s",
            (user_id, obra_nome),
        )
        return cur.fetchone() is not None
    finally:
        cur.close()
        put_db(conn)


def user_can_use_obra_id(user_id: int, obra_id: int) -> bool:
    try:
        obra_id = int(obra_id)
    except:
        return False
    if obra_id <= 0:
        return False
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT 1 FROM usuario_obras WHERE user_id=%s AND obra_id=%s",
            (user_id, obra_id),
        )
        return cur.fetchone() is not None
    finally:
        cur.close()
        put_db(conn)


# -------------------------
# DB init
# -------------------------
def init_db():
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "CREATE TABLE IF NOT EXISTS usuarios ("
            "id SERIAL PRIMARY KEY, "
            "nome TEXT, "
            "email TEXT UNIQUE, "
            "senha TEXT, "
            "tipo TEXT"
            ");"
        )
        cur.execute("CREATE TABLE IF NOT EXISTS obras (id SERIAL PRIMARY KEY, nome TEXT UNIQUE);")
        cur.execute("CREATE TABLE IF NOT EXISTS servicos (id SERIAL PRIMARY KEY, nome TEXT UNIQUE, unidade TEXT);")
        cur.execute("CREATE TABLE IF NOT EXISTS insumos (id SERIAL PRIMARY KEY, nome TEXT UNIQUE, unidade TEXT);")
        cur.execute("CREATE TABLE IF NOT EXISTS funcoes (id SERIAL PRIMARY KEY, nome TEXT UNIQUE);")

        cur.execute(
            "CREATE TABLE IF NOT EXISTS usuario_obras ("
            "id SERIAL PRIMARY KEY, "
            "user_id INTEGER, "
            "obra_id INTEGER, "
            "UNIQUE(user_id, obra_id)"
            ");"
        )

        cur.execute(
            "CREATE TABLE IF NOT EXISTS produtividade ("
            "id SERIAL PRIMARY KEY, "
            "obra TEXT, "
            "servico TEXT, "
            "servico_unidade TEXT, "
            "quantidade DOUBLE PRECISION, "
            "horas DOUBLE PRECISION, "
            "funcoes JSONB, "
            "insumos JSONB, "
            "observacao TEXT, "
            "data TEXT, "
            "user_id INTEGER"
            ");"
        )

        cur.execute(
            "CREATE TABLE IF NOT EXISTS custos_funcao_obra ("
            "id SERIAL PRIMARY KEY, "
            "obra_id INTEGER, "
            "funcao_id INTEGER, "
            "custo_mensal DOUBLE PRECISION, "
            "UNIQUE(obra_id, funcao_id)"
            ");"
        )

        cur.execute(
            "CREATE TABLE IF NOT EXISTS efetivo_planejamentos ("
            "id SERIAL PRIMARY KEY, "
            "obra_id INTEGER, "
            "user_id INTEGER, "
            "mes TEXT, "
            "itens JSONB, "
            "total DOUBLE PRECISION, "
            "created_at TEXT, "
            "updated_at TEXT, "
            "UNIQUE(user_id, obra_id, mes)"
            ");"
        )

        ae = os.environ.get("ADMIN_EMAIL")
        ap = os.environ.get("ADMIN_PASSWORD")
        if ae and ap:
            ok, _ = password_policy_ok(ap)
            if ok:
                cur.execute(
                    "INSERT INTO usuarios (nome,email,senha,tipo) VALUES (%s,%s,%s,%s) "
                    "ON CONFLICT (email) DO NOTHING;",
                    ("Administrador", ae.strip().lower(), generate_password_hash(ap), "admin"),
                )

        conn.commit()
    finally:
        cur.close()
        put_db(conn)


init_db()


# -------------------------
# Routes básicas / PWA
# -------------------------
@app.route("/ping")
def ping():
    return "ok", 200


@app.route("/sw.js")
def service_worker():
    return app.send_static_file("sw.js"), 200, {
        "Content-Type": "application/javascript",
        "Service-Worker-Allowed": "/",
    }


@app.route("/manifest.json")
def manifest():
    data = {
        "name": "Reobote Home",
        "short_name": "Reobote Produtividade",
        "description": "Registro de Produtividade e Consumo",
        "start_url": "/dashboard",
        "display": "standalone",
        "background_color": "#1a1a2e",
        "theme_color": "#1a1a2e",
        "icons": [
            {"src": "/static/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "/static/icon-512.png", "sizes": "512x512", "type": "image/png"},
        ],
    }
    return Response(json.dumps(data), mimetype="application/json")


# -------------------------
# UI Template
# -------------------------
BASE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Reobote Home</title>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<meta name="theme-color" content="#1a1a2e">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Reobote">
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/static/icon-192.png">
<link href="https://cdn.jsdelivr.net/npm/tom-select@2.3.1/dist/css/tom-select.css" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/tom-select@2.3.1/dist/js/tom-select.complete.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Rajdhani:wght@600;700&display=swap" rel="stylesheet">
<style>
:root{
  --gold:#C9933A;--gold-l:#E0B060;--gold-d:#8a6420;
  --dark:#1a1a2e;--dark2:#16213e;
  --bg:#f2f0ec;--card:#fff;
  --text:#1a1a2e;--muted:#6b7280;
  --border:#e5e7eb;
  --ok:#16a34a;--err:#dc2626;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);font-family:'Inter',sans-serif;color:var(--text);min-height:100vh;-webkit-font-smoothing:antialiased;}

/* HEADER */
.hdr{background:var(--dark);padding:6px 12px 0;position:sticky;top:0;z-index:200;border-bottom:2px solid var(--gold);}
.hdr-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;}
.hdr-brand{display:flex;align-items:center;gap:8px;flex-shrink:0;}
.hdr-logo{height:26px;width:auto;}
.hdr-name{font-family:'Rajdhani',sans-serif;font-weight:700;font-size:.9rem;color:var(--gold);letter-spacing:.07em;line-height:1;}
.hdr-sub{font-size:.5rem;color:rgba(255,255,255,.3);letter-spacing:.1em;text-transform:uppercase;display:block;margin-top:1px;}
.hdr-utils{display:flex;gap:2px;align-items:center;}
.hdr-uname{font-size:.65rem;font-weight:600;color:rgba(255,255,255,.6);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:130px;margin-right:6px;}
.hdr-utils a.util{font-size:.63rem;color:rgba(255,255,255,.4);text-decoration:none;font-weight:500;padding:3px 6px;white-space:nowrap;}
.hdr-utils a.util:hover{color:rgba(255,255,255,.8);}
.hdr-utils .hdr-sep{color:rgba(255,255,255,.15);font-size:.65rem;}
.hdr-tabs{display:flex;gap:6px;padding-bottom:8px;}
.hdr-tabs a.tab{flex:1;text-align:center;font-size:.75rem;color:var(--gold);text-decoration:none;font-weight:700;padding:7px 6px;border-radius:8px;border:1.5px solid var(--gold-d);background:rgba(201,147,58,.1);letter-spacing:.04em;white-space:nowrap;transition:background .15s,color .15s;}
.hdr-tabs a.tab:hover,.hdr-tabs a.tab.active{background:var(--gold);color:var(--dark);border-color:var(--gold);}

/* WRAP */
.wrap{max-width:540px;margin:0 auto;padding:14px 12px 72px;}

/* CARD */
.card{background:var(--card);border-radius:14px;padding:18px 15px;margin-bottom:13px;border:1px solid var(--border);box-shadow:0 1px 3px rgba(0,0,0,.06),0 6px 20px rgba(0,0,0,.04);}
.ctitle{font-family:'Rajdhani',sans-serif;font-weight:700;font-size:1rem;letter-spacing:.07em;text-transform:uppercase;color:var(--dark);margin-bottom:15px;display:flex;align-items:center;gap:8px;}
.ctitle::before{content:'';display:block;width:3px;height:16px;background:var(--gold);border-radius:2px;flex-shrink:0;}
.stitle{font-family:'Rajdhani',sans-serif;font-weight:700;font-size:.75rem;letter-spacing:.1em;text-transform:uppercase;color:var(--gold-d);margin:17px 0 9px;padding-top:15px;border-top:1px solid var(--border);}

/* FIELDS */
.fl{margin-bottom:10px;}
.fl-label{display:block;font-size:.7rem;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;}
input[type=text],input[type=email],input[type=password],input[type=number],input[type=month],select,textarea{
  width:100%;padding:12px 13px;border:1.5px solid var(--border);border-radius:10px;
  font-size:.97rem;font-family:'Inter',sans-serif;color:var(--text);background:#fff;
  transition:border-color .15s,box-shadow .15s;-webkit-appearance:none;appearance:none;}
input:focus,select:focus,textarea:focus{border-color:var(--gold);box-shadow:0 0 0 3px rgba(201,147,58,.15);outline:none;}
input[readonly],input[disabled]{background:#f9f7f3;color:var(--muted);}
textarea{resize:vertical;min-height:74px;}

/* TOM SELECT */
.ts-wrapper{width:100%;}
.ts-control{border:1.5px solid var(--border)!important;border-radius:10px!important;padding:11px 13px!important;font-size:.97rem!important;font-family:'Inter',sans-serif!important;min-height:47px!important;background:#fff!important;box-shadow:none!important;}
.ts-wrapper.focus .ts-control{border-color:var(--gold)!important;box-shadow:0 0 0 3px rgba(201,147,58,.15)!important;}
.ts-dropdown{border:1.5px solid var(--border)!important;border-radius:10px!important;font-family:'Inter',sans-serif!important;box-shadow:0 8px 24px rgba(0,0,0,.12)!important;margin-top:4px!important;}
.ts-dropdown .option{padding:11px 13px!important;font-size:.93rem!important;}
.ts-dropdown .option:hover,.ts-dropdown .option.active{background:#fdf5e8!important;color:var(--gold-d)!important;}

/* ADD ROW */
.add-row{display:flex;gap:8px;align-items:center;}
.add-row .ts-wrapper{flex:1 1 0;min-width:0;}
.qty-box{width:66px!important;flex-shrink:0;text-align:center;padding:12px 6px!important;}
.upill{flex-shrink:0;min-width:40px;height:47px;display:flex;align-items:center;justify-content:center;background:#fdf5e8;border:1.5px solid #f0d5a0;border-radius:10px;font-size:.73rem;font-weight:700;color:var(--gold-d);letter-spacing:.04em;padding:0 8px;}
.btn-plus{flex-shrink:0;width:47px;height:47px;border-radius:10px;border:none;background:var(--dark);color:var(--gold);font-size:1.5rem;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .15s;}
.btn-plus:hover{background:var(--dark2);}

/* TAGS */
.tags{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;min-height:4px;}
.tag{display:inline-flex;align-items:center;gap:5px;background:#fdf5e8;border:1px solid #f0d5a0;border-radius:20px;padding:5px 8px 5px 11px;font-size:.81rem;color:var(--gold-d);font-weight:600;}
.tag-x{width:19px;height:19px;border-radius:50%;border:none;background:#f0d5a0;color:var(--gold-d);font-size:.68rem;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;flex-shrink:0;}
.tag-x:hover{background:var(--gold);color:#fff;}

/* BUTTONS */
.btn{display:block;width:100%;padding:13px;border:none;border-radius:10px;font-family:'Rajdhani',sans-serif;font-weight:700;font-size:.97rem;letter-spacing:.07em;text-transform:uppercase;cursor:pointer;text-align:center;text-decoration:none;transition:filter .15s,transform .1s;}
.btn:active{transform:scale(.98);}
.btn:hover{filter:brightness(1.08);}
.btn-gold{background:var(--gold);color:var(--dark);}
.btn-dark{background:var(--dark);color:var(--gold);}
.btn-green{background:#16a34a;color:#fff;margin-top:6px;}
.btn-warn{background:#d97706;color:#fff;}
.btn-info{background:#0891b2;color:#fff;}
.btn-purple{background:#7c3aed;color:#fff;}
.btn-sm{display:inline-flex;align-items:center;padding:5px 11px;border-radius:8px;font-size:.76rem;font-weight:600;cursor:pointer;border:1.5px solid;background:transparent;font-family:'Inter',sans-serif;text-decoration:none;white-space:nowrap;}
.btn-edit{border-color:var(--gold);color:var(--gold-d);}
.btn-edit:hover{background:#fdf5e8;}
.btn-del{border-color:var(--err);color:var(--err);}
.btn-del:hover{background:#fef2f2;}

/* ALERTS */
.alert{padding:10px 13px;border-radius:10px;font-size:.87rem;font-weight:500;margin-bottom:12px;}
.a-ok{background:#dcfce7;border:1px solid #86efac;color:#15803d;}
.a-err{background:#fef9c3;border:1px solid #fde047;color:#854d0e;}
.a-danger{background:#fee2e2;border:1px solid #fca5a5;color:#991b1b;}

/* REG CARDS */
.rc{background:#fafaf8;border:1px solid var(--border);border-radius:10px;padding:12px 13px;margin-bottom:8px;}
.rc-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:3px;}
.rc-obra{font-weight:700;font-size:.91rem;color:var(--dark);}
.rc-date{font-size:.7rem;color:var(--muted);white-space:nowrap;}
.rc-srv{font-size:.81rem;color:var(--muted);margin-bottom:5px;line-height:1.3;}
.rc-nums{display:flex;gap:13px;font-size:.77rem;color:var(--muted);margin-bottom:7px;}
.rc-nums strong{color:var(--dark);font-weight:600;}
.rc-act{display:flex;gap:7px;align-items:center;flex-wrap:wrap;}

/* ADMIN GRID */
.agrid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:4px;}

/* LOGIN */
.login-wrap{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:24px 16px;background:var(--dark);}
.login-logo{height:52px;margin-bottom:6px;}
.login-sub{font-family:'Rajdhani',sans-serif;font-size:.7rem;letter-spacing:.18em;color:rgba(255,255,255,.3);text-transform:uppercase;margin-bottom:28px;}
.login-card{background:#fff;border-radius:16px;padding:26px 22px;width:100%;max-width:350px;box-shadow:0 20px 60px rgba(0,0,0,.4);}
.login-card h4{font-family:'Rajdhani',sans-serif;font-size:1.15rem;font-weight:700;letter-spacing:.06em;color:var(--dark);margin-bottom:18px;text-transform:uppercase;}

.back{display:inline-flex;align-items:center;gap:5px;font-size:.81rem;font-weight:600;color:var(--gold-d);text-decoration:none;margin-bottom:12px;}
.back:hover{color:var(--gold);}
.mini{font-size:.74rem;color:var(--muted);}
.clist{list-style:none;padding:0;margin-top:9px;}
.clist li{padding:8px 11px;background:#fafaf8;border:1px solid var(--border);border-radius:8px;font-size:.87rem;margin-bottom:5px;}
.utable{width:100%;border-collapse:collapse;font-size:.83rem;}
.utable th{padding:7px 9px;text-align:left;font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);border-bottom:2px solid var(--border);}
.utable td{padding:9px;border-bottom:1px solid var(--border);vertical-align:middle;}
.utable tr:last-child td{border-bottom:none;}
.acc{margin-top:10px;}
.acc-btn{width:100%;display:flex;justify-content:space-between;align-items:center;padding:9px 12px;background:#fafaf8;border:1px solid var(--border);border-radius:8px;cursor:pointer;font-size:.84rem;font-weight:600;color:var(--text);font-family:'Inter',sans-serif;}
.acc-btn:hover{background:#fdf5e8;border-color:#f0d5a0;}
.acc-arr{font-size:.7rem;color:var(--muted);transition:transform .2s;}
.acc-arr.open{transform:rotate(180deg);}
.acc-body{padding-top:4px;}

/* EFETIVO */
.line{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
.kpi{display:flex;gap:12px;flex-wrap:wrap;}
.kpi .box{background:#fafaf8;border:1px solid var(--border);border-radius:10px;padding:10px 12px;min-width:140px;}
.kpi .box .lbl{font-size:.68rem;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);font-weight:700;}
.kpi .box .val{margin-top:4px;font-weight:800;color:var(--dark);font-size:1.05rem;}
</style>
</head>
<body>
{% if show_header %}
<header class="hdr">
  <div class="hdr-top">
    <div class="hdr-brand">
      <img src="/static/logo.png" class="hdr-logo" alt="Reobote">
      <div>
        <div class="hdr-name">REOBOTE HOME</div>
        <span class="hdr-sub">Produtividade &amp; Consumo</span>
      </div>
    </div>
    <div class="hdr-utils">
      <span class="hdr-uname">{{ session.get("nome","") }}</span>
      <a href="/minha_senha" class="util">Senha</a>
      <span class="hdr-sep">|</span>
      <a href="/logout" class="util">Sair</a>
    </div>
  </div>
  <div class="hdr-tabs">
    <a href="/dashboard" class="tab {% if current_page=='produtividade' %}active{% endif %}">Produtividade</a>
    <a href="/efetivo" class="tab {% if current_page=='efetivo' %}active{% endif %}">Custo de Efetivo</a>
  </div>
</header>
{% endif %}
<div class="{% if show_header %}wrap{% else %}login-wrap{% endif %}">
  {{ conteudo|safe }}
</div>
<script>
function togAcc(id){
  const el=document.getElementById(id);
  if(!el)return;
  const body=el.querySelector('.acc-body');
  const arr=el.querySelector('.acc-arr');
  if(body.style.display==='none'){body.style.display='block';arr.classList.add('open');}
  else{body.style.display='none';arr.classList.remove('open');}
}

let tsS=null,tsF=null,tsI=null;
function setSrvUnid(){
  const s=document.getElementById("servico_select"),u=document.getElementById("servico_unidade"),b=document.getElementById("sb");
  if(!s||!u)return;
  const o=s.options[s.selectedIndex],v=(o&&o.dataset&&o.dataset.unidade)?o.dataset.unidade:"";
  u.value=v;if(b)b.textContent=v||"—";
}
function updInsUnid(){
  const s=document.getElementById("ins_select"),b=document.getElementById("ib");
  if(!s||!b)return;
  const o=Array.from(s.options).find(x=>x.value===s.value);
  b.textContent=(o&&o.dataset&&o.dataset.unidade)?o.dataset.unidade:"un";
}
function addFunc(){
  const s=document.getElementById("func_select"),n=s?s.value:"",q=document.getElementById("fq");
  const v=q?(q.value||"0"):"0";if(!n)return;
  const t=document.createElement("div");t.className="tag";
  t.innerHTML=`<span>${n} &mdash; ${v}</span><button type="button" class="tag-x" onclick="this.closest('.tag').remove()">x</button><input type="hidden" name="funcoes_nome[]" value="${n}"><input type="hidden" name="funcoes_qtd[]" value="${v}">`;
  document.getElementById("fl").appendChild(t);
  if(tsF)tsF.clear(true);if(q)q.value="1";
}
function addIns(){
  const s=document.getElementById("ins_select"),n=s?s.value:"";
  const q=document.getElementById("iq"),b=document.getElementById("ib");
  const v=q?(q.value||"0"):"0",u=b?(b.textContent||"un"):"un";
  if(!n)return;
  const t=document.createElement("div");t.className="tag";
  t.innerHTML=`<span>${n} &mdash; ${v} ${u}</span><button type="button" class="tag-x" onclick="this.closest('.tag').remove()">x</button><input type="hidden" name="insumos_nome[]" value="${n}"><input type="hidden" name="insumos_qtd[]" value="${v}"><input type="hidden" name="insumos_unid[]" value="${u}">`;
  document.getElementById("il").appendChild(t);
  if(tsI)tsI.clear(true);if(q)q.value="1";
}

function fmtMoneyBR(v){
  try{
    const n=Number(v||0);
    return n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
  }catch(e){ return 'R$ 0,00'; }
}
function recalcEfetivo(){
  const rows=document.querySelectorAll('[data-ef-row]');
  let total=0;
  rows.forEach(r=>{
    const custo=Number(r.getAttribute('data-custo')||0);
    const inp=r.querySelector('input[name="qtd[]"]');
    const qtd=Number(inp?inp.value:0)||0;
    const sub=custo*qtd;
    const out=r.querySelector('[data-sub]');
    if(out) out.textContent = fmtMoneyBR(sub);
    total += sub;
  });
  const t=document.getElementById('ef_total');
  if(t) t.textContent = fmtMoneyBR(total);
}

document.addEventListener("DOMContentLoaded",()=>{
  setSrvUnid();
  const ss=document.getElementById("servico_select");
  const fs=document.getElementById("func_select");
  const is_=document.getElementById("ins_select");
  if(ss&&!tsS)tsS=new TomSelect("#servico_select",{create:false,placeholder:"Buscar servico...",allowEmptyOption:true,maxItems:1,sortField:{field:"text",direction:"asc"},onChange:function(){setSrvUnid();}});
  if(fs&&!tsF)tsF=new TomSelect("#func_select",{create:false,placeholder:"Buscar funcao...",allowEmptyOption:true,maxItems:1,sortField:{field:"text",direction:"asc"}});
  if(is_&&!tsI)tsI=new TomSelect("#ins_select",{create:false,placeholder:"Buscar insumo...",allowEmptyOption:true,maxItems:1,sortField:{field:"text",direction:"asc"},onChange:function(){updInsUnid();}});
  updInsUnid();

  const ef=document.getElementById('ef_form');
  if(ef){
    ef.addEventListener('input', (e)=>{
      if(e.target && e.target.name==='qtd[]') recalcEfetivo();
    });
    recalcEfetivo();
  }

  if('serviceWorker' in navigator){
    navigator.serviceWorker.register('/sw.js');
  }
});
</script>
</body>
</html>"""


def render(conteudo, show_header=True, page=""):
    return render_template_string(BASE, conteudo=conteudo, show_header=show_header, session=session, current_page=page)


def alert(msg):
    if not msg:
        return ""
    if ":" not in msg:
        return f'<div class="alert a-err">⚠ {msg}</div>'
    tipo, texto = msg.split(":", 1)
    cls = "a-ok" if tipo == "ok" else "a-err"
    icon = "✓" if tipo == "ok" else "⚠"
    return f'<div class="alert {cls}">{icon} {texto}</div>'


# -------------------------
# Auth
# -------------------------
@app.route("/", methods=["GET", "POST"])
def login():
    msg = ""
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        senha = request.form.get("senha", "").strip()
        conn = get_db()
        cur = conn.cursor()
        try:
            cur.execute("SELECT id,nome,email,senha,tipo FROM usuarios WHERE email=%s", (email,))
            user = cur.fetchone()
        finally:
            cur.close()
            put_db(conn)

        if user and check_password_hash(user[3], senha):
            session["user_id"] = user[0]
            session["nome"] = user[1] or ""
            session["tipo"] = user[4] or "usuario"
            csrf_token()
            return redirect("/dashboard")
        msg = "E-mail ou senha incorretos."

    c = f"""<img src="/static/logo.png" class="login-logo" alt="Reobote">
    <div class="login-sub">Produtividade &amp; Consumo</div>
    <div class="login-card">
      <h4>Acesso</h4>
      {"<div class='alert a-danger'>"+msg+"</div>" if msg else ""}
      <form method="POST">
        <div class="fl"><label class="fl-label">E-mail</label><input type="email" name="email" placeholder="seu@email.com" required></div>
        <div class="fl"><label class="fl-label">Senha</label><input type="password" name="senha" placeholder="••••••••" required></div>
        <button class="btn btn-gold" style="margin-top:6px;">Entrar</button>
      </form>
      <div class="mini" style="margin-top:11px;text-align:center;">Troque sua senha apos o primeiro acesso.</div>
    </div>"""
    return render(c, show_header=False)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/minha_senha", methods=["GET", "POST"])
def minha_senha():
    if not require_login():
        return redirect("/")
    msg = ""
    if request.method == "POST":
        require_csrf()
        sa = request.form.get("senha_atual", "").strip()
        n = request.form.get("nova_senha", "").strip()
        n2 = request.form.get("nova_senha2", "").strip()
        if n != n2:
            msg = "err:As senhas nao conferem."
        else:
            ok, m = password_policy_ok(n)
            if not ok:
                msg = f"err:{m}"
            else:
                conn = get_db()
                cur = conn.cursor()
                try:
                    cur.execute("SELECT senha FROM usuarios WHERE id=%s", (session["user_id"],))
                    row = cur.fetchone()
                    if not row or not check_password_hash(row[0], sa):
                        msg = "err:Senha atual incorreta."
                    else:
                        cur.execute(
                            "UPDATE usuarios SET senha=%s WHERE id=%s",
                            (generate_password_hash(n), session["user_id"]),
                        )
                        conn.commit()
                        msg = "ok:Senha alterada!"
                finally:
                    cur.close()
                    put_db(conn)

    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card">
      <div class="ctitle">Minha Senha</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div class="fl"><label class="fl-label">Senha atual</label><input type="password" name="senha_atual" required></div>
        <div class="fl"><label class="fl-label">Nova senha</label><input type="password" name="nova_senha" required></div>
        <div class="fl"><label class="fl-label">Confirmar nova senha</label><input type="password" name="nova_senha2" required></div>
        <div class="mini" style="margin-bottom:11px;">Minimo 8 caracteres, pelo menos 1 letra e 1 numero.</div>
        <button class="btn btn-gold">Salvar</button>
      </form>
    </div>"""
    return render(c)


# -------------------------
# Dropdowns / Permissões
# -------------------------
def load_dropdowns_admin():
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT nome FROM obras ORDER BY nome")
        obras = [o[0] for o in cur.fetchall()]
        cur.execute("SELECT nome,COALESCE(unidade,'un') FROM servicos ORDER BY nome")
        servicos = cur.fetchall()
        cur.execute("SELECT nome FROM funcoes ORDER BY nome")
        funcoes = [f[0] for f in cur.fetchall()]
        cur.execute("SELECT nome,COALESCE(unidade,'un') FROM insumos ORDER BY nome")
        insumos = cur.fetchall()
        return obras, servicos, funcoes, insumos
    finally:
        cur.close()
        put_db(conn)


def load_dropdowns_user(user_id: int):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT o.nome FROM obras o "
            "JOIN usuario_obras uo ON uo.obra_id=o.id "
            "WHERE uo.user_id=%s ORDER BY o.nome",
            (user_id,),
        )
        obras = [o[0] for o in cur.fetchall()]
        cur.execute("SELECT nome,COALESCE(unidade,'un') FROM servicos ORDER BY nome")
        servicos = cur.fetchall()
        cur.execute("SELECT nome FROM funcoes ORDER BY nome")
        funcoes = [f[0] for f in cur.fetchall()]
        cur.execute("SELECT nome,COALESCE(unidade,'un') FROM insumos ORDER BY nome")
        insumos = cur.fetchall()
        return obras, servicos, funcoes, insumos
    finally:
        cur.close()
        put_db(conn)


def can_edit(uid):
    if is_admin():
        return True
    return int(uid or 0) == int(session.get("user_id", 0))


def form_funcoes_insumos(funcoes, insumos):
    fo = "".join([f'<option value="{f}">{f}</option>' for f in funcoes])
    io = "".join([f'<option value="{n}" data-unidade="{u}">{n} ({u})</option>' for n, u in insumos])
    return fo, io


# -------------------------
# Dashboard (Produtividade)
# -------------------------
@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if not require_login():
        return redirect("/")

    msg = ""
    if is_admin():
        obras, servicos, funcoes, insumos = load_dropdowns_admin()
    else:
        obras, servicos, funcoes, insumos = load_dropdowns_user(session["user_id"])

    conn = get_db()
    cur = conn.cursor()
    try:
        if request.method == "POST":
            require_csrf()
            obra = request.form.get("obra", "").strip()
            servico = request.form.get("servico", "").strip()
            su = request.form.get("servico_unidade", "").strip()
            obs = request.form.get("observacao", "").strip()

            try:
                qtd = float(request.form.get("quantidade", "0"))
            except:
                qtd = 0.0
            try:
                hrs = float(request.form.get("horas", "0"))
            except:
                hrs = 0.0

            if not is_admin():
                if not user_can_use_obra_name(session["user_id"], obra):
                    msg = "err:Voce nao tem permissao para registrar nessa obra."

            fn = request.form.getlist("funcoes_nome[]")
            fq = request.form.getlist("funcoes_qtd[]")
            fu = {}
            for n, q in zip(fn, fq):
                n = (n or "").strip()
                try:
                    qi = int(q)
                except:
                    qi = 0
                if n and qi > 0:
                    fu[n] = fu.get(n, 0) + qi

            ins_n = request.form.getlist("insumos_nome[]")
            ins_q = request.form.getlist("insumos_qtd[]")
            ins_u = request.form.getlist("insumos_unid[]")
            iu = {}
            for n, q, u in zip(ins_n, ins_q, ins_u):
                n = (n or "").strip()
                u = (u or "").strip() or "un"
                try:
                    qf = float(q)
                except:
                    qf = 0.0
                if n and qf > 0:
                    if n in iu:
                        iu[n]["quantidade"] = float(iu[n]["quantidade"]) + qf
                    else:
                        iu[n] = {"quantidade": qf, "unidade": u}

            if msg == "":
                if obra and servico and su and qtd > 0 and hrs > 0:
                    cur.execute(
                        "INSERT INTO produtividade "
                        "(obra,servico,servico_unidade,quantidade,horas,funcoes,insumos,observacao,data,user_id) "
                        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                        (
                            obra,
                            servico,
                            su,
                            qtd,
                            hrs,
                            psycopg2.extras.Json(fu),
                            psycopg2.extras.Json(iu),
                            obs,
                            datetime.now().strftime("%d/%m/%Y"),
                            session["user_id"],
                        ),
                    )
                    conn.commit()
                    msg = "ok:Registro salvo!"
                else:
                    msg = "err:Preencha Obra, Servico, Quantidade e Horas."

        if is_admin():
            cur.execute(
                "SELECT p.id,p.data,p.obra,p.servico,p.quantidade,p.horas,u.nome "
                "FROM produtividade p LEFT JOIN usuarios u ON u.id=p.user_id "
                "ORDER BY p.id DESC LIMIT 20"
            )
        else:
            cur.execute(
                "SELECT p.id,p.data,p.obra,p.servico,p.quantidade,p.horas,u.nome "
                "FROM produtividade p LEFT JOIN usuarios u ON u.id=p.user_id "
                "WHERE p.user_id=%s ORDER BY p.id DESC LIMIT 20",
                (session["user_id"],),
            )
        recent = cur.fetchall()
    finally:
        cur.close()
        put_db(conn)

    oo = "".join([f'<option value="{o}">{o}</option>' for o in obras])
    so = "".join([f'<option value="{s}" data-unidade="{u}">{s} ({u})</option>' for s, u in servicos])
    fo, io = form_funcoes_insumos(funcoes, insumos)

    rcs = ""
    for rid, d, ob, sv, qt, hr, un in recent:
        df = (
            f"<form method='POST' action='/produtividade/excluir/{rid}' style='display:inline;'>"
            f"<input type='hidden' name='_csrf' value='{csrf_token()}'>"
            f"<button class='btn-sm btn-del' onclick=\"return confirm('Excluir?')\">Excluir</button>"
            f"</form>"
        )
        ut = f"<span class='mini' style='margin-left:auto;'>{un or ''}</span>" if is_admin() else ""

        qts = fmt_num(qt, 2) or "—"
        hrs_ = fmt_num(hr, 1) or "—"

        rcs += (
            f"<div class='rc'>"
            f"<div class='rc-top'><div class='rc-obra'>{ob or ''}</div><div class='rc-date'>{d or ''}</div></div>"
            f"<div class='rc-srv'>{sv or ''}</div>"
            f"<div class='rc-nums'><div>Qtd: <strong>{qts}</strong></div><div>Horas: <strong>{hrs_}</strong></div></div>"
            f"<div class='rc-act'><a class='btn-sm btn-edit' href='/produtividade/editar/{rid}'>Editar</a>{df}{ut}</div>"
            f"</div>"
        )

    adm = ""
    if is_admin():
        adm = (
            "<div class='card'><div class='ctitle'>Administracao</div>"
            "<div class='agrid'>"
            "<a href='/exportar' class='btn btn-dark' style='font-size:.82rem;padding:11px;'>Exportar Excel</a>"
            "<a href='/criar_usuario' class='btn btn-warn' style='font-size:.82rem;padding:11px;'>Criar Usuario</a>"
            "<a href='/cadastros' class='btn btn-info' style='font-size:.82rem;padding:11px;'>Cadastros</a>"
            "<a href='/usuarios' class='btn btn-purple' style='font-size:.82rem;padding:11px;'>Usuarios</a>"
            "<a href='/efetivo/custos' class='btn btn-dark' style='font-size:.82rem;padding:11px;'>Custos de Efetivo</a>"
            "</div></div>"
        )

    if not obras and not is_admin():
        aviso_obras = "<div class='alert a-err'>⚠ Nenhuma obra liberada para voce. Peça ao admin para liberar.</div>"
    else:
        aviso_obras = ""

    c = f"""{alert(msg)}{aviso_obras}
    <div class="card">
      <div class="ctitle">Novo Registro</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div class="fl"><label class="fl-label">Obra</label><select name="obra" required><option value="">Selecione a obra...</option>{oo}</select></div>
        <div class="fl"><label class="fl-label">Servico</label><select id="servico_select" name="servico" required><option value="">Selecione o servico...</option>{so}</select></div>
        <div style="display:flex;gap:8px;margin-bottom:10px;">
          <div style="flex:1;"><label class="fl-label">Quantidade executada</label><input type="number" name="quantidade" step="0.01" placeholder="0,00" required></div>
          <div style="width:76px;"><label class="fl-label">Unidade</label><div class="upill" id="sb" style="width:100%;">—</div><input type="hidden" name="servico_unidade" id="servico_unidade"></div>
        </div>
        <div class="fl"><label class="fl-label">Horas trabalhadas</label><input type="number" name="horas" step="0.1" placeholder="0,0" required></div>

        <div class="stitle">Funcoes</div>
        <div class="add-row"><select id="func_select"><option value=""></option>{fo}</select><input class="qty-box" type="number" id="fq" min="1" step="1" value="1"><button type="button" class="btn-plus" onclick="addFunc()">+</button></div>
        <div class="tags" id="fl"></div>

        <div class="stitle">Insumos</div>
        <div class="add-row"><select id="ins_select"><option value=""></option>{io}</select><input class="qty-box" type="number" id="iq" min="0" step="0.01" value="1"><div class="upill" id="ib">un</div><button type="button" class="btn-plus" onclick="addIns()">+</button></div>
        <div class="tags" id="il"></div>

        <div class="fl" style="margin-top:13px;"><label class="fl-label">Observacoes</label><textarea name="observacao" placeholder="Opcional..."></textarea></div>
        <button class="btn btn-green">Salvar Registro</button>
      </form>
    </div>

    <div class="card">
      <div class="ctitle">Registros Recentes</div>
      <div class="mini" style="margin-bottom:9px;">{"Ultimos 20 (todos os usuarios)." if is_admin() else "Seus ultimos 20 registros."}</div>
      {rcs if rcs else '<div class="mini" style="padding:10px 0;">Nenhum registro ainda.</div>'}
    </div>
    {adm}"""
    return render(c, page="produtividade")


@app.route("/produtividade/editar/<int:rid>", methods=["GET", "POST"])
def produtividade_editar(rid):
    if not require_login():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT id,obra,servico,servico_unidade,quantidade,horas,funcoes,insumos,observacao,data,user_id "
            "FROM produtividade WHERE id=%s",
            (rid,),
        )
        reg = cur.fetchone()
        if not reg:
            abort(404)

        (rid, obra, servico, su, qtd, hrs, fr, ir, obs, data, uid) = reg
        if not can_edit(uid):
            abort(403)

        if is_admin():
            obras, servicos, funcoes, insumos = load_dropdowns_admin()
        else:
            obras, servicos, funcoes, insumos = load_dropdowns_user(session["user_id"])

        msg = ""
        if request.method == "POST":
            require_csrf()
            on = request.form.get("obra", "").strip()
            sn = request.form.get("servico", "").strip()
            sun = request.form.get("servico_unidade", "").strip()
            on2 = request.form.get("observacao", "").strip()

            try:
                qn = float(request.form.get("quantidade", "0"))
            except:
                qn = 0.0
            try:
                hn = float(request.form.get("horas", "0"))
            except:
                hn = 0.0

            if not is_admin():
                if not user_can_use_obra_name(session["user_id"], on):
                    msg = "err:Voce nao tem permissao para registrar nessa obra."

            fn = request.form.getlist("funcoes_nome[]")
            fq = request.form.getlist("funcoes_qtd[]")
            fu = {}
            for n, q in zip(fn, fq):
                n = (n or "").strip()
                try:
                    qi = int(q)
                except:
                    qi = 0
                if n and qi > 0:
                    fu[n] = fu.get(n, 0) + qi

            ins_n = request.form.getlist("insumos_nome[]")
            ins_q = request.form.getlist("insumos_qtd[]")
            ins_u = request.form.getlist("insumos_unid[]")
            iu = {}
            for n, q, u in zip(ins_n, ins_q, ins_u):
                n = (n or "").strip()
                u = (u or "").strip() or "un"
                try:
                    qf = float(q)
                except:
                    qf = 0.0
                if n and qf > 0:
                    if n in iu:
                        iu[n]["quantidade"] = float(iu[n]["quantidade"]) + qf
                    else:
                        iu[n] = {"quantidade": qf, "unidade": u}

            if msg == "":
                if on and sn and sun and qn > 0 and hn > 0:
                    cur.execute(
                        "UPDATE produtividade SET obra=%s,servico=%s,servico_unidade=%s,quantidade=%s,horas=%s,funcoes=%s,insumos=%s,observacao=%s "
                        "WHERE id=%s",
                        (on, sn, sun, qn, hn, psycopg2.extras.Json(fu), psycopg2.extras.Json(iu), on2, rid),
                    )
                    conn.commit()
                    msg = "ok:Atualizado!"
                    obra, servico, su, qtd, hrs, obs = on, sn, sun, qn, hn, on2
                    fr, ir = fu, iu
                else:
                    msg = "err:Preencha todos os campos obrigatorios."
    finally:
        cur.close()
        put_db(conn)

    oo = "".join([f'<option value="{o}" {"selected" if o==obra else ""}>{o}</option>' for o in obras])
    so = "".join(
        [
            f'<option value="{s}" data-unidade="{u}" {"selected" if s==servico else ""}>{s} ({u})</option>'
            for s, u in servicos
        ]
    )
    fo, io = form_funcoes_insumos(funcoes, insumos)

    fd = safe_json_load(fr) or {}
    id_ = safe_json_load(ir) or {}

    fb = ""
    if isinstance(fd, dict):
        for n, q in fd.items():
            try:
                qi = int(q)
            except:
                qi = 0
            if n and qi > 0:
                fb += (
                    "<div class='tag'>"
                    f"<span>{n} — {qi}</span>"
                    "<button type='button' class='tag-x' onclick=\"this.closest('.tag').remove()\">x</button>"
                    f"<input type='hidden' name='funcoes_nome[]' value='{n}'>"
                    f"<input type='hidden' name='funcoes_qtd[]' value='{qi}'>"
                    "</div>"
                )

    ib = ""
    if isinstance(id_, dict):
        for n, v in id_.items():
            if not n:
                continue
            if isinstance(v, dict):
                qf = v.get("quantidade", 0)
                uu = v.get("unidade", "un")
            else:
                qf = v
                uu = "un"
            try:
                qf = float(qf)
            except:
                qf = 0.0
            if qf > 0:
                ib += (
                    "<div class='tag'>"
                    f"<span>{n} — {fmt_num(qf,2)} {uu}</span>"
                    "<button type='button' class='tag-x' onclick=\"this.closest('.tag').remove()\">x</button>"
                    f"<input type='hidden' name='insumos_nome[]' value='{n}'>"
                    f"<input type='hidden' name='insumos_qtd[]' value='{qf}'>"
                    f"<input type='hidden' name='insumos_unid[]' value='{uu}'>"
                    "</div>"
                )

    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card">
      <div class="ctitle">Editar #{rid}</div>
      <div class="mini" style="margin-bottom:13px;">Data: {data or ""}</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div class="fl"><label class="fl-label">Obra</label><select name="obra" required><option value="">Selecione...</option>{oo}</select></div>
        <div class="fl"><label class="fl-label">Servico</label><select id="servico_select" name="servico" required><option value="">Selecione...</option>{so}</select></div>
        <div style="display:flex;gap:8px;margin-bottom:10px;">
          <div style="flex:1;"><label class="fl-label">Quantidade executada</label><input type="number" name="quantidade" step="0.01" value="{fmt_num(qtd,2)}" required></div>
          <div style="width:76px;"><label class="fl-label">Unidade</label><div class="upill" id="sb" style="width:100%;">{su or "—"}</div><input type="hidden" name="servico_unidade" id="servico_unidade" value="{su or ""}"></div>
        </div>
        <div class="fl"><label class="fl-label">Horas trabalhadas</label><input type="number" name="horas" step="0.1" value="{fmt_num(hrs,1)}" required></div>

        <div class="stitle">Funcoes</div>
        <div class="add-row"><select id="func_select"><option value=""></option>{fo}</select><input class="qty-box" type="number" id="fq" min="1" step="1" value="1"><button type="button" class="btn-plus" onclick="addFunc()">+</button></div>
        <div class="tags" id="fl">{fb}</div>

        <div class="stitle">Insumos</div>
        <div class="add-row"><select id="ins_select"><option value=""></option>{io}</select><input class="qty-box" type="number" id="iq" min="0" step="0.01" value="1"><div class="upill" id="ib">un</div><button type="button" class="btn-plus" onclick="addIns()">+</button></div>
        <div class="tags" id="il">{ib}</div>

        <div class="fl" style="margin-top:13px;"><label class="fl-label">Observacoes</label><textarea name="observacao">{obs or ""}</textarea></div>
        <button class="btn btn-gold">Salvar Alteracoes</button>
      </form>
    </div>"""
    return render(c, page="produtividade")


@app.route("/produtividade/excluir/<int:rid>", methods=["POST"])
def produtividade_excluir(rid):
    if not require_login():
        return redirect("/")
    require_csrf()

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT user_id FROM produtividade WHERE id=%s", (rid,))
        row = cur.fetchone()
        if not row:
            return redirect("/dashboard")
        if not can_edit(row[0]):
            abort(403)

        cur.execute("DELETE FROM produtividade WHERE id=%s", (rid,))
        conn.commit()
        return redirect("/dashboard")
    finally:
        cur.close()
        put_db(conn)


# -------------------------
# Admin: Criar usuário + atribuir obras
# -------------------------
@app.route("/criar_usuario", methods=["GET", "POST"])
def criar_usuario():
    if not require_login() or not is_admin():
        return redirect("/")
    msg = ""

    obras_all = get_all_obras()

    if request.method == "POST":
        require_csrf()
        nome = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip().lower()
        senha = request.form.get("senha", "").strip()
        tipo = request.form.get("tipo", "usuario")
        obras_sel = request.form.getlist("obras_sel[]")

        ok, m = password_policy_ok(senha)
        if not ok:
            msg = f"err:{m}"
        else:
            conn = get_db()
            cur = conn.cursor()
            try:
                cur.execute(
                    "INSERT INTO usuarios (nome,email,senha,tipo) VALUES (%s,%s,%s,%s) RETURNING id",
                    (nome, email, generate_password_hash(senha), tipo),
                )
                uid = cur.fetchone()[0]

                for oid in obras_sel:
                    try:
                        oid_i = int(oid)
                    except:
                        continue
                    if oid_i > 0:
                        cur.execute(
                            "INSERT INTO usuario_obras (user_id, obra_id) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                            (uid, oid_i),
                        )

                conn.commit()
                msg = "ok:Usuario criado!"
            except:
                conn.rollback()
                msg = "err:E-mail ja cadastrado."
            finally:
                cur.close()
                put_db(conn)

    checks = ""
    if obras_all:
        for oid, on in obras_all:
            checks += f'<label style="display:block;margin:6px 0;"><input type="checkbox" name="obras_sel[]" value="{oid}"> {on}</label>'
    else:
        checks = "<div class='mini'>Nenhuma obra cadastrada ainda.</div>"

    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card">
      <div class="ctitle">Criar Usuario</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div class="fl"><label class="fl-label">Nome</label><input type="text" name="nome" placeholder="Nome completo" required></div>
        <div class="fl"><label class="fl-label">E-mail</label><input type="email" name="email" placeholder="email@exemplo.com" required></div>
        <div class="fl"><label class="fl-label">Senha</label><input type="password" name="senha" placeholder="Minimo 8 caracteres" required></div>
        <div class="mini" style="margin-bottom:9px;">Minimo 8 caracteres, pelo menos 1 letra e 1 numero.</div>
        <div class="fl"><label class="fl-label">Perfil</label>
          <select name="tipo"><option value="usuario">Usuario</option><option value="admin">Administrador</option></select>
        </div>
        <div class="fl"><label class="fl-label">Obras permitidas</label>{checks}</div>
        <button class="btn btn-gold">Criar</button>
      </form>
    </div>"""
    return render(c)


@app.route("/usuarios")
def usuarios():
    if not require_login() or not is_admin():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT id,nome,email,tipo FROM usuarios ORDER BY id")
        users = cur.fetchall()
    finally:
        cur.close()
        put_db(conn)

    rows = ""
    for uid, nome, email, tipo in users:
        dis = "disabled" if uid == session["user_id"] else ""
        cor = "#C9933A" if tipo == "admin" else "#6b7280"
        rows += (
            "<tr>"
            f"<td>{nome}</td>"
            f"<td style='font-size:.75rem;color:#6b7280;'>{email}</td>"
            f"<td><span style='font-size:.7rem;font-weight:700;text-transform:uppercase;color:{cor};'>{tipo}</span></td>"
            "<td><div style='display:flex;gap:5px;flex-wrap:wrap;'>"
            f"<a class='btn-sm btn-edit' href='/usuarios/reset_senha/{uid}'>Senha</a>"
            f"<a class='btn-sm btn-edit' href='/usuarios/obras/{uid}'>Obras</a>"
            "<form method='POST' action='/usuarios/excluir' style='display:inline;'>"
            f"<input type='hidden' name='_csrf' value='{csrf_token()}'>"
            f"<input type='hidden' name='id' value='{uid}'>"
            f"<button class='btn-sm btn-del' {dis} onclick=\"return confirm('Excluir?')\">Excluir</button>"
            "</form></div></td></tr>"
        )

    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    <div class="card"><div class="ctitle">Usuarios</div>
      <div style="overflow-x:auto;"><table class="utable"><thead><tr><th>Nome</th><th>E-mail</th><th>Perfil</th><th>Acoes</th></tr></thead><tbody>{rows}</tbody></table></div>
    </div>"""
    return render(c)


@app.route("/usuarios/excluir", methods=["POST"])
def usuarios_excluir():
    if not require_login() or not is_admin():
        return redirect("/")
    require_csrf()
    try:
        uid = int(request.form.get("id", "0"))
    except:
        uid = 0
    if uid <= 0 or uid == session["user_id"]:
        return redirect("/usuarios")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM usuario_obras WHERE user_id=%s", (uid,))
        cur.execute("DELETE FROM usuarios WHERE id=%s", (uid,))
        conn.commit()
    finally:
        cur.close()
        put_db(conn)

    return redirect("/usuarios")


@app.route("/usuarios/obras/<int:uid>", methods=["GET", "POST"])
def usuarios_obras(uid):
    if not require_login() or not is_admin():
        return redirect("/")

    msg = ""
    obras_all = get_all_obras()

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT id,nome,email,tipo FROM usuarios WHERE id=%s", (uid,))
        user = cur.fetchone()
        if not user:
            abort(404)

        if request.method == "POST":
            require_csrf()
            sel = request.form.getlist("obras_sel[]")
            cur.execute("DELETE FROM usuario_obras WHERE user_id=%s", (uid,))
            for oid in sel:
                try:
                    oid_i = int(oid)
                except:
                    continue
                if oid_i > 0:
                    cur.execute(
                        "INSERT INTO usuario_obras (user_id, obra_id) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                        (uid, oid_i),
                    )
            conn.commit()
            msg = "ok:Permissoes atualizadas!"

        cur.execute("SELECT obra_id FROM usuario_obras WHERE user_id=%s", (uid,))
        assigned = {r[0] for r in cur.fetchall()}
    finally:
        cur.close()
        put_db(conn)

    checks = ""
    if obras_all:
        for oid, on in obras_all:
            ck = "checked" if oid in assigned else ""
            checks += f'<label style="display:block;margin:6px 0;"><input type="checkbox" name="obras_sel[]" value="{oid}" {ck}> {on}</label>'
    else:
        checks = "<div class='mini'>Nenhuma obra cadastrada ainda.</div>"

    c = f"""<a href="/usuarios" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card">
      <div class="ctitle">Obras do Usuario</div>
      <div class="mini" style="margin-bottom:10px;">{user[1]} — {user[2]}</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        {checks}
        <button class="btn btn-gold" style="margin-top:12px;">Salvar</button>
      </form>
    </div>"""
    return render(c)


@app.route("/usuarios/reset_senha/<int:uid>", methods=["GET", "POST"])
def usuarios_reset_senha(uid):
    if not require_login() or not is_admin():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT id,nome,email,tipo FROM usuarios WHERE id=%s", (uid,))
        user = cur.fetchone()
        if not user:
            abort(404)

        msg = ""
        if request.method == "POST":
            require_csrf()
            n = request.form.get("nova_senha", "").strip()
            n2 = request.form.get("nova_senha2", "").strip()
            if n != n2:
                msg = "err:As senhas nao conferem."
            else:
                ok, m = password_policy_ok(n)
                if not ok:
                    msg = f"err:{m}"
                else:
                    cur.execute("UPDATE usuarios SET senha=%s WHERE id=%s", (generate_password_hash(n), uid))
                    conn.commit()
                    msg = "ok:Senha resetada!"
    finally:
        cur.close()
        put_db(conn)

    c = f"""<a href="/usuarios" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card"><div class="ctitle">Resetar Senha</div>
      <div class="mini" style="margin-bottom:14px;">Usuario: <strong>{user[1]}</strong> — {user[2]}</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div class="fl"><label class="fl-label">Nova senha</label><input type="password" name="nova_senha" required></div>
        <div class="fl"><label class="fl-label">Confirmar</label><input type="password" name="nova_senha2" required></div>
        <div class="mini" style="margin-bottom:11px;">Minimo 8 caracteres, pelo menos 1 letra e 1 numero.</div>
        <button class="btn btn-gold">Resetar</button>
      </form>
    </div>"""
    return render(c)


# -------------------------
# Cadastros (admin)
# -------------------------
@app.route("/cadastros", methods=["GET", "POST"])
def cadastros():
    if not require_login() or not is_admin():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    msg = ""
    try:
        if request.method == "POST":
            require_csrf()
            if request.form.get("obra", "").strip():
                try:
                    cur.execute("INSERT INTO obras (nome) VALUES (%s)", (request.form["obra"].strip(),))
                    conn.commit()
                    msg = "ok:Obra cadastrada!"
                except:
                    conn.rollback()
                    msg = "err:Obra ja existe."
            if request.form.get("servico", "").strip() and request.form.get("unidade", "").strip():
                try:
                    cur.execute(
                        "INSERT INTO servicos (nome,unidade) VALUES (%s,%s)",
                        (request.form["servico"].strip(), request.form["unidade"].strip()),
                    )
                    conn.commit()
                    msg = "ok:Servico cadastrado!"
                except:
                    conn.rollback()
                    msg = "err:Servico ja existe."
            if request.form.get("insumo", "").strip() and request.form.get("unidade_insumo", "").strip():
                try:
                    cur.execute(
                        "INSERT INTO insumos (nome,unidade) VALUES (%s,%s)",
                        (request.form["insumo"].strip(), request.form["unidade_insumo"].strip()),
                    )
                    conn.commit()
                    msg = "ok:Insumo cadastrado!"
                except:
                    conn.rollback()
                    msg = "err:Insumo ja existe."
            if request.form.get("funcao", "").strip():
                try:
                    cur.execute("INSERT INTO funcoes (nome) VALUES (%s)", (request.form["funcao"].strip(),))
                    conn.commit()
                    msg = "ok:Funcao cadastrada!"
                except:
                    conn.rollback()
                    msg = "err:Funcao ja existe."

        cur.execute("SELECT nome FROM obras ORDER BY nome")
        obras = [o[0] for o in cur.fetchall()]
        cur.execute("SELECT nome,COALESCE(unidade,'un') FROM servicos ORDER BY nome")
        servicos = cur.fetchall()
        cur.execute("SELECT nome,COALESCE(unidade,'un') FROM insumos ORDER BY nome")
        insumos = cur.fetchall()
        cur.execute("SELECT nome FROM funcoes ORDER BY nome")
        funcoes = [f[0] for f in cur.fetchall()]
    finally:
        cur.close()
        put_db(conn)

    def lst(items, lbl=False, sid=""):
        count = len(items)
        if not items:
            return '<div class="mini" style="margin-top:8px;">Nenhum cadastrado.</div>'
        ih = "<ul class='clist'>" + "".join(
            [f"<li>{(i[0]+' ('+i[1]+')') if lbl else i}</li>" for i in items]
        ) + "</ul>"
        return f'<div class="acc" id="acc-{sid}"><button type="button" class="acc-btn" onclick="togAcc(\'acc-{sid}\')"><span>{count} cadastrado{"s" if count!=1 else ""}</span><span class="acc-arr">▼</span></button><div class="acc-body" style="display:none;">{ih}</div></div>'

    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card"><div class="ctitle">Obras</div>
      <form method="POST"><input type="hidden" name="_csrf" value="{csrf_token()}">
        <div style="display:flex;gap:8px;"><input type="text" name="obra" placeholder="Nome da obra" required style="margin-bottom:0;"><button class="btn btn-gold" style="width:auto;padding:12px 16px;white-space:nowrap;">+ Add</button></div></form>
      {lst(obras, sid="obras")}</div>
    <div class="card"><div class="ctitle">Servicos</div>
      <form method="POST"><input type="hidden" name="_csrf" value="{csrf_token()}">
        <input type="text" name="servico" placeholder="Nome do servico" required>
        <div style="display:flex;gap:8px;"><input type="text" name="unidade" placeholder="Unidade (m2, m3, un...)" required style="margin-bottom:0;"><button class="btn btn-gold" style="width:auto;padding:12px 16px;white-space:nowrap;">+ Add</button></div></form>
      {lst(servicos, lbl=True, sid="servicos")}</div>
    <div class="card"><div class="ctitle">Insumos</div>
      <form method="POST"><input type="hidden" name="_csrf" value="{csrf_token()}">
        <input type="text" name="insumo" placeholder="Nome do insumo" required>
        <div style="display:flex;gap:8px;"><input type="text" name="unidade_insumo" placeholder="Unidade (kg, m, saco...)" required style="margin-bottom:0;"><button class="btn btn-gold" style="width:auto;padding:12px 16px;white-space:nowrap;">+ Add</button></div></form>
      {lst(insumos, lbl=True, sid="insumos")}</div>
    <div class="card"><div class="ctitle">Funcoes</div>
      <form method="POST"><input type="hidden" name="_csrf" value="{csrf_token()}">
        <div style="display:flex;gap:8px;"><input type="text" name="funcao" placeholder="Ex: Armador, Pedreiro" required style="margin-bottom:0;"><button class="btn btn-gold" style="width:auto;padding:12px 16px;white-space:nowrap;">+ Add</button></div></form>
      {lst(funcoes, sid="funcoes")}</div>"""
    return render(c)


# -------------------------
# Custo de Efetivo (Mês) — Engenheiro
# -------------------------
@app.route("/efetivo")
def efetivo_index():
    if not require_login():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    try:
        if is_admin():
            cur.execute(
                "SELECT e.id,e.mes,e.total,e.updated_at,o.nome,u.nome "
                "FROM efetivo_planejamentos e "
                "JOIN obras o ON o.id=e.obra_id "
                "JOIN usuarios u ON u.id=e.user_id "
                "ORDER BY e.mes DESC, e.id DESC LIMIT 50"
            )
        else:
            cur.execute(
                "SELECT e.id,e.mes,e.total,e.updated_at,o.nome,u.nome "
                "FROM efetivo_planejamentos e "
                "JOIN obras o ON o.id=e.obra_id "
                "JOIN usuarios u ON u.id=e.user_id "
                "WHERE e.user_id=%s "
                "ORDER BY e.mes DESC, e.id DESC LIMIT 50",
                (session["user_id"],),
            )
        rows = cur.fetchall()
    finally:
        cur.close()
        put_db(conn)

    cards = ""
    for rid, mes, total, upd, obra_nome, user_nome in rows:
        who = f"<span class='mini' style='margin-left:auto;'>{user_nome or ''}</span>" if is_admin() else ""

        # Botão excluir — apenas para admin
        del_btn = ""
        if is_admin():
            del_btn = (
                f"<form method='POST' action='/efetivo/excluir/{rid}' style='display:inline;'>"
                f"<input type='hidden' name='_csrf' value='{csrf_token()}'>"
                f"<button class='btn-sm btn-del' onclick=\"return confirm('Excluir planejamento {obra_nome} — {mes}?')\">Excluir</button>"
                f"</form>"
            )

        cards += (
            "<div class='rc'>"
            f"<div class='rc-top'><div class='rc-obra'>{obra_nome}</div><div class='rc-date'>{mes}</div></div>"
            f"<div class='rc-srv'>Total do mês: <strong>R$ {fmt_num(total,2) or '0'}</strong></div>"
            f"<div class='rc-nums'><div class='mini'>Atualizado: {upd or ''}</div>{who}</div>"
            f"<div class='rc-act'>"
            f"<a class='btn-sm btn-edit' href='/efetivo/editar/{rid}'>Abrir</a>"
            f"<a class='btn-sm btn-edit' href='/efetivo/pdf/{rid}'>PDF</a>"
            f"{del_btn}"
            f"</div></div>"
        )

    novo_btn = "<a class='btn btn-gold' href='/efetivo/novo'>+ Novo Mês</a>"
    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    <div class="card">
      <div class="ctitle">Custo de Efetivo (Mês)</div>
      <div class="mini" style="margin-bottom:12px;">Planejamento de custo fixo de mao de obra propria por obra e por mes.</div>
      {novo_btn}
    </div>
    <div class="card">
      <div class="ctitle">Meus Planejamentos</div>
      {cards if cards else "<div class='mini'>Nenhum planejamento ainda.</div>"}
    </div>"""
    return render(c, page="efetivo")


# -------------------------
# Excluir Planejamento (admin only)
# -------------------------
@app.route("/efetivo/excluir/<int:eid>", methods=["POST"])
def efetivo_excluir(eid):
    if not require_login() or not is_admin():
        abort(403)
    require_csrf()

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM efetivo_planejamentos WHERE id=%s", (eid,))
        conn.commit()
    finally:
        cur.close()
        put_db(conn)

    return redirect("/efetivo")


@app.route("/efetivo/novo", methods=["GET", "POST"])
def efetivo_novo():
    if not require_login():
        return redirect("/")

    msg = ""
    if is_admin():
        obras = get_all_obras()
    else:
        obras = get_allowed_obras(session["user_id"])

    if request.method == "POST":
        require_csrf()
        try:
            obra_id = int(request.form.get("obra_id", "0"))
        except:
            obra_id = 0
        mes = parse_month(request.form.get("mes", ""))

        if obra_id <= 0 or not mes:
            msg = "err:Selecione Obra e Mes."
        elif (not is_admin()) and (not user_can_use_obra_id(session["user_id"], obra_id)):
            msg = "err:Voce nao tem permissao para essa obra."
        else:
            conn = get_db()
            cur = conn.cursor()
            try:
                cur.execute(
                    "SELECT id FROM efetivo_planejamentos WHERE user_id=%s AND obra_id=%s AND mes=%s",
                    (session["user_id"], obra_id, mes),
                )
                r = cur.fetchone()
                if r:
                    return redirect(f"/efetivo/editar/{r[0]}")

                now = datetime.now().strftime("%d/%m/%Y %H:%M")
                cur.execute(
                    "INSERT INTO efetivo_planejamentos (obra_id,user_id,mes,itens,total,created_at,updated_at) "
                    "VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING id",
                    (obra_id, session["user_id"], mes, psycopg2.extras.Json([]), 0.0, now, now),
                )
                new_id = cur.fetchone()[0]
                conn.commit()
                return redirect(f"/efetivo/editar/{new_id}")
            except:
                conn.rollback()
                msg = "err:Nao foi possivel criar."
            finally:
                cur.close()
                put_db(conn)

    opts = "".join([f"<option value='{oid}'>{on}</option>" for oid, on in obras])
    if not obras and not is_admin():
        aviso = "<div class='alert a-err'>⚠ Nenhuma obra liberada para voce.</div>"
    else:
        aviso = ""

    c = f"""<a href="/efetivo" class="back">← Voltar</a>
    {alert(msg)}{aviso}
    <div class="card">
      <div class="ctitle">Novo Planejamento</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div class="fl"><label class="fl-label">Obra</label>
          <select name="obra_id" required><option value="">Selecione...</option>{opts}</select>
        </div>
        <div class="fl"><label class="fl-label">Mes</label>
          <input type="month" name="mes" required>
          <div class="mini" style="margin-top:6px;">Exemplo: 2026-04 (Abril/2026).</div>
        </div>
        <button class="btn btn-gold">Criar / Abrir</button>
      </form>
    </div>"""
    return render(c, page="efetivo")


def load_costs_for_obra(obra_id: int):
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("SELECT id,nome FROM funcoes ORDER BY nome")
        funs = cur.fetchall()
        cur.execute("SELECT funcao_id,COALESCE(custo_mensal,0) FROM custos_funcao_obra WHERE obra_id=%s", (obra_id,))
        costs = {r[0]: float(r[1] or 0) for r in cur.fetchall()}
        out = []
        for fid, fn in funs:
            out.append((fid, fn, float(costs.get(fid, 0))))
        return out
    finally:
        cur.close()
        put_db(conn)


@app.route("/efetivo/editar/<int:eid>", methods=["GET", "POST"])
def efetivo_editar(eid):
    if not require_login():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT id,obra_id,user_id,mes,itens,total,created_at,updated_at "
            "FROM efetivo_planejamentos WHERE id=%s",
            (eid,),
        )
        reg = cur.fetchone()
        if not reg:
            abort(404)

        (eid, obra_id, user_id, mes, itens, total, created_at, updated_at) = reg

        if not is_admin() and int(user_id) != int(session["user_id"]):
            abort(403)

        if not is_admin():
            if not user_can_use_obra_id(session["user_id"], obra_id):
                abort(403)

        cur.execute("SELECT nome FROM obras WHERE id=%s", (obra_id,))
        obra_nome = (cur.fetchone() or [""])[0]

        funcs = load_costs_for_obra(obra_id)

        itens_list = safe_json_load(itens) or []
        qtd_map = {}
        custo_snapshot_map = {}
        if isinstance(itens_list, list):
            for it in itens_list:
                try:
                    fid = int(it.get("funcao_id", 0))
                except:
                    fid = 0
                if fid > 0:
                    try:
                        qtd_map[fid] = int(it.get("qtd", 0) or 0)
                    except:
                        qtd_map[fid] = 0
                    try:
                        custo_snapshot_map[fid] = float(it.get("custo_mensal", 0) or 0)
                    except:
                        custo_snapshot_map[fid] = 0.0

        msg = ""
        if request.method == "POST":
            require_csrf()
            fids = request.form.getlist("funcao_id[]")
            custos = request.form.getlist("custo_mensal[]")
            qtds = request.form.getlist("qtd[]")

            itens_new = []
            total_new = 0.0

            for a, b, c in zip(fids, custos, qtds):
                try:
                    fid = int(a)
                except:
                    continue
                fn = next((x[1] for x in funcs if int(x[0]) == fid), "")
                try:
                    custo = float(b)
                except:
                    custo = 0.0
                try:
                    qtd = int(float(c))
                except:
                    qtd = 0
                if qtd < 0:
                    qtd = 0
                subtotal = float(custo) * float(qtd)
                total_new += subtotal

                itens_new.append(
                    {
                        "funcao_id": fid,
                        "funcao_nome": fn,
                        "custo_mensal": float(custo),
                        "qtd": int(qtd),
                        "subtotal": float(subtotal),
                    }
                )

            now = datetime.now().strftime("%d/%m/%Y %H:%M")
            cur.execute(
                "UPDATE efetivo_planejamentos SET itens=%s,total=%s,updated_at=%s WHERE id=%s",
                (psycopg2.extras.Json(itens_new), float(total_new), now, eid),
            )
            conn.commit()
            msg = "ok:Planejamento salvo!"

            itens_list = itens_new
            total = total_new
            updated_at = now
            qtd_map = {it["funcao_id"]: it["qtd"] for it in itens_new}
            custo_snapshot_map = {it["funcao_id"]: it["custo_mensal"] for it in itens_new}

    finally:
        cur.close()
        put_db(conn)

    lines = ""
    for fid, fn, custo_atual in funcs:
        custo = custo_snapshot_map.get(fid, custo_atual)
        qtd = qtd_map.get(fid, 0)
        sub = float(custo) * float(qtd)
        lines += (
            f"<tr data-ef-row='1' data-custo='{custo}'>"
            f"<td style='font-weight:700;'>{fn}</td>"
            f"<td style='white-space:nowrap;'>R$ {fmt_num(custo,2) or '0'}</td>"
            f"<td style='width:90px;'><input type='number' name='qtd[]' min='0' step='1' value='{qtd}' style='padding:10px 10px;border-radius:10px;border:1.5px solid var(--border);width:90px;'></td>"
            f"<td style='white-space:nowrap;font-weight:800;color:var(--dark);' data-sub>R$ {fmt_num(sub,2) or '0'}</td>"
            f"<input type='hidden' name='funcao_id[]' value='{fid}'>"
            f"<input type='hidden' name='custo_mensal[]' value='{custo}'>"
            f"</tr>"
        )

    # Botão excluir dentro do planejamento — apenas admin
    del_btn_inline = ""
    if is_admin():
        del_btn_inline = (
            f"<form method='POST' action='/efetivo/excluir/{eid}' style='display:inline;margin-top:10px;'>"
            f"<input type='hidden' name='_csrf' value='{csrf_token()}'>"
            f"<button class='btn btn-warn' style='margin-top:10px;' onclick=\"return confirm('Excluir este planejamento permanentemente?')\">Excluir Planejamento</button>"
            f"</form>"
        )

    pdf_btn = f"<a class='btn btn-dark' href='/efetivo/pdf/{eid}' style='margin-top:10px;'>Gerar PDF</a>"
    c = f"""<a href="/efetivo" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card">
      <div class="ctitle">Custo de Efetivo — {mes}</div>
      <div class="mini" style="margin-bottom:8px;">Obra: <strong>{obra_nome}</strong></div>
      <div class="mini">Criado: {created_at or ""} • Atualizado: {updated_at or ""}</div>
    </div>

    <div class="card">
      <div class="ctitle">Efetivo do Mes</div>
      <form method="POST" id="ef_form">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <div style="overflow-x:auto;">
          <table class="utable">
            <thead><tr>
              <th>Funcao</th><th>Custo Mensal</th><th>Qtd</th><th>Subtotal</th>
            </tr></thead>
            <tbody>{lines}</tbody>
          </table>
        </div>

        <div class="kpi" style="margin-top:12px;">
          <div class="box">
            <div class="lbl">Total do mes</div>
            <div class="val" id="ef_total">R$ {fmt_num(total,2) or "0"}</div>
          </div>
        </div>

        <button class="btn btn-gold" style="margin-top:12px;">Salvar</button>
      </form>
      {pdf_btn}
      {del_btn_inline}
      <div class="mini" style="margin-top:10px;">Obs: o custo usado aqui fica congelado (snapshot) para este mes, mesmo que o admin altere os custos futuramente.</div>
    </div>"""
    return render(c, page="efetivo")


# -------------------------
# Admin: configurar custos por obra x função
# -------------------------
@app.route("/efetivo/custos", methods=["GET", "POST"])
def efetivo_custos():
    if not require_login() or not is_admin():
        return redirect("/")

    msg = ""
    obras = get_all_obras()
    obra_id = 0

    if request.method == "POST":
        require_csrf()
        try:
            obra_id = int(request.form.get("obra_id", "0"))
        except:
            obra_id = 0
        action = request.form.get("action", "select")

        if action == "save" and obra_id > 0:
            fids = request.form.getlist("funcao_id[]")
            custos = request.form.getlist("custo[]")
            conn = get_db()
            cur = conn.cursor()
            try:
                for a, b in zip(fids, custos):
                    try:
                        fid = int(a)
                    except:
                        continue
                    try:
                        custo = float((b or "0").replace(",", "."))
                    except:
                        custo = 0.0
                    if custo < 0:
                        custo = 0.0
                    cur.execute(
                        "INSERT INTO custos_funcao_obra (obra_id,funcao_id,custo_mensal) VALUES (%s,%s,%s) "
                        "ON CONFLICT (obra_id,funcao_id) DO UPDATE SET custo_mensal=EXCLUDED.custo_mensal",
                        (obra_id, fid, float(custo)),
                    )
                conn.commit()
                msg = "ok:Custos salvos!"
            except:
                conn.rollback()
                msg = "err:Nao foi possivel salvar."
            finally:
                cur.close()
                put_db(conn)

    else:
        try:
            obra_id = int(request.args.get("obra_id", "0"))
        except:
            obra_id = 0

    opts = "".join([f"<option value='{oid}' {'selected' if oid==obra_id else ''}>{on}</option>" for oid, on in obras])

    table_html = ""
    obra_nome = ""
    if obra_id > 0:
        conn = get_db()
        cur = conn.cursor()
        try:
            cur.execute("SELECT nome FROM obras WHERE id=%s", (obra_id,))
            obra_nome = (cur.fetchone() or [""])[0]
        finally:
            cur.close()
            put_db(conn)

        funcs = load_costs_for_obra(obra_id)
        rows = ""
        for fid, fn, custo in funcs:
            rows += (
                "<tr>"
                f"<td style='font-weight:700;'>{fn}</td>"
                f"<td style='width:180px;'><input type='number' step='0.01' min='0' name='custo[]' value='{fmt_num(custo,2)}' placeholder='0,00'></td>"
                f"<td class='mini'>por pessoa / mes</td>"
                f"<input type='hidden' name='funcao_id[]' value='{fid}'>"
                "</tr>"
            )
        table_html = f"""
        <div class="card">
          <div class="ctitle">Custos — {obra_nome}</div>
          <form method="POST">
            <input type="hidden" name="_csrf" value="{csrf_token()}">
            <input type="hidden" name="obra_id" value="{obra_id}">
            <input type="hidden" name="action" value="save">
            <div style="overflow-x:auto;">
              <table class="utable">
                <thead><tr><th>Funcao</th><th>Custo mensal</th><th></th></tr></thead>
                <tbody>{rows}</tbody>
              </table>
            </div>
            <button class="btn btn-gold" style="margin-top:12px;">Salvar Custos</button>
          </form>
          <div class="mini" style="margin-top:10px;">Dica: alterar aqui afeta somente planejamentos novos. Planejamentos antigos mantêm o custo do mês (snapshot).</div>
        </div>
        """

    c = f"""<a href="/dashboard" class="back">← Voltar</a>
    {alert(msg)}
    <div class="card">
      <div class="ctitle">Custos de Efetivo</div>
      <div class="mini" style="margin-bottom:10px;">Defina o custo mensal por pessoa, por obra e por funcao. (Somente admin)</div>
      <form method="POST">
        <input type="hidden" name="_csrf" value="{csrf_token()}">
        <input type="hidden" name="action" value="select">
        <div class="fl"><label class="fl-label">Selecionar Obra</label>
          <select name="obra_id" required><option value="">Selecione...</option>{opts}</select>
        </div>
        <button class="btn btn-dark">Abrir</button>
      </form>
    </div>
    {table_html}
    """
    return render(c, page="efetivo")


# -------------------------
# PDF do Efetivo (Mês)
# -------------------------
def _money_br(v):
    try:
        v = float(v or 0)
    except:
        v = 0.0
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


@app.route("/efetivo/pdf/<int:eid>")
def efetivo_pdf(eid):
    if not require_login():
        return redirect("/")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT e.id,e.mes,e.total,e.created_at,e.updated_at,e.itens,e.obra_id,e.user_id,o.nome,u.nome "
            "FROM efetivo_planejamentos e "
            "JOIN obras o ON o.id=e.obra_id "
            "JOIN usuarios u ON u.id=e.user_id "
            "WHERE e.id=%s",
            (eid,),
        )
        row = cur.fetchone()
        if not row:
            abort(404)

        (
            _id,
            mes,
            total,
            created_at,
            updated_at,
            itens,
            obra_id,
            user_id,
            obra_nome,
            user_nome,
        ) = row

        if not is_admin() and int(user_id) != int(session["user_id"]):
            abort(403)

        if not is_admin():
            if not user_can_use_obra_id(session["user_id"], obra_id):
                abort(403)

        itens_list = safe_json_load(itens) or []
    finally:
        cur.close()
        put_db(conn)

    out = BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    W, H = A4

    x0 = 18 * mm
    y = H - 18 * mm

    def draw_line(text, size=11, bold=False, dy=6 * mm):
        nonlocal y
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(x0, y, text)
        y -= dy

    draw_line("REOBOTE HOME", 14, True, 7 * mm)
    draw_line("Custo de Efetivo (Mês)", 12, True, 8 * mm)

    draw_line(f"Obra: {obra_nome}", 11, False, 6 * mm)
    draw_line(f"Mês: {mes}", 11, False, 6 * mm)
    draw_line(f"Usuário: {user_nome}", 11, False, 6 * mm)
    draw_line(f"Criado: {created_at or ''}    Atualizado: {updated_at or ''}", 10, False, 8 * mm)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(x0, y, "Função")
    c.drawString(x0 + 85 * mm, y, "Custo mensal")
    c.drawString(x0 + 125 * mm, y, "Qtd")
    c.drawString(x0 + 145 * mm, y, "Subtotal")
    y -= 6 * mm
    c.setLineWidth(0.6)
    c.line(x0, y, W - x0, y)
    y -= 6 * mm

    c.setFont("Helvetica", 10)

    def ensure_space(min_y=20 * mm):
        nonlocal y
        if y < min_y:
            c.showPage()
            y = H - 18 * mm
            c.setFont("Helvetica-Bold", 10)
            c.drawString(x0, y, "Função")
            c.drawString(x0 + 85 * mm, y, "Custo mensal")
            c.drawString(x0 + 125 * mm, y, "Qtd")
            c.drawString(x0 + 145 * mm, y, "Subtotal")
            y -= 6 * mm
            c.line(x0, y, W - x0, y)
            y -= 6 * mm
            c.setFont("Helvetica", 10)

    for it in itens_list if isinstance(itens_list, list) else []:
        ensure_space()
        fn = str(it.get("funcao_nome", "") or "")
        custo = it.get("custo_mensal", 0) or 0
        qtd = it.get("qtd", 0) or 0
        sub = it.get("subtotal", 0) or (float(custo) * float(qtd))

        c.drawString(x0, y, fn[:45])
        c.drawRightString(x0 + 120 * mm, y, _money_br(custo))
        c.drawRightString(x0 + 138 * mm, y, str(int(qtd)))
        c.drawRightString(W - x0, y, _money_br(sub))
        y -= 6 * mm

    y -= 3 * mm
    c.setLineWidth(0.6)
    c.line(x0, y, W - x0, y)
    y -= 8 * mm
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(W - x0, y, f"TOTAL: {_money_br(total)}")

    c.showPage()
    c.save()

    out.seek(0)
    filename = f"custo_efetivo_{obra_nome}_{mes}.pdf".replace(" ", "_")
    return send_file(out, download_name=filename, as_attachment=True)


# -------------------------
# Exportar Excel (admin)
# -------------------------
@app.route("/exportar")
def exportar():
    if not require_login() or not is_admin():
        return redirect("/dashboard")

    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT p.id,p.obra,p.servico,p.servico_unidade,p.quantidade,p.horas,p.funcoes,p.insumos,p.observacao,p.data,p.user_id,u.nome "
            "FROM produtividade p LEFT JOIN usuarios u ON u.id=p.user_id ORDER BY p.id DESC"
        )
        dados = cur.fetchall()
    finally:
        cur.close()
        put_db(conn)

    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    ws.append(
        [
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
            "homem_hora",
        ]
    )

    for row in dados:
        (rid, obra, servico, su, qtd, hrs, fr, ir, obs, data, uid, unom) = row
        try:
            hpu = (float(hrs) / float(qtd)) if float(qtd) != 0 else 0
        except:
            hpu = 0

        fu = safe_json_load(fr) or {}
        iu = safe_json_load(ir) or {}
        wrote = False

        if isinstance(fu, dict):
            for fn, fq in fu.items():
                try:
                    qi = int(fq)
                except:
                    qi = 0
                if not fn or qi <= 0:
                    continue
                try:
                    hh = float(hrs) * float(qi)
                except:
                    hh = 0
                ws.append([rid, data, obra, servico, su, qtd, hrs, hpu, obs, uid, unom or "", "FUNCAO", str(fn), qi, "", hh])
                wrote = True

        if isinstance(iu, dict):
            for inn, iv in iu.items():
                if not inn:
                    continue
                if isinstance(iv, dict):
                    iq = iv.get("quantidade", "")
                    iu2 = iv.get("unidade", "")
                else:
                    parts = str(iv).strip().split()
                    iq = parts[0] if parts else ""
                    iu2 = " ".join(parts[1:]) if len(parts) > 1 else ""
                try:
                    qn = float(iq)
                except:
                    qn = 0.0
                if qn <= 0:
                    continue
                ws.append([rid, data, obra, servico, su, qtd, hrs, hpu, obs, uid, unom or "", "INSUMO", str(inn), qn, str(iu2), ""])
                wrote = True

        if not wrote:
            ws.append([rid, data, obra, servico, su, qtd, hrs, hpu, obs, uid, unom or "", "REGISTRO", "", "", "", ""])

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, download_name="produtividade_powerbi.xlsx", as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)





