"""Consulta de Parcelas por CPF/CNPJ.

Reproduz somente o painel "Informações Parcelas Filtradas".
A planilha (.xlsx ou .csv) é só a fonte: apenas as colunas necessárias
para os indicadores são lidas e gravadas no SQLite (parcelas.db).
"""
import csv
import datetime as dt
import os
import re
import sqlite3
import threading
import unicodedata
import webbrowser

from flask import Flask, jsonify, render_template_string, request

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "parcelas.db")
UPLOADS = os.path.join(BASE, "_uploads")
os.makedirs(UPLOADS, exist_ok=True)

app = Flask(__name__)

# ----------------------------------------------------------------------------
# Mapeamento de colunas (nomes normalizados: sem acento, minúsculo)
# Só estes campos são lidos da planilha.
# ----------------------------------------------------------------------------
ALIASES = {
    "doc": ["cnpj cpf", "cpf cnpj", "cnpj", "cpf"],
    "venc": ["data vencimento"],
    "orig": ["valor original"],
    "saldo": ["valor saldo"],
    "recebido": ["valor total recebido"],
    "juros": ["juros calculado"],
    "multa": ["multa calculada"],
    "sit": ["situacao"],
    "numdoc": ["no documento", "n documento", "numero documento"],
    "parcela": ["parcela"],
    # colunas exibidas na lista de linhas
    "filial": ["filial"],
    "codcli": ["codigo cliente"],
    "nome": ["nome"],
    "emissao": ["emissao"],
    "venc_orig": ["vencto original"],
    "cobranca": ["tipo de cobranca"],
    "dias": ["dias atraso"],
    "liquidacao": ["data da liquidacao"],
}
REQUIRED = ["doc", "venc", "orig", "saldo", "sit"]


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def map_header(header):
    """Devolve {campo: índice} usando a PRIMEIRA coluna com cada nome."""
    normed = [norm(h) if h is not None else "" for h in header]
    mapping = {}
    for field, names in ALIASES.items():
        for name in names:
            if name in normed:
                mapping[field] = normed.index(name)
                break
    missing = [f for f in REQUIRED if f not in mapping]
    if missing:
        raise ValueError(
            "Colunas obrigatórias não encontradas: %s. Cabeçalho lido: %s"
            % (", ".join(missing), [h for h in header if h])
        )
    return mapping


# ----------------------------------------------------------------------------
# Conversões
# ----------------------------------------------------------------------------
def money(v):
    """Valor -> centavos (inteiro). None se vazio."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v * 100))
    s = str(v).strip().replace("R$", "").strip()
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return int(round(float(s) * 100))
    except ValueError:
        return None


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()[:10]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def only_digits(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):  # número puro: recupera zeros à esquerda
        d = str(int(v))
        return d.zfill(11) if len(d) <= 11 else d.zfill(14)
    return re.sub(r"\D", "", str(v))


def to_int(v):
    try:
        return int(float(str(v).strip().replace(",", ".")))
    except (ValueError, TypeError):
        return None


def to_text(v):
    return None if v is None or str(v).strip() == "" else str(v).strip()


def classify(sit):
    n = norm(sit or "")
    if n.startswith("liquid"):
        return "liquidado"
    if n.startswith("cancel"):
        return "cancelado"
    if n.startswith("renegoc"):
        return "renegociado"
    return "aberto"


# ----------------------------------------------------------------------------
# Leitura de planilha (streaming)
# ----------------------------------------------------------------------------
def iter_rows(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl

        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        yield from ws.iter_rows(values_only=True)
        wb.close()
    else:
        for enc in ("utf-8-sig", "latin-1"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    sample = f.read(4096)
                    f.seek(0)
                    delim = ";" if sample.count(";") > sample.count(",") else ","
                    yield from csv.reader(f, delimiter=delim)
                return
            except UnicodeDecodeError:
                continue


STATUS = {"state": "idle", "msg": "", "rows": 0, "situacoes": {}}


def do_import(path):
    STATUS.update(state="running", msg="Lendo planilha...", rows=0, situacoes={})
    try:
        rows = iter_rows(path)
        mapping = map_header(next(rows))
        get = lambda r, k: r[mapping[k]] if k in mapping and mapping[k] < len(r) else None

        con = sqlite3.connect(DB)
        con.execute("DROP TABLE IF EXISTS parcelas_new")
        con.execute(
            """CREATE TABLE parcelas_new (
                doc TEXT NOT NULL, numdoc TEXT, parcela TEXT, venc TEXT,
                orig INTEGER, saldo INTEGER, recebido INTEGER,
                juros INTEGER, multa INTEGER, situacao TEXT, status TEXT,
                filial TEXT, codcli TEXT, nome TEXT, emissao TEXT, venc_orig TEXT,
                cobranca TEXT, dias INTEGER, liquidacao TEXT)"""
        )
        sql = "INSERT INTO parcelas_new VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        batch, total, sits = [], 0, {}
        for r in rows:
            doc = only_digits(get(r, "doc"))
            if not doc:
                continue
            sit = str(get(r, "sit") or "").strip()
            st = classify(sit)
            sits[sit] = sits.get(sit, 0) + 1
            batch.append((
                doc,
                None if get(r, "numdoc") is None else str(get(r, "numdoc")).strip(),
                None if get(r, "parcela") is None else str(get(r, "parcela")).strip(),
                to_date(get(r, "venc")),
                money(get(r, "orig")) or 0,
                money(get(r, "saldo")) or 0,
                money(get(r, "recebido")),
                money(get(r, "juros")) or 0,
                money(get(r, "multa")) or 0,
                sit, st,
                to_text(get(r, "filial")), to_text(get(r, "codcli")),
                to_text(get(r, "nome")), to_date(get(r, "emissao")),
                to_date(get(r, "venc_orig")), to_text(get(r, "cobranca")),
                to_int(get(r, "dias")), to_date(get(r, "liquidacao")),
            ))
            total += 1
            if len(batch) >= 5000:
                con.executemany(sql, batch)
                batch = []
                STATUS.update(rows=total, msg="Importando... %d linhas" % total)
        if batch:
            con.executemany(sql, batch)
        con.execute("CREATE INDEX idx_new_doc ON parcelas_new(doc)")
        con.execute("DROP TABLE IF EXISTS parcelas")
        con.execute("ALTER TABLE parcelas_new RENAME TO parcelas")
        con.execute("DROP INDEX IF EXISTS idx_new_doc")
        con.execute("CREATE INDEX idx_parcelas_doc ON parcelas(doc)")
        con.execute("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)")
        con.execute("INSERT OR REPLACE INTO meta VALUES ('importado_em', ?)",
                    (dt.datetime.now().strftime("%d/%m/%Y %H:%M"),))
        con.execute("INSERT OR REPLACE INTO meta VALUES ('arquivo', ?)",
                    (os.path.basename(path),))
        con.commit()
        con.close()
        STATUS.update(state="done", rows=total, situacoes=sits,
                      msg="Importação concluída: %d parcelas." % total)
    except Exception as e:  # noqa: BLE001
        STATUS.update(state="error", msg="Erro na importação: %s" % e)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


# ----------------------------------------------------------------------------
# Consulta
# ----------------------------------------------------------------------------
FILTROS = {"todas": None, "aberto": "aberto", "liquidadas": "liquidado",
           "canceladas": "cancelado", "renegociadas": "renegociado"}


def db_info():
    if not os.path.exists(DB):
        return None
    con = sqlite3.connect(DB)
    try:
        n = con.execute("SELECT COUNT(*) FROM parcelas").fetchone()[0]
        meta = dict(con.execute("SELECT k, v FROM meta").fetchall())
        return {"parcelas": n, "importado_em": meta.get("importado_em"),
                "arquivo": meta.get("arquivo")}
    except sqlite3.Error:
        return None
    finally:
        con.close()


@app.route("/api/consulta")
def consulta():
    tipo = request.args.get("tipo", "cpf")
    doc = re.sub(r"\D", "", request.args.get("doc", ""))
    filtro = request.args.get("filtro", "todas")
    de = to_date(request.args.get("de"))
    ate = to_date(request.args.get("ate"))

    esperado = 11 if tipo == "cpf" else 14
    if len(doc) != esperado:
        return jsonify(erro="Informe um %s com %d dígitos." % (tipo.upper(), esperado)), 400
    if not os.path.exists(DB) or not db_info():
        return jsonify(erro="Nenhuma planilha importada ainda."), 400

    where, params = ["doc = ?"], [doc]
    if FILTROS.get(filtro):
        where.append("status = ?")
        params.append(FILTROS[filtro])
    if de:
        where.append("venc >= ?")
        params.append(de)
    if ate:
        where.append("venc <= ?")
        params.append(ate)

    con = sqlite3.connect(DB)
    existentes = {c[1] for c in con.execute("PRAGMA table_info(parcelas)")}
    base = ["numdoc", "parcela", "venc", "situacao", "status",
            "orig", "saldo", "recebido", "juros", "multa"]
    extras = ["filial", "codcli", "nome", "emissao", "venc_orig",
              "cobranca", "dias", "liquidacao"]
    keys = ["id"] + base + extras
    cols = ["rowid"] + base + [c if c in existentes else "NULL" for c in extras]
    q = ("SELECT " + ", ".join(cols) + " FROM parcelas WHERE " +
         " AND ".join(where) + " ORDER BY venc, rowid")
    rows = con.execute(q, params).fetchall()
    con.close()
    return jsonify(parcelas=[dict(zip(keys, r)) for r in rows],
                   reimportar=not set(extras) <= existentes)


@app.route("/api/importar", methods=["POST"])
def importar():
    if STATUS["state"] == "running":
        return jsonify(erro="Já existe uma importação em andamento."), 409
    f = request.files.get("arquivo")
    if not f or not f.filename:
        return jsonify(erro="Selecione um arquivo."), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return jsonify(erro="Use um arquivo .xlsx ou .csv."), 400
    path = os.path.join(UPLOADS, "import" + ext)
    f.save(path)
    STATUS.update(state="running", msg="Iniciando...", rows=0, situacoes={})
    threading.Thread(target=do_import, args=(path,), daemon=True).start()
    return jsonify(ok=True)


@app.route("/api/status")
def status():
    return jsonify(STATUS=STATUS, base=db_info())


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
PAGE = r"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Consulta de Parcelas v4</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f4f4;color:#111;font-size:14px}
 .wrap{max-width:1100px;margin:0 auto;padding:16px}
 .box{background:#fff;border:1px solid #ccc;padding:12px 14px;margin-bottom:12px}
 .row{display:flex;gap:14px;align-items:center;flex-wrap:wrap}
 input[type=text],input[type=date]{padding:6px 8px;border:1px solid #999;font-size:14px}
 input#doc{width:220px}
 button{padding:6px 14px;font-size:14px;cursor:pointer}
 h2{font-size:14px;margin:0 0 10px}
 .base{font-size:12px;color:#555}
 .erro{color:#b00020;margin-top:8px}
 .grid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px 22px}
 .f{display:flex;align-items:center;justify-content:space-between;gap:8px}
 .f label{white-space:nowrap}
 .f output{display:block;min-width:105px;text-align:right;background:#ececec;
   border:1px solid #888;padding:3px 6px}
 .total label,.total output{font-weight:700}
 .total output{font-size:15px}
 @media(max-width:850px){.grid{grid-template-columns:1fr 1fr}}
 @media(max-width:520px){.grid{grid-template-columns:1fr}}
 #imp{font-size:12px;margin-top:6px}
 .tw{max-height:55vh;min-height:140px;overflow:auto;border:1px solid #bbb}
 table{border-collapse:collapse;width:100%;font-size:13px}
 th,td{padding:4px 8px;border-bottom:1px solid #e3e3e3;white-space:nowrap}
 th{position:sticky;top:0;background:#e9e9e9;text-align:left}
 td.n,th.n{text-align:right}
 tbody tr{cursor:pointer}
 tbody tr:hover{background:#f0f6ff}
 tbody tr.sel{background:#bcd8ff;box-shadow:inset 4px 0 0 #1a64d6;font-weight:600}
 .bd{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;background:#e6e6e6}
 tr.cancelado .bd{background:#fbd5d5}tr.aberto .bd{background:#d8f0d8}
 td.c,th.c{width:28px;text-align:center}
 tr.liquidado td,tr.renegociado td{color:#777}
 tr.cancelado td{color:#c00000}
 .sub{font-size:12px;color:#555;margin:6px 0}
</style></head><body><div class="wrap">

<div class="box">
 <div class="row">
  <label><input type="radio" name="tipo" value="cpf" checked> CPF</label>
  <label><input type="radio" name="tipo" value="cnpj"> CNPJ</label>
  <input id="doc" type="text" placeholder="Somente números" inputmode="numeric" autofocus>
  <button id="btn">Consultar</button>
 </div>
 <div class="row" style="margin-top:10px">
  <b>Mostrar:</b>
  <label><input type="radio" name="filtro" value="todas" checked> Todas as parcelas</label>
  <label><input type="radio" name="filtro" value="aberto"> Em aberto</label>
  <label><input type="radio" name="filtro" value="liquidadas"> Liquidadas</label>
  <label><input type="radio" name="filtro" value="canceladas"> Canceladas</label>
  <label><input type="radio" name="filtro" value="renegociadas"> Renegociadas</label>
 </div>
 <div class="row" style="margin-top:10px">
  <b>Vencimento:</b> de <input id="de" type="date"> até <input id="ate" type="date">
  <button id="limpa" type="button">Limpar período</button>
 </div>
 <div id="erro" class="erro"></div>
</div>

<div class="box">
 <h2>Informações Parcelas Filtradas</h2>
 <div class="grid">
  <div class="f"><label>Qtde Parcelas:</label><output id="qtde">0</output></div>
  <div class="f"><label>Total Original:</label><output id="original"></output></div>
  <div class="f"><label>Juros a Receber:</label><output id="juros"></output></div>
  <div class="f total"><label>Total a Receber:</label><output id="a_receber"></output></div>

  <div class="f"><label>Primeiro Vencimento:</label><output id="primeiro"></output></div>
  <div class="f"><label>Total Saldo:</label><output id="saldo"></output></div>
  <div class="f"><label>Multa a Receber:</label><output id="multa"></output></div>
  <div></div>

  <div class="f"><label>Último Vencimento:</label><output id="ultimo"></output></div>
  <div class="f"><label>Total Liquidado:</label><output id="liquidado"></output></div>
  <div class="f"><label>Total Cancelado:</label><output id="cancelado"></output></div>
  <div class="f"><label>Total Renegociado:</label><output id="renegociado"></output></div>
 </div>
</div>

<div class="box">
 <h2>Linhas da planilha — documento informado no período selecionado <span id="cont" class="base"></span></h2>
 <div class="sub"><span id="sel-info"></span> <button id="sel-clr" type="button">Limpar seleção</button>
  <button id="sel-all" type="button">Selecionar todas</button></div>
 <div id="aviso" class="erro"></div>
 <div class="tw"><table>
  <thead><tr><th class="c"></th><th>Filial</th><th>Cód. Cliente</th><th>Nome</th><th>Nº Documento</th>
  <th>Parcela</th><th>Emissão</th><th>Vencimento</th><th>Vencto Original</th><th>Tipo de Cobrança</th>
  <th>Situação</th><th class="n">Valor Original</th><th class="n">Valor Saldo</th>
  <th class="n">Dias Atraso</th><th class="n">Juros</th><th class="n">Multa</th>
  <th>Data Liquidação</th><th class="n">Valor Liquidado</th></tr></thead>
  <tbody id="lista"></tbody></table></div>
</div>

<div class="box base">
 v4 · <span id="base">Carregando...</span>
 &nbsp; <input type="file" id="arq" accept=".xlsx,.xlsm,.csv">
 <button id="imp-btn" type="button">Importar / atualizar planilha</button>
 <div id="imp"></div>
</div>
</div>

<script>
const $=id=>document.getElementById(id);
const money=v=>'R$ '+Number(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const dateBR=s=>s?s.split('-').reverse().join('/'):'';
const radio=n=>document.querySelector('input[name='+n+']:checked').value;

let PARC=[],SEL=new Set();
const liqv=p=>p.status==='liquidado'?(p.recebido!==null?p.recebido:p.orig):0;
const aberto=p=>p.status==='aberto';
function pinta(){
  const base=SEL.size?PARC.filter(p=>SEL.has(p.id)):PARC;
  const sum=f=>base.reduce((t,p)=>t+f(p),0)/100;
  const v=base.map(p=>p.venc).filter(Boolean).sort();
  $('qtde').textContent=base.length;
  $('primeiro').textContent=dateBR(v[0]);$('ultimo').textContent=dateBR(v[v.length-1]);
  $('original').textContent=money(sum(p=>p.orig));
  $('saldo').textContent=money(sum(p=>p.saldo));
  $('liquidado').textContent=money(sum(liqv));
  $('juros').textContent=money(sum(p=>aberto(p)?p.juros:0));
  $('multa').textContent=money(sum(p=>aberto(p)?p.multa:0));
  $('a_receber').textContent=money(sum(p=>aberto(p)?p.saldo+p.juros+p.multa:0));
  $('cancelado').textContent=money(sum(p=>p.status==='cancelado'?p.orig:0));
  $('renegociado').textContent=money(sum(p=>p.status==='renegociado'?p.orig:0));
  $('sel-info').textContent=SEL.size?(SEL.size+' de '+PARC.length+' selecionada(s) - o painel mostra só a seleção.'):
    (PARC.length+' parcela(s) - o painel mostra todas. Clique numa linha para ver só ela (Ctrl+clique ou caixas para várias).');
  document.querySelectorAll('#lista tr[data-id]').forEach(tr=>{const on=SEL.has(+tr.dataset.id);
    tr.classList.toggle('sel',on);tr.querySelector('input').checked=on;});
}
const esc=t=>String(t??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
let MSG='Informe o CPF/CNPJ e clique em Consultar.';
function lista(){
  const col=18;
  if(!PARC.length){$('lista').innerHTML='<tr><td colspan="'+col+'" style="cursor:default;color:#555;padding:14px">'+esc(MSG)+'</td></tr>';$('cont').textContent='';return;}
  $('cont').textContent='('+PARC.length+' linha(s))';
  const t=x=>esc(x);
  $('lista').innerHTML=PARC.map(p=>'<tr class="'+p.status+'" data-id="'+p.id+'"><td class="c"><input type="checkbox"></td>'+
   '<td>'+t(p.filial)+'</td><td>'+t(p.codcli)+'</td><td>'+t(p.nome)+'</td><td>'+t(p.numdoc)+'</td><td>'+t(p.parcela)+
   '</td><td>'+dateBR(p.emissao)+'</td><td>'+dateBR(p.venc)+'</td><td>'+dateBR(p.venc_orig)+'</td><td>'+t(p.cobranca)+
   '</td><td><span class="bd">'+t(p.situacao)+'</span></td><td class="n">'+money(p.orig/100)+
   '</td><td class="n">'+money(p.saldo/100)+'</td><td class="n">'+t(p.dias)+
   '</td><td class="n">'+money(p.juros/100)+'</td><td class="n">'+money(p.multa/100)+
   '</td><td>'+dateBR(p.liquidacao)+'</td><td class="n">'+money(liqv(p)/100)+'</td></tr>').join('');
}
$('lista').onclick=e=>{
  const tr=e.target.closest('tr');if(!tr||!tr.dataset.id)return;
  const id=+tr.dataset.id;
  if(e.ctrlKey||e.metaKey||e.target.type==='checkbox'){SEL.has(id)?SEL.delete(id):SEL.add(id);}
  else if(SEL.size===1&&SEL.has(id)){SEL.clear();}
  else{SEL.clear();SEL.add(id);}
  pinta();
};
$('sel-clr').onclick=()=>{SEL.clear();pinta();};
$('sel-all').onclick=()=>{PARC.forEach(p=>SEL.add(p.id));pinta();};
function zera(m){PARC=[];SEL.clear();if(m)MSG=m;$('aviso').textContent='';lista();pinta();}
async function consulta(){
  $('erro').textContent='';
  const doc=$('doc').value.replace(/\D/g,'');
  if(!doc){zera('Informe o CPF/CNPJ e clique em Consultar.');return;}
  const p=new URLSearchParams({tipo:radio('tipo'),doc,filtro:radio('filtro'),de:$('de').value,ate:$('ate').value});
  const r=await fetch('/api/consulta?'+p);const d=await r.json();
  if(!r.ok){zera('Sem linhas para exibir.');$('erro').textContent=d.erro;return;}
  PARC=d.parcelas;SEL.clear();
  MSG='Nenhuma parcela encontrada para este documento no período/filtro selecionado.';
  $('aviso').textContent=(d.reimportar&&PARC.length)?'Reimporte a planilha (botão no rodapé) para exibir Filial, Nome, Emissão e demais colunas.':'';
  lista();pinta();
}
$('btn').onclick=consulta;
$('doc').addEventListener('keydown',e=>{if(e.key==='Enter')consulta();});
document.querySelectorAll('input[name=filtro],input[name=tipo]').forEach(e=>e.onchange=consulta);
$('de').onchange=$('ate').onchange=consulta;
$('limpa').onclick=()=>{$('de').value='';$('ate').value='';consulta();};

async function base(){
  const d=await (await fetch('/api/status')).json();
  const b=d.base;
  $('base').textContent=b?('Base: '+b.parcelas.toLocaleString('pt-BR')+' parcelas · '+b.arquivo+' · importado em '+b.importado_em):'Nenhuma planilha importada.';
  const s=d.STATUS;
  if(s.state==='running'){$('imp').textContent=s.msg;setTimeout(base,1000);}
  else if(s.state==='done'){
    $('imp').textContent=s.msg+' Situações encontradas: '+Object.entries(s.situacoes).map(([k,v])=>k+' ('+v+')').join(', ');
  } else if(s.state==='error'){$('imp').textContent=s.msg;}
}
$('imp-btn').onclick=async()=>{
  const f=$('arq').files[0];if(!f){$('imp').textContent='Escolha o arquivo primeiro.';return;}
  const fd=new FormData();fd.append('arquivo',f);
  $('imp').textContent='Enviando arquivo...';
  const r=await fetch('/api/importar',{method:'POST',body:fd});const d=await r.json();
  if(!r.ok){$('imp').textContent=d.erro;return;}
  base();
};
zera();base();
</script></body></html>"""


@app.route("/")
def index():
    resp = app.make_response(render_template_string(PAGE))
    resp.headers["Cache-Control"] = "no-store"
    return resp


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    if not os.environ.get("NO_BROWSER"):
        threading.Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:%d" % port)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
