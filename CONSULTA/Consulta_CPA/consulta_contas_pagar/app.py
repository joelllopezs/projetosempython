"""Consulta de Contas a Pagar por CPF/CNPJ.

Reproduz o painel "Informações Parcelas Filtradas" do contas a pagar
(Total de Parcelas, Total Original, Total Saldo, Primeiro e Último Vencimento)
e lista as linhas da planilha do documento informado.
A planilha (.csv ou .xlsx) é só a fonte; os dados ficam em contas_pagar.db.
"""
import csv
import datetime as dt
import json
import os
import re
import sqlite3
import threading
import unicodedata
import webbrowser

from flask import Flask, jsonify, render_template_string, request

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "contas_pagar.db")
UPLOADS = os.path.join(BASE, "_uploads")
os.makedirs(UPLOADS, exist_ok=True)

app = Flask(__name__)

# Mapeamento de colunas (nome normalizado: sem acento, minúsculo, sem pontuação)
ALIASES = {
    # usadas nos indicadores / pesquisa / filtros
    "doc": ["cnpj cpf", "cpf cnpj", "cnpj", "cpf"],
    "venc": ["vcto saldo", "vencto saldo", "vencimento saldo"],
    "orig": ["valor original"],
    "saldo": ["valor saldo"],
    "sit": ["situacao"],
    # exibidas na lista de linhas
    "codfor": ["cod fornec"],
    "filial": ["filial"],
    "nome": ["nome"],
    "numdoc": ["no documento", "n documento", "numero documento"],
    "parcela": ["parcela"],
    "numap": ["numero ap"],
    "entrada": ["entrada"],
    "venc_orig": ["vencto original"],
    "liquidacao": ["data liquidacao"],
    "historico": ["historico"],
    "despesa": ["descricao tipo despesa"],
}
REQUIRED = ["doc", "venc", "orig", "saldo", "sit"]


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def map_header(header):
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


# --------------------------- conversões ------------------------------------
_WRAP = re.compile(r'^="(.*)"$', re.S)


def clean(v):
    """Remove o invólucro ="texto" que o ERP coloca em campos de texto."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        m = _WRAP.match(v)
        if m:
            v = m.group(1).strip()
        return v or None
    return v


def money(v):
    v = clean(v)
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v * 100))
    s = str(v).replace("R$", "").strip()
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return int(round(float(s) * 100))
    except ValueError:
        return None


def to_date(v):
    v = clean(v)
    if v is None:
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
    v = clean(v)
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        d = str(int(v))
        return d.zfill(11) if len(d) <= 11 else d.zfill(14)
    return re.sub(r"\D", "", str(v))


def text(v):
    v = clean(v)
    return None if v is None else str(v)


# --------------------------- leitura (streaming) ----------------------------
def detect_encoding(path):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as f:
                while f.read(1 << 20):
                    pass
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"


def iter_rows(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl

        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        yield from wb.worksheets[0].iter_rows(values_only=True)
        wb.close()
        return
    enc = detect_encoding(path)
    with open(path, encoding=enc, newline="") as f:
        first = f.readline()
        f.seek(0)
        delim = ";" if first.count(";") >= first.count(",") else ","
        yield from csv.reader(f, delimiter=delim)


STATUS = {"state": "idle", "msg": "", "rows": 0, "situacoes": {}, "sem_doc": 0}
COLS = ["doc", "venc", "orig", "saldo", "sit", "codfor", "filial", "nome", "numdoc",
        "parcela", "numap", "entrada", "venc_orig", "liquidacao", "historico", "despesa"]


def do_import(path):
    STATUS.update(state="running", msg="Lendo arquivo...", rows=0, situacoes={}, sem_doc=0)
    try:
        rows = iter_rows(path)
        mapping = map_header(next(rows))
        get = lambda r, k: r[mapping[k]] if k in mapping and mapping[k] < len(r) else None

        con = sqlite3.connect(DB)
        con.execute("DROP TABLE IF EXISTS parcelas_new")
        con.execute(
            """CREATE TABLE parcelas_new (
                doc TEXT NOT NULL, venc TEXT, orig INTEGER, saldo INTEGER, sit TEXT,
                codfor TEXT, filial TEXT, nome TEXT, numdoc TEXT, parcela TEXT,
                numap TEXT, entrada TEXT, venc_orig TEXT, liquidacao TEXT,
                historico TEXT, despesa TEXT)"""
        )
        sql = "INSERT INTO parcelas_new VALUES (%s)" % ",".join("?" * len(COLS))
        batch, total, sem_doc, sits = [], 0, 0, {}
        for r in rows:
            if not any(r):
                continue
            doc = only_digits(get(r, "doc"))
            sit = text(get(r, "sit")) or ""
            if not doc:  # sem CPF/CNPJ não há como pesquisar
                sem_doc += 1
                continue
            sits[sit] = sits.get(sit, 0) + 1
            batch.append((
                doc, to_date(get(r, "venc")),
                money(get(r, "orig")) or 0, money(get(r, "saldo")) or 0, sit,
                text(get(r, "codfor")), text(get(r, "filial")), text(get(r, "nome")),
                text(get(r, "numdoc")), text(get(r, "parcela")), text(get(r, "numap")),
                to_date(get(r, "entrada")), to_date(get(r, "venc_orig")),
                to_date(get(r, "liquidacao")), text(get(r, "historico")),
                text(get(r, "despesa")),
            ))
            total += 1
            if len(batch) >= 5000:
                con.executemany(sql, batch)
                batch = []
                STATUS.update(rows=total, msg="Importando... %d linhas" % total)
        if batch:
            con.executemany(sql, batch)
        con.execute("DROP TABLE IF EXISTS parcelas")
        con.execute("ALTER TABLE parcelas_new RENAME TO parcelas")
        con.execute("CREATE INDEX idx_parcelas_doc ON parcelas(doc)")
        con.execute("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)")
        meta = {
            "importado_em": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "arquivo": os.path.basename(path),
            "situacoes": json.dumps(sits, ensure_ascii=False),
            "sem_doc": str(sem_doc),
        }
        for k, v in meta.items():
            con.execute("INSERT OR REPLACE INTO meta VALUES (?, ?)", (k, v))
        con.commit()
        con.close()
        STATUS.update(state="done", rows=total, situacoes=sits, sem_doc=sem_doc,
                      msg="Importação concluída: %d parcelas." % total)
    except Exception as e:  # noqa: BLE001
        STATUS.update(state="error", msg="Erro na importação: %s" % e)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


# --------------------------- consulta ---------------------------------------
def db_info():
    if not os.path.exists(DB):
        return None
    con = sqlite3.connect(DB)
    try:
        n = con.execute("SELECT COUNT(*) FROM parcelas").fetchone()[0]
        meta = dict(con.execute("SELECT k, v FROM meta").fetchall())
        return {"parcelas": n, "importado_em": meta.get("importado_em"),
                "arquivo": meta.get("arquivo"),
                "situacoes": json.loads(meta.get("situacoes") or "{}"),
                "sem_doc": int(meta.get("sem_doc") or 0)}
    except (sqlite3.Error, ValueError):
        return None
    finally:
        con.close()


@app.route("/api/consulta")
def consulta():
    tipo = request.args.get("tipo", "cnpj")
    doc = re.sub(r"\D", "", request.args.get("doc", ""))
    filtro = request.args.get("filtro", "todas")
    de = to_date(request.args.get("de"))
    ate = to_date(request.args.get("ate"))

    esperado = 11 if tipo == "cpf" else 14
    if len(doc) != esperado:
        return jsonify(erro="Informe um %s com %d dígitos." % (tipo.upper(), esperado)), 400
    if not db_info():
        return jsonify(erro="Nenhum arquivo importado ainda."), 400

    # também casa o documento gravado sem os zeros à esquerda
    candidatos = sorted({doc, doc.lstrip("0")})
    where = ["doc IN (%s)" % ",".join("?" * len(candidatos))]
    params = list(candidatos)
    if filtro != "todas":
        where.append("sit = ?")
        params.append(filtro)
    if de:
        where.append("venc >= ?")
        params.append(de)
    if ate:
        where.append("venc <= ?")
        params.append(ate)

    keys = ["id"] + COLS[1:]
    q = ("SELECT rowid, " + ", ".join(COLS[1:]) + " FROM parcelas WHERE " +
         " AND ".join(where) + " ORDER BY venc, rowid")
    con = sqlite3.connect(DB)
    rows = con.execute(q, params).fetchall()
    con.close()
    return jsonify(parcelas=[dict(zip(keys, r)) for r in rows])


@app.route("/api/importar", methods=["POST"])
def importar():
    if STATUS["state"] == "running":
        return jsonify(erro="Já existe uma importação em andamento."), 409
    f = request.files.get("arquivo")
    if not f or not f.filename:
        return jsonify(erro="Selecione um arquivo."), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return jsonify(erro="Use um arquivo .csv ou .xlsx."), 400
    path = os.path.join(UPLOADS, "import" + ext)
    f.save(path)
    STATUS.update(state="running", msg="Iniciando...", rows=0, situacoes={}, sem_doc=0)
    threading.Thread(target=do_import, args=(path,), daemon=True).start()
    return jsonify(ok=True)


@app.route("/api/status")
def status():
    return jsonify(STATUS=STATUS, base=db_info())


# --------------------------- interface --------------------------------------
PAGE = r"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Consulta Contas a Pagar</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f4f4;color:#111;font-size:14px}
 .wrap{max-width:1200px;margin:0 auto;padding:16px}
 .box{background:#fff;border:1px solid #ccc;padding:12px 14px;margin-bottom:12px}
 .row{display:flex;gap:14px;align-items:center;flex-wrap:wrap}
 input[type=text],input[type=date]{padding:6px 8px;border:1px solid #999;font-size:14px}
 input#doc{width:220px}
 button{padding:6px 14px;font-size:14px;cursor:pointer}
 h2{font-size:14px;margin:0 0 10px}
 .base{font-size:12px;color:#555}
 .erro{color:#b00020;margin-top:8px}
 .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px 28px;max-width:900px}
 .f{display:flex;align-items:center;justify-content:space-between;gap:8px}
 .f label{white-space:nowrap}
 .f output{display:block;min-width:130px;text-align:right;background:#ececec;
   border:1px solid #888;padding:3px 6px}
 @media(max-width:700px){.grid{grid-template-columns:1fr}}
 #imp{font-size:12px;margin-top:6px}
 .tw{max-height:55vh;min-height:140px;overflow:auto;border:1px solid #bbb}
 table{border-collapse:collapse;width:100%;font-size:13px}
 th,td{padding:4px 8px;border-bottom:1px solid #e3e3e3;white-space:nowrap}
 th{position:sticky;top:0;background:#e9e9e9;text-align:left}
 td.n,th.n{text-align:right}
 td.c,th.c{width:28px;text-align:center}
 tbody tr{cursor:pointer}
 tbody tr:hover{background:#f0f6ff}
 tbody tr.sel{background:#bcd8ff;box-shadow:inset 4px 0 0 #1a64d6;font-weight:600}
 .bd{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;background:#e6e6e6}
 tr.s-aberto .bd{background:#d8f0d8}tr.s-parcial .bd{background:#fff0c2}
 tr.s-cancelado .bd{background:#fbd5d5}
 tr.s-liquidado td,tr.s-renegociado td{color:#777}
 tr.s-cancelado td{color:#c00000}
 .sub{font-size:12px;color:#555;margin:6px 0}
 td.h{max-width:320px;overflow:hidden;text-overflow:ellipsis}
</style></head><body><div class="wrap">

<div class="box">
 <div class="row">
  <b>Contas a Pagar</b>
  <label><input type="radio" name="tipo" value="cnpj" checked> CNPJ</label>
  <label><input type="radio" name="tipo" value="cpf"> CPF</label>
  <input id="doc" type="text" placeholder="Somente números" inputmode="numeric" autofocus>
  <button id="btn">Consultar</button>
 </div>
 <div class="row" style="margin-top:10px"><b>Mostrar:</b><span id="filtros">
  <label><input type="radio" name="filtro" value="todas" checked> Todas as parcelas</label></span></div>
 <div class="row" style="margin-top:10px">
  <b>Vencimento:</b> de <input id="de" type="date"> até <input id="ate" type="date">
  <button id="limpa" type="button">Limpar período</button>
 </div>
 <div id="erro" class="erro"></div>
</div>

<div class="box">
 <h2>Informações Parcelas Filtradas</h2>
 <div class="grid">
  <div class="f"><label>Total de Parcelas:</label><output id="qtde">0</output></div>
  <div class="f"><label>Total Original:</label><output id="original"></output></div>
  <div class="f"><label>Primeiro Vencimento:</label><output id="primeiro"></output></div>
  <div></div>
  <div class="f"><label>Total Saldo:</label><output id="saldo"></output></div>
  <div class="f"><label>Último Vencimento:</label><output id="ultimo"></output></div>
 </div>
</div>

<div class="box">
 <h2>Linhas da planilha — documento informado no período selecionado <span id="cont" class="base"></span></h2>
 <div class="sub"><span id="sel-info"></span> <button id="sel-clr" type="button">Limpar seleção</button>
  <button id="sel-all" type="button">Selecionar todas</button></div>
 <div class="tw"><table>
  <thead><tr><th class="c"></th><th>Cod. Fornec.</th><th>Filial</th><th>Nome</th><th>Nº Documento</th>
  <th>Número AP</th><th>Parcela</th><th>Entrada</th><th>Vencto Original</th><th>Vcto Saldo</th>
  <th class="n">Valor Original</th><th class="n">Valor Saldo</th><th>Situação</th>
  <th>Data Liquidação</th><th>Histórico</th><th>Tipo Despesa</th></tr></thead>
  <tbody id="lista"></tbody></table></div>
</div>

<div class="box base">
 <span id="base">Carregando...</span>
 &nbsp; <input type="file" id="arq" accept=".csv,.xlsx,.xlsm">
 <button id="imp-btn" type="button">Importar / atualizar arquivo</button>
 <div id="imp"></div>
</div>
</div>

<script>
const $=id=>document.getElementById(id);
const money=v=>'R$ '+Number(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});
const dateBR=s=>s?s.split('-').reverse().join('/'):'';
const esc=t=>String(t??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const radio=n=>document.querySelector('input[name='+n+']:checked').value;
const slug=s=>String(s).toLowerCase().normalize('NFD').replace(/[^a-z0-9]/g,'');
let PARC=[],SEL=new Set(),MSG='Informe o CNPJ/CPF e clique em Consultar.';

function pinta(){
  const base=SEL.size?PARC.filter(p=>SEL.has(p.id)):PARC;
  const sum=f=>base.reduce((t,p)=>t+f(p),0)/100;
  const v=base.map(p=>p.venc).filter(Boolean).sort();
  $('qtde').textContent=base.length.toLocaleString('pt-BR');
  $('primeiro').textContent=dateBR(v[0]);$('ultimo').textContent=dateBR(v[v.length-1]);
  $('original').textContent=money(sum(p=>p.orig));
  $('saldo').textContent=money(sum(p=>p.saldo));
  $('sel-info').textContent=SEL.size?(SEL.size+' de '+PARC.length+' selecionada(s) - o painel mostra só a seleção.'):
    (PARC.length+' parcela(s) - o painel mostra todas. Clique numa linha para ver só ela (Ctrl+clique ou caixas para várias).');
  document.querySelectorAll('#lista tr[data-id]').forEach(tr=>{const on=SEL.has(+tr.dataset.id);
    tr.classList.toggle('sel',on);tr.querySelector('input').checked=on;});
}
function lista(){
  if(!PARC.length){$('lista').innerHTML='<tr><td colspan="16" style="cursor:default;color:#555;padding:14px">'+esc(MSG)+'</td></tr>';$('cont').textContent='';return;}
  $('cont').textContent='('+PARC.length+' linha(s))';
  $('lista').innerHTML=PARC.map(p=>'<tr class="s-'+slug(p.sit)+'" data-id="'+p.id+'"><td class="c"><input type="checkbox"></td>'+
   '<td>'+esc(p.codfor)+'</td><td>'+esc(p.filial)+'</td><td>'+esc(p.nome)+'</td><td>'+esc(p.numdoc)+
   '</td><td>'+esc(p.numap)+'</td><td>'+esc(p.parcela)+'</td><td>'+dateBR(p.entrada)+
   '</td><td>'+dateBR(p.venc_orig)+'</td><td>'+dateBR(p.venc)+
   '</td><td class="n">'+money(p.orig/100)+'</td><td class="n">'+money(p.saldo/100)+
   '</td><td><span class="bd">'+esc(p.sit)+'</span></td><td>'+dateBR(p.liquidacao)+
   '</td><td class="h" title="'+esc(p.historico)+'">'+esc(p.historico)+'</td><td>'+esc(p.despesa)+'</td></tr>').join('');
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
function zera(m){PARC=[];SEL.clear();if(m)MSG=m;lista();pinta();}

async function consulta(){
  $('erro').textContent='';
  const doc=$('doc').value.replace(/\D/g,'');
  if(!doc){zera('Informe o CNPJ/CPF e clique em Consultar.');return;}
  const p=new URLSearchParams({tipo:radio('tipo'),doc,filtro:radio('filtro'),de:$('de').value,ate:$('ate').value});
  const r=await fetch('/api/consulta?'+p);const d=await r.json();
  if(!r.ok){zera('Sem linhas para exibir.');$('erro').textContent=d.erro;return;}
  PARC=d.parcelas;SEL.clear();
  MSG='Nenhuma parcela encontrada para este documento no período/filtro selecionado.';
  lista();pinta();
}
$('btn').onclick=consulta;
$('doc').addEventListener('keydown',e=>{if(e.key==='Enter')consulta();});
document.querySelectorAll('input[name=tipo]').forEach(e=>e.onchange=consulta);
$('de').onchange=$('ate').onchange=consulta;
$('limpa').onclick=()=>{$('de').value='';$('ate').value='';consulta();};

const ORDEM=['Aberto','Parcial','Liquidado','Renegociado','Cancelado'];
function filtros(sits){
  const atual=radio('filtro');
  const nomes=Object.keys(sits).sort((a,b)=>{const ia=ORDEM.indexOf(a),ib=ORDEM.indexOf(b);
    return (ia<0?99:ia)-(ib<0?99:ib)||a.localeCompare(b);});
  $('filtros').innerHTML='<label><input type="radio" name="filtro" value="todas"> Todas as parcelas</label> '+
    nomes.map(n=>'<label><input type="radio" name="filtro" value="'+esc(n)+'"> '+esc(n||'(sem situação)')+'</label>').join(' ');
  const alvo=[...document.querySelectorAll('input[name=filtro]')].find(i=>i.value===atual)||document.querySelector('input[name=filtro]');
  alvo.checked=true;
  document.querySelectorAll('input[name=filtro]').forEach(e=>e.onchange=consulta);
}
let ultSit='';
async function base(){
  const d=await (await fetch('/api/status')).json();
  const b=d.base;
  $('base').textContent=b?('Base: '+b.parcelas.toLocaleString('pt-BR')+' parcelas · '+b.arquivo+' · importado em '+b.importado_em+
    (b.sem_doc?(' · '+b.sem_doc.toLocaleString('pt-BR')+' linhas sem CNPJ/CPF não puderam ser importadas'):'')):'Nenhum arquivo importado.';
  if(b){const k=JSON.stringify(b.situacoes);if(k!==ultSit){ultSit=k;filtros(b.situacoes);}}
  const s=d.STATUS;
  if(s.state==='running'){$('imp').textContent=s.msg;setTimeout(base,1000);}
  else if(s.state==='done'){$('imp').textContent=s.msg+' Situações: '+Object.entries(s.situacoes).map(([k,v])=>k+' ('+v+')').join(', ');}
  else if(s.state==='error'){$('imp').textContent=s.msg;}
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
    port = int(os.environ.get("PORT", 5001))
    if not os.environ.get("NO_BROWSER"):
        threading.Timer(1.0, lambda: webbrowser.open("http://127.0.0.1:%d" % port)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
