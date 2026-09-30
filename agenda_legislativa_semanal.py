# -*- coding: utf-8 -*-
"""
Agenda Legislativa Semanal — recorte da pauta do Congresso por temas de interesse.

Lê a agenda do Congresso Nacional, da Câmara dos Deputados e do Senado Federal
(plenários, comissões e outros eventos), abre a pauta de cada reunião quando ela
existe, filtra pelas palavras-chave de cada perfil de interesse e gera um relatório
em HTML — opcionalmente enviado por e-mail.

Fontes (dados abertos oficiais, sem chave de acesso):
  Câmara    https://dadosabertos.camara.leg.br/api/v2/eventos  + /eventos/{id}/pauta
  Senado    https://legis.senado.leg.br/dadosabertos/plenario/agenda/mes/{AAAAMM01}
            https://legis.senado.leg.br/dadosabertos/comissao/agenda/{ini}/{fim}
  Congresso https://legis.senado.leg.br/dadosabertos/plenario/agenda/cn/{ini}/{fim}

Uso:
  python agenda_legislativa_semanal.py                          semana corrente (seg a sex)
  python agenda_legislativa_semanal.py 2026-09-28               a semana que contém essa data
  python agenda_legislativa_semanal.py 2026-09-28 2026-10-02    intervalo exato
  python agenda_legislativa_semanal.py --proxima                próxima semana (seg a sex)
  python agenda_legislativa_semanal.py --dias=10                de hoje até daqui a 10 dias

Opções:
  --sem-email            só gera o HTML (padrão de fato, se não houver senha configurada)
  --cliente=NOME         roda um perfil só
  --sem-pauta            não abre a pauta de cada reunião (mais rápido, menos preciso)
  --ajuda                mostra este texto

E-mail: configure por variáveis de ambiente, nunca no código —
  AGENDA_EMAIL_REMETENTE, AGENDA_EMAIL_DESTINO (separados por vírgula), AGENDA_SENHA_APP
No Gmail, AGENDA_SENHA_APP é uma "senha de app" de 16 letras, não a senha da conta.

Requisito: pip install requests
Licença: MIT
"""
import sys, os, re, html, unicodedata, datetime, smtplib
from email.mime.text import MIMEText
from concurrent.futures import ThreadPoolExecutor
import requests

# No Windows o terminal costuma abrir em cp1252 e qualquer acento na saída derruba o
# programa com UnicodeEncodeError. Isto força UTF-8 na saída, sem depender do terminal.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ==== E-MAIL (por variável de ambiente) ====
EMAIL_REMETENTE = os.environ.get("AGENDA_EMAIL_REMETENTE", "")
EMAIL_DESTINO = [e.strip() for e in os.environ.get("AGENDA_EMAIL_DESTINO", "").split(",") if e.strip()]
SENHA_APP = os.environ.get("AGENDA_SENHA_APP", "")
SMTP_SERVIDOR = os.environ.get("AGENDA_SMTP_SERVIDOR", "smtp.gmail.com")
SMTP_PORTA = int(os.environ.get("AGENDA_SMTP_PORTA", "465"))

# ==== PERFIS DE INTERESSE (EDITE AQUI) ====
# Cada perfil é um conjunto de palavras-chave. Pode escrever com acento e maiúscula:
# o programa normaliza tudo ao carregar. O termo casa com pedaço de palavra
# ("mineracao" pega "Mineração" e "de mineração").
#
#   obrigatorias -> nome próprio (empresa, órgão, pessoa): se aparecer, entra sempre
#   temas        -> termos próprios do perfil: valem em qualquer comissão ou plenário
#   temas_orgao  -> termos genéricos: só valem nas comissões listadas em "comissoes"
#                   ou quando vierem acompanhados de um termo próprio
#   temas_fracos -> termos muito genéricos: só entram acompanhados de outro tema
#   comissoes    -> siglas das comissões de interesse (Câmara e Senado). O Plenário
#                   das três Casas é sempre lido, para todos os perfis.
#
# Os dois perfis abaixo são exemplos — troque pelos seus.
CLIENTES = {
    "Energia": {
        "obrigatorias": [   # nomes próprios que você acompanha
            "nome da empresa", "nome da subsidiaria",
        ],
        "temas": [   # termos próprios: valem em qualquer colegiado
            "eletricidade", "fontes de energia", "aneel", "distribuição de energia",
            "energia solar", "energia renovavel",
        ],
        "temas_orgao": [   # genéricos: só nas comissões de interesse
            "medida provisoria", "reforma tributaria", "audiencia publica",
            "energia eletrica", "tarifa", "subvencao",
        ],
        "comissoes": [   # Senado e Câmara
            "cme", "ci", "cma", "cae", "ccj", "cmads", "cft", "ccjc", "cdeics", "cmmc",
        ],
    },

    "Saneamento": {
        "obrigatorias": [
            "nome da empresa",
        ],
        "temas": [
            "saneamento basico", "esgotamento sanitario", "abastecimento de agua",
            "marco do saneamento", "residuos solidos", "economia circular",
            "reciclagem", "logistica reversa", "construcao civil", "habitacao",
        ],
        "temas_orgao": [
            "medida provisoria", "reforma tributaria", "audiencia publica",
            "fgts", "licitacao", "norma tecnica",
        ],
        "comissoes": [
            "ci", "cma", "cae", "ccj", "cdu", "cmads", "cft", "ccjc", "cdc",
        ],
    },
}
ORDEM_CLIENTES = ["Energia", "Saneamento"]

# Pauta protocolar: sozinho, isso não faz o item entrar
RUIDO = [
    "requerimento de informacao", "voto de aplauso", "homenagem",
    "sessao solene", "posse de membros", "eleicao de presidente",
]

# ---------------------------------------------------------------- utilidades

def normaliza(t):
    """Tira acento e caixa: o texto do Congresso e as palavras-chave são comparados assim."""
    t = unicodedata.normalize("NFKD", t or "").encode("ascii", "ignore").decode()
    return t.lower().strip()


def prepara_clientes(clientes):
    """Normaliza as listas uma vez, no carregamento: assim você pode escrever as
    palavras-chave com acento e maiúscula que o programa entende igual."""
    prontos = {}
    for nome, cli in clientes.items():
        prontos[nome] = {k: [normaliza(v) for v in val] if isinstance(val, list) else val
                         for k, val in cli.items()}
    return prontos


CLIENTES = prepara_clientes(CLIENTES)
RUIDO = [normaliza(r) for r in RUIDO]
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Accept": "application/json"}
API_CAMARA = "https://dadosabertos.camara.leg.br/api/v2"
API_SENADO = "https://legis.senado.leg.br/dadosabertos"


def pega_json(url, params=None):
    try:
        r = requests.get(url, headers=UA, params=params, timeout=60)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"  aviso: falhou {url.split('/')[-1]} ({e})")
        return {}


def lista(x):
    """A API do Senado devolve 1 item como dict e vários como lista."""
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


def periodo(argv):
    """Decide o intervalo lido, conforme o que você passar na linha de comando."""
    hoje = datetime.date.today()
    datas = sorted(a for a in argv if re.fullmatch(r"\d{4}-\d{2}-\d{2}", a))
    dias = next((int(a.split("=")[1]) for a in argv if a.startswith("--dias=")), None)

    if len(datas) >= 2:
        return (datetime.date.fromisoformat(datas[0]), datetime.date.fromisoformat(datas[-1]))
    if dias:
        return (hoje, hoje + datetime.timedelta(days=dias))
    base = datetime.date.fromisoformat(datas[0]) if datas else hoje
    if "--proxima" in argv:
        base = base + datetime.timedelta(days=7)
    segunda = base - datetime.timedelta(days=base.weekday())
    return (segunda, segunda + datetime.timedelta(days=4))


# ---------------------------------------------------------------- coleta

def agenda_camara(ini, fim, com_pauta=True):
    """Eventos da Câmara: plenário, comissões e outros eventos, com a pauta de cada um."""
    d = pega_json(f"{API_CAMARA}/eventos", {
        "dataInicio": ini.isoformat(), "dataFim": fim.isoformat(),
        "ordem": "ASC", "ordenarPor": "dataHoraInicio", "itens": 200})
    eventos = d.get("dados", []) or []

    def pauta(ev):
        itens = pega_json(f"{API_CAMARA}/eventos/{ev['id']}/pauta").get("dados", []) or []
        linhas = []
        for it in itens:
            p = it.get("proposicao_") or {}
            ident = " ".join(str(x) for x in (p.get("siglaTipo"), p.get("numero")) if x)
            if p.get("ano"):
                ident += f"/{p['ano']}"
            texto = (p.get("ementa") or it.get("topico") or "").strip()
            rel = (it.get("relator") or {}).get("nome")
            if rel:
                texto += f" (Relator: {rel})"
            par = (ident.strip(), texto)
            if (ident or texto) and par not in linhas:
                linhas.append(par)
        return linhas

    if com_pauta and eventos:
        with ThreadPoolExecutor(8) as ex:
            pautas = list(ex.map(pauta, eventos))
    else:
        pautas = [[] for _ in eventos]

    saida = []
    for ev, itens in zip(eventos, pautas):
        org = (ev.get("orgaos") or [{}])[0]
        sigla = (org.get("sigla") or "").upper()
        apelido = normaliza(org.get("apelido") or "")
        if sigla == "PLEN":
            bloco = "Plenário"
        elif sigla in ("EVENTOS", "") or apelido.startswith("outros"):
            bloco = "Outros eventos"
        else:
            bloco = "Comissões"
        local = (ev.get("localCamara") or {}).get("nome") or ev.get("localExterno") or ""
        saida.append({
            "casa": "Câmara dos Deputados", "bloco": bloco, "sigla": sigla,
            "quando": ev.get("dataHoraInicio", "")[:16].replace("T", " "),
            "colegiado": org.get("nome") or sigla,
            "titulo": (ev.get("descricao") or ev.get("descricaoTipo") or "").strip(),
            "situacao": ev.get("situacao", ""), "local": local,
            "itens": itens,
            "link": f"https://www.camara.leg.br/evento-legislativo/{ev['id']}",
        })
    return saida


def agenda_senado_plenario(ini, fim):
    """Sessões do Plenário do Senado, com as matérias da pauta."""
    saida, meses = [], set()
    d = ini
    while d <= fim:
        meses.add(d.strftime("%Y%m01"))
        d += datetime.timedelta(days=1)
    for mes in sorted(meses):
        dados = pega_json(f"{API_SENADO}/plenario/agenda/mes/{mes}")
        for s in lista((dados.get("AgendaPlenario", {}).get("Sessoes") or {}).get("Sessao")):
            try:
                data = datetime.date.fromisoformat(s.get("Data", "")[:10])
            except ValueError:
                continue
            if not (ini <= data <= fim):
                continue
            itens = []
            for m in lista((s.get("Materias") or {}).get("Materia")):
                ident = (m.get("DescricaoIdentificacaoMateria") or "").strip()
                texto = (m.get("Ementa") or m.get("EmentaPapeleta") or "").strip()
                if m.get("Parecer"):
                    texto += " — " + m["Parecer"]
                if (ident, texto) not in itens:
                    itens.append((ident, texto))
            saida.append({
                "casa": "Senado Federal", "bloco": "Plenário", "sigla": "PLEN",
                "quando": f"{data.isoformat()} {s.get('Hora','')}",
                "colegiado": "Plenário do Senado Federal",
                "titulo": " ".join(x for x in (s.get("NumeroSessao"), s.get("TipoSessao")) if x).strip(),
                "situacao": s.get("SituacaoSessao", ""), "local": s.get("LocalSessao", ""),
                "itens": itens, "extra": (s.get("Comunicado") or "").strip(),
                "link": ("https://www25.senado.leg.br/web/atividade/sessao-plenaria/-/pauta/"
                         + str(s.get("CodigoSessao"))),
            })
    return saida


def agenda_senado_comissoes(ini, fim):
    """Reuniões das comissões do Senado e das comissões mistas, com finalidade e convidados."""
    d = pega_json(f"{API_SENADO}/comissao/agenda/{ini.strftime('%Y%m%d')}/{fim.strftime('%Y%m%d')}")
    saida = []
    for r in lista((d.get("AgendaReuniao", {}).get("reunioes") or {}).get("reuniao")):
        col = r.get("colegiadoCriador") or {}
        casa = "Congresso Nacional" if col.get("siglaCasa") == "CN" else "Senado Federal"
        itens = []
        for parte in lista(r.get("partes")):
            ev = parte.get("evento") or {}
            fin = (ev.get("finalidade") or parte.get("nome") or "").strip()
            if fin and (parte.get("descricaoTipo") or "", fin) not in itens:
                itens.append((parte.get("descricaoTipo") or "", fin))
            for c in lista(ev.get("convidados")) + lista(ev.get("participantes")):
                nome = " — ".join(x for x in (c.get("nome"), c.get("cargo")) if x)
                if nome and ("Convidado", nome) not in itens:
                    itens.append(("Convidado", nome))
            for it in lista(parte.get("itens")) + lista(parte.get("item")):
                mat = it.get("materia") or it
                ident = (mat.get("identificacao")
                         or mat.get("descricaoIdentificacaoMateria") or "").strip()
                texto = (mat.get("ementa") or it.get("texto") or "").strip()
                if (ident or texto) and (ident, texto) not in itens:
                    itens.append((ident, texto))
        tipo = (r.get("tipo") or {}).get("descricao") or ""
        bloco = "Outros eventos" if "outros eventos" in normaliza(tipo) else "Comissões"
        saida.append({
            "casa": casa, "bloco": bloco, "sigla": (col.get("sigla") or "").upper(),
            "quando": (r.get("dataInicio") or "")[:16].replace("T", " "),
            "colegiado": col.get("nome") or col.get("sigla") or "",
            "titulo": " — ".join(x for x in (r.get("titulo"), r.get("informacao")) if x).strip(),
            "situacao": r.get("situacao", ""), "local": r.get("local", ""),
            "itens": itens,
            "link": (r.get("urlUltimaPautaCheiaPublicada")
                     or f"https://legis.senado.leg.br/comissoes/reuniao?reuniao={r.get('codigo')}"),
        })
    return saida


def agenda_congresso(ini, fim):
    """Sessões conjuntas do Congresso Nacional."""
    d = pega_json(f"{API_SENADO}/plenario/agenda/cn/{ini.strftime('%Y%m%d')}/{fim.strftime('%Y%m%d')}")
    raiz = d.get("AgendaPlenarioCN", {})
    saida = []
    for s in lista((raiz.get("Sessoes") or {}).get("Sessao")) or lista(raiz.get("Sessao")):
        itens = []
        for m in lista((s.get("Materias") or {}).get("Materia")):
            par = ((m.get("DescricaoIdentificacaoMateria") or "").strip(),
                   (m.get("Ementa") or "").strip())
            if any(par) and par not in itens:
                itens.append(par)
        saida.append({
            "casa": "Congresso Nacional", "bloco": "Plenário", "sigla": "CN",
            "quando": f"{s.get('Data','')[:10]} {s.get('Hora','')}",
            "colegiado": "Plenário do Congresso Nacional",
            "titulo": " ".join(x for x in (s.get("NumeroSessao"), s.get("TipoSessao")) if x).strip(),
            "situacao": s.get("SituacaoSessao", ""), "local": s.get("LocalSessao", ""),
            "itens": itens, "extra": (s.get("Comunicado") or "").strip(),
            "link": ("https://www.congressonacional.leg.br/pt_BR/sessoes/"
                     "agenda-do-congresso-senado-e-camara"),
        })
    return saida


def coleta(ini, fim, com_pauta=True):
    """Junta as três Casas e descarta a mesma reunião vinda por duas fontes."""
    eventos = agenda_congresso(ini, fim)
    print(f"  Congresso Nacional: {len(eventos)} sessoes")
    n = len(eventos)
    eventos += agenda_camara(ini, fim, com_pauta)
    print(f"  Camara: {len(eventos) - n} eventos")
    n = len(eventos)
    eventos += agenda_senado_plenario(ini, fim)
    eventos += agenda_senado_comissoes(ini, fim)
    print(f"  Senado: {len(eventos) - n} sessoes e reunioes")

    unicos, vistos = [], set()
    for ev in eventos:
        chave = (ev["quando"][:16], normaliza(ev["colegiado"])[:40], normaliza(ev["titulo"])[:60])
        if chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(ev)
    if len(unicos) != len(eventos):
        print(f"  ({len(eventos) - len(unicos)} repetidos descartados)")
    return unicos


# ---------------------------------------------------------------- filtro

def texto_do_item(ev):
    partes = [ev.get("titulo", ""), ev.get("colegiado", ""), ev.get("extra", "")]
    for ident, txt in ev.get("itens", []):
        partes += [ident, txt]
    return normaliza(" ".join(partes))


def relevante(ev, cli):
    """Nome do cliente entra sempre; termo próprio vale em qualquer colegiado; termo
    genérico só na comissão de interesse ou junto de um termo próprio."""
    texto = texto_do_item(ev)
    for termo in cli["obrigatorias"]:
        if termo and termo in texto:
            return [termo]

    de_interesse = (ev.get("bloco") == "Plenário"
                    or normaliza(ev.get("sigla", "")) in cli.get("comissoes", []))
    proprios = [t for t in cli["temas"] if t in texto]
    genericos = [t for t in cli.get("temas_orgao", []) if t in texto]
    if not proprios and not (de_interesse and genericos):
        return []

    # pauta apenas protocolar (homenagem, voto de aplauso, posse) sem termo próprio: fora
    if not proprios and any(r in texto for r in RUIDO):
        return []

    achados = proprios + (genericos if (de_interesse or proprios) else [])
    achados += [t for t in cli.get("temas_fracos", []) if t in texto]
    return achados


def itens_que_casam(ev, cli):
    """Destaca, dentro da pauta, as matérias que bateram com as palavras-chave."""
    termos = [t for t in cli["obrigatorias"] + cli["temas"] + cli.get("temas_orgao", []) if t]
    marcados = []
    for ident, txt in ev.get("itens", []):
        alvo = normaliza(f"{ident} {txt}")
        if any(t in alvo for t in termos):
            marcados.append((ident, txt))
    return marcados or ev.get("itens", [])[:3]


# ---------------------------------------------------------------- saída

def encurta(txt, limite=400):
    txt = re.sub(r"\s+", " ", txt or "").strip()
    if len(txt) <= limite:
        return txt
    corte = txt[:limite]
    espaco = corte.rfind(" ")
    return (corte[:espaco] if espaco > limite * 0.6 else corte).rstrip(" ,;.-") + "..."


DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
        "sexta-feira", "sábado", "domingo"]


def formata_quando(q):
    q = (q or "").strip()
    try:
        d = datetime.datetime.strptime(q[:16], "%Y-%m-%d %H:%M")
        hora = d.strftime("%Hh%M").replace("h00", "h")
        return f"{DIAS[d.weekday()]}, {d.strftime('%d/%m')} às {hora}"
    except ValueError:
        try:
            d = datetime.date.fromisoformat(q[:10])
            return f"{DIAS[d.weekday()]}, {d.strftime('%d/%m')}"
        except ValueError:
            return q


VAZIO = "Não há nenhuma pauta de interesse agendada até o momento."
VAZIO_EVENTO = "Não há nenhum evento de interesse agendado até o momento."


def bloco_html(eventos, cli):
    L = []
    for ev in sorted(eventos, key=lambda e: e.get("quando", "")):
        cab = f"{formata_quando(ev.get('quando'))} — {html.escape(ev.get('colegiado',''))}"
        if ev.get("sigla") and ev["sigla"] not in ("PLEN", "CN"):
            cab += f" ({html.escape(ev['sigla'])})"
        L.append(f"<p style='margin:10px 0 2px'><b>{cab}</b></p>")
        if ev.get("titulo"):
            linha = html.escape(encurta(ev["titulo"], 300))
            if ev.get("local"):
                linha += f" — {html.escape(ev['local'])}"
            L.append(f"<p style='margin:2px 0'>{linha}</p>")
        if ev.get("extra"):
            L.append(f"<p style='margin:2px 0'><i>{html.escape(encurta(ev['extra'], 600))}</i></p>")
        vistos = set()
        for ident, txt in itens_que_casam(ev, cli)[:12]:
            chave = (ident.strip(), encurta(txt, 120))
            if chave in vistos:
                continue
            vistos.add(chave)
            prefixo = f"<b>{html.escape(ident)}</b> — " if ident else ""
            L.append(f"<p style='margin:2px 0 2px 18px'>{prefixo}{html.escape(encurta(txt))}</p>")
        if ev.get("link"):
            L.append(f"<p style='margin:2px 0 8px 18px'>"
                     f"<a href='{ev['link']}' target='_blank'>pauta / mais informações</a></p>")
    return L


def monta_html(recortes, ini, fim, clientes):
    titulo = (f"PRÉVIA AGENDA LEGISLATIVA SEMANAL "
              f"({ini.strftime('%d/%m')}-{fim.strftime('%d/%m')})")
    L = [f"<h1 style='font-size:20px'>{titulo}</h1>"]
    casas = [("Congresso Nacional", "1) Congresso Nacional"),
             ("Câmara dos Deputados", "2) Câmara dos Deputados"),
             ("Senado Federal", "3) Senado Federal")]
    for nome in clientes:
        cli = CLIENTES[nome]
        meus = recortes.get(nome, [])
        L.append(f"<h2 style='font-size:19px;margin-top:28px;color:#073763'>"
                 f"{html.escape(nome)}</h2>")
        for casa, rotulo in casas:
            L.append(f"<p style='margin:16px 0 4px'><b>{rotulo}</b></p>")
            for bloco in ("Plenário", "Comissões"):
                L.append(f"<p style='margin:8px 0 2px'><b>{bloco}</b></p>")
                sel = [e for e in meus if e["casa"] == casa and e["bloco"] == bloco]
                L += bloco_html(sel, cli) if sel else [f"<p style='margin:2px 0'>{VAZIO}</p>"]
        L.append("<p style='margin:16px 0 4px'><b>OUTROS EVENTOS</b></p>")
        sel = [e for e in meus if e["bloco"] == "Outros eventos"]
        L += bloco_html(sel, cli) if sel else [f"<p style='margin:2px 0'>{VAZIO_EVENTO}</p>"]
    return ("<meta charset='utf-8'><body style='font-family:Source Sans Pro,Calibri,sans-serif;"
            "font-size:15px;line-height:1.45;max-width:900px;margin:auto;text-align:justify'>"
            + "".join(L) + "</body>")


def envia_email(corpo, ini, fim, total):
    if not SENHA_APP:
        print("SENHA_APP vazia: defina a variavel de ambiente ou preencha no script. "
              "E-mail nao enviado.")
        return
    msg = MIMEText(corpo, "html", "utf-8")
    msg["Subject"] = (f"Prévia Agenda Legislativa Semanal "
                      f"({ini.strftime('%d/%m')}-{fim.strftime('%d/%m')}) — {total} itens")
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = ", ".join(EMAIL_DESTINO)
    with smtplib.SMTP_SSL(SMTP_SERVIDOR, SMTP_PORTA, timeout=60) as s:
        s.login(EMAIL_REMETENTE, SENHA_APP.replace(" ", ""))
        s.sendmail(EMAIL_REMETENTE, EMAIL_DESTINO, msg.as_string())
    print("E-mail enviado para", ", ".join(EMAIL_DESTINO))


# ---------------------------------------------------------------- principal

def main():
    argv = sys.argv[1:]
    if "--ajuda" in argv or "-h" in argv:
        print(__doc__)
        return
    ini, fim = periodo(argv)
    so_cliente = next((a.split("=", 1)[1] for a in argv if a.startswith("--cliente=")), None)

    print(f"Agenda de {ini.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}...")
    eventos = coleta(ini, fim, com_pauta="--sem-pauta" not in argv)

    clientes = [c for c in ORDEM_CLIENTES if c in CLIENTES]
    if so_cliente:
        clientes = [c for c in clientes if normaliza(c) == normaliza(so_cliente)]
        if not clientes:
            print(f"Cliente '{so_cliente}' nao esta na lista. Opcoes: "
                  + ", ".join(ORDEM_CLIENTES))
            return

    recortes, total = {}, 0
    for nome in clientes:
        sel = [e for e in eventos if relevante(e, CLIENTES[nome])]
        recortes[nome] = sel
        total += len(sel)
        print(f"  {nome}: {len(sel)} itens de interesse")

    corpo = monta_html(recortes, ini, fim, clientes)
    saida = f"agenda_{ini.isoformat()}_a_{fim.isoformat()}.html"
    with open(saida, "w", encoding="utf-8") as f:
        f.write(corpo)
    print(f"{total} itens -> {saida}")

    if "--sem-email" not in argv:
        envia_email(corpo, ini, fim, total)


if __name__ == "__main__":
    main()
