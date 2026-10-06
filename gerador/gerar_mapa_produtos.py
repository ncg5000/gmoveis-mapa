#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador do Mapa dos Produtos G. Móveis  (v0 — 06/10/2026)

Lê a tabela mestre (+Tabela_gmoveis.xlsx, aba Representante) e a lista de fotos de
público Catálogo\\Fundo_infinito, monta um card por modelo e escreve:

  <saida>/index.html            site público (SEM preço)
  <saida>/dados/produtos.json   dados públicos dos cards
  <saida>/gerador/fotos_lista.json  origem -> slug (para o script de redução de fotos)
  <saida>/notas/pendencias.md   fotos faltando, nomes fora da regra, refs sem card
  <saida>/notas/nomes.csv       nome da tabela -> nome amigável (para revisão)
  --interno <arquivo.html>      versão interna COM preço, fotos por caminho relativo

Uso (casa, Anaconda):   python gerar_mapa_produtos.py
Uso (nuvem):            python gerar_mapa_produtos.py --tabela t.xlsx --listagem lista.json --saida ../
Nunca grava preço dentro da pasta de saída pública.
"""
import argparse, csv, json, os, re, sys, unicodedata, zipfile
from collections import OrderedDict, defaultdict
from datetime import date
import xml.etree.ElementTree as ET

AQUI = os.path.dirname(os.path.abspath(__file__))
ONEDRIVE = os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")),
                        "OneDrive - G MOVEIS LTDA", "Dados - Documentos")
TABELA_PADRAO = os.path.join(ONEDRIVE, "gm", "TABELA - PRODUTOS", "+Tabela_gmoveis.xlsx")
FOTOS_PADRAO = os.path.join(ONEDRIVE, "gm", "Imagens", "público Catálogo", "Fundo_infinito")
CATEGORIAS = ["CADEIRA", "BANQUETA", "POLTRONA", "MESA", "APARADOR"]
PECAS = ["ASSENTO", "ENCOSTO", "ANATOMICO"]
PASTA_CAT = {"CADEIRA": "Cadeira", "BANQUETA": "Banqueta", "POLTRONA": "Poltrona",
             "MESA": "Mesa", "APARADOR": "Mesa"}
CAT_NOME = {"CADEIRA": "Cadeiras", "BANQUETA": "Banquetas", "POLTRONA": "Poltronas",
            "MESA": "Mesas", "APARADOR": "Aparadores"}
CORES = {"azul": "#1F4E57", "bege": "#D1C2A3"}

# ---------- leitura da tabela (OOXML strict ou normal, sem dependências) ----------
def _ln(tag):  # local-name
    return tag.split("}", 1)[-1]

def ler_tabela(caminho):
    z = zipfile.ZipFile(caminho)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2alvo = {r.get("Id"): r.get("Target") for r in rels}
    alvo = None
    for s in wb.iter():
        if _ln(s.tag) == "sheet" and s.get("name") == "Representante":
            rid = [v for k, v in s.attrib.items() if _ln(k) == "id"][0]
            alvo = rid2alvo[rid]
    if not alvo:
        sys.exit("aba Representante não encontrada")
    alvo = alvo.lstrip("/")
    if not alvo.startswith("xl/"):
        alvo = "xl/" + alvo
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
            ss.append("".join(t.text or "" for t in si.iter() if _ln(t.tag) == "t"))
    linhas = []
    for row in ET.fromstring(z.read(alvo)).iter():
        if _ln(row.tag) != "row":
            continue
        cels = {}
        for c in row:
            if _ln(c.tag) != "c":
                continue
            col = re.match(r"[A-Z]+", c.get("r")).group(0)
            t = c.get("t")
            v = None
            for ch in c:
                if _ln(ch.tag) == "v":
                    v = ch.text
                elif _ln(ch.tag) == "is":
                    v = "".join(x.text or "" for x in ch.iter() if _ln(x.tag) == "t")
            if v is None:
                continue
            if t == "s":
                v = ss[int(v)]
            elif t in (None, "n"):
                try:
                    v = float(v)
                except ValueError:
                    pass
            cels[col] = v
        linhas.append(cels)
    return linhas

def produtos_da_tabela(linhas):
    prods = OrderedDict()
    validade = None
    for r in linhas:
        if r.get("L") == "Valido até:" and r.get("M") is not None:
            validade = r["M"]
        ref, nome = r.get("C"), r.get("D")
        if ref is None or nome is None:
            continue
        ref = str(int(ref)) if isinstance(ref, float) else str(ref).strip()
        if not re.fullmatch(r"\d{3,5}", ref):
            continue
        p = prods.setdefault(ref, {"ref": ref, "nome": re.sub(r"\s+", " ", str(nome)).strip(),
                                   "A": None, "L": None, "C": None, "tp": [], "preco": {}})
        A = r.get("E")
        if isinstance(A, float):
            p["A"], p["L"], p["C"] = A, r.get("F"), r.get("G")
        tp = r.get("H")
        if tp is not None and str(tp) not in p["tp"]:
            p["tp"].append(str(tp))
        if tp is not None and isinstance(r.get("J"), float):
            p["preco"][str(tp)] = round(r["J"], 2)   # À VISTA — só sai na versão interna
    return prods, validade

# ---------- nomes amigáveis ----------
ABREV = [
    (r"\bENC\.? TODO ESTOF\.?", "encosto todo estofado"),
    (r"\bENC\.? TODO MAD\.?", "encosto todo madeira"),
    (r"\bENC\.? ESTOF\.? ?\+ ?MAD\.?", "encosto estofado e madeira"),
    (r"\bESTOF\.? ?\+ ?MAD\.?", "estofado e madeira"),
    (r"\bTODA MAD\.?", "toda madeira"),
    (r"\bASSENTO ESTOF\.?", "assento estofado"),
    (r"\bSEM BRACO\b", "sem braço"), (r"\bCOM BRAÇO\b|\bCOM BRACO\b", "com braço"),
    (r"\bC\.? ?BASE\b", "com base"), (r"\bC\.? ?PES\b|\bC\.? ?PÉS\b", "com pés"),
    (r"\bC/GIR\b", "centro giratório"),
    (r"\bGIRATÓRIA\b", "giratória"), (r"\bGIR\b", "giratória"), (r"\bFIXA\b", "fixa"),
    (r"\bFIXO\b", "fixo"),
    (r"\bTP/MADEIRA\b", "tampo madeira"), (r"\bTP/VIDRO\b", "tampo vidro"),
    (r"\bVID\.? ?DET\.? ?MAD\.?", "vidro com detalhe madeira"),
    (r"\bVID NA BORDA\b", "vidro na borda"), (r"\bVID NO CENTRO\b", "vidro no centro"),
    (r"\bVID\.?", "vidro"), (r"\bMAD\.?", "madeira"), (r"\bESTOF\.?", "estofado"),
    (r"\bENC\.?", "encosto"), (r"\bJOGOS\b", "jogos"), (r"\bQUADRADA\b", "quadrada"),
    (r"\bREDONDA\b", "redonda"),
]
def capitaliza(s):
    return " ".join(LINHA_NOMES.get(w, w if w in ("II", "III") else w.capitalize()) for w in s.split())

def amigavel(cat, linha, resto):
    """'CADEIRA','CAPRI','ENC TODO ESTOF' -> ('Cadeira Capri', 'encosto todo estofado')"""
    titulo = f"{cat.capitalize()} {capitaliza(linha)}"
    d = resto.replace("!", "").strip()
    for pat, sub in ABREV:
        d = re.sub(pat, sub, d)
    d = re.sub(r"\s+", " ", d).strip(" .")
    return titulo, d

LINHA_NOMES = {"PERUIBE": "Peruíbe", "MONACO": "Mônaco", "COPENHAGE": "Copenhague"}

def dividir_nome(nome):
    """'MESA IPANEMA II 1,60X0,90 MAD!' -> cat, linha, tamanho, resto
       'CADEIRA CAPRI ENC TODO ESTOF'  -> CADEIRA, CAPRI, None, 'ENC TODO ESTOF'"""
    toks = nome.replace("!", "").split()
    cat = toks[0]
    linha = toks[1:2]
    i = 2
    if i < len(toks) and toks[i] in ("II", "III"):
        linha.append(toks[i]); i += 1
    rest = " ".join(toks[i:])
    m = re.search(r"(\d,\d\d)\s*(?:X\s*(\d,\d\d))?", rest)
    tam = None
    if m:
        tam = m.group(1) + (" × " + m.group(2) if m.group(2) else "")
        rest = (rest[:m.start()] + " " + rest[m.end():])
    return cat, " ".join(linha), tam, re.sub(r"\s+", " ", rest).strip()

# ---------- fotos ----------
def slug(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-").lower()
    return s[:80]

def refs_da_foto(nome):
    base = re.sub(r"\.[^.]+$", "", nome)
    m = re.match(r"\(?([\d\-\s]+)\)?", base)
    refs = re.findall(r"\d{4}", m.group(1)) if m else []
    extra = bool(re.search(r"_\d+|\(\d\)|_[A-Z]+\d*$|_\d+_$", base)) or \
        bool(re.search(r"\b[ABC]\d{3}\b|\b0\d\d-[A-Z]+", base))
    return refs, extra

def listar_fotos(pasta_fundo=None, listagem=None):
    """-> {categoria_pasta: [nome_arquivo, ...]} (só a raiz de cada categoria)"""
    out = defaultdict(list)
    if listagem:
        d = json.load(open(listagem, encoding="utf-8"))
        for e in d["entries"]:
            p = e["name"].replace("/", "\\").split("\\")
            if e["type"] == "file" and len(p) == 2 and p[1].lower().endswith((".jpg", ".jpeg", ".png")):
                out[p[0]].append(p[1])
        return out
    for cat in sorted(set(PASTA_CAT.values())):
        d = os.path.join(pasta_fundo, cat)
        if os.path.isdir(d):
            out[cat] = sorted(f for f in os.listdir(d)
                              if f.lower().endswith((".jpg", ".jpeg", ".png")))
    return out

# ---------- cards ----------
def montar_cards(prods, fotos, overrides):
    cards = OrderedDict()
    pend = {"sem_foto": [], "foto_sem_card": [], "ref_duplicada": [], "nome_diverge": []}
    for p in prods.values():
        cat, linha, tam, resto = dividir_nome(p["nome"])
        if cat not in CATEGORIAS:
            continue
        if cat in ("MESA", "APARADOR"):
            chave = f"{cat}|{linha}|{resto}"
        else:
            chave = p["ref"]
        c = cards.get(chave)
        if not c:
            titulo, desc = amigavel(cat, linha, resto)
            c = cards[chave] = {"id": chave, "cat": cat, "linha": capitaliza(linha), "titulo": titulo,
                                "desc": desc, "refs": [], "itens": [], "classes": [],
                                "fotos": [], "principal": None}
        c["refs"].append(p["ref"])
        c["itens"].append({"ref": p["ref"], "nome": p["nome"], "tam": tam,
                           "A": p["A"], "L": p["L"], "C": p["C"], "preco": p["preco"]})
        for t in p["tp"]:
            if t not in c["classes"]:
                c["classes"].append(t)
    # fotos
    ref2card = {}
    for c in cards.values():
        for r in c["refs"]:
            ref2card[r] = c["id"]
    lista = []
    for pasta, arqs in fotos.items():
        for a in arqs:
            refs, extra = refs_da_foto(a)
            alvo = {ref2card[r] for r in refs if r in ref2card}
            if len(alvo) > 1:
                pend["ref_duplicada"].append(f"{pasta}\\{a} → refs em mais de um card ({', '.join(sorted(alvo))})")
            if not alvo:
                pend["foto_sem_card"].append(f"{pasta}\\{a}")
                continue
            cid = alvo.pop()
            s = slug(re.sub(r"\.[^.]+$", "", a))
            if any(f["slug"] == s for f in cards[cid]["fotos"]):
                continue   # mesma foto em .jpg e .png
            cards[cid]["fotos"].append({"origem": f"{pasta}\\{a}", "slug": s, "extra": extra, "nrefs": len(refs)})
            lista.append({"origem": f"{pasta}\\{a}", "slug": s})
            # nome da foto diz GIR/FIXA mas o produto não (ou vice-versa) → provável troca de ref
            c = cards[cid]
            if len(c["refs"]) == 1:
                for tok in ("GIR", "FIXA"):
                    if (re.search(rf"\b{tok}", a.upper()) is not None) != (re.search(rf"\b{tok}", c["itens"][0]["nome"]) is not None):
                        pend["nome_diverge"].append(f"{pasta}\\{a} ↔ tabela {c['refs'][0]} = {c['itens'][0]['nome']}")
                        break
    for c in cards.values():
        if not c["fotos"]:
            pend["sem_foto"].append(f"{c['titulo']} — {c['desc']} ({', '.join(c['refs'])})")
            continue
        ov = overrides.get(c["refs"][0]) or overrides.get(c["id"])
        cand = [f for f in c["fotos"] if ov and f["origem"].endswith(ov)]
        if not cand:
            cand = sorted([f for f in c["fotos"] if not f["extra"]] or c["fotos"],
                          key=lambda f: (-f["nrefs"], len(f["origem"]), f["origem"]))
        c["principal"] = cand[0]["slug"]
        c["fotos"].sort(key=lambda f: (f["slug"] != c["principal"], f["origem"]))
    return cards, pend, lista

# ---------- quadros de combinações (internos) ----------
AMOSTRAS_ACAB = {"018": "018-Pinhao.jpg", "032": "032-Capuccino.jpg", "007": "007-Tabaco.jpg", "099": "099-Amêndoa.jpg",
                 "017": "017-Branco-cor.jpg", "030": "030-OffWhite.jpg", "029": "029-Preto.jpg"}

def listar_celulas(pasta_fundo=None, listagem=None):
    """-> {ref: {sku: (caminho relativo a Fundo_infinito, data 'dd/mm/aaaa')}}, {ttt: caminho amostra tecido}"""
    cel, tec = defaultdict(dict), {}
    def guarda(partes, mtime):
        # partes = [Cat, 'Variações_Acabamentos', X, arquivo]
        if len(partes) != 4 or partes[1] != "Variações_Acabamentos":
            return
        rel = "\\".join(partes)
        if partes[2] == "tecidos":
            m = re.fullmatch(r"(\d{3})\.(png|jpg)", partes[3], re.I)
            if m:
                tec[m.group(1)] = rel
        elif re.fullmatch(r"\d{4}", partes[2]):
            m = re.fullmatch(r"(\d{4}\d{3}-\d{3})\.(jpg|jpeg|png)", partes[3], re.I)
            if m:
                from datetime import datetime
                cel[partes[2]][m.group(1)] = (rel, datetime.fromtimestamp(mtime / 1000).strftime("%d/%m/%Y") if mtime else "")
    if listagem:
        for e in json.load(open(listagem, encoding="utf-8"))["entries"]:
            if e["type"] == "file":
                guarda(e["name"].replace("/", "\\").split("\\"), e.get("mtimeMs"))
        return cel, tec
    for cat in sorted(set(PASTA_CAT.values())):
        va = os.path.join(pasta_fundo, cat, "Variações_Acabamentos")
        if not os.path.isdir(va):
            continue
        for sub in os.listdir(va):
            d = os.path.join(va, sub)
            if os.path.isdir(d):
                for f in os.listdir(d):
                    p = os.path.join(d, f)
                    if os.path.isfile(p):
                        guarda([cat, "Variações_Acabamentos", sub, f], os.path.getmtime(p) * 1000)
    return cel, tec

def gerar_quadros(cards, cel, tec, cfg, shopify, pasta_interna, meta):
    """Escreve <pasta_interna>/mapa_produtos/quadro_<ref>.html; devolve {ref: {n, geradas, shopify}} e a lista de miniaturas."""
    q = cfg["quadro"]; fam = q["familias"]; acabs = q["acabamentos"]
    todos_tec = [t for f in fam.values() for t in f]
    total = len(todos_tec) * len(acabs)
    out_dir = os.path.join(pasta_interna, "mapa_produtos")
    os.makedirs(out_dir, exist_ok=True)
    tpl = open(os.path.join(AQUI, "modelo_quadro.html"), encoding="utf-8").read()
    resumo, lista = {}, []
    for a, arq in AMOSTRAS_ACAB.items():
        lista.append({"origem": "Acabamentos\\" + arq, "destino": f"amostras\\a{a}.jpg", "px": q["amostra_px"]})
    for t, rel in tec.items():
        lista.append({"origem": "Fundo_infinito\\" + rel, "destino": f"amostras\\t{t}.jpg", "px": q["amostra_px"]})
    for ref, skus in sorted(cel.items()):
        card = next((c for c in cards.values() if ref in c["refs"]), None)
        if not card:
            continue
        loja = set(shopify.get("produtos", {}).get(ref, {}).get("skus", []))
        celulas = []
        for t in todos_tec:
            for a in acabs:
                sku = f"{ref}{t}-{a}"
                arq = skus.get(sku)
                if arq:
                    lista.append({"origem": "Fundo_infinito\\" + arq[0], "destino": f"{ref}\\{sku}.jpg", "px": q["celula_px"]})
                celulas.append({"sku": sku, "tecido": t, "acab": a, "arquivo": bool(arq),
                                "original": "../Fundo_infinito/" + arq[0].replace("\\", "/") if arq else "",
                                "mini": f"celulas/{ref}/{sku}.jpg", "data": arq[1] if arq else "", "shopify": sku in loja})
        dados = {"ref": ref, "total": total, "gerado": meta["gerado"], "pasta_celulas": f"Variações_Acabamentos\\{ref}\\",
                 "shopify_consultado": shopify.get("consultado", "") if ref in shopify.get("produtos", {}) else "",
                 "familias": [{"nome": n, "tecidos": [{"cod": t, "amostra": f"celulas/amostras/t{t}.jpg"} for t in ts]} for n, ts in fam.items()],
                 "acabamentos": [{"cod": a, "nome": n, "amostra": f"celulas/amostras/a{a}.jpg"} for a, n in acabs.items()],
                 "celulas": celulas}
        foto = next((f for f in card["fotos"] if f["slug"] == card["principal"]), None)
        html_q = (tpl.replace("/*QUADRO*/{}", json.dumps(dados, ensure_ascii=False, separators=(",", ":")))
                  .replace("/*TITULO*/", card["titulo"]).replace("/*DESC*/", card["desc"]).replace("/*REF*/", ref)
                  .replace("/*LOGO*/", "../" + LOGO_INTERNO)
                  .replace("/*FOTO_PRODUTO*/", "../Fundo_infinito/" + foto["origem"].replace("\\", "/") if foto else ""))
        open(os.path.join(out_dir, f"quadro_{ref}.html"), "w", encoding="utf-8").write(html_q)
        resumo[ref] = {"total": total, "geradas": sum(1 for c in celulas if c["arquivo"]),
                       "shopify": sum(1 for c in celulas if c["shopify"]), "pagina": f"mapa_produtos/quadro_{ref}.html"}
    json.dump(lista, open(os.path.join(AQUI, "celulas_lista.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    return resumo

# ---------- HTML ----------
LOGO_INTERNO = "../Logo G Móveis/2026-Logo-600px-gmoveis-%231F4E57.png"  # relativo a 'público Catálogo\'

def html(cards, meta, interno=False, prefixo_fotos="fotos/", fotos_originais=None, quadros=None):
    dados = []
    for c in cards.values():
        d = {k: c[k] for k in ("id", "cat", "linha", "titulo", "desc", "refs", "classes", "principal")}
        d["catNome"] = CAT_NOME[c["cat"]]
        qd = next((quadros[r] for r in c["refs"] if quadros and r in quadros), None)
        if qd:
            d["quadro"] = {k: v for k, v in qd.items() if interno or k != "pagina"}
        d["itens"] = [{k: v for k, v in it.items() if interno or k != "preco"} for it in c["itens"]]
        if fotos_originais:
            d["fotos"] = [{"src": fotos_originais(f["origem"]), "g": fotos_originais(f["origem"])} for f in c["fotos"]]
        else:
            d["fotos"] = [{"src": f"{prefixo_fotos}{f['slug']}.jpg", "g": f"{prefixo_fotos}g/{f['slug']}.jpg"} for f in c["fotos"]]
        dados.append(d)
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":"))
    mj = json.dumps(meta, ensure_ascii=False)
    tpl = open(os.path.join(AQUI, "modelo.html"), encoding="utf-8").read()
    out = (tpl.replace("/*DADOS*/[]", js).replace("/*META*/{}", mj)
              .replace("<!--INTERNO-->", "1" if interno else "0"))
    if interno:
        out = out.replace('src="assets/logo.png"', f'src="{LOGO_INTERNO}"').replace('href="assets/logo.png"', f'href="{LOGO_INTERNO}"')
        out = out.replace("<title>Mapa dos Produtos · G. Móveis</title>", "<title>Mapa dos Produtos · G. Móveis — INTERNO</title>")
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tabela", default=TABELA_PADRAO)
    ap.add_argument("--fotos", default=FOTOS_PADRAO, help="pasta Fundo_infinito")
    ap.add_argument("--listagem", help="json de listagem (uso na nuvem) em vez da pasta")
    ap.add_argument("--saida", default=os.path.join(AQUI, ".."))
    ap.add_argument("--interno", help="grava também a versão interna COM preço neste arquivo")
    a = ap.parse_args()
    cfg = json.load(open(os.path.join(AQUI, "config.json"), encoding="utf-8"))
    linhas = ler_tabela(a.tabela)
    prods, validade = produtos_da_tabela(linhas)
    fotos = listar_fotos(a.fotos, a.listagem)
    cards, pend, lista = montar_cards(prods, fotos, cfg.get("fotos_principais", {}))
    saida = os.path.abspath(a.saida)
    meta = {"gerado": date.today().isoformat(), "tabela": os.path.basename(a.tabela),
            "validade": str(validade)[:10] if validade else "", "site": cfg.get("site", ""),
            "n_cards": len(cards)}
    os.makedirs(os.path.join(saida, "dados"), exist_ok=True)
    os.makedirs(os.path.join(saida, "notas"), exist_ok=True)
    # quadros de combinações (só na saída interna)
    quadros = None
    if a.interno:
        cel, tec = listar_celulas(a.fotos, a.listagem)
        shop_path = os.path.join(saida, "dados", "shopify_midias.json")
        shopify = json.load(open(shop_path, encoding="utf-8")) if os.path.exists(shop_path) else {}
        quadros = gerar_quadros(cards, cel, tec, cfg, shopify, os.path.dirname(os.path.abspath(a.interno)), meta)
        json.dump(quadros, open(os.path.join(saida, "dados", "quadros.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    else:
        qp = os.path.join(saida, "dados", "quadros.json")
        quadros = json.load(open(qp, encoding="utf-8")) if os.path.exists(qp) else None
    # público
    pub = []
    for c in cards.values():
        pub.append({**{k: c[k] for k in ("id", "cat", "linha", "titulo", "desc", "refs", "classes", "principal")},
                    "itens": [{k: v for k, v in it.items() if k != "preco"} for it in c["itens"]],
                    "fotos": [f["slug"] for f in c["fotos"]]})
    json.dump({"meta": meta, "cards": pub}, open(os.path.join(saida, "dados", "produtos.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    json.dump(lista, open(os.path.join(AQUI, "fotos_lista.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    open(os.path.join(saida, "index.html"), "w", encoding="utf-8").write(html(cards, meta, quadros=quadros))
    with open(os.path.join(saida, "notas", "nomes.csv"), "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ref", "nome na tabela", "titulo", "descricao amigavel"])
        for c in cards.values():
            for it in c["itens"]:
                w.writerow([it["ref"], it["nome"], c["titulo"], c["desc"]])
    with open(os.path.join(saida, "notas", "pendencias.md"), "w", encoding="utf-8") as f:
        f.write(f"# Pendências do mapa — gerado {meta['gerado']}\n\n")
        for k, tit in (("sem_foto", "Cards sem foto"), ("foto_sem_card", "Fotos sem produto na tabela"),
                       ("ref_duplicada", "Fotos com refs em mais de um card"),
                       ("nome_diverge", "Nome da foto diverge da tabela (GIR/FIXA) — conferir ref")):
            f.write(f"## {tit} ({len(pend[k])})\n" + "".join(f"- {x}\n" for x in pend[k]) + "\n")
    if a.interno:
        def rel(origem):  # arquivo interno fica em 'público Catálogo\'; fotos em Fundo_infinito\<cat>\
            return "Fundo_infinito/" + origem.replace("\\", "/")
        open(a.interno, "w", encoding="utf-8").write(html(cards, meta, interno=True, fotos_originais=rel, quadros=quadros))
    print(f"{len(cards)} cards · {sum(len(c['fotos']) for c in cards.values())} fotos · "
          f"sem foto {len(pend['sem_foto'])} · fotos sem card {len(pend['foto_sem_card'])}")

if __name__ == "__main__":
    main()
