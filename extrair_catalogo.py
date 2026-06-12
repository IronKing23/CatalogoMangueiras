# -*- coding: utf-8 -*-
"""Extrai sistemas, tabelas e diagramas do catálogo Copecar CH570 (versão definitiva).
Trata: arestas duplas/triplas, tabelas-fantasma, rodapés, linhas-fantasma de edição
(3 regras) e células entrelaçadas (reparo por referência cruzada + correções manuais)."""
import io, json, re, sys
from pathlib import Path
import fitz, pdfplumber
from PIL import Image

PDF = sys.argv[1] if len(sys.argv) > 1 else 'catalogo.pdf'
COLS = ["ITEM","OEM","MANGUEIRA","COMPRIMENTO","CONEXAO1","CONEXAO2","PROTECAO"]
RE_ST = re.compile(r"\bST\d{5,}\b")
HDR = {"ITEM","OEM","MANGUEIRA","AEROQUIP","COMPRIMENTO","CORTE","CONEXÃO","PROTEÇÃO"}
# correções com evidência visual (raster da página) p/ texto entrelaçado sem referência limpa
CORRECOES = {
    ('ST932564','AXT19552'): {'MANGUEIRA':'GH781-20','COMPRIMENTO':'0,48'},
    ('ST932564','AXT19554'): {'COMPRIMENTO':'1,74'},
}

def cluster(vals, tol=8.0):
    out=[]
    for v in sorted(vals):
        if out and v-out[-1][-1]<=tol: out[-1].append(v)
        else: out.append([v])
    return [sum(g)/len(g) for g in out]

def extrai_titulo(texto):
    linhas=[l.strip() for l in (texto or '').splitlines() if l.strip()]
    buf=[]
    for l in linhas[:6]:
        buf.append(l); m=RE_ST.search(l)
        if m: return re.sub(r"\s+"," "," ".join(buf)).strip(), m.group(0)
    return None,None

def contida(a,b):
    return a[0]>=b[0]-2 and a[1]>=b[1]-2 and a[2]<=b[2]+2 and a[3]<=b[3]+2 and a!=b

def linhas_da_pagina(pg, avisos):
    tbs=pg.find_tables()
    tbs=[t for t in tbs if not any(contida(t.bbox,o.bbox) for o in tbs)]
    out=[]
    for tb in sorted(tbs,key=lambda t:t.bbox[1]):
        xs={round(c[0],1) for c in tb.cells}|{round(c[2],1) for c in tb.cells}
        ed=cluster(xs)
        if len(ed)!=8:
            avisos.append((pg.page_number,f"{len(ed)} arestas")); continue
        ytop,ybot=tb.bbox[1],tb.bbox[3]
        ws=[w for w in pg.extract_words()
            if ed[0]-3<=(w['x0']+w['x1'])/2<=ed[-1]+3 and ytop-2<=w['top']<=ybot+16]
        ws.sort(key=lambda w:(w['top'],w['x0']))
        grupos=[]
        for w in ws:
            if grupos and abs(w['top']-grupos[-1][0])<=3.5: grupos[-1][1].append(w)
            else: grupos.append([w['top'],[w]])
        kept=[]
        for top,g in grupos:
            if any(w['text'].upper() in HDR for w in g): continue
            row=['']*7
            for w in g:
                xc=(w['x0']+w['x1'])/2
                for j in range(7):
                    if ed[j]-3<=xc<ed[j+1]+(3 if j==6 else 0):
                        row[j]=(row[j]+' '+w['text']).strip(); break
            cheios=sum(1 for v in row if v)
            if cheios==0: continue
            if cheios==1 and row[0]: continue              # nº de página do rodapé
            if not row[0]:
                # regra (a): quebra legítima nunca toca OEM/MANGUEIRA/COMPRIMENTO
                if row[1] or row[2] or row[3]:
                    avisos.append((pg.page_number,('fantasma-a',row))); continue
                # regra (b): valores que duplicam células já vistas na página = fantasma
                vistos={(j,seg) for r in kept for j,v in enumerate(r) if v
                        for seg in v.split(' / ')}
                vals=[(j,v) for j,v in enumerate(row) if v]
                if vals and all((j,v) in vistos for j,v in vals):
                    avisos.append((pg.page_number,('fantasma-b',row))); continue
                # regra (c): quebra legítima -> mescla na linha anterior
                if kept:
                    for j,v in enumerate(row):
                        if v: kept[-1][j]=(kept[-1][j]+' / '+v) if kept[-1][j] else v
                continue
            kept.append(row)
        out.extend(kept)
    return out

def celula_corrompida(l):
    for k,v in l.items():
        if k.startswith('_'): continue
        if ',,' in v or '--' in v: return True
        if k in ('CONEXAO1','CONEXAO2','MANGUEIRA'):
            for seg in v.split(' / '):
                if ' ' in seg.strip(): return True
    return False

sistemas, avisos = [], []
with pdfplumber.open(PDF) as pdf:
    for n in range(6, len(pdf.pages)):
        pg=pdf.pages[n]
        titulo,st=extrai_titulo(pg.extract_text())
        if titulo: sistemas.append({"titulo":titulo,"st":st,"paginas":[n+1],"linhas":[]})
        elif sistemas: sistemas[-1]["paginas"].append(n+1)
        else: continue
        for row in linhas_da_pagina(pg,avisos):
            sistemas[-1]["linhas"].append(dict(zip(COLS,row)))

# correções manuais (texto entrelaçado, evidência: raster)
for s in sistemas:
    for l in s['linhas']:
        fix=CORRECOES.get((s['st'],l['OEM']))
        if fix and celula_corrompida(l):
            l.update(fix)
            l['_obs']='Valores corrigidos por leitura visual da página (texto sobreposto no PDF original)'

# reparo por referência cruzada de OEM
limpos={}
for s in sistemas:
    for l in s['linhas']:
        if not celula_corrompida(l) and l['OEM'] and ' / ' not in l['OEM']:
            limpos.setdefault(l['OEM'], l)
reparos=[]
for s in sistemas:
    for l in s['linhas']:
        if celula_corrompida(l):
            ref=limpos.get(l['OEM'])
            if ref:
                for k in ('MANGUEIRA','COMPRIMENTO','CONEXAO1','CONEXAO2','PROTECAO'):
                    l[k]=ref[k]
                l['_obs']='Valores restaurados de ocorrência limpa do mesmo OEM (texto sobreposto no PDF original)'
                reparos.append(('ok',s['st'],l['OEM']))
            else:
                reparos.append(('PENDENTE',s['st'],l['OEM'],dict(l)))

# imagens
doc=fitz.open(PDF); uso={}
for pg in doc:
    for img in pg.get_images(full=True): uso[img[0]]=uso.get(img[0],0)+1
fundo={x for x,q in uso.items() if q>len(doc)*0.5}
Path('diagramas').mkdir(exist_ok=True)
for f in Path('diagramas').glob('*.jpg'): f.unlink()
for s in sistemas:
    cont=0
    for npg in s['paginas']:
        for img in doc[npg-1].get_images(full=True):
            xref,w=img[0],img[2]
            if xref in fundo or w<300: continue
            cont+=1
            raw=doc.extract_image(xref)['image']
            im=Image.open(io.BytesIO(raw)).convert('RGB')
            if im.width>1000: im=im.resize((1000,int(im.height*1000/im.width)),Image.LANCZOS)
            nome=f"{s['st']}.jpg" if cont==1 else f"{s['st']}_{cont}.jpg"
            im.save(Path('diagramas')/nome,'JPEG',quality=72,optimize=True)
    s['n_imagens']=cont
pix=doc[5].get_pixmap(dpi=110)
Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB').save('diagramas/legenda.jpg','JPEG',quality=75,optimize=True)
doc.close()

Path('dados.json').write_text(json.dumps(sistemas,ensure_ascii=False,indent=1),encoding='utf-8')
total=sum(len(s['linhas']) for s in sistemas)
fa=[a for a in avisos if isinstance(a[1],tuple) and a[1][0]=='fantasma-a']
fb=[a for a in avisos if isinstance(a[1],tuple) and a[1][0]=='fantasma-b']
outros=[a for a in avisos if not isinstance(a[1],tuple)]
print(f"Sistemas: {len(sistemas)} | Linhas: {total}")
print(f"Fantasmas regra-a: {len(fa)} | regra-b: {len(fb)} | outros avisos: {outros}")
print(f"Reparos: {reparos}")
