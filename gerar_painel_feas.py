
import pandas as pd
import os, json, warnings, re
from datetime import date
warnings.filterwarnings('ignore')

BASE_PATH   = r"C:\Users\catha\OneDrive - CAMG\SGFEAS - FEAS_BI"
SCRIPT_PATH = r"C:\Users\catha\OneDrive - CAMG\SGFEAS - FEAS_BI\FEAS_Scripts"
UO = 4251

ARQ = {
    "orc": os.path.join(BASE_PATH, "RAW_Orcamento", "orcamento.xlsx"),
    "des": os.path.join(BASE_PATH, "RAW_Despesa",   "despesa.xlsx"),
    "qlik": os.path.join(BASE_PATH, "PCASP_QlikView","Saldos_Contas FEAS.xls"),
}
HTML_SAIDA = os.path.join(BASE_PATH, "painel_feas.html")
TEMPLATE   = os.path.join(SCRIPT_PATH, "template_feas.html")

def ler(path, header=3):
    df = pd.read_excel(path, header=header)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")].dropna(how="all")
    if "Unidade Orçamentária - Código" in df.columns:
        df = df[df["Unidade Orçamentária - Código"] == UO]
    return df

print("Lendo bases...")
orc = ler(ARQ["orc"]); des = ler(ARQ["des"])
print(f"Orçamento: {len(orc)} | Despesa: {len(des)}")

def ti(v):
    try: return int(float(v))
    except: return 0

def fe(cod, desc):
    c = str(cod).replace('.0','').strip()
    d = str(desc).strip()
    if c in ['0','nan','']: return ''
    if d in ['nan','SEM INFORMACAO','']: return c
    return f"{c} - {d}"

def fs(v):
    try: return str(int(float(v)))
    except: return ''

def fmt_ag(v):
    try: s=str(int(float(v))); return s[:-1]+'-'+s[-1] if len(s)>1 else s
    except: return ''

def fmt_ct(v):
    try: s=str(int(float(v))); return s[:-1]+'-'+s[-1] if len(s)>1 else s
    except: return ''

def fmt_bk(v):
    try: return str(int(float(v)))
    except: return ''

DIMS = ['Projeto_Atividade - Código','Fonte Recurso - Código','Grupo Despesa - Código',
        'Modalidade Aplicação - Código','Elemento Item Despesa - Código',
        'Elemento Item Despesa - Descrição','Mês - Numérico']

orc_g = orc.groupby(DIMS, dropna=False).agg(ini=('Valor Crédito Inicial','sum'),aut=('Valor Crédito Autorizado','sum'),cota=('Valor Cota Descentralizada','sum')).reset_index().fillna(0)
des_dims = DIMS + (['ContratoConvênio Entrada'] if 'ContratoConvênio Entrada' in des.columns else [])
des_g = des.groupby(des_dims, dropna=False).agg(emp=('Valor Despesa Empenhada','sum'),liq=('Valor Despesa Liquidada','sum'),pago=('Valor Pago Financeiro','sum')).reset_index().fillna(0)

dados_orc = [{'acao':ti(r['Projeto_Atividade - Código']),'fonte':ti(r['Fonte Recurso - Código']),'grupo':ti(r['Grupo Despesa - Código']),'mod':ti(r['Modalidade Aplicação - Código']),'mes':ti(r['Mês - Numérico']),'elem':fe(r['Elemento Item Despesa - Código'],r['Elemento Item Despesa - Descrição']),'ini':round(float(r['ini']),2),'aut':round(float(r['aut']),2),'cota':round(float(r['cota']),2)} for _, r in orc_g.iterrows()]

dados_des = []
for _, r in des_g.iterrows():
    sv = ''
    if 'ContratoConvênio Entrada' in des_g.columns:
        try: sv = fs(r['ContratoConvênio Entrada']) if float(str(r['ContratoConvênio Entrada']).replace('nan','0') or '0') > 0 else ''
        except: sv = ''
    dados_des.append({'acao':ti(r['Projeto_Atividade - Código']),'fonte':ti(r['Fonte Recurso - Código']),'grupo':ti(r['Grupo Despesa - Código']),'mod':ti(r['Modalidade Aplicação - Código']),'mes':ti(r['Mês - Numérico']),'elem':fe(r['Elemento Item Despesa - Código'],r['Elemento Item Despesa - Descrição']),'siafi':sv,'emp':round(float(r['emp']),2),'liq':round(float(r['liq']),2),'pago':round(float(r['pago']),2)})

MESES_L={1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
por_mes = des.groupby('Mês - Numérico').agg(emp=('Valor Despesa Empenhada','sum'),liq=('Valor Despesa Liquidada','sum'),pago=('Valor Pago Financeiro','sum')).reset_index()
dados_mes = [{'mes':ti(r['Mês - Numérico']),'label':MESES_L.get(ti(r['Mês - Numérico']),str(ti(r['Mês - Numérico']))),'emp':round(float(r['emp']),2),'liq':round(float(r['liq']),2),'pago':round(float(r['pago']),2)} for _, r in por_mes.iterrows()]

acoes_orc=sorted(set(ti(v) for v in orc_g['Projeto_Atividade - Código']))
acoes_des=sorted(set(ti(v) for v in des_g['Projeto_Atividade - Código']))
fontes_orc=sorted(set(ti(v) for v in orc_g['Fonte Recurso - Código']))
fontes_des=sorted(set(ti(v) for v in des_g['Fonte Recurso - Código']))
grupos_orc=sorted(set(ti(v) for v in orc_g['Grupo Despesa - Código']))
grupos_des=sorted(set(ti(v) for v in des_g['Grupo Despesa - Código']))
mods_orc=sorted(set(ti(v) for v in orc_g['Modalidade Aplicação - Código']))
mods_des=sorted(set(ti(v) for v in des_g['Modalidade Aplicação - Código']))
meses=sorted(set(ti(v) for v in des_g['Mês - Numérico']))
mes_max=max(meses) if meses else 4
elems_des=sorted(set(fe(r['Elemento Item Despesa - Código'],r['Elemento Item Despesa - Descrição']) for _,r in des_g.iterrows())-{''})
siafi_vals=sorted(set(d['siafi'] for d in dados_des if d['siafi']))

qlik_rows=[]
if os.path.exists(ARQ["qlik"]):
    try:
        qk=pd.read_excel(ARQ["qlik"],engine="xlrd").dropna(how="all")
        qk=qk[qk['Nº SIAFI'].notna()]
        sg=qk.groupby('Nº SIAFI').agg(saldo=('Saldo Atual (R$)','sum'),banco=('Banco','first'),agencia=('Agência Bancária','first'),conta=('Conta Bancária','first')).reset_index()
        ex={}
        if 'ContratoConvênio Entrada' in des.columns:
            for _,r in des.iterrows():
                s=fs(r.get('ContratoConvênio Entrada',''))
                if not s or s in ['0','nan']: continue
                if s not in ex: ex[s]={'emp':0,'liq':0,'pago':0}
                ex[s]['emp']+=float(r.get('Valor Despesa Empenhada',0) or 0)
                ex[s]['liq']+=float(r.get('Valor Despesa Liquidada',0) or 0)
                ex[s]['pago']+=float(r.get('Valor Pago Financeiro',0) or 0)
        for _,r in sg.iterrows():
            s=fs(r['Nº SIAFI']); e=ex.get(s,{'emp':0,'liq':0,'pago':0})
            qlik_rows.append({'siafi':s,'banco':fmt_bk(r.get('banco','')),'agencia':fmt_ag(r.get('agencia','')),'conta':fmt_ct(r.get('conta','')),'emp':round(e['emp'],2),'liq':round(e['liq'],2),'pago':round(e['pago'],2),'saldo':round(float(r.get('saldo',0)),2)})
        print(f"QlikView: {len(qlik_rows)} SIAFIs")
    except Exception as e: print(f"QlikView erro: {e}")

siafi_html="\n".join([f"<tr><td>{r['siafi']}</td><td>{r['banco']}</td><td>{r['agencia']}</td><td>{r['conta']}</td><td>R$ {r['emp']:,.2f}</td><td>R$ {r['liq']:,.2f}</td><td>R$ {r['pago']:,.2f}</td><td>R$ {r['saldo']:,.2f}</td></tr>" for r in qlik_rows]) if qlik_rows else "<tr><td colspan='8' style='text-align:center;padding:20px;color:#96a882;font-style:italic;'>Carregue o arquivo do QlikView na pasta PCASP_QlikView</td></tr>"

aut_t=orc['Valor Crédito Autorizado'].sum(); ini_t=orc['Valor Crédito Inicial'].sum(); cota_t=orc['Valor Cota Descentralizada'].sum()
emp_t=des['Valor Despesa Empenhada'].sum(); liq_t=des['Valor Despesa Liquidada'].sum(); pago_t=des['Valor Pago Financeiro'].sum()
def pct(a,b): return f"{a/b*100:.1f}%" if b>0 else "0.0%"

def alertas():
    al=[]
    pf={}
    for r in dados_orc:
        f=r['fonte']
        if f not in pf: pf[f]={'aut':0,'emp':0}
        pf[f]['aut']+=r['aut']
    for r in dados_des:
        f=r['fonte']
        if f not in pf: pf[f]={'aut':0,'emp':0}
        pf[f]['emp']+=r['emp']
    for f,v in sorted(pf.items()):
        if v['aut']>0:
            p=v['emp']/v['aut']
            if p<0.10: al.append(f'<div class="alerta"><span class="alerta-icon danger">&#9888;</span><span class="alerta-text"><strong>Fonte {f} com execu&ccedil;&atilde;o de {p*100:.1f}%</strong> &mdash; R$ {v["aut"]/1e6:.1f} M autorizados. Risco elevado de sobra ao final do exerc&iacute;cio.</span></div>')
            elif p<0.25: al.append(f'<div class="alerta"><span class="alerta-icon warn">&#9888;</span><span class="alerta-text"><strong>Fonte {f} com execu&ccedil;&atilde;o de {p*100:.1f}%</strong> &mdash; verificar capacidade de absor&ccedil;&atilde;o.</span></div>')
    pa={}
    for r in dados_orc:
        a=r['acao']
        if a not in pa: pa[a]={'aut':0,'emp':0}
        pa[a]['aut']+=r['aut']
    for r in dados_des:
        a=r['acao']
        if a not in pa: pa[a]={'aut':0,'emp':0}
        pa[a]['emp']+=r['emp']
    baixas=[a for a,v in pa.items() if v['aut']>100000 and v['emp']/v['aut']<0.10]
    if baixas: al.append(f'<div class="alerta"><span class="alerta-icon warn">&#9888;</span><span class="alerta-text"><strong>A&ccedil;&otilde;es com execu&ccedil;&atilde;o abaixo de 10%: {", ".join(map(str,sorted(baixas)))}</strong> &mdash; verificar capacidade de absor&ccedil;&atilde;o no 2&ordm; semestre.</span></div>')
    sem=[r['siafi'] for r in qlik_rows if r['saldo']>0 and r['emp']==0]
    if sem: al.append(f'<div class="alerta"><span class="alerta-icon warn">&#9888;</span><span class="alerta-text"><strong>Conv&ecirc;nios com saldo em conta sem execu&ccedil;&atilde;o vinculada: {", ".join(sem[:5])}</strong> &mdash; verificar pend&ecirc;ncias.</span></div>')
    if not al: al.append('<div class="alerta"><span class="alerta-icon ok">&#10003;</span><span class="alerta-text">Nenhum alerta identificado no per&iacute;odo.</span></div>')
    return "\n".join(al)

def ms_opts(vals, ms_id, cb_fn):
    return "\n".join([f'<label class="ms-item"><input type="checkbox" value="{v}" onchange="updateMSLabel(\'{ms_id}\');{cb_fn}();"> <span>{v}</span></label>' for v in vals])

with open(TEMPLATE,encoding='utf-8') as f: html=f.read()

html=html.replace("{{DATA_ATUALIZACAO}}",date.today().strftime("%d/%m/%Y"))
html=html.replace("{{DADOS_ORC_JS}}",json.dumps(dados_orc,ensure_ascii=True))
html=html.replace("{{DADOS_DES_JS}}",json.dumps(dados_des,ensure_ascii=True))
html=html.replace("{{DADOS_MES_JS}}",json.dumps(dados_mes,ensure_ascii=True))
html=html.replace("{{OPTS_G_ACAO}}",ms_opts(acoes_orc,'ms-g-acao','renderG'))
html=html.replace("{{OPTS_G_FONTE}}",ms_opts(fontes_orc,'ms-g-fonte','renderG'))
html=html.replace("{{OPTS_G_MES}}",ms_opts(meses,'ms-g-mes','renderG'))
html=html.replace("{{OPTS_O_ACAO}}",ms_opts(acoes_orc,'ms-o-acao','renderO'))
html=html.replace("{{OPTS_O_FONTE}}",ms_opts(fontes_orc,'ms-o-fonte','renderO'))
html=html.replace("{{OPTS_O_GRUPO}}",ms_opts(grupos_orc,'ms-o-grupo','renderO'))
html=html.replace("{{OPTS_O_MOD}}",ms_opts(mods_orc,'ms-o-mod','renderO'))
html=html.replace("{{OPTS_E_ACAO}}",ms_opts(acoes_des,'ms-e-acao','renderE'))
html=html.replace("{{OPTS_E_FONTE}}",ms_opts(fontes_des,'ms-e-fonte','renderE'))
html=html.replace("{{OPTS_E_GRUPO}}",ms_opts(grupos_des,'ms-e-grupo','renderE'))
html=html.replace("{{OPTS_E_MOD}}",ms_opts(mods_des,'ms-e-mod','renderE'))
html=html.replace("{{OPTS_E_ELEM}}",ms_opts(elems_des,'ms-e-elem','renderE'))
html=html.replace("{{OPTS_E_SIAFI}}",ms_opts(siafi_vals,'ms-e-siafi','renderE'))
html=html.replace("{{OPTS_E_MES}}",ms_opts(meses,'ms-e-mes','renderE'))
html=html.replace("{{SIAFI_ROWS}}",siafi_html)
html=html.replace("{{ALERTAS}}",alertas())
html=html.replace("{{AUT_TOTAL}}",f"R$ {aut_t/1e6:.1f} M")
html=html.replace("{{INI_TOTAL}}",f"R$ {ini_t/1e6:.1f} M")
html=html.replace("{{COTA_TOTAL}}",f"R$ {cota_t/1e6:.1f} M")
html=html.replace("{{EMP_TOTAL}}",f"R$ {emp_t/1e6:.1f} M")
html=html.replace("{{LIQ_TOTAL}}",f"R$ {liq_t/1e6:.1f} M")
html=html.replace("{{PAGO_TOTAL}}",f"R$ {pago_t/1e6:.1f} M")
html=html.replace("{{PCT_EMP}}",pct(emp_t,aut_t))
html=html.replace("{{PCT_LIQ}}",pct(liq_t,aut_t))
html=html.replace("{{PCT_PAGO}}",pct(pago_t,aut_t))
html=html.replace("{{PCT_COTA}}",pct(cota_t,aut_t))

restantes=re.findall(r'\{\{[A-Z_]+\}\}',html)
if restantes: print(f"AVISO: {set(restantes)}")

with open(HTML_SAIDA,"w",encoding="utf-8") as f: f.write(html)
print(f"\nPainel atualizado: {HTML_SAIDA}")
print(f"Meses: {[MESES_L.get(m) for m in sorted(meses)]}")
print(f"SIAFI: {[r['siafi'] for r in qlik_rows]}")
