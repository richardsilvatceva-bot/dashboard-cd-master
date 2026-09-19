import pandas as pd
import json
import os
import math
import urllib.request
import urllib.error
import base64
from datetime import datetime

# ==========================================
# CONFIGURAÇÕES DO GITHUB
# ==========================================
GITHUB_TOKEN = "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"
GITHUB_REPO = "richardsilvatceva-bot/dashboard-cd-master" # <-- Se o nome do repositório não for este, altere só o final!

# ==========================================
# FUNÇÕES AUXILIARES E TRATAMENTO
# ==========================================
def get_col_letter(col_idx):
    result = ""
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def mapear_erros(xls):
    erros = []
    error_strings = ['#N/D', '#VALOR!', '#REF!', '#DIV/0!', '#NÚM!', '#NOME?', '#NULO!', '#N/A', '#VALUE!', '#NUM!', '#NULL!']
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                if str(value).strip().upper() in error_strings:
                    erros.append({"aba": sheet_name, "celula": f"{get_col_letter(col_idx)}{row_idx + 2}", "erro": str(value).strip().upper()})
    return erros

def formatar_data(row, ano, col_mes='MÊS', col_dia='DIA'):
    meses = {'JANEIRO': 1, 'FEVEREIRO': 2, 'MARÇO': 3, 'ABRIL': 4, 'MAIO': 5, 'JUNHO': 6, 'JULHO': 7, 'AGOSTO': 8, 'SETEMBRO': 9, 'OUTUBRO': 10, 'NOVEMBRO': 11, 'DEZEMBRO': 12}
    mes = row.get(col_mes)
    dia = row.get(col_dia)
    if pd.isnull(mes) or pd.isnull(dia) or str(dia).strip().lower() == '(vazio)': return None
    try: return f"{ano}-{meses.get(str(mes).strip().upper(), 1):02d}-{int(float(dia)):02d}"
    except: return None

def tratar_valor_monetario(val):
    if pd.isnull(val) or val == '(vazio)': return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).replace('R$', '').replace('\xa0', '').strip()
    if val_str in ['', '-']: return 0.0
    if ',' in val_str and '.' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.') if val_str.rfind(',') > val_str.rfind('.') else val_str.replace(',', '')
    elif ',' in val_str: val_str = val_str.replace(',', '.')
    try: return float(val_str)
    except: return 0.0

# ==========================================
# FUNÇÃO PRINCIPAL DE LEITURA
# ==========================================
def gerar_json():
    arquivo_excel = 'Matris.xlsx'
    if not os.path.exists(arquivo_excel):
        print(f"Erro: Arquivo '{arquivo_excel}' não encontrado na pasta atual.")
        return

    print("A analisar dados da planilha...")
    xls = pd.ExcelFile(arquivo_excel)
    erros_encontrados = mapear_erros(xls)
    
    # 1. CÍCLICO
    df_cic = pd.read_excel(xls, 'Cíclico') if 'Cíclico' in xls.sheet_names else pd.DataFrame()
    ciclico_dados, offenders, tratativas = [], [], []
    if not df_cic.empty:
        df_cic['DATA_STR'] = df_cic.apply(lambda r: formatar_data(r, 2026), axis=1)
        for data, group in df_cic.dropna(subset=['DATA_STR']).groupby('DATA_STR'):
            ciclico_dados.append({"data": data, "valorContagem": tratar_valor_monetario(group['Soma de VALOR CONTAGEM'].sum()), "itensUnicos": group['ITEM 1° CONTAGEM'].nunique(), "fisicoFinal": tratar_valor_monetario(group['Soma de FISICO FINAL'].sum()), "ajustePos": tratar_valor_monetario(group[group['Soma de VALOR DO AJUSTE'] > 0]['Soma de VALOR DO AJUSTE'].sum()), "ajusteNeg": tratar_valor_monetario(group[group['Soma de VALOR DO AJUSTE'] < 0]['Soma de VALOR DO AJUSTE'].sum())})
        of_group = df_cic.groupby('ITEM 1° CONTAGEM').agg({'Soma de VALOR DO AJUSTE': 'sum'}).reset_index()
        top_of = of_group.assign(abs=of_group['Soma de VALOR DO AJUSTE'].abs()).sort_values('abs', ascending=False).head(50)
        for _, row in top_of.iterrows(): offenders.append({"sku": str(row['ITEM 1° CONTAGEM']), "valorAjuste": tratar_valor_monetario(row['Soma de VALOR DO AJUSTE'])})
        try:
            for _, row in df_cic.iloc[1:, 12:16].dropna(subset=[df_cic.columns[13]]).iterrows():
                if str(row.iloc[1]).strip() == 'CAUSA' or pd.isna(row.iloc[2]): continue
                try: tratativas.append({"mes": str(row.iloc[0]).strip().upper(), "causa": str(row.iloc[1]).strip(), "qtd": int(row.iloc[2])})
                except: pass
        except: pass

    # 2. NET (E EXTRAÇÃO DO YTD)
    df_net = pd.read_excel(xls, 'NET') if 'NET' in xls.sheet_names else pd.DataFrame()
    net_dados = []
    net_ytd = {"netVal": 0.0, "netPct": 0.0}
    if not df_net.empty:
        ytd_col_idx = None
        for col_idx in range(len(df_net.columns)):
            for row_idx in range(min(5, len(df_net))):
                if 'YTD' in str(df_net.iloc[row_idx, col_idx]).upper():
                    ytd_col_idx = col_idx
                    break
            if ytd_col_idx is not None: break
            
        if ytd_col_idx is not None:
            for r in range(len(df_net)):
                val_col1 = str(df_net.iloc[r, 1]).strip().upper()
                if val_col1 == 'NET': net_ytd["netVal"] = tratar_valor_monetario(df_net.iloc[r, ytd_col_idx])
                elif val_col1 == '% NET': net_ytd["netPct"] = tratar_valor_monetario(df_net.iloc[r, ytd_col_idx])

        meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
        idx = df_net.set_index(df_net.columns[1])
        def ex(ch, m_id):
            try: return tratar_valor_monetario(idx.loc[ch].iloc[m_id])
            except: return 0.0
        for i, mes in enumerate(meses):
            net_dados.append({"mes": mes, "divcic": ex('DIVCIC', i+1), "ajusteNeg": ex('Ajuste Cíclico Negativo', i+1), "ajustePos": ex('Ajuste Cíclico Positivo', i+1), "sobras": ex('Sobras Processadas', i+1), "avaria": ex('Avaria Processo', i+1), "netVal": ex('NET', i+1), "netPct": ex('% NET', i+1), "grossVal": ex('GROSS', i+1), "grossPct": ex('% GROSS', i+1)})

    # 3. PLANEJAMENTO
    df_plan = pd.read_excel(xls, 'Planejamento') if 'Planejamento' in xls.sheet_names else pd.DataFrame()
    plan_dados, plan_curvas = [], {}
    if not df_plan.empty:
        df_plan['DATA_STR'] = pd.to_datetime(df_plan['DATA'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_plan['Curva ABC'] = df_plan['Curva ABC'].fillna('Sem Curva')
        for data, group in df_plan.dropna(subset=['DATA_STR']).groupby(['DATA_STR', 'Curva ABC']).size().reset_index(name='qtd').groupby('DATA_STR'):
            plan_dados.append({"data": data, "total": int(group['qtd'].sum()), "curvas": {r['Curva ABC']: int(r['qtd']) for _, r in group.iterrows()}})
        if 'OBSERVAÇÃO' in df_plan.columns:
            for c in df_plan['Curva ABC'].unique():
                df_c = df_plan[df_plan['Curva ABC'] == c]
                obs = df_c['OBSERVAÇÃO'].astype(str).str.upper()
                plan_curvas[str(c)] = {"contado": int(len(df_c[obs == 'PLANEJADO'])), "pendente": int(len(df_c[obs == 'PENDENTE']))}

    # 4. REPICKING
    df_rep = pd.read_excel(xls, 'Repicking') if 'Repicking' in xls.sheet_names else pd.DataFrame()
    rep_dados, rep_turnos, rep_kpis, rep_lanc = [], [], {}, {}
    if not df_rep.empty:
        df_rep['DATA_STR'] = pd.to_datetime(df_rep['DATA'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_rep['TRAT_UP'] = df_rep['TRATATIVA'].astype(str).str.upper()
        df_rep['LANC_UP'] = df_rep['LANÇAMENTO'].astype(str).str.upper().str.strip()
        rep_kpis = {"apontamentos": len(df_rep), "encontrados": int(len(df_rep[df_rep['TRAT_UP'].str.contains('ENCONTRADO ANTES', na=False)]))}
        for l in ['REPICKING', 'FALTA', 'CORTE']:
            df_l = df_rep[df_rep['LANC_UP'] == l]
            rep_lanc[l] = {"qtd": len(df_l), "valor": sum(tratar_valor_monetario(v) for v in df_l['VALOR'])}
        for data, group in df_rep.dropna(subset=['DATA_STR']).groupby('DATA_STR'):
            rep_dados.append({"data": data, "repickFeito": int(len(group[group['Status'].astype(str).str.contains('REPICKING', na=False)])), "valorRepicking": sum(tratar_valor_monetario(v) for v in group[group['LANC_UP'] == 'REPICKING']['VALOR'])})
        for _, row in df_rep.groupby('TURNO').agg(ocorrencias=('CHAVE', 'count'), itens=('Quantidade ', 'sum')).reset_index().iterrows():
            if str(row['TURNO']) not in ['0', 'nan']: rep_turnos.append({"turno": str(row['TURNO']), "ocorrencias": int(row['ocorrencias']), "itens": tratar_valor_monetario(row['itens'])})

    # 5. CORTES
    aba_cortes = next((s for s in xls.sheet_names if 'cortes' in s.lower()), None)
    df_cortes = pd.read_excel(xls, aba_cortes) if aba_cortes else pd.DataFrame()
    cortes_dados = []
    if not df_cortes.empty:
        df_cortes.columns = [str(c).strip() for c in df_cortes.columns]
        c_data = next((c for c in df_cortes.columns if 'rótulos' in c.lower() or 'data' in c.lower()), df_cortes.columns[0])
        c_item = next((c for c in df_cortes.columns if 'item' in c.lower()), df_cortes.columns[1])
        c_peca = next((c for c in df_cortes.columns if 'pç' in c.lower() or 'peça' in c.lower() or 'qtd' in c.lower()), df_cortes.columns[2])
        c_val = next((c for c in df_cortes.columns if 'valor' in c.lower()), df_cortes.columns[3])
        for _, row in df_cortes.iterrows():
            dr = row.get(c_data)
            if pd.isna(dr) or 'total' in str(dr).lower() or '(vazio)' in str(dr).lower(): continue
            try:
                dstr = pd.to_datetime(dr, dayfirst=True, errors='coerce').strftime('%Y-%m-%d')
                cortes_dados.append({"data": dstr, "itens": int(tratar_valor_monetario(row.get(c_item))), "pecas": int(tratar_valor_monetario(row.get(c_peca))), "valor": tratar_valor_monetario(row.get(c_val))})
            except: pass

    # 6. SOBRAS
    aba_sobras = next((s for s in xls.sheet_names if 'sobras' in s.lower()), None)
    df_sobras = pd.read_excel(xls, aba_sobras) if aba_sobras else pd.DataFrame()
    sobras_dados = []
    if not df_sobras.empty:
        df_sobras.columns = [str(c).strip() for c in df_sobras.columns]
        c_st = next((c for c in df_sobras.columns if 'status' in c.lower()), df_sobras.columns[0])
        c_mc = next((c for c in df_sobras.columns if 'marca' in c.lower()), df_sobras.columns[1])
        c_loc = next((c for c in df_sobras.columns if 'loc' in c.lower()), df_sobras.columns[2])
        c_val = next((c for c in df_sobras.columns if 'valor' in c.lower()), df_sobras.columns[3])
        if c_st in df_sobras.columns:
            df_sobras[c_st] = df_sobras[c_st].ffill()
            for _, row in df_sobras.iterrows():
                st, mc = str(row.get(c_st)).strip(), str(row.get(c_mc)).strip()
                if st.lower() in ['nan', 'none', ''] or 'total' in st.lower(): continue
                if mc.lower() in ['(vazio)', 'nan', 'none', '']: mc = 'SEM MARCA'
                sobras_dados.append({"status": st.upper(), "marca": mc.upper(), "locacoes": int(tratar_valor_monetario(row.get(c_loc))), "valor": tratar_valor_monetario(row.get(c_val))})

    # 7. PICLINHA E DIVCIC
    aba_pic = next((s for s in xls.sheet_names if 'piclinha' in s.lower() or 'divcic' in s.lower() and 'net' not in s.lower()), None)
    df_pic = pd.read_excel(xls, aba_pic, header=None) if aba_pic else pd.DataFrame()
    pic_dados = []
    if not df_pic.empty:
        header_idx = 0
        for i, r in df_pic.iterrows():
            if any('FAMILIA' in str(v).upper() or 'FAMÍLIA' in str(v).upper() for v in r.values):
                header_idx = i
                break
        df_pic.columns = df_pic.iloc[header_idx]
        df_pic = df_pic.iloc[header_idx+1:].reset_index(drop=True)
        fam_cols = [i for i, c in enumerate(df_pic.columns) if 'FAMILIA' in str(c).upper() or 'FAMÍLIA' in str(c).upper()]
        for f_idx in fam_cols:
            setor = 'PICLINHA' if f_idx < 5 else 'DIVCIC'
            for _, row in df_pic.iterrows():
                try:
                    f, a, v = str(row.iloc[f_idx]).strip(), row.iloc[f_idx + 1], tratar_valor_monetario(row.iloc[f_idx + 3])
                    if f.lower() in ['nan', 'none', '(vazio)', 'total', ''] or pd.isna(a): continue
                    pic_dados.append({"setor": setor, "familia": f, "aging": int(float(a)), "valor": v})
                except: pass

    # 8. AVARIAS
    aba_avarias = next((s for s in xls.sheet_names if 'avarias' in s.lower()), None)
    df_avarias = pd.read_excel(xls, aba_avarias) if aba_avarias else pd.DataFrame()
    avarias_dados = []
    if not df_avarias.empty:
        df_avarias.columns = [str(c).strip() for c in df_avarias.columns]
        c_data = next((c for c in df_avarias.columns if 'data' in c.lower() or 'rótulos' in c.lower()), df_avarias.columns[0])
        c_mot = next((c for c in df_avarias.columns if 'motivo' in c.lower()), df_avarias.columns[1] if len(df_avarias.columns)>1 else None)
        c_oco = next((c for c in df_avarias.columns if 'ocorrência' in c.lower() or 'ocorrencia' in c.lower()), df_avarias.columns[2] if len(df_avarias.columns)>2 else None)
        c_val = next((c for c in df_avarias.columns if 'soma' in c.lower() or 'r$' in c.lower() or 'valor' in c.lower()), df_avarias.columns[3] if len(df_avarias.columns)>3 else None)
        
        if c_data in df_avarias.columns:
            df_avarias[c_data] = df_avarias[c_data].ffill()
            for _, row in df_avarias.iterrows():
                dr, mot, oc = row.get(c_data), str(row.get(c_mot)).strip(), str(row.get(c_oco)).strip()
                if pd.isna(dr) or 'total' in str(dr).lower() or '(vazio)' in str(dr).lower() or mot.lower() in ['nan', 'none', '']: continue
                try:
                    dstr = pd.to_datetime(dr, dayfirst=True, errors='coerce').strftime('%Y-%m-%d')
                    if dstr != 'NaT':
                        avarias_dados.append({"data": dstr, "motivo": mot.upper(), "ocorrencia": oc.upper() if oc.lower() not in ['nan','none',''] else 'OUTROS', "valor": tratar_valor_monetario(row.get(c_val))})
                except: pass

    # ==========================================
    # SALVAR LOCALMENTE O JSON
    # ==========================================
    json_final = {
        "erros": erros_encontrados, 
        "ciclico": {"daily": ciclico_dados, "offenders": offenders, "tratativas": tratativas},
        "net": net_dados, "net_ytd": net_ytd, 
        "planejamento": {"daily": plan_dados, "curvas": plan_curvas},
        "repicking": {"daily": rep_dados, "turnos": rep_turnos, "kpis": rep_kpis, "lancamentos": rep_lanc},
        "cortes": cortes_dados, "sobras": sobras_dados, "pic_div": pic_dados, "avarias": avarias_dados
    }

    with open('dados.json', 'w', encoding='utf-8') as f:
        json.dump(json_final, f, ensure_ascii=False, indent=4)
        
    print(f"✓ Ficheiro 'dados.json' gerado localmente.")

    # ==========================================
    # ENVIO AUTOMÁTICO PARA O GITHUB
    # ==========================================
    try:
        print(f"\nA enviar o novo 'dados.json' para o GitHub ({GITHUB_REPO})...")
        with open('dados.json', 'r', encoding='utf-8') as f:
            conteudo = f.read()
            
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/dados.json"
        headers = {
            "Authorization": f"token {GITHUB_TOKEN}", 
            "Accept": "application/vnd.github.v3+json", 
            "User-Agent": "Python-Upload-Script"
        }
        
        # Verifica se o ficheiro já existe para obter a chave SHA
        sha = None
        try:
            req_get = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req_get) as response:
                sha = json.loads(response.read().decode())['sha']
        except urllib.error.HTTPError as e:
            if e.code != 404: raise e

        # Constrói o payload para envio
        payload = {
            "message": "Atualização automática pelo Python 🚀", 
            "content": base64.b64encode(conteudo.encode('utf-8')).decode('utf-8')
        }
        if sha: payload["sha"] = sha
        
        req_put = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method="PUT")
        with urllib.request.urlopen(req_put) as response:
            print("✅ SUCESSO: O Dashboard online foi atualizado com os novos dados!")
            
    except urllib.error.HTTPError as e:
        print(f"❌ Erro do GitHub: {e.code} - {e.reason}")
        print(f"Verifique se o repositório '{GITHUB_REPO}' existe e se o nome está correto na linha 14.")
    except Exception as e:
        print(f"❌ Falha ao tentar conectar ao GitHub: {e}")

if __name__ == "__main__":
    gerar_json()