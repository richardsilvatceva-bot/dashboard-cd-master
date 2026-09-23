import pandas as pd
import json
import os
import urllib.request
import urllib.error
import base64

# ==========================================
# CONFIGURAÇÕES DO GITHUB
# ==========================================
GITHUB_TOKEN = "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"
GITHUB_REPO = "richardsilvatceva-bot/dashboard-cd-master" 

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

    # 2. NET
    df_net = pd.read_excel(xls, 'NET') if 'NET' in xls.sheet_names else pd.DataFrame()
    net_dados = []
    net_ytd = {"netVal": 0.0, "netPct": 0.0}
    if not df_net.empty:
        ytd_col_idx = None
        for col_idx in range(len(df_net.columns)):
            for row_idx in range(min(5, len(df_net))):
                if 'YTD' in str(df_net.iloc[row_idx, col_idx]).upper():
                    ytd_col_idx = col_idx; break
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
    rep_dados, rep_turnos = [], []
    if not df_rep.empty:
        df_rep['DATA_STR'] = pd.to_datetime(df_rep['DATA'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_rep['TRAT_UP'] = df_rep['TRATATIVA'].astype(str).str.upper()
        df_rep['LANC_UP'] = df_rep['LANÇAMENTO'].astype(str).str.upper().str.strip()
        
        for data, group in df_rep.dropna(subset=['DATA_STR']).groupby('DATA_STR'):
            lanc_diario = {}
            for l in ['REPICKING', 'FALTA', 'CORTE']:
                df_l = group[group['LANC_UP'] == l]
                lanc_diario[l] = {"qtd": len(df_l), "valor": sum(tratar_valor_monetario(v) for v in df_l['VALOR'])}
                
            rep_dados.append({
                "data": data, 
                "apontamentos": len(group),
                "encontrados": int(len(group[group['TRAT_UP'].str.contains('ENCONTRADO ANTES', na=False)])),
                "repickFeito": int(len(group[group['Status'].astype(str).str.contains('REPICKING', na=False)])), 
                "valorRepicking": lanc_diario['REPICKING']['valor'],
                "lancamentos": lanc_diario
            })
            
        for _, row in df_rep.groupby('TURNO').agg(ocorrencias=('CHAVE', 'count'), itens=('Quantidade ', 'sum')).reset_index().iterrows():
            if str(row['TURNO']) not in ['0', 'nan']: rep_turnos.append({"turno": str(row['TURNO']), "ocorrencias": int(row['ocorrencias']), "itens": tratar_valor_monetario(row['itens'])})

    # 5. CORTES (AGORA COM O MOTIVO)
    aba_cortes = next((s for s in xls.sheet_names if 'cortes' in s.lower()), None)
    df_cortes = pd.read_excel(xls, aba_cortes) if aba_cortes else pd.DataFrame()
    cortes_dados = []
    if not df_cortes.empty:
        df_cortes.columns = [str(c).strip() for c in df_cortes.columns]
        c_data = next((c for c in df_cortes.columns if 'rótulos' in c.lower() or 'data' in c.lower()), df_cortes.columns[0])
        c_mot = next((c for c in df_cortes.columns if 'motivo' in c.lower()), None)
        c_item = next((c for c in df_cortes.columns if 'item' in c.lower()), df_cortes.columns[1] if c_mot is None else df_cortes.columns[2])
        c_peca = next((c for c in df_cortes.columns if 'pç' in c.lower() or 'peça' in c.lower() or 'qtd' in c.lower()), df_cortes.columns[2] if c_mot is None else df_cortes.columns[3])
        c_val = next((c for c in df_cortes.columns if 'valor' in c.lower()), df_cortes.columns[3] if c_mot is None else df_cortes.columns[4])
        
        if c_data in df_cortes.columns:
            df_cortes[c_data] = df_cortes[c_data].ffill()
            
        for _, row in df_cortes.iterrows():
            dr = row.get(c_data)
            if pd.isna(dr) or 'total' in str(dr).lower() or '(vazio)' in str(dr).lower(): continue
            try:
                dstr = pd.to_datetime(dr, dayfirst=True, errors='coerce').strftime('%Y-%m-%d')
                if dstr == 'NaT': continue
                motivo = str(row.get(c_mot)).strip().upper() if c_mot else "OUTROS"
                if motivo in ['NAN', 'NONE', '']: motivo = "OUTROS"
                
                cortes_dados.append({
                    "data": dstr, 
                    "motivo": motivo,
                    "itens": int(tratar_valor_monetario(row.get(c_item))), 
                    "pecas": int(tratar_valor_monetario(row.get(c_peca))), 
                    "valor": tratar_valor_monetario(row.get(c_val))
                })
            except: pass

    # 6. SOBRAS
    aba_sobras = next((s for s in xls.sheet_names if 'sobras' in s.lower()), None)
    df_sobras = pd.read_excel(xls, aba_sobras, header=None) if aba_sobras else pd.DataFrame()
    sobras_dados = []
    sobras_detalhe = []
    sku_map = {}
    
    if not df_sobras.empty:
        header_1 = -1
        idx_st1 = idx_prod1 = idx_mc1 = idx_loc1 = idx_val1 = -1
        for r in range(min(15, len(df_sobras))):
            row_vals = [str(v).upper().strip() for v in df_sobras.iloc[r].values[:6]]
            if 'STATUS' in row_vals and 'MARCA' in row_vals:
                header_1 = r
                idx_st1 = row_vals.index('STATUS') if 'STATUS' in row_vals else -1
                idx_prod1 = row_vals.index('PRODUTO') if 'PRODUTO' in row_vals else -1
                idx_mc1 = row_vals.index('MARCA') if 'MARCA' in row_vals else -1
                idx_loc1 = next((i for i, v in enumerate(row_vals) if 'LOCA' in v), -1)
                idx_val1 = next((i for i, v in enumerate(row_vals) if 'VALOR' in v and i >= idx_mc1), -1)
                break
                
        if header_1 != -1:
            curr_st = "INDEFINIDO"
            curr_mc = "SEM MARCA"
            for r in range(header_1 + 1, len(df_sobras)):
                if idx_st1 != -1:
                    val = str(df_sobras.iloc[r, idx_st1]).strip()
                    if val.lower() not in ['nan', 'none', '', '(vazio)']:
                        curr_st = val.upper()
                if 'TOTAL' in curr_st: continue
                if idx_mc1 != -1:
                    val = str(df_sobras.iloc[r, idx_mc1]).strip()
                    if val.lower() not in ['nan', 'none', '', '(vazio)']:
                        curr_mc = val.upper()
                prod = ""
                if idx_prod1 != -1:
                    prod = str(df_sobras.iloc[r, idx_prod1]).strip()
                    if prod.lower() in ['nan', 'none', '', '(vazio)'] or 'TOTAL' in prod.upper(): continue
                        
                loc_val = tratar_valor_monetario(df_sobras.iloc[r, idx_loc1]) if idx_loc1 != -1 else 0
                val_val = tratar_valor_monetario(df_sobras.iloc[r, idx_val1]) if idx_val1 != -1 else 0
                
                if prod: sku_map[prod] = {"status": curr_st, "marca": curr_mc}
                if curr_st != 'RESGATADO': sobras_dados.append({"status": curr_st, "marca": curr_mc, "locacoes": int(loc_val), "valor": val_val})

        header_2 = -1
        idx_prod2 = idx_qtd2 = idx_val2 = -1
        for r in range(min(15, len(df_sobras))):
            row_vals = [str(v).upper().strip() for v in df_sobras.iloc[r].values]
            if any('QTD' in v for v in row_vals):
                header_2 = r
                for i, v in enumerate(row_vals):
                    if 'PRODUTO' in v and i >= 5: idx_prod2 = i
                    elif 'QTD' in v: idx_qtd2 = i
                    elif 'VALOR' in v and i > 4: idx_val2 = i
                break
                
        if header_2 != -1 and idx_prod2 != -1:
            for r in range(header_2 + 1, len(df_sobras)):
                prod = str(df_sobras.iloc[r, idx_prod2]).strip()
                if prod.lower() in ['nan', 'none', '', '(vazio)', 'total geral', 'total']: continue
                
                info = sku_map.get(prod, {"status": "INDEFINIDO", "marca": "SEM MARCA"})
                if info["status"] == 'RESGATADO': continue
                
                qtd = tratar_valor_monetario(df_sobras.iloc[r, idx_qtd2]) if idx_qtd2 != -1 else 0
                val = tratar_valor_monetario(df_sobras.iloc[r, idx_val2]) if idx_val2 != -1 else 0
                
                sobras_detalhe.append({"status": info["status"], "produto": prod, "qtd": int(qtd), "valor": val})

    # 7. PICLINHA E DIVCIC
    aba_pic = next((s for s in xls.sheet_names if 'piclinha' in s.lower() or 'divcic' in s.lower() and 'net' not in s.lower()), None)
    df_pic_raw = pd.read_excel(xls, aba_pic, header=None) if aba_pic else pd.DataFrame()
    pic_dados = []
    saldos_divcic = {}
    saldos_piclinha = {}
    
    if not df_pic_raw.empty:
        header_idx = -1
        for i, r in df_pic_raw.iterrows():
            if any('FAMILIA' in str(v).upper() or 'FAMÍLIA' in str(v).upper() for v in r.values):
                header_idx = i; break
                
        if header_idx != -1:
            df_pic_classico = df_pic_raw.iloc[header_idx+1:].reset_index(drop=True)
            df_pic_classico.columns = df_pic_raw.iloc[header_idx]
            fam_cols = [i for i, c in enumerate(df_pic_classico.columns) if 'FAMILIA' in str(c).upper() or 'FAMÍLIA' in str(c).upper()]
            for f_idx in fam_cols:
                setor = 'PICLINHA' if f_idx < 5 else 'DIVCIC'
                for _, row in df_pic_classico.iterrows():
                    try:
                        f, a, v = str(row.iloc[f_idx]).strip(), row.iloc[f_idx + 1], tratar_valor_monetario(row.iloc[f_idx + 3])
                        if f.lower() in ['nan', 'none', '(vazio)', 'total', ''] or pd.isna(a): continue
                        pic_dados.append({"setor": setor, "familia": f, "aging": int(float(a)), "valor": v})
                    except: pass

        for r in range(min(15, len(df_pic_raw))):
            row_vals = [str(v).upper().strip() for v in df_pic_raw.iloc[r].values]
            if 'SALDO DIVCIC' in row_vals or 'SALDO PICLINHA' in row_vals:
                idx_div = row_vals.index('SALDO DIVCIC') if 'SALDO DIVCIC' in row_vals else -1
                idx_pic = row_vals.index('SALDO PICLINHA') if 'SALDO PICLINHA' in row_vals else -1
                
                if idx_div != -1:
                    for i in range(r+2, len(df_pic_raw)):
                        sku = str(df_pic_raw.iloc[i, idx_div]).strip()
                        if sku.lower() in ['nan', 'none', '', '(vazio)', 'total geral']: continue
                        if pd.isna(df_pic_raw.iloc[i, idx_div]): break
                        saldos_divcic[sku] = int(tratar_valor_monetario(df_pic_raw.iloc[i, idx_div+1]))
                        
                if idx_pic != -1:
                    for i in range(r+2, len(df_pic_raw)):
                        sku = str(df_pic_raw.iloc[i, idx_pic]).strip()
                        if sku.lower() in ['nan', 'none', '', '(vazio)', 'total geral']: continue
                        if pd.isna(df_pic_raw.iloc[i, idx_pic]): break
                        saldos_piclinha[sku] = int(tratar_valor_monetario(df_pic_raw.iloc[i, idx_pic+1]))
                break

    cruzamento_sobras = []
    for item in sobras_detalhe:
        sku = item['produto']
        if sku in saldos_divcic or sku in saldos_piclinha:
            cruzamento_sobras.append({
                "status": item['status'],
                "sku": sku,
                "qtd_sobra": item['qtd'],
                "valor_sobra": item['valor'],
                "saldo_divcic": saldos_divcic.get(sku, 0),
                "saldo_piclinha": saldos_piclinha.get(sku, 0)
            })
    cruzamento_sobras = sorted(cruzamento_sobras, key=lambda x: x['valor_sobra'], reverse=True)

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

    json_final = {
        "erros": erros_encontrados, 
        "ciclico": {"daily": ciclico_dados, "offenders": offenders, "tratativas": tratativas},
        "net": net_dados, "net_ytd": net_ytd, 
        "planejamento": {"daily": plan_dados, "curvas": plan_curvas},
        "repicking": {"daily": rep_dados, "turnos": rep_turnos},
        "cortes": cortes_dados, 
        "sobras": sobras_dados,
        "sobras_cruzamento": cruzamento_sobras, 
        "pic_div": pic_dados, 
        "avarias": avarias_dados
    }

    with open('dados.json', 'w', encoding='utf-8') as f:
        json.dump(json_final, f, ensure_ascii=False, indent=4)
        
    print(f"✓ Ficheiro 'dados.json' gerado localmente com as novas atualizações!")

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
        
        sha = None
        try:
            req_get = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req_get) as response:
                sha = json.loads(response.read().decode())['sha']
        except urllib.error.HTTPError as e:
            if e.code != 404: raise e

        payload = {
            "message": "Atualização automática (Gráfico Motivos de Corte) 🚀", 
            "content": base64.b64encode(conteudo.encode('utf-8')).decode('utf-8')
        }
        if sha: payload["sha"] = sha
        
        req_put = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method="PUT")
        with urllib.request.urlopen(req_put) as response:
            print("✅ SUCESSO: O Dashboard online foi atualizado!")
            
    except urllib.error.HTTPError as e:
        print(f"❌ Erro do GitHub: {e.code} - {e.reason}")
    except Exception as e:
        print(f"❌ Falha ao tentar conectar ao GitHub: {e}")

if __name__ == "__main__":
    gerar_json()