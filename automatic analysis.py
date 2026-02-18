# =======================
# ANALISE VAREJO - CAMPANHA HONDA (CSV + Excel)
# =======================
import pandas as pd
import re
import glob
import os
from datetime import datetime

# =======================
# FUNÇÕES AUXILIARES
# =======================
def get_files(folder, patterns=("*.csv",)):
    """Retorna todos os arquivos dentro da pasta de acordo com o padrão"""
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(folder, pattern)))
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo encontrado em {folder}")
    return files

def get_latest_file(folder, patterns=("*.xlsx", "*.xls")):
    """Retorna o arquivo mais recente dentro da pasta"""
    files = get_files(folder, patterns)
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def normalize_phone(num):
    """Normaliza número de telefone (remove caracteres e código do país)"""
    if pd.isna(num): 
        return None
    s = re.sub(r"\D", "", str(num))
    if s.startswith("55") and len(s) > 10:
        s = s[2:]
    return s

def normalize_columns(df):
    """Normaliza nomes das colunas: remove espaços e deixa minúsculo"""
    df.columns = df.columns.str.strip().str.lower()
    return df

# =======================
# CAMINHOS DAS PASTAS - HONDA
# =======================
PASTA_CHAMADAS = r"PASTA COM ARQUIVOS PARA CRUZAMENTO"
PASTA_OCORRENCIAS = r"De X Para"
PASTA_TELEFONES = r"PASTA COM ARQUIVOS ORIGINAIS"
PASTA_SAIDA = r"PASTA DESTINO"

# =======================
# ARQUIVOS FIXOS
# =======================
ARQ_CHAMADAS = get_latest_file(PASTA_CHAMADAS, patterns=("*.xlsx", "*.xls"))
ARQ_OCORRENCIAS = get_latest_file(PASTA_OCORRENCIAS, patterns=("*.xlsx", "*.xls"))

# =======================
# INÍCIO EXECUÇÃO
# =======================
inicio_total = datetime.now()
print(f"🚀 Início da execução: {inicio_total.strftime('%Y-%m-%d %H:%M:%S')}")

# =======================
# LOOP SOBRE ARQUIVOS DE TELEFONES (CSV)
# =======================
arquivos_telefones = get_files(PASTA_TELEFONES, patterns=("*.csv",))

for idx, ARQ_TELEFONES in enumerate(arquivos_telefones, start=1):
    inicio_arquivo = datetime.now()
    print("="*60)
    print(f"⏳ [{inicio_arquivo.strftime('%H:%M:%S')}] Processando arquivo {idx} de {len(arquivos_telefones)}: {ARQ_TELEFONES}")

    nome_base_tel = os.path.splitext(os.path.basename(ARQ_TELEFONES))[0]
    SAIDA = os.path.join(PASTA_SAIDA, f"ANALISE_{nome_base_tel}.xlsx")

    # =======================
    # LEITURA DE BASES
    # =======================
    df_ch = normalize_columns(pd.read_excel(ARQ_CHAMADAS))
    df_map = normalize_columns(pd.read_excel(ARQ_OCORRENCIAS))
    df_tel = pd.read_csv(ARQ_TELEFONES, sep=";", quotechar='"')
    df_tel = normalize_columns(df_tel)

    print("Colunas do CSV:", df_tel.columns.tolist())  # DEBUG: mostra colunas

    col_numero = "número" if "número" in df_ch.columns else "numero"
    df_ch["numero_norm"] = df_ch[col_numero].apply(normalize_phone)

    # =======================
    # TRATAR BASE TELEFONES - DETECÇÃO FLEXÍVEL
    # =======================
    tel_list = []

    ID_COL = "cpf"  # Usar CPF como identificador

    # Layout múltiplo: ddd01/telefone01, ddd02/telefone02...
    for i in range(1, 11):
        ddd_col = next((c for c in df_tel.columns if c.strip().lower() == f"ddd{i:02d}"), None)
        tel_col = next((c for c in df_tel.columns if c.strip().lower() == f"telefone{i:02d}"), None)
        if ddd_col and tel_col and ID_COL in df_tel.columns:
            temp = df_tel[[ID_COL, ddd_col, tel_col]].copy()
            temp["numero_norm"] = (temp[ddd_col].fillna("").astype(str) + temp[tel_col].fillna("").astype(str))
            temp["numero_norm"] = temp["numero_norm"].apply(normalize_phone)
            temp = temp.dropna(subset=["numero_norm"])
            tel_list.append(temp[[ID_COL, "numero_norm"]])

    # Layout simples: ddd + numero (ou variações possíveis)
    ddd_cols = [c for c in df_tel.columns if "ddd" in c.lower()]
    num_cols = [c for c in df_tel.columns if "numero" in c.lower() or "tel" in c.lower()]
    if ddd_cols and num_cols and ID_COL in df_tel.columns:
        ddd_col = ddd_cols[0]
        num_col = num_cols[0]
        temp = df_tel[[ID_COL, ddd_col, num_col]].copy()
        temp["numero_norm"] = (temp[ddd_col].fillna("").astype(str) + temp[num_col].fillna("").astype(str))
        temp["numero_norm"] = temp["numero_norm"].apply(normalize_phone)
        temp = temp.dropna(subset=["numero_norm"])
        tel_list.append(temp[[ID_COL, "numero_norm"]])

    if len(tel_list) == 0:
        print("❌ Nenhum layout de telefones reconhecido nesse arquivo.")
        print("Colunas disponíveis no CSV:", df_tel.columns.tolist())
        continue  # não quebra a execução, passa para o próximo CSV

    df_tel_long = pd.concat(tel_list, ignore_index=True).drop_duplicates()

    # =======================
    # CRUZAMENTO TELEFONE ↔ CHAMADA
    # =======================
    df_merge = df_ch.merge(df_tel_long, left_on="numero_norm", right_on="numero_norm", how="inner")

    # =======================
    # CRUZAR COM DE X PARA
    # =======================
    df_merge = df_merge.merge(df_map, left_on="acionamento", right_on="ocorrência", how="left")

    # Log de erros
    nao_localizados = df_merge[df_merge["tentativa"].isna()]["acionamento"].unique()
    if len(nao_localizados) > 0:
        print("❌ ERRO: Acionamentos sem correspondência no De x Para:")
        for ac in nao_localizados:
            print("   -", ac)
    else:
        print("✅ Todos os acionamentos localizados no De x Para.")

    # =======================
    # RESUMO POR TELEFONE
    # =======================
    resumo_tel = (
        df_merge.groupby("numero_norm")
        .agg(
            total_chamadas=("número", "count"),
            tentativas=("tentativa", "sum"),
            alo=("alo", "sum"),
            cpc=("cpc", "sum"),
            promessa=("promessa", "sum"),
        )
        .reset_index()
    )

    # =======================
    # PIVOT POR SERVIÇO
    # =======================
    df_unique = df_merge.groupby(["serviço", "numero_norm"]).agg(
        promessa=("promessa", "max"),
        cpc=("cpc", "max"),
        alo=("alo", "max"),
        tentativa=("tentativa", "max"),
    ).reset_index()

    def define_prioridade(row):
        if row["promessa"] and row["promessa"] > 0:
            return "promessa"
        elif row["cpc"] and row["cpc"] > 0:
            return "cpc"
        elif row["alo"] and row["alo"] > 0:
            return "alo"
        elif row["tentativa"] and row["tentativa"] > 0:
            return "tentativa"
        else:
            return "nenhuma"

    df_unique["ocorrencia_final"] = df_unique.apply(define_prioridade, axis=1)

    pivot_servico = (
        df_unique.groupby(["serviço", "ocorrencia_final"])["numero_norm"]
        .nunique()
        .unstack(fill_value=0)
        .reset_index()
    )

    pivot_servico["qtd_telefones"] = df_unique.groupby("serviço")["numero_norm"].nunique().reindex(pivot_servico["serviço"]).values

    colunas_final = ["serviço", "qtd_telefones", "alo", "cpc", "promessa"]
    for col in colunas_final:
        if col not in pivot_servico.columns:
            pivot_servico[col] = 0
    pivot_servico = pivot_servico[colunas_final]

    totais = {
        "serviço": "TOTAL GERAL",
        "qtd_telefones": pivot_servico["qtd_telefones"].sum(),
        "alo": pivot_servico["alo"].sum(),
        "cpc": pivot_servico["cpc"].sum(),
        "promessa": pivot_servico["promessa"].sum(),
    }
    pivot_servico = pd.concat([pivot_servico, pd.DataFrame([totais])], ignore_index=True)

    # =======================
    # EXPORTAÇÃO
    # =======================
    with pd.ExcelWriter(SAIDA, engine="openpyxl") as writer:
        df_merge.to_excel(writer, sheet_name="DETALHE", index=False)
        resumo_tel.to_excel(writer, sheet_name="POR_TELEFONE", index=False)
        pivot_servico.to_excel(writer, sheet_name="RESUMO_PIVOT", index=False)

    fim_arquivo = datetime.now()
    print(f"✅ Arquivo gerado: {SAIDA} ({(fim_arquivo - inicio_arquivo).total_seconds():.2f}s)")

# =======================
# FIM EXECUÇÃO
# =======================
fim_total = datetime.now()
print(f"🏁 Execução finalizada: {fim_total.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"⏱ Duração total: {(fim_total - inicio_total).total_seconds():.2f}s")

