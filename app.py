from flask import Flask, render_template, request, session, redirect, url_for
import pandas as pd
import os
import glob
import re

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "troque-esta-chave-por-uma-chave-segura-no-render"
)

USUARIO_LOGIN = os.environ.get(
    "USUARIO_LOGIN",
    "FNP-ADMIN"
)

SENHA_LOGIN = os.environ.get(
    "SENHA_LOGIN",
    "R0ta!Frota#2026@Segura"
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_DADOS = os.path.join(BASE_DIR, "dados")


# ============================================================
# LOCALIZAR EXCEL
# ============================================================

def localizar_excel():
    """Localiza a planilha de ocorrências, ignorando a planilha de manutenção."""
    arquivos = []
    arquivos.extend(glob.glob(os.path.join(PASTA_DADOS, "*.xlsx")))
    arquivos.extend(glob.glob(os.path.join(PASTA_DADOS, "*.xls")))
    arquivos = [
        arquivo for arquivo in arquivos
        if not os.path.basename(arquivo).startswith("~$")
        and "manuten" not in os.path.basename(arquivo).lower()
    ]
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum arquivo Excel de ocorrências foi encontrado. Pasta procurada: {PASTA_DADOS}"
        )
    return max(arquivos, key=os.path.getmtime)


def localizar_excel_manutencao():
    """Localiza automaticamente a planilha mais recente de manutenção."""
    arquivos = []
    arquivos.extend(glob.glob(os.path.join(PASTA_DADOS, "*.xlsx")))
    arquivos.extend(glob.glob(os.path.join(PASTA_DADOS, "*.xls")))
    arquivos = [
        arquivo for arquivo in arquivos
        if not os.path.basename(arquivo).startswith("~$")
        and "manuten" in os.path.basename(arquivo).lower()
    ]
    if not arquivos:
        raise FileNotFoundError(
            "Nenhuma planilha de manutenção foi encontrada. Coloque o arquivo de manutenção dentro da pasta 'dados'."
        )
    return max(arquivos, key=os.path.getmtime)


# ============================================================
# IDENTIFICAR ABA
# ============================================================

def localizar_aba_excel(caminho_excel):

    excel = pd.ExcelFile(caminho_excel)

    palavras_chave = [
        "notification",
        "trigger",
        "velocidade",
        "excesso"
    ]

    for aba in excel.sheet_names:

        nome_aba = str(aba).lower()

        for palavra in palavras_chave:
            if palavra in nome_aba:
                return aba

    return excel.sheet_names[0]


# ============================================================
# LOCALIZAR COLUNA
# ============================================================

def localizar_coluna(colunas, termos):

    for coluna in colunas:

        nome = str(coluna).strip().lower()

        for termo in termos:
            if termo.lower() in nome:
                return coluna

    return None


# ============================================================
# EXTRAIR VELOCIDADES
# ============================================================

def extrair_velocidades(texto):

    if pd.isna(texto):
        return None, None

    texto = str(texto)

    velocidade_atual = None
    velocidade_limite = None

    padrao_atual = re.search(
        r"velocidade\s*atual\s*[:=]\s*(\d+(?:[.,]\d+)?)",
        texto,
        re.IGNORECASE
    )

    if padrao_atual:
        velocidade_atual = float(
            padrao_atual.group(1).replace(",", ".")
        )

    padrao_limite = re.search(
        r"velocidade\s*limite\s*[:=]\s*(\d+(?:[.,]\d+)?)",
        texto,
        re.IGNORECASE
    )

    if padrao_limite:
        velocidade_limite = float(
            padrao_limite.group(1).replace(",", ".")
        )

    return velocidade_atual, velocidade_limite


# ============================================================
# CLASSIFICAR GRAVIDADE
# ============================================================

def classificar_gravidade(excesso, cerca):

    if pd.isna(excesso):
        return "NÃO CLASSIFICADO"

    try:
        excesso = float(excesso)
    except:
        return "NÃO CLASSIFICADO"

    # PARAMETRIZAÇÃO ATUAL:
    # SEDE: 31-35 LEVE | 36-40 MÉDIA | 41-50 GRAVE | acima de 50 GRAVÍSSIMA
    # NOVA PIRATININGA: 71-75 LEVE | 76-80 MÉDIA | 81-90 GRAVE | acima de 90 GRAVÍSSIMA
    #
    # Nas duas cercas atuais, a faixa máxima GRAVE é +20 km/h
    # sobre o limite. Portanto, excesso acima de 20 km/h é GRAVÍSSIMA.
    # Não deixamos mais nenhuma ocorrência como "FORA DA FAIXA".

    if excesso <= 5:
        return "LEVE"
    elif excesso <= 10:
        return "MÉDIA"
    elif excesso <= 20:
        return "GRAVE"
    else:
        return "GRAVÍSSIMA"

# ============================================================
# IDENTIFICAR CERCA
# ============================================================

def identificar_cerca(texto):

    if pd.isna(texto):
        return "NÃO IDENTIFICADA"

    texto = str(texto).upper()

    cercas = [
        "SEDE",
        "NOVA PIRATININGA",
        "FAZENDA",
        "RETRO",
        "RETIRO"
    ]

    for cerca in cercas:
        if cerca in texto:
            return cerca

    return "NÃO IDENTIFICADA"


# ============================================================
# PROCESSAR DADOS
# ============================================================

def processar_dados():

    caminho_excel = localizar_excel()
    nome_arquivo = os.path.basename(caminho_excel)

    aba = localizar_aba_excel(caminho_excel)

    df_bruto = pd.read_excel(
        caminho_excel,
        sheet_name=aba,
        header=None
    )

    linha_cabecalho = None

    palavras_procuradas = [
        "agrupamento",
        "notification",
        "trigger",
        "driver",
        "texto"
    ]

    for indice, linha in df_bruto.iterrows():

        texto_linha = " ".join(
            str(valor).lower()
            for valor in linha.values
            if pd.notna(valor)
        )

        quantidade = sum(
            palavra in texto_linha
            for palavra in palavras_procuradas
        )

        if quantidade >= 2:
            linha_cabecalho = indice
            break

    if linha_cabecalho is None:
        raise ValueError(
            "Não foi possível identificar o cabeçalho da planilha."
        )

    df = pd.read_excel(
        caminho_excel,
        sheet_name=aba,
        header=linha_cabecalho
    )

    df = df.dropna(how="all")

    df.columns = [
        str(coluna).strip()
        for coluna in df.columns
    ]

    # ========================================================
    # LOCALIZAR COLUNAS
    # ========================================================

    coluna_data_hora = localizar_coluna(
        df.columns,
        [
            "notification trigger time",
            "trigger time",
            "notification"
        ]
    )

    coluna_alerta = localizar_coluna(
        df.columns,
        [
            "texto do alerta",
            "alerta",
            "alert text"
        ]
    )

    coluna_frota = localizar_coluna(
        df.columns,
        [
            "agrupamento",
            "frota",
            "grupo"
        ]
    )

    coluna_driver = localizar_coluna(
        df.columns,
        [
            "driver",
            "condutor"
        ]
    )

    if coluna_data_hora is None:
        raise ValueError(
            f"Não foi possível localizar a coluna de data/hora. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    if coluna_alerta is None:
        raise ValueError(
            f"Não foi possível localizar a coluna de texto do alerta. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    # ========================================================
    # DATA E HORA
    # ========================================================

    df["DataHora"] = pd.to_datetime(
        df[coluna_data_hora],
        errors="coerce",
        dayfirst=True
    )

    df = df.dropna(subset=["DataHora"])

    df["Data"] = df["DataHora"].dt.strftime("%d/%m/%Y")
    df["Hora"] = df["DataHora"].dt.strftime("%H:%M:%S")
    df["HoraNumero"] = df["DataHora"].dt.hour.astype(int)

    # ========================================================
    # FROTA
    # ========================================================

    if coluna_frota is not None:

        df["Frota"] = (
            df[coluna_frota]
            .fillna("NÃO IDENTIFICADA")
            .astype(str)
            .str.strip()
        )

    else:
        df["Frota"] = "NÃO IDENTIFICADA"

    # ========================================================
    # DRIVER
    # ========================================================

    if coluna_driver is not None:

        df["Driver"] = (
            df[coluna_driver]
            .fillna("-----")
            .astype(str)
            .str.strip()
        )

    else:
        df["Driver"] = "-----"

    # ========================================================
    # VELOCIDADES
    # ========================================================

    velocidades = df[coluna_alerta].apply(
        extrair_velocidades
    )

    df["Velocidade Atual"] = velocidades.apply(
        lambda x: x[0]
    )

    df["Velocidade Limite"] = velocidades.apply(
        lambda x: x[1]
    )

    df = df.dropna(
        subset=[
            "Velocidade Atual",
            "Velocidade Limite"
        ]
    )

    # ========================================================
    # EXCESSO
    # ========================================================

    df["Excesso"] = (
        df["Velocidade Atual"]
        - df["Velocidade Limite"]
    )

    # ========================================================
    # CERCA E GRAVIDADE
    # ========================================================

    df["Cerca"] = df[coluna_alerta].apply(
        identificar_cerca
    )

    df["Gravidade"] = df.apply(
        lambda linha: classificar_gravidade(
            linha["Excesso"],
            linha["Cerca"]
        ),
        axis=1
    )

    # ========================================================
    # DADOS FINAIS
    # ========================================================

    colunas_final = [
        "DataHora",
        "Data",
        "Hora",
        "Driver",
        "Frota",
        "Velocidade Atual",
        "Velocidade Limite",
        "Excesso",
        "Cerca",
        "Gravidade"
    ]

    df_final = df[colunas_final].copy()

    df_final = df_final.sort_values(
        by="DataHora",
        ascending=False
    ).reset_index(drop=True)

    data_min = df_final["DataHora"].min()
    data_max = df_final["DataHora"].max()

    periodo = (
        f"{data_min.strftime('%d/%m/%Y')} até "
        f"{data_max.strftime('%d/%m/%Y')}"
    )

    return {
        "arquivo": nome_arquivo,
        "aba": aba,
        "periodo": periodo,
        "dados": df_final
    }




# ============================================================
# PROCESSAR MANUTENÇÃO
# ============================================================

def extrair_km_manutencao(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).strip().lower()
    if texto in ("", "-----", "nan"):
        return None
    encontrado = re.search(r"(-?\d[\d\.,]*)\s*km", texto)
    if not encontrado:
        return None
    numero = encontrado.group(1).replace(".", "").replace(",", ".")
    try:
        return float(numero)
    except ValueError:
        return None


# Frotas cadastradas como motocicletas
MOTOS = {"FNP-0186", "FNP-0372"}


def classificar_status_manutencao(estado, km, frota):
    texto = str(estado).strip().lower()
    frota_normalizada = str(frota).strip().upper()

    if "expir" in texto:
        return "VENCIDA"

    if km is None or pd.isna(km):
        return "SEM DADO"

    # Motos: ciclo de manutenção de 1.000 km
    if frota_normalizada in MOTOS:
        if km <= 100:
            return "CRÍTICA"
        elif km <= 500:
            return "PRÓXIMA"
        return "EM DIA"

    # Demais veículos: ciclo de referência de 10.000 km
    if km <= 1000:
        return "CRÍTICA"
    elif km <= 5000:
        return "PRÓXIMA"
    return "EM DIA"


def processar_manutencao():
    caminho_excel = localizar_excel_manutencao()
    nome_arquivo = os.path.basename(caminho_excel)
    excel = pd.ExcelFile(caminho_excel)
    aba = next(
        (
            nome for nome in excel.sheet_names
            if "próxima manutenção" in str(nome).lower()
            or "proxima manutencao" in str(nome).lower()
            or "manuten" in str(nome).lower()
        ),
        excel.sheet_names[0]
    )

    df = pd.read_excel(caminho_excel, sheet_name=aba)
    df = df.dropna(how="all")
    df.columns = [str(coluna).strip() for coluna in df.columns]

    coluna_frota = localizar_coluna(df.columns, ["agrupamento", "frota", "grupo"])
    coluna_manutencao = localizar_coluna(df.columns, ["manutenção", "manutencao"])
    coluna_estado = localizar_coluna(
        df.columns,
        ["estado por quilometragem", "quilometragem", "km restantes"]
    )
    coluna_descricao = localizar_coluna(
        df.columns,
        ["descrição", "descricao", "serviço", "servico"]
    )

    if coluna_frota is None or coluna_estado is None:
        raise ValueError(
            "Não foi possível localizar as colunas obrigatórias da manutenção. "
            f"Colunas encontradas: {list(df.columns)}"
        )

    dados = pd.DataFrame()
    dados["Frota"] = df[coluna_frota].fillna("").astype(str).str.strip()
    dados["Manutenção"] = (
        df[coluna_manutencao].fillna("").astype(str).str.strip()
        if coluna_manutencao is not None else ""
    )
    dados["Estado por quilometragem"] = (
        df[coluna_estado].fillna("").astype(str).str.strip()
    )
    dados["Descrição"] = (
        df[coluna_descricao].fillna("").astype(str).str.strip()
        if coluna_descricao is not None else ""
    )

    dados = dados[
        (dados["Frota"] != "")
        & (dados["Frota"].str.lower() != "nan")
        & (dados["Estado por quilometragem"] != "-----")
        & (dados["Estado por quilometragem"] != "")
    ].copy()

    dados["KM"] = dados["Estado por quilometragem"].apply(extrair_km_manutencao)
    dados["Status"] = dados.apply(
        lambda linha: classificar_status_manutencao(
            linha["Estado por quilometragem"],
            linha["KM"],
            linha["Frota"]
        ),
        axis=1
    )

    ordem_status = {
        "VENCIDA": 0,
        "CRÍTICA": 1,
        "PRÓXIMA": 2,
        "EM DIA": 3,
        "SEM DADO": 4
    }
    dados["OrdemStatus"] = dados["Status"].map(ordem_status).fillna(99)

    dados = dados.sort_values(
        by=["OrdemStatus", "KM", "Frota"],
        ascending=[True, True, True],
        na_position="last"
    ).reset_index(drop=True)

    return {"arquivo": nome_arquivo, "aba": aba, "dados": dados}


def normalizar_texto(texto):
    texto = str(texto).upper()
    texto = re.sub(r"[ÁÀÃÂÄ]", "A", texto)
    texto = re.sub(r"[ÉÈÊË]", "E", texto)
    texto = re.sub(r"[ÍÌÎÏ]", "I", texto)
    texto = re.sub(r"[ÓÒÕÔÖ]", "O", texto)
    texto = re.sub(r"[ÚÙÛÜ]", "U", texto)
    texto = texto.replace("Ç", "C")
    return texto


def identificar_categoria_manutencao(manutencao, descricao):
    texto = normalizar_texto(f"{manutencao} {descricao}")

    # Primeiro identifica pneus para evitar misturar serviços específicos.
    if any(palavra in texto for palavra in [
        "PNEU", "PNEUS", "RODAGEM", "RODIZIO", "ALINHAMENTO",
        "BALANCEAMENTO", "CAMBAGEM"
    ]):
        return "pneus"

    if any(palavra in texto for palavra in [
        "OLEO", "LUBRIFIC", "LUBRIFICACAO", "FILTRO DE OLEO"
    ]):
        return "oleo"

    return "outros"


def dados_dashboard_manutencao(categoria="geral"):
    resultado = processar_manutencao()
    dados = resultado["dados"].copy()

    categoria = str(categoria or "geral").strip().lower()
    categorias_validas = {"geral", "oleo", "pneus"}
    if categoria not in categorias_validas:
        categoria = "geral"

    dados["Categoria"] = dados.apply(
        lambda linha: identificar_categoria_manutencao(
            linha["Manutenção"], linha["Descrição"]
        ),
        axis=1
    )

    if categoria in {"oleo", "pneus"}:
        dados = dados[dados["Categoria"] == categoria].copy()

    total = len(dados)
    frotas = int(dados["Frota"].nunique()) if not dados.empty else 0
    vencidas = int((dados["Status"] == "VENCIDA").sum())
    criticas = int((dados["Status"] == "CRÍTICA").sum())
    proximas = int((dados["Status"] == "PRÓXIMA").sum())
    em_dia = int((dados["Status"] == "EM DIA").sum())

    ranking = dados.sort_values(
        by=["OrdemStatus", "KM", "Frota"],
        ascending=[True, True, True],
        na_position="last"
    ).head(10)

    cores = {
        "VENCIDA": "#C62828",
        "CRÍTICA": "#D97706",
        "PRÓXIMA": "#E0AA16",
        "EM DIA": "#1F7A45",
        "SEM DADO": "#64748B"
    }

    nomes_categoria = {
        "geral": "GERAL",
        "oleo": "TROCA DE ÓLEO",
        "pneus": "PNEUS"
    }

    return {
        "arquivo": resultado["arquivo"],
        "aba": resultado["aba"],
        "categoria": categoria,
        "categoria_nome": nomes_categoria[categoria],
        "total": total,
        "frotas": frotas,
        "vencidas": vencidas,
        "criticas": criticas,
        "proximas": proximas,
        "em_dia": em_dia,
        "grafico_labels": ranking["Frota"].tolist(),
        "grafico_valores": [
            int(valor) if pd.notna(valor) else 0
            for valor in ranking["KM"].tolist()
        ],
        "grafico_cores": [
            cores.get(status, "#64748B")
            for status in ranking["Status"].tolist()
        ],
        "tabela": dados[
            ["Frota", "Manutenção", "Descrição", "Estado por quilometragem", "Status"]
        ].copy()
    }


# ============================================================
# LOGIN
# ============================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    erro = ""

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        senha = request.form.get("senha", "")

        if usuario == USUARIO_LOGIN and senha == SENHA_LOGIN:
            session["autenticado"] = True

            proxima_url = request.args.get("next")
            if proxima_url:
                return redirect(proxima_url)

            return redirect(url_for("index"))

        erro = "Usuário ou senha incorretos."

    return render_template("login.html", erro=erro)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ============================================================
# ROTA PRINCIPAL
# ============================================================

@app.route("/")
def index():

    if not session.get("autenticado"):
        return redirect(url_for("login", next=request.url))

    try:

        painel = request.args.get("painel", "ocorrencias").strip().lower()

        if painel == "manutencao":
            categoria_manutencao = request.args.get("categoria", "geral")
            manutencao = dados_dashboard_manutencao(categoria_manutencao)
            return render_template(
                "index.html",
                painel="manutencao",
                manutencao=manutencao
            )

        resultado = processar_dados()

        dados_completos = resultado["dados"].copy()

        # ====================================================
        # OPÇÕES DOS FILTROS
        # ====================================================

        opcoes_frota = sorted(
            dados_completos["Frota"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        opcoes_condutor = sorted(
            dados_completos["Driver"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        opcoes_cerca = sorted(
            dados_completos["Cerca"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        # Opções exibidas no filtro de gravidade.
        # A categoria antiga "FORA DA FAIXA" não aparece mais.
        opcoes_gravidade = [
            "LEVE",
            "MÉDIA",
            "GRAVE",
            "GRAVÍSSIMA"
        ]

        # ====================================================
        # RECEBER FILTROS
        # ====================================================

        data_inicio = request.args.get("data_inicio", "")
        data_fim = request.args.get("data_fim", "")
        frota = request.args.get("frota", "")
        condutor = request.args.get("condutor", "")
        cerca = request.args.get("cerca", "")
        gravidade = request.args.get("gravidade", "")
        faixa_excesso = request.args.get("faixa_excesso", "")
        hora = request.args.get("hora", "")

        # Compatibilidade com links/filtros antigos.
        if str(gravidade).strip().upper() == "FORA DA FAIXA":
            gravidade = "GRAVÍSSIMA"

        filtros = {
            "data_inicio": data_inicio,
            "data_fim": data_fim,
            "frota": frota,
            "condutor": condutor,
            "cerca": cerca,
            "gravidade": gravidade,
            "faixa_excesso": faixa_excesso,
            "hora": hora
        }

        dados = dados_completos.copy()

        # ====================================================
        # FILTRO DE DATA
        # ====================================================

        if data_inicio:

            inicio = pd.to_datetime(
                data_inicio,
                errors="coerce"
            )

            if pd.notna(inicio):

                dados = dados[
                    dados["DataHora"] >= inicio
                ]

        if data_fim:

            fim = pd.to_datetime(
                data_fim,
                errors="coerce"
            )

            if pd.notna(fim):

                fim = fim + pd.Timedelta(days=1)

                dados = dados[
                    dados["DataHora"] < fim
                ]

        # ====================================================
        # FILTRO DE FROTA
        # ====================================================

        if frota:

            dados = dados[
                dados["Frota"].astype(str) == frota
            ]

        # ====================================================
        # FILTRO DE CONDUTOR
        # ====================================================

        if condutor:

            dados = dados[
                dados["Driver"].astype(str) == condutor
            ]

        # ====================================================
        # FILTRO DE CERCA
        # ====================================================

        if cerca:

            dados = dados[
                dados["Cerca"].astype(str) == cerca
            ]

        # ====================================================
        # FILTRO DE GRAVIDADE
        # ====================================================

        if gravidade:

            dados = dados[
                dados["Gravidade"].astype(str) == gravidade
            ]

        # ====================================================
        # FILTRO DE FAIXA DE EXCESSO
        # ====================================================

        if faixa_excesso:

            excesso_numerico = pd.to_numeric(
                dados["Excesso"],
                errors="coerce"
            )

            if faixa_excesso == "1 a 5 km/h":

                dados = dados[
                    (excesso_numerico >= 1) &
                    (excesso_numerico <= 5)
                ]

            elif faixa_excesso == "6 a 10 km/h":

                dados = dados[
                    (excesso_numerico > 5) &
                    (excesso_numerico <= 10)
                ]

            elif faixa_excesso == "11 a 20 km/h":

                dados = dados[
                    (excesso_numerico > 10) &
                    (excesso_numerico <= 20)
                ]

            elif faixa_excesso == "Acima de 20 km/h":

                dados = dados[
                    excesso_numerico > 20
                ]

        # ====================================================
        # FILTRO DE HORÁRIO
        # ====================================================

        if hora:
            try:
                hora_numero = int(str(hora).split(":")[0])
                if 0 <= hora_numero <= 23:
                    dados = dados[
                        dados["DataHora"].dt.hour == hora_numero
                    ]
            except (ValueError, TypeError):
                pass

        # ====================================================
        # INDICADORES
        # ====================================================

        total = len(dados)

        condutores_validos = dados["Driver"].astype(str)

        condutores_validos = condutores_validos[
            (condutores_validos != "-----") &
            (condutores_validos.str.lower() != "nan") &
            (condutores_validos != "") &
            (condutores_validos.str.upper() != "NÃO IDENTIFICADO") &
            (condutores_validos.str.upper() != "NÃO IDENTIFICADA")
        ]

        quantidade_condutores = condutores_validos.nunique()

        frotas_validas = dados["Frota"].astype(str)

        frotas_validas = frotas_validas[
            (frotas_validas.str.lower() != "nan") &
            (frotas_validas != "") &
            (frotas_validas.str.upper() != "NÃO IDENTIFICADA")
        ]

        quantidade_frotas = frotas_validas.nunique()

        leves = int(
            (dados["Gravidade"] == "LEVE").sum()
        )

        medias = int(
            (dados["Gravidade"] == "MÉDIA").sum()
        )

        graves = int(
            (dados["Gravidade"] == "GRAVE").sum()
        )

        gravissimas = int(
            (dados["Gravidade"] == "GRAVÍSSIMA").sum()
        )

        fora_faixa = int(
            (dados["Gravidade"] == "GRAVÍSSIMA").sum()
        )

        # ====================================================
        # RANKING POR CONDUTOR
        # ====================================================

        ranking_condutores = (
            condutores_validos
            .value_counts()
            .sort_values(ascending=False)
        )

        grafico_condutores_labels = [
            str(nome)
            for nome in ranking_condutores.index.tolist()
        ]

        grafico_condutores_valores = [
            int(valor)
            for valor in ranking_condutores.values.tolist()
        ]

        # ====================================================
        # RANKING POR FROTA
        # ====================================================

        ranking_frotas = (
            dados["Frota"]
            .astype(str)
            .value_counts()
            .sort_values(ascending=False)
        )

        grafico_frotas_labels = [
            str(nome)
            for nome in ranking_frotas.index.tolist()
        ]

        grafico_frotas_valores = [
            int(valor)
            for valor in ranking_frotas.values.tolist()
        ]

        # ====================================================
        # GRÁFICO DE GRAVIDADE
        # ====================================================

        grafico_gravidade_labels = [
            "LEVE",
            "MÉDIA",
            "GRAVE",
            "GRAVÍSSIMA"
        ]

        grafico_gravidade_valores = [
            leves,
            medias,
            graves,
            gravissimas
        ]

        # ====================================================
        # GRÁFICO FAIXA DE EXCESSO
        # ====================================================

        excesso_numerico = pd.to_numeric(
            dados["Excesso"],
            errors="coerce"
        )

        faixa_1_5 = int(
            (
                (excesso_numerico >= 1) &
                (excesso_numerico <= 5)
            ).sum()
        )

        faixa_6_10 = int(
            (
                (excesso_numerico > 5) &
                (excesso_numerico <= 10)
            ).sum()
        )

        faixa_11_20 = int(
            (
                (excesso_numerico > 10) &
                (excesso_numerico <= 20)
            ).sum()
        )

        faixa_acima_20 = int(
            (excesso_numerico > 20).sum()
        )

        grafico_excesso_labels = [
            "1 a 5 km/h",
            "6 a 10 km/h",
            "11 a 20 km/h",
            "Acima de 20 km/h"
        ]

        grafico_excesso_valores = [
            faixa_1_5,
            faixa_6_10,
            faixa_11_20,
            faixa_acima_20
        ]

        # ====================================================
        # GRÁFICO DE OCORRÊNCIAS POR HORÁRIO
        # ====================================================

        horas_completas = list(range(24))

        ocorrencias_por_hora = (
            dados["DataHora"]
            .dt.hour
            .value_counts()
            .reindex(horas_completas, fill_value=0)
            .sort_index()
        )

        grafico_horario_labels = [
            f"{hora:02d}:00"
            for hora in horas_completas
        ]

        grafico_horario_valores = [
            int(ocorrencias_por_hora.loc[hora])
            for hora in horas_completas
        ]

        if total > 0:
            hora_critica_numero = int(ocorrencias_por_hora.idxmax())
            hora_critica = f"{hora_critica_numero:02d}:00"
            hora_critica_quantidade = int(ocorrencias_por_hora.max())
        else:
            hora_critica = "--:--"
            hora_critica_quantidade = 0

        # ====================================================
        # GRÁFICO HORÁRIO × GRAVIDADE
        # ====================================================

        ordem_gravidade = [
            "LEVE",
            "MÉDIA",
            "GRAVE",
            "GRAVÍSSIMA"
        ]

        dados_hora_gravidade = pd.crosstab(
            dados["DataHora"].dt.hour,
            dados["Gravidade"]
        ).reindex(
            index=horas_completas,
            columns=ordem_gravidade,
            fill_value=0
        )

        grafico_hora_gravidade_leve = [
            int(dados_hora_gravidade.loc[h, "LEVE"])
            for h in horas_completas
        ]

        grafico_hora_gravidade_media = [
            int(dados_hora_gravidade.loc[h, "MÉDIA"])
            for h in horas_completas
        ]

        grafico_hora_gravidade_grave = [
            int(dados_hora_gravidade.loc[h, "GRAVE"])
            for h in horas_completas
        ]

        grafico_hora_gravidade_gravissima = [
            int(dados_hora_gravidade.loc[h, "GRAVÍSSIMA"])
            for h in horas_completas
        ]

        # ====================================================
        # PERÍODO APÓS FILTROS
        # ====================================================

        if not dados.empty:

            data_min = dados["DataHora"].min()
            data_max = dados["DataHora"].max()

            periodo = (
                f"{data_min.strftime('%d/%m/%Y')} até "
                f"{data_max.strftime('%d/%m/%Y')}"
            )

        else:

            periodo = (
                "SEM OCORRÊNCIAS PARA OS FILTROS SELECIONADOS"
            )

        # ====================================================
        # TABELA
        # ====================================================

        colunas_tabela = [
            "Data",
            "Hora",
            "Driver",
            "Frota",
            "Velocidade Atual",
            "Velocidade Limite",
            "Excesso",
            "Cerca",
            "Gravidade"
        ]

        tabela = dados[
            colunas_tabela
        ].head(50)

        # ====================================================
        # ENVIAR PARA O HTML
        # ====================================================

        return render_template(
            "index.html",
            painel="ocorrencias",

            arquivo=resultado["arquivo"],
            aba=resultado["aba"],
            periodo=periodo,

            total=total,
            condutores=quantidade_condutores,
            frotas=quantidade_frotas,

            leves=leves,
            medias=medias,
            graves=graves,
            gravissimas=gravissimas,
            fora_faixa=fora_faixa,

            tabela=tabela,

            filtros=filtros,

            opcoes_frota=opcoes_frota,
            opcoes_condutor=opcoes_condutor,
            opcoes_cerca=opcoes_cerca,
            opcoes_gravidade=opcoes_gravidade,

            grafico_condutores_labels=grafico_condutores_labels,
            grafico_condutores_valores=grafico_condutores_valores,

            grafico_frotas_labels=grafico_frotas_labels,
            grafico_frotas_valores=grafico_frotas_valores,

            grafico_gravidade_labels=grafico_gravidade_labels,
            grafico_gravidade_valores=grafico_gravidade_valores,

            grafico_excesso_labels=grafico_excesso_labels,
            grafico_excesso_valores=grafico_excesso_valores,

            grafico_horario_labels=grafico_horario_labels,
            grafico_horario_valores=grafico_horario_valores,

            grafico_hora_gravidade_leve=grafico_hora_gravidade_leve,
            grafico_hora_gravidade_media=grafico_hora_gravidade_media,
            grafico_hora_gravidade_grave=grafico_hora_gravidade_grave,
            grafico_hora_gravidade_gravissima=grafico_hora_gravidade_gravissima,

            hora_critica=hora_critica,
            hora_critica_quantidade=hora_critica_quantidade
        )

    except Exception as erro:

        return (
            "<h1>CENTRAL DE MONITORAMENTO - FROTAS LEVES</h1>"
            "<h2>Erro ao carregar dados</h2>"
            f"<p>Erro: {erro}</p>"
            f"<p><b>Pasta de dados:</b> {PASTA_DADOS}</p>"
        )


# ============================================================
# INICIAR SISTEMA
# ============================================================

if __name__ == "__main__":

    app.run(
        debug=True,
        host="127.0.0.1",
        port=5000
    )