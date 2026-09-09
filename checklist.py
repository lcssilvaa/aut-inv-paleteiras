from pathlib import Path
from datetime import datetime, timedelta
import base64
import html
import os
import re
import unicodedata

import pandas as pd
from dotenv import load_dotenv

from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

SAIDA_DIR = BASE_DIR / "saida"

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

SHAREPOINT_DIR = Path(os.getenv("SHAREPOINT_DIR", ""))

ARQUIVO_EQUIPAMENTOS = SHAREPOINT_DIR / os.getenv(
    "CHECKLIST_ARQUIVO_EQUIPAMENTOS",
    "Base Checklist.xlsx",
)

ARQUIVO_LEITURAS = SHAREPOINT_DIR / os.getenv(
    "CHECKLIST_ARQUIVO_LEITURAS",
    "Especifico Cliente - Relatorio BI - Checklist Veicular.xlsx",
)

PRAZO_DIAS = int(
    os.getenv(
        "CHECKLIST_PRAZO_DIAS",
        "7",
    )
)


def caminho_configuracao(nome_variavel, padrao):

    valor = os.getenv(nome_variavel, padrao)

    caminho = Path(valor)

    if not caminho.is_absolute():
        caminho = BASE_DIR / caminho

    return caminho


CREDENTIALS_FILE = caminho_configuracao(
    "GMAIL_CREDENTIALS_FILE",
    "credentials.json",
)

TOKEN_FILE = caminho_configuracao(
    "GMAIL_TOKEN_FILE",
    "token.json",
)


def carregar_destinatarios():

    destinatarios = {}

    filiais = [
        "AJU",
        "BACKUP",
        "BAU",
        "BHZ",
        "BNU",
        "BRA",
        "CAW",
        "CCM",
        "CEDOC",
        "CON",
        "CPQ",
        "CTN",
        "CWB",
        "DIQ",
        "ESTOQUE",
        "FEC",
        "FLN",
        "FOR",
        "GVR",
        "IPN",
        "ITN",
        "JDF",
        "JMV",
        "JOI",
        "LDB",
        "MHA",
        "MOC",
        "NSE",
        "PNZ",
        "PPB",
        "PPY",
        "PREDIAL",
        "PSW",
        "PTM",
        "QDF",
        "QHG",
        "QHV",
        "QXD",
        "RAO",
        "RIO",
        "SAJ",
        "SALVADOS",
        "SÃO",
        "SEI",
        "SJK",
        "SJP",
        "SSA",
        "TFL",
        "UBA",
        "UDI",
        "VAG",
        "VDC",
        "VIX",
    ]

    for filial in filiais:

        emails = []

        for numero in range(1, 10):

            email = os.getenv(f"{filial}_EMAIL_{numero}")

            if email:
                email = email.strip()

                if email:
                    emails.append(email)

        if emails:
            destinatarios[filial] = emails

    return destinatarios


DESTINATARIOS = carregar_destinatarios()


def limpar_texto(valor):

    if pd.isna(valor):
        return ""

    return " ".join(str(valor).strip().split())


def remover_acentos(valor):

    valor = limpar_texto(valor)

    return "".join(
        caractere
        for caractere in unicodedata.normalize(
            "NFKD",
            valor,
        )
        if not unicodedata.combining(caractere)
    )


def normalizar_texto(valor):

    return remover_acentos(valor).upper()


def normalizar_placa(valor):

    return normalizar_texto(valor)


def normalizar_codigo(valor):

    if pd.isna(valor):
        return None

    valor = str(valor).strip()

    if valor.endswith(".0"):
        valor = valor[:-2]

    valor = re.sub(
        r"\s+",
        "",
        valor,
    )

    return valor or None


def normalizar_tipo(valor):

    return normalizar_texto(valor)


def interpretar_ativo(valor):

    valor_normalizado = normalizar_texto(valor)

    if valor_normalizado in {"", "SIM", "S", "TRUE", "VERDADEIRO", "1", "1.0"}:
        return True

    if valor_normalizado in {"NAO", "N", "FALSE", "FALSO", "0", "0.0"}:
        return False

    raise ValueError(
        f"Valor inválido na coluna Ativo: {valor!r}. "
        "Use Sim ou Não; vazio mantém o equipamento ativo."
    )


def formatar_data(valor):

    if pd.isna(valor):
        return "-"

    return pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")


def formatar_valor(valor):

    if pd.isna(valor):
        return "-"

    return html.escape(str(valor))


def validar_bases():

    arquivos = {
        "Base de equipamentos": ARQUIVO_EQUIPAMENTOS,
        "Histórico de checklists": ARQUIVO_LEITURAS,
    }

    for nome, arquivo in arquivos.items():

        if not arquivo.exists():

            raise FileNotFoundError(f"{nome} não encontrado:\n{arquivo}")

        if arquivo.stat().st_size == 0:

            raise ValueError(f"{nome} está vazio.")

        modificacao = datetime.fromtimestamp(arquivo.stat().st_mtime)

        print(f"{nome}: " f"{modificacao.strftime('%d/%m/%Y %H:%M:%S')}")


def carregar_equipamentos():

    print("\n[1/5] Carregando cadastro de equipamentos...")

    df = pd.read_excel(ARQUIVO_EQUIPAMENTOS)

    df.columns = df.columns.astype(str).str.strip()

    colunas = [
        "Placa",
        "Código Filial",
        "Filial",
        "Tipo de Equipamento",
    ]

    faltantes = [coluna for coluna in colunas if coluna not in df.columns]

    if faltantes:

        raise ValueError("Colunas não encontradas na Base Checklist: " f"{faltantes}")

    if "Ativo" in df.columns:
        colunas.append("Ativo")

    df = df[colunas].copy()

    df = df.dropna(subset=["Placa"])

    df["ATIVO"] = df["Ativo"].apply(interpretar_ativo) if "Ativo" in df else True

    df["PLACA_NORMALIZADA"] = df["Placa"].apply(normalizar_placa)

    df["CODIGO_FILIAL_NORMALIZADO"] = df["Código Filial"].apply(normalizar_codigo)

    df["FILIAL"] = df["Filial"].apply(limpar_texto).str.upper()

    df["TIPO_ESPERADO_NORMALIZADO"] = df["Tipo de Equipamento"].apply(normalizar_tipo)

    equipamentos = []
    desabilitados = 0

    for placa, grupo in df.groupby(
        "PLACA_NORMALIZADA",
        sort=False,
    ):

        # A desativação vale para a placa inteira, incluindo todos os seus tipos.
        if not grupo["ATIVO"].all():
            desabilitados += 1
            continue

        placas = [
            limpar_texto(valor) for valor in grupo["Placa"] if limpar_texto(valor)
        ]

        codigos = [
            normalizar_codigo(valor)
            for valor in grupo["Código Filial"]
            if normalizar_codigo(valor)
        ]

        filiais = [
            limpar_texto(valor).upper()
            for valor in grupo["Filial"]
            if limpar_texto(valor)
        ]

        tipos_exibicao = list(
            dict.fromkeys(
                limpar_texto(valor)
                for valor in grupo["Tipo de Equipamento"]
                if limpar_texto(valor)
            )
        )

        tipos_normalizados = set(
            normalizar_tipo(valor)
            for valor in grupo["Tipo de Equipamento"]
            if normalizar_tipo(valor)
        )

        codigos_unicos = list(dict.fromkeys(codigos))

        filiais_unicas = list(dict.fromkeys(filiais))

        if len(codigos_unicos) > 1:

            print(
                f"ATENÇÃO: {placa} possui mais de um "
                f"Código Filial na Base Checklist: "
                f"{codigos_unicos}"
            )

        if len(filiais_unicas) > 1:

            print(
                f"ATENÇÃO: {placa} possui mais de uma "
                f"Filial na Base Checklist: "
                f"{filiais_unicas}"
            )

        equipamentos.append(
            {
                "PLACA_NORMALIZADA": placa,
                "Placa": placas[0] if placas else placa,
                "Código Filial": grupo["Código Filial"].iloc[0],
                "CODIGO_FILIAL_NORMALIZADO": (
                    codigos_unicos[0] if codigos_unicos else None
                ),
                "FILIAL": (filiais_unicas[0] if filiais_unicas else ""),
                "TIPOS_ACEITOS": tipos_exibicao,
                "TIPOS_ACEITOS_NORMALIZADOS": tipos_normalizados,
            }
        )

    equipamentos = pd.DataFrame(
        equipamentos,
        columns=[
            "PLACA_NORMALIZADA",
            "Placa",
            "Código Filial",
            "CODIGO_FILIAL_NORMALIZADO",
            "FILIAL",
            "TIPOS_ACEITOS",
            "TIPOS_ACEITOS_NORMALIZADOS",
        ],
    )

    equipamentos["TIPOS_ACEITOS_TEXTO"] = equipamentos["TIPOS_ACEITOS"].apply(
        lambda tipos: " / ".join(tipos)
    )

    print(f"    Equipamentos ativos: {len(equipamentos)}")
    print(f"    Equipamentos desabilitados: {desabilitados}")

    print("    Registros de placa/tipo na base: " f"{len(df)}")

    return equipamentos


def converter_data_resposta(
    data_resposta,
    anomes,
):

    if pd.isna(data_resposta):
        return pd.NaT

    if isinstance(
        data_resposta,
        (pd.Timestamp, datetime),
    ):
        return pd.Timestamp(data_resposta)

    valor = limpar_texto(data_resposta)

    if not valor:
        return pd.NaT

    ano_esperado = None
    mes_esperado = None

    if not pd.isna(anomes):

        texto_anomes = str(anomes).strip()

        encontrado = re.search(
            r"(\d{4})[/\-](\d{1,2})",
            texto_anomes,
        )

        if encontrado:

            ano_esperado = int(encontrado.group(1))

            mes_esperado = int(encontrado.group(2))

    candidatos = []

    for dayfirst in [False, True]:

        try:

            data = pd.to_datetime(
                valor,
                dayfirst=dayfirst,
                errors="coerce",
            )

            if not pd.isna(data):

                data = pd.Timestamp(data)

                if data not in candidatos:
                    candidatos.append(data)

        except Exception:
            pass

    if ano_esperado is not None and mes_esperado is not None:

        for candidato in candidatos:

            if candidato.year == ano_esperado and candidato.month == mes_esperado:
                return candidato

    if candidatos:
        return candidatos[0]

    return pd.NaT


def carregar_historico():

    print("\n[2/5] Carregando histórico de checklists...")

    df = pd.read_excel(
        ARQUIVO_LEITURAS,
        header=5,
    )

    df.columns = df.columns.astype(str).str.strip()

    colunas_necessarias = [
        "CODIGOFILIALRESPOSTA",
        "NOMECHECKLIST",
        "DATARESPOSTA",
        "ANOMES",
        "DATA",
        "RESPONDIDOPOR",
        "PLACAVEICULO",
        "RESPOSTALOG",
    ]

    faltantes = [coluna for coluna in colunas_necessarias if coluna not in df.columns]

    if faltantes:

        raise ValueError("Colunas não encontradas no histórico: " f"{faltantes}")

    df = df[colunas_necessarias].copy()

    df = df.dropna(subset=["PLACAVEICULO"])

    df["PLACA_NORMALIZADA"] = df["PLACAVEICULO"].apply(normalizar_placa)

    df["CODIGO_FILIAL_RESPOSTA_NORMALIZADO"] = df["CODIGOFILIALRESPOSTA"].apply(
        normalizar_codigo
    )

    df["TIPO_REALIZADO_NORMALIZADO"] = df["NOMECHECKLIST"].apply(normalizar_tipo)

    df["DATARESPOSTA_CONVERTIDA"] = df.apply(
        lambda row: converter_data_resposta(
            row["DATARESPOSTA"],
            row["ANOMES"],
        ),
        axis=1,
    )

    print(f"    Registros históricos: {len(df)}")

    print("    Datas inválidas: " f"{df['DATARESPOSTA_CONVERTIDA'].isna().sum()}")

    return df


def obter_ultimos_checklists(
    historico,
):

    historico_valido = historico[historico["DATARESPOSTA_CONVERTIDA"].notna()].copy()

    historico_valido = historico_valido.sort_values(
        "DATARESPOSTA_CONVERTIDA"
    ).drop_duplicates(
        subset=["PLACA_NORMALIZADA"],
        keep="last",
    )

    return historico_valido


def processar(
    equipamentos,
    historico,
):

    print("\n[3/5] Processando checklists...")

    ultimos = obter_ultimos_checklists(historico)

    ultimos = ultimos[
        [
            "PLACA_NORMALIZADA",
            "CODIGOFILIALRESPOSTA",
            "CODIGO_FILIAL_RESPOSTA_NORMALIZADO",
            "NOMECHECKLIST",
            "TIPO_REALIZADO_NORMALIZADO",
            "DATARESPOSTA_CONVERTIDA",
            "RESPONDIDOPOR",
            "RESPOSTALOG",
        ]
    ].copy()

    ultimos = ultimos.rename(
        columns={
            "DATARESPOSTA_CONVERTIDA": "ULTIMA_RESPOSTA",
            "RESPONDIDOPOR": "ULTIMO_RESPONSAVEL",
            "NOMECHECKLIST": "CHECKLIST_REALIZADO",
        }
    )

    resultado = equipamentos.merge(
        ultimos,
        on="PLACA_NORMALIZADA",
        how="left",
    )

    agora = pd.Timestamp.now()

    resultado["DATA_ANALISE"] = agora

    resultado["VENCIMENTO"] = resultado["ULTIMA_RESPOSTA"] + timedelta(days=PRAZO_DIAS)

    resultado["STATUS_PRAZO"] = "PENDENTE"

    mascara_ok = resultado["ULTIMA_RESPOSTA"].notna() & (
        agora < resultado["VENCIMENTO"]
    )

    resultado.loc[
        mascara_ok,
        "STATUS_PRAZO",
    ] = "OK"

    resultado["MOTIVO_PENDENCIA"] = ""

    mascara_nunca = resultado["ULTIMA_RESPOSTA"].isna()

    resultado.loc[
        mascara_nunca,
        "MOTIVO_PENDENCIA",
    ] = "Nunca realizado"

    mascara_vencido = resultado["ULTIMA_RESPOSTA"].notna() & (
        resultado["STATUS_PRAZO"] == "PENDENTE"
    )

    resultado.loc[
        mascara_vencido,
        "MOTIVO_PENDENCIA",
    ] = f"Prazo superior a {PRAZO_DIAS} dias"

    diferenca = agora - resultado["ULTIMA_RESPOSTA"]

    resultado["HORAS_SEM_CHECKLIST"] = (diferenca.dt.total_seconds() / 3600).round(2)

    resultado["DIAS_SEM_CHECKLIST"] = (resultado["HORAS_SEM_CHECKLIST"] / 24).round(2)

    resultado["HORAS_RESTANTES"] = (
        (resultado["VENCIMENTO"] - agora).dt.total_seconds() / 3600
    ).round(2)

    def verificar_inconsistencia_tipo(row):

        if pd.isna(row["ULTIMA_RESPOSTA"]):
            return False

        tipo_realizado = row["TIPO_REALIZADO_NORMALIZADO"]

        tipos_aceitos = row["TIPOS_ACEITOS_NORMALIZADOS"]

        if not tipo_realizado:
            return True

        return tipo_realizado not in tipos_aceitos

    resultado["INCONSISTENCIA_TIPO"] = pd.Series(
        [verificar_inconsistencia_tipo(row) for _, row in resultado.iterrows()],
        index=resultado.index,
        dtype=bool,
    )

    resultado["INCONSISTENCIA_FILIAL"] = False

    possui_resposta = resultado["ULTIMA_RESPOSTA"].notna()

    resultado.loc[
        possui_resposta,
        "INCONSISTENCIA_FILIAL",
    ] = (
        resultado.loc[
            possui_resposta,
            "CODIGO_FILIAL_NORMALIZADO",
        ]
        != resultado.loc[
            possui_resposta,
            "CODIGO_FILIAL_RESPOSTA_NORMALIZADO",
        ]
    )

    def descricao_inconsistencia(row):

        erros = []

        if row["INCONSISTENCIA_TIPO"]:

            erros.append(
                "Tipo(s) aceito(s): "
                f"{row['TIPOS_ACEITOS_TEXTO']} | "
                "Checklist realizado: "
                f"{row['CHECKLIST_REALIZADO']}"
            )

        if row["INCONSISTENCIA_FILIAL"]:

            erros.append(
                "Código da filial esperado: "
                f"{row['Código Filial']} | "
                "Código informado: "
                f"{row['CODIGOFILIALRESPOSTA']}"
            )

        return " ; ".join(erros)

    resultado["DESCRICAO_INCONSISTENCIA"] = pd.Series(
        [descricao_inconsistencia(row) for _, row in resultado.iterrows()],
        index=resultado.index,
        dtype="str",
    )

    resultado["TEM_INCONSISTENCIA"] = (
        resultado["INCONSISTENCIA_TIPO"] | resultado["INCONSISTENCIA_FILIAL"]
    )

    resultado["TEM_PROBLEMA"] = (resultado["STATUS_PRAZO"] == "PENDENTE") | resultado[
        "TEM_INCONSISTENCIA"
    ]

    print(f"    Equipamentos analisados: " f"{len(resultado)}")

    print("    Pendentes: " f"{(resultado['STATUS_PRAZO'] == 'PENDENTE').sum()}")

    print("    Inconsistências de tipo: " f"{resultado['INCONSISTENCIA_TIPO'].sum()}")

    print(
        "    Inconsistências de filial: " f"{resultado['INCONSISTENCIA_FILIAL'].sum()}"
    )

    return resultado


def gerar_resumo(
    resultado,
):

    print("\n[4/5] Gerando resumo...")

    resumo = (
        resultado.groupby("FILIAL")
        .agg(
            TOTAL=(
                "PLACA_NORMALIZADA",
                "count",
            ),
            OK=(
                "STATUS_PRAZO",
                lambda x: (x == "OK").sum(),
            ),
            PENDENTES=(
                "STATUS_PRAZO",
                lambda x: (x == "PENDENTE").sum(),
            ),
            INCONSISTENCIAS_TIPO=(
                "INCONSISTENCIA_TIPO",
                "sum",
            ),
            INCONSISTENCIAS_FILIAL=(
                "INCONSISTENCIA_FILIAL",
                "sum",
            ),
        )
        .reset_index()
    )

    resumo["PERCENTUAL_OK"] = (resumo["OK"] / resumo["TOTAL"] * 100).round(2)

    return resumo


def salvar_resultados(
    resultado,
    resumo,
):

    print("\n[5/5] Salvando resultados...")

    SAIDA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    colunas = [
        "FILIAL",
        "Código Filial",
        "Placa",
        "TIPOS_ACEITOS_TEXTO",
        "CHECKLIST_REALIZADO",
        "CODIGOFILIALRESPOSTA",
        "ULTIMO_RESPONSAVEL",
        "ULTIMA_RESPOSTA",
        "VENCIMENTO",
        "STATUS_PRAZO",
        "MOTIVO_PENDENCIA",
        "DIAS_SEM_CHECKLIST",
        "HORAS_RESTANTES",
        "INCONSISTENCIA_TIPO",
        "INCONSISTENCIA_FILIAL",
        "DESCRICAO_INCONSISTENCIA",
        "DATA_ANALISE",
    ]

    resultado_final = (
        resultado[colunas]
        .copy()
        .sort_values(
            [
                "FILIAL",
                "STATUS_PRAZO",
                "Placa",
            ]
        )
    )

    resultado_final = resultado_final.rename(
        columns={"TIPOS_ACEITOS_TEXTO": "TIPOS_ACEITOS"}
    )

    resultado_final.to_excel(
        SAIDA_DIR / "Resultado_Checklist.xlsx",
        index=False,
    )

    problemas = resultado[resultado["TEM_PROBLEMA"]][colunas].copy()

    problemas = problemas.rename(columns={"TIPOS_ACEITOS_TEXTO": "TIPOS_ACEITOS"})

    problemas.to_excel(
        SAIDA_DIR / "Problemas_Checklist.xlsx",
        index=False,
    )

    resumo.to_excel(
        SAIDA_DIR / "Resumo_Checklist.xlsx",
        index=False,
    )


def gerar_html_email(
    filial,
    dados,
):

    pendencias = dados[dados["STATUS_PRAZO"] == "PENDENTE"].copy()

    inconsistencias = dados[dados["TEM_INCONSISTENCIA"]].copy()

    partes = []

    if not pendencias.empty:

        linhas = []

        for _, row in pendencias.iterrows():

            linhas.append(f"""
                <tr>
                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["Placa"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["TIPOS_ACEITOS_TEXTO"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["CHECKLIST_REALIZADO"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_data(row["ULTIMA_RESPOSTA"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_data(row["VENCIMENTO"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["MOTIVO_PENDENCIA"])}
                    </td>
                </tr>
                """)

        partes.append(f"""
            <h3>Checklists pendentes</h3>

            <p>
                Os equipamentos abaixo não possuem
                checklist realizado dentro do prazo de
                {PRAZO_DIAS} dias.
            </p>

            <table style="
                border-collapse:collapse;
                width:100%;
                font-size:13px;
            ">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #ddd;padding:8px;">
                            Placa
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Tipos aceitos
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Último checklist
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Última realização
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Vencimento
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Motivo
                        </th>
                    </tr>
                </thead>

                <tbody>
                    {''.join(linhas)}
                </tbody>
            </table>
            """)

    if not inconsistencias.empty:

        linhas = []

        for _, row in inconsistencias.iterrows():

            linhas.append(f"""
                <tr>
                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["Placa"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["TIPOS_ACEITOS_TEXTO"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["CHECKLIST_REALIZADO"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["Código Filial"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["CODIGOFILIALRESPOSTA"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_data(row["ULTIMA_RESPOSTA"])}
                    </td>

                    <td style="border:1px solid #ddd;padding:8px;">
                        {formatar_valor(row["DESCRICAO_INCONSISTENCIA"])}
                    </td>
                </tr>
                """)

        partes.append(
            """
            <h3>Inconsistências identificadas</h3>

            <p>
                Os registros abaixo foram considerados
                como realizados, porém apresentam divergências
                entre o cadastro oficial do equipamento e
                as informações utilizadas no checklist.
            </p>

            <table style="
                border-collapse:collapse;
                width:100%;
                font-size:13px;
            ">
                <thead>
                    <tr style="background:#f2f2f2;">
                        <th style="border:1px solid #ddd;padding:8px;">
                            Placa
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Tipos aceitos
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Checklist realizado
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Cód. filial esperado
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Cód. filial informado
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Data
                        </th>
                        <th style="border:1px solid #ddd;padding:8px;">
                            Problema
                        </th>
                    </tr>
                </thead>

                <tbody>
            """
            + "".join(linhas)
            + """
                </tbody>
            </table>
            """
        )

    return f"""
    <html>
    <body style="
        font-family:Arial,sans-serif;
        color:#333;
    ">

        <div style="
            max-width:1100px;
            margin:auto;
        ">

            <h2>
                Checklist de Equipamentos -
                {html.escape(filial)}
            </h2>

            <p>Olá,</p>

            <p>
                Foram identificadas pendências e/ou
                inconsistências nos checklists de
                equipamentos da filial
                <strong>{html.escape(filial)}</strong>.
            </p>

            {''.join(partes)}

            <p style="margin-top:25px;">
                Favor verificar os registros relacionados.
            </p>

            <p>
                Atenciosamente,<br>
                <strong>
                    Automação - Patrus Transportes
                </strong>
            </p>

        </div>

    </body>
    </html>
    """


def autenticar():

    credenciais = None

    if TOKEN_FILE.exists():

        credenciais = Credentials.from_authorized_user_file(
            TOKEN_FILE,
            SCOPES,
        )

    if not credenciais or not credenciais.valid:

        if credenciais and credenciais.expired and credenciais.refresh_token:

            credenciais.refresh(Request())

        else:

            fluxo = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE,
                SCOPES,
            )

            credenciais = fluxo.run_local_server(port=0)

        TOKEN_FILE.write_text(
            credenciais.to_json(),
            encoding="utf-8",
        )

    return credenciais


def enviar_email(
    destinatarios,
    assunto,
    corpo_html,
):

    if isinstance(
        destinatarios,
        str,
    ):
        destinatarios = [destinatarios]

    destinatarios = [
        email.strip() for email in destinatarios if email and email.strip()
    ]

    if not destinatarios:

        raise ValueError("Nenhum destinatário informado.")

    credenciais = autenticar()

    servico = build(
        "gmail",
        "v1",
        credentials=credenciais,
    )

    mensagem = MIMEText(
        corpo_html,
        "html",
        "utf-8",
    )

    mensagem["to"] = ", ".join(destinatarios)

    mensagem["subject"] = assunto

    raw = base64.urlsafe_b64encode(mensagem.as_bytes()).decode()

    resultado = (
        servico.users()
        .messages()
        .send(
            userId="me",
            body={"raw": raw},
        )
        .execute()
    )

    print("E-mail enviado para: " + ", ".join(destinatarios))

    return resultado


def enviar_emails(
    resultado,
):

    filiais_base = sorted(resultado["FILIAL"].dropna().astype(str).str.strip().unique())

    for filial in filiais_base:

        destinatarios = DESTINATARIOS.get(filial)

        if not destinatarios:

            print(f"{filial}: sem destinatário " "configurado no .env.")

            continue

        dados = resultado[
            (resultado["FILIAL"].astype(str).str.strip() == filial)
            & (resultado["TEM_PROBLEMA"])
        ].copy()

        if dados.empty:

            print(f"{filial}: sem pendências " "ou inconsistências.")

            continue

        qtd_pendentes = dados["STATUS_PRAZO"].eq("PENDENTE").sum()

        qtd_tipo = dados["INCONSISTENCIA_TIPO"].sum()

        qtd_filial = dados["INCONSISTENCIA_FILIAL"].sum()

        qtd_inconsistencias = dados["TEM_INCONSISTENCIA"].sum()

        assunto = (
            "[CHECKLIST] "
            f"{filial} - "
            f"{qtd_pendentes} pendência(s) / "
            f"{qtd_inconsistencias} inconsistência(s)"
        )

        print(f"\n{filial}:")

        print(f"    Pendências: {qtd_pendentes}")

        print(f"    Inconsistências de tipo: {qtd_tipo}")

        print(f"    Inconsistências de filial: {qtd_filial}")

        corpo = gerar_html_email(
            filial,
            dados,
        )

        enviar_email(
            destinatarios,
            assunto,
            corpo,
        )


def main():

    print("=" * 60)

    print("AUTOMAÇÃO DE CHECKLISTS DE EQUIPAMENTOS")

    print("=" * 60)

    try:

        print("\nValidando bases...")

        validar_bases()

        equipamentos = carregar_equipamentos()

        historico = carregar_historico()

        resultado = processar(
            equipamentos,
            historico,
        )

        resumo = gerar_resumo(resultado)

        salvar_resultados(
            resultado,
            resumo,
        )

        enviar_emails(resultado)

        print("\nProcessamento concluído.")

        print(f"Resultado: " f"{SAIDA_DIR / 'Resultado_Checklist.xlsx'}")

        print(f"Problemas: " f"{SAIDA_DIR / 'Problemas_Checklist.xlsx'}")

        print(f"Resumo: " f"{SAIDA_DIR / 'Resumo_Checklist.xlsx'}")

    except Exception as erro:

        print("\nERRO DURANTE O PROCESSAMENTO")

        print(f"{type(erro).__name__}: {erro}")

        raise


if __name__ == "__main__":
    main()
