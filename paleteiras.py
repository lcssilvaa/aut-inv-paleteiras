from pathlib import Path
from datetime import datetime, timedelta
import html
import base64
import os

import pandas as pd

from dotenv import load_dotenv
from automacao.relatorios_pdf import gerar_pdf_paleteiras, salvar_pdf_paleteiras
from automacao.mensagem_email import criar_mensagem

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ============================================================
# CONFIGURAÇÕES
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

load_dotenv(BASE_DIR / ".env")

SAIDA_DIR = BASE_DIR / "saida"

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


# ============================================================
# GMAIL
# ============================================================


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


# ============================================================
# BASES
# ============================================================

SHAREPOINT_DIR = Path(os.getenv("SHAREPOINT_DIR", ""))

ARQUIVO_BASE_1 = SHAREPOINT_DIR / os.getenv(
    "ARQUIVO_BASE_1",
    "Gestão Manutenção de Frotas - Controle de Paleteiras.xlsx",
)

ARQUIVO_BASE_2 = SHAREPOINT_DIR / os.getenv(
    "ARQUIVO_BASE_2",
    "Planilha de Paleteiras Disktrans.xlsx",
)

ABA_BASE_2 = os.getenv(
    "ABA_BASE_2",
    "Base",
)

PRAZO_DIAS = int(
    os.getenv(
        "PRAZO_DIAS",
        "7",
    )
)

# ============================================================
# DESTINATÁRIOS
# ============================================================


def carregar_destinatarios():

    destinatarios = {}

    filiais = [
        "AJU",
        "BAU",
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


# ============================================================
# VALIDAÇÃO
# ============================================================


def validar_bases():

    arquivos = {
        "Base 1": ARQUIVO_BASE_1,
        "Base 2": ARQUIVO_BASE_2,
    }

    for nome, arquivo in arquivos.items():

        if not arquivo.exists():
            raise FileNotFoundError(f"{nome} não encontrado:\n{arquivo}")

        tamanho_mb = arquivo.stat().st_size / (1024 * 1024)

        if tamanho_mb == 0:
            raise ValueError(f"{nome} está vazio.")

        ultima_modificacao = datetime.fromtimestamp(arquivo.stat().st_mtime)

        print(f"{nome}: {ultima_modificacao.strftime('%d/%m/%Y %H:%M:%S')}")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================


def normalizar_qr(valor):

    if pd.isna(valor):
        return None

    valor = str(valor).strip()

    if "-" in valor:
        valor = valor.split("-")[0]

    if valor.endswith(".0"):
        valor = valor[:-2]

    valor = "".join(caractere for caractere in valor if caractere.isdigit())

    if not valor:
        return None

    return valor.zfill(5)


def converter_custo(valor):

    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    valor = str(valor).strip()
    valor = valor.replace("R$", "").replace(" ", "")

    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except ValueError:
        return 0.0


def formatar_data(valor):

    if pd.isna(valor):
        return "-"

    return pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")


def formatar_moeda(valor):

    valor = float(valor or 0)

    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ============================================================
# LEITURA DAS BASES
# ============================================================


def carregar_base_1():

    df = pd.read_excel(ARQUIVO_BASE_1, header=6)

    df.columns = df.columns.astype(str).str.strip()

    colunas_necessarias = [
        "CODIGO",
        "FILIAL",
        "NOME",
        "DATAHORA",
    ]

    faltantes = [coluna for coluna in colunas_necessarias if coluna not in df.columns]

    if faltantes:
        raise ValueError(f"Colunas não encontradas na Base 1: {faltantes}")

    df = df.dropna(how="all")

    df["CODIGO"] = df["CODIGO"].astype(str).str.strip()

    df["FILIAL"] = df["FILIAL"].astype(str).str.strip()

    df["NOME"] = df["NOME"].astype(str).str.strip()

    df["QR_NORMALIZADO"] = df["CODIGO"].apply(normalizar_qr)

    def converter_data(valor):

        if pd.isna(valor):
            return pd.NaT

        if isinstance(valor, (pd.Timestamp, datetime)):
            return pd.Timestamp(valor)

        valor = str(valor).strip()

        if not valor:
            return pd.NaT

        if len(valor) >= 10 and valor[4] == "-":
            return pd.to_datetime(valor, errors="coerce")

        return pd.to_datetime(valor, dayfirst=True, errors="coerce")

    df["DATAHORA"] = df["DATAHORA"].apply(converter_data)

    df = df[df["QR_NORMALIZADO"].notna()].copy()

    df = df[df["DATAHORA"].notna()].copy()

    return df


def carregar_base_2():

    df = pd.read_excel(
        ARQUIVO_BASE_2,
        sheet_name=ABA_BASE_2,
        dtype=str,
        engine="openpyxl",
    )

    df.columns = df.columns.astype(str).str.strip()

    colunas_necessarias = [
        "FILIAL",
        "QR Code",
        "NR_DISK",
        "MODELO",
        "CUSTO modelo",
    ]

    faltantes = [coluna for coluna in colunas_necessarias if coluna not in df.columns]

    if faltantes:
        raise ValueError(f"Colunas não encontradas na Base 2: {faltantes}")

    df = df.dropna(how="all")

    df["QR_NORMALIZADO"] = df["QR Code"].apply(normalizar_qr)

    df["CUSTO_NUMERICO"] = df["CUSTO modelo"].apply(converter_custo)

    df = df[df["QR_NORMALIZADO"].notna()].copy()

    return df


# ============================================================
# PROCESSAMENTO
# ============================================================


def processar(base_1, base_2):

    leituras = base_1[
        [
            "QR_NORMALIZADO",
            "CODIGO",
            "NOME",
            "DATAHORA",
        ]
    ].copy()

    leituras = leituras.rename(
        columns={
            "CODIGO": "CODIGO_LEITURA",
            "NOME": "ULTIMO_INVENTARIADOR",
            "DATAHORA": "ULTIMA_LEITURA",
        }
    )

    leituras["ULTIMA_LEITURA"] = pd.to_datetime(
        leituras["ULTIMA_LEITURA"],
        errors="coerce",
    )

    resultado = base_2.merge(
        leituras,
        on="QR_NORMALIZADO",
        how="left",
    )

    resultado["VENCIMENTO"] = pd.NaT

    mascara_lido = resultado["ULTIMA_LEITURA"].notna()

    resultado.loc[mascara_lido, "VENCIMENTO"] = resultado.loc[
        mascara_lido, "ULTIMA_LEITURA"
    ] + timedelta(days=PRAZO_DIAS)

    agora = pd.Timestamp.now()

    resultado["DATA_ANALISE"] = agora

    def determinar_status(row):

        if pd.isna(row["ULTIMA_LEITURA"]):
            return "NUNCA LIDO"

        if agora < row["VENCIMENTO"]:
            return "LIDO"

        return "PENDENTE"

    resultado["STATUS"] = resultado.apply(
        determinar_status,
        axis=1,
    )

    diferenca = agora - resultado["ULTIMA_LEITURA"]

    resultado["HORAS_SEM_LEITURA"] = (diferenca.dt.total_seconds() / 3600).round(2)

    resultado["DIAS_SEM_LEITURA"] = (resultado["HORAS_SEM_LEITURA"] / 24).round(2)

    resultado["HORAS_RESTANTES"] = (
        (resultado["VENCIMENTO"] - agora).dt.total_seconds() / 3600
    ).round(2)

    return resultado


def gerar_resumo(resultado):

    resumo = (
        resultado.groupby("FILIAL")
        .agg(
            TOTAL=("QR_NORMALIZADO", "count"),
            LIDOS=("STATUS", lambda x: (x == "LIDO").sum()),
            PENDENTES=("STATUS", lambda x: (x == "PENDENTE").sum()),
            NUNCA_LIDOS=("STATUS", lambda x: (x == "NUNCA LIDO").sum()),
            VALOR_PENDENTE=("CUSTO_NUMERICO", lambda x: 0),
        )
        .reset_index()
    )

    valor_pendente = (
        resultado[resultado["STATUS"].isin(["PENDENTE", "NUNCA LIDO"])]
        .groupby("FILIAL")["CUSTO_NUMERICO"]
        .sum()
    )

    resumo["VALOR_PENDENTE"] = resumo["FILIAL"].map(valor_pendente).fillna(0)

    resumo["LIDOS"] = resumo["LIDOS"].astype(int)
    resumo["PENDENTES"] = resumo["PENDENTES"].astype(int)
    resumo["NUNCA_LIDOS"] = resumo["NUNCA_LIDOS"].astype(int)

    resumo["NAO_LIDOS"] = resumo["PENDENTES"] + resumo["NUNCA_LIDOS"]

    resumo["PERCENTUAL_LIDOS"] = (resumo["LIDOS"] / resumo["TOTAL"] * 100).round(2)

    resumo["PERCENTUAL_NAO_LIDOS"] = (
        resumo["NAO_LIDOS"] / resumo["TOTAL"] * 100
    ).round(2)

    return resumo


# ============================================================
# E-MAIL
# ============================================================


def gerar_html_email(filial, pendencias):

    total = len(pendencias)

    valor_total = pendencias["CUSTO_NUMERICO"].sum()

    pendentes = (pendencias["STATUS"] == "PENDENTE").sum()

    nunca_lidos = (pendencias["STATUS"] == "NUNCA LIDO").sum()

    linhas = []

    for _, row in pendencias.iterrows():

        qr = html.escape(str(row["QR Code"]))

        nr_disk = html.escape(str(row["NR_DISK"]))

        modelo = html.escape(str(row["MODELO"]))

        status = html.escape(str(row["STATUS"]))

        data_leitura = formatar_data(row["ULTIMA_LEITURA"])

        dias_sem_leitura = "-"

        if not pd.isna(row["DIAS_SEM_LEITURA"]):
            dias_sem_leitura = f"{row['DIAS_SEM_LEITURA']:.2f}".replace(".", ",")

        custo = formatar_moeda(row["CUSTO_NUMERICO"])

        linhas.append(f"""
            <tr>
                <td>{qr}</td>
                <td>{nr_disk}</td>
                <td>{modelo}</td>
                <td>{data_leitura}</td>
                <td>{dias_sem_leitura}</td>
                <td>{custo}</td>
                <td>{status}</td>
            </tr>
            """)

    tabela = "".join(linhas)

    valor_formatado = formatar_moeda(valor_total)

    return f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">

        <div style="max-width: 1100px; margin: auto;">

            <h2>
                Inventário de Paleteiras - {html.escape(filial)}
            </h2>

            <p>
                Olá,
            </p>

            <p>
                Identificamos
                <strong>{total} equipamentos</strong>
                pendentes de inventário na filial
                <strong>{html.escape(filial)}</strong>.
            </p>

            <div style="
                display: flex;
                gap: 15px;
                margin: 20px 0;
            ">

                <div style="
                    padding: 15px;
                    border: 1px solid #ddd;
                    border-radius: 6px;
                ">
                    <strong>{total}</strong><br>
                    Pendências
                </div>

                <div style="
                    padding: 15px;
                    border: 1px solid #ddd;
                    border-radius: 6px;
                ">
                    <strong>{pendentes}</strong><br>
                    Pendentes
                </div>

                <div style="
                    padding: 15px;
                    border: 1px solid #ddd;
                    border-radius: 6px;
                ">
                    <strong>{nunca_lidos}</strong><br>
                    Nunca lidos
                </div>

                <div style="
                    padding: 15px;
                    border: 1px solid #ddd;
                    border-radius: 6px;
                ">
                    <strong>{valor_formatado}</strong><br>
                    Valor total
                </div>

            </div>

            <p>
                Favor realizar o inventário dos equipamentos
                relacionados abaixo para regularização do controle.
            </p>

            <table style="
                border-collapse: collapse;
                width: 100%;
                font-size: 13px;
            ">

                <thead>

                    <tr style="background-color: #f2f2f2;">

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            QR Code
                        </th>

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            NR Disk
                        </th>

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            Modelo
                        </th>

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            Última leitura
                        </th>

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            Dias sem leitura
                        </th>

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            Custo
                        </th>

                        <th style="border: 1px solid #ddd; padding: 8px;">
                            Status
                        </th>

                    </tr>

                </thead>

                <tbody>
                    {tabela}
                </tbody>

            </table>

            <p style="margin-top: 25px;">
                Após a realização das leituras, os equipamentos
                serão atualizados no próximo processamento do relatório.
            </p>

            <p>
                Atenciosamente,<br>
                <strong>Automação - Patrus Transportes</strong>
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


def enviar_email(destinatarios, assunto, corpo_html, anexo_pdf=None):

    if isinstance(destinatarios, str):
        destinatarios = [destinatarios]

    destinatarios = [
        email.strip() for email in destinatarios if email and email.strip()
    ]

    if not destinatarios:
        raise ValueError("Nenhum destinatário informado para o e-mail.")

    credenciais = autenticar()

    servico = build(
        "gmail",
        "v1",
        credentials=credenciais,
    )

    mensagem = criar_mensagem(destinatarios, assunto, corpo_html, anexo_pdf)

    mensagem_codificada = base64.urlsafe_b64encode(mensagem.as_bytes()).decode()

    corpo = {"raw": mensagem_codificada}

    resultado = (
        servico.users()
        .messages()
        .send(
            userId="me",
            body=corpo,
        )
        .execute()
    )

    print(f"E-mail enviado para: {', '.join(destinatarios)}")

    print(f"ID da mensagem: {resultado['id']}")

    return resultado


def enviar_emails(resultado):

    status_pendencia = [
        "PENDENTE",
        "NUNCA LIDO",
    ]

    for filial, destinatarios in DESTINATARIOS.items():

        print(f"\nProcessando e-mail da filial {filial}...")

        pendencias = resultado[
            (resultado["FILIAL"].astype(str).str.strip() == filial)
            & (resultado["STATUS"].isin(status_pendencia))
        ].copy()

        if pendencias.empty:

            print(
                f"Nenhuma pendência encontrada para {filial}. "
                f"E-mail não será enviado."
            )

            continue

        corpo_html = gerar_html_email(
            filial,
            pendencias,
        )

        assunto = f"[PENDÊNCIAS] " f"Inventário de Paleteiras - {filial}"

        enviar_email(
            destinatarios=destinatarios,
            assunto=assunto,
            corpo_html=corpo_html,
            anexo_pdf=(
                f"Resumo_Paleteiras_{filial}.pdf",
                gerar_pdf_paleteiras(resultado, PRAZO_DIAS, filial=filial),
            ),
        )


# ============================================================
# SALVAR RESULTADOS
# ============================================================


def salvar_resultados(resultado, resumo):

    SAIDA_DIR.mkdir(parents=True, exist_ok=True)

    resultado_final = resultado[
        [
            "FILIAL",
            "QR Code",
            "NR_DISK",
            "MODELO",
            "CUSTO_NUMERICO",
            "CODIGO_LEITURA",
            "ULTIMO_INVENTARIADOR",
            "ULTIMA_LEITURA",
            "VENCIMENTO",
            "STATUS",
            "DIAS_SEM_LEITURA",
            "HORAS_RESTANTES",
            "DATA_ANALISE",
        ]
    ].copy()

    resultado_final = resultado_final.rename(
        columns={
            "CUSTO_NUMERICO": "CUSTO",
            "ULTIMO_INVENTARIADOR": "NOME",
            "ULTIMA_LEITURA": "DATAHORA",
        }
    )

    mascara_nunca_lido = resultado_final["STATUS"] == "NUNCA LIDO"

    resultado_final.loc[
        mascara_nunca_lido,
        [
            "CODIGO_LEITURA",
            "NOME",
            "DATAHORA",
            "VENCIMENTO",
            "DIAS_SEM_LEITURA",
            "HORAS_RESTANTES",
        ],
    ] = None

    ordem_status = {
        "PENDENTE": 1,
        "NUNCA LIDO": 2,
        "LIDO": 3,
    }

    resultado_final["_ORDEM"] = resultado_final["STATUS"].map(ordem_status)

    resultado_final = resultado_final.sort_values(
        [
            "FILIAL",
            "_ORDEM",
            "DIAS_SEM_LEITURA",
        ],
        ascending=[
            True,
            True,
            False,
        ],
        na_position="last",
    ).drop(columns=["_ORDEM"])

    resultado_final.to_excel(
        SAIDA_DIR / "Resultado_Inventario.xlsx",
        index=False,
    )

    resumo.to_excel(
        SAIDA_DIR / "Resumo_Filiais.xlsx",
        index=False,
    )

    pdf = salvar_pdf_paleteiras(resultado, SAIDA_DIR, PRAZO_DIAS)
    print(f"Resumo visual em PDF: {pdf}")


# ============================================================
# EXECUÇÃO
# ============================================================


def main():

    print("=" * 60)
    print("AUTOMAÇÃO DE INVENTÁRIO DE PALETEIRAS")
    print("=" * 60)

    try:

        print("\nValidando bases...")
        validar_bases()

        print("Processando dados...")
        base_1 = carregar_base_1()
        base_2 = carregar_base_2()

        resultado = processar(
            base_1,
            base_2,
        )

        resumo = gerar_resumo(resultado)

        salvar_resultados(
            resultado,
            resumo,
        )

        enviar_emails(resultado)

        total = len(resultado)
        lidos = (resultado["STATUS"] == "LIDO").sum()
        pendentes = (resultado["STATUS"] == "PENDENTE").sum()
        nunca_lidos = (resultado["STATUS"] == "NUNCA LIDO").sum()

        print("\nProcessamento concluído.")
        print(f"Total: {total}")
        print(f"Lidos: {lidos}")
        print(f"Pendentes: {pendentes}")
        print(f"Nunca lidos: {nunca_lidos}")

    except Exception as erro:

        print("\nERRO DURANTE O PROCESSAMENTO")
        print(f"{type(erro).__name__}: {erro}")

        raise


if __name__ == "__main__":
    main()
