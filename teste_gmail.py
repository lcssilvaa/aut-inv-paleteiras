import base64
from email.mime.text import MIMEText
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

DESTINATARIO = "seu-email@gmail.com"


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


def enviar_email():
    credenciais = autenticar()

    servico = build(
        "gmail",
        "v1",
        credentials=credenciais,
    )

    mensagem = MIMEText(
        """
        <h2>Teste - Automação de Inventário de Paleteiras</h2>

        <p>Olá, Lucas.</p>

        <p>
        Este é um teste de envio automático utilizando a
        <strong>Gmail API</strong>.
        </p>

        <p>
        Se você recebeu este e-mail, a integração com o Gmail
        está funcionando corretamente.
        </p>

        <p>
        <strong>Automação de Inventário de Paleteiras</strong>
        </p>
        """,
        "html",
        "utf-8",
    )

    mensagem["to"] = DESTINATARIO
    mensagem["subject"] = "TESTE - Automação de Inventário de Paleteiras"

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

    print("E-mail enviado com sucesso!")
    print("ID da mensagem:", resultado["id"])


if __name__ == "__main__":
    enviar_email()
