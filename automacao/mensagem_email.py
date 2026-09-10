"""Monta mensagens HTML com um anexo PDF opcional, sem realizar o envio."""

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def criar_mensagem(destinatarios, assunto, corpo_html, anexo_pdf=None):
    corpo = MIMEText(corpo_html, "html", "utf-8")
    if anexo_pdf is None:
        mensagem = corpo
    else:
        nome, conteudo = anexo_pdf
        mensagem = MIMEMultipart("mixed")
        mensagem.attach(corpo)
        anexo = MIMEApplication(conteudo, _subtype="pdf")
        anexo.add_header("Content-Disposition", "attachment", filename=nome)
        mensagem.attach(anexo)
    mensagem["to"] = ", ".join(destinatarios)
    mensagem["subject"] = assunto
    return mensagem
