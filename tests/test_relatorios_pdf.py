import base64
from email import policy
from email.parser import BytesParser
from pathlib import Path
import re
import sys
import unittest
from unittest.mock import patch
import zlib

import pandas as pd

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import checklist
import paleteiras
from automacao.mensagem_email import criar_mensagem
from automacao.relatorios_pdf import (
    dados_relatorio,
    gerar_pdf_checklist,
    gerar_pdf_paleteiras,
    resumir,
)


def texto_pdf(pdf):
    """Lê os streams produzidos pelo ReportLab para conferir a presença dos dados."""
    streams = []
    for encontrado in re.finditer(rb"stream\r?\n(.*?)endstream", pdf, re.S):
        cabecalho = pdf[pdf.rfind(b"obj", 0, encontrado.start()) : encontrado.start()]
        tamanho = int(re.search(rb"/Length\s+(\d+)", cabecalho).group(1))
        stream = encontrado.group(1)[:tamanho]
        if b"/ASCII85Decode" in cabecalho:
            stream = base64.a85decode(stream.strip(), adobe=True)
        if b"/FlateDecode" in cabecalho:
            stream = zlib.decompress(stream)
        streams.append(stream)
    return b"\n".join(streams)


def base_checklist():
    linhas = []
    for placa, filial, dias, erro in [
        ("REGULAR_FROTAS", "FROTAS", 1, False),
        ("ERRADO_FROTAS", "FROTAS", 1, True),
        ("VENCIDO_FROTAS", "FROTAS", 10, False),
        ("SEM_FROTAS", "FROTAS", None, False),
        ("OUTRA_FILIAL_CON", "CON", 1, False),
    ]:
        ultima = (
            pd.Timestamp("2026-09-09") - pd.Timedelta(days=dias)
            if dias is not None
            else pd.NaT
        )
        linhas.append(
            {
                "Placa": placa,
                "PLACA_NORMALIZADA": placa,
                "FILIAL": filial,
                "Código Filial": 2,
                "CODIGOFILIALRESPOSTA": 2,
                "TIPOS_ACEITOS_TEXTO": "Bateria",
                "CHECKLIST_REALIZADO": "Automatic" if erro else "Bateria",
                "ULTIMO_RESPONSAVEL": (
                    "Ana <Silva> & Souza" if dias is not None else None
                ),
                "ULTIMA_RESPOSTA": ultima,
                "VENCIMENTO": ultima + pd.Timedelta(days=7),
                "DIAS_SEM_CHECKLIST": dias,
                "HORAS_RESTANTES": None,
                "STATUS_PRAZO": "PENDENTE" if dias is None or dias >= 7 else "OK",
                "STATUS_CHECKLIST": "PENDENTE" if dias is None or dias >= 7 else "OK",
                "PENDENCIA_BATERIA": False,
                "INCONSISTENCIA_TIPO": erro,
                "INCONSISTENCIA_FILIAL": False,
                "TEM_INCONSISTENCIA": erro,
                "TEM_PROBLEMA": erro or dias is None or dias >= 7,
                "FILIAL_ESPERADA": "CON",
                "FILIAL_RESPOSTA": "CON",
                "DESCRICAO_INCONSISTENCIA": "Tipo incorreto" if erro else "",
                "MOTIVO_PENDENCIA": "Nunca realizado" if dias is None else "",
                "DATA_ANALISE": pd.Timestamp("2026-09-09 08:00"),
            }
        )
    return pd.DataFrame(linhas)


def base_paleteiras():
    linhas = []
    for indice, status, dias, custo in [
        (1, "LIDO", 1, 2000),
        (2, "PENDENTE", 10, 1500),
        (3, "NUNCA LIDO", None, 2500),
        (4, "NUNCA LIDO", None, None),
    ]:
        ultima = (
            pd.Timestamp("2026-09-09") - pd.Timedelta(days=dias)
            if dias is not None
            else pd.NaT
        )
        linhas.append(
            {
                "QR Code": f"QR{indice:03}",
                "QR_NORMALIZADO": f"QR{indice:03}",
                "NR_DISK": f"DISK{indice:03}",
                "FILIAL": "CON",
                "MODELO": "Manual",
                "CUSTO_NUMERICO": custo,
                "CODIGO_LEITURA": f"QR{indice:03}",
                "ULTIMO_INVENTARIADOR": "Ana" if dias is not None else None,
                "ULTIMA_LEITURA": ultima,
                "VENCIMENTO": ultima + pd.Timedelta(days=7),
                "STATUS": status,
                "DIAS_SEM_LEITURA": dias,
                "HORAS_RESTANTES": None,
                "DATA_ANALISE": pd.Timestamp("2026-09-09 08:00"),
            }
        )
    return pd.DataFrame(linhas)


class RelatoriosPDFTest(unittest.TestCase):
    def test_contagens_e_filtragem_checklist(self):
        dados = dados_relatorio(base_checklist(), 7, "checklist", "FROTAS")
        resumo = resumir(dados["registros"])
        self.assertEqual(resumo["total"], 4)
        self.assertEqual(resumo["regular"], 1)
        self.assertEqual(resumo["no_prazo"], 2)
        self.assertEqual(resumo["pendentes"], 2)
        self.assertEqual(resumo["inconsistencias"], 1)
        self.assertTrue(all(r["filial"] == "FROTAS" for r in dados["registros"]))

    def test_pdf_real_tem_graficos_resumo_detalhes_e_paginacao(self):
        resultado = pd.concat([base_checklist()] * 30, ignore_index=True)
        resultado["Placa"] = [f"PLACA_{i:04}" for i in range(len(resultado))]
        pdf = gerar_pdf_checklist(resultado, 7)
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertTrue(pdf.rstrip().endswith(b"%%EOF"))
        self.assertGreater(len(re.findall(rb"/Type\s*/Page\b", pdf)), 3)
        conteudo = texto_pdf(pdf)
        self.assertIn(b"PLACA_0000", conteudo)
        self.assertIn(b"PLACA_0149", conteudo)

    def test_paleteiras_soma_apenas_custos_pendentes(self):
        dados = dados_relatorio(base_paleteiras(), 7, "paleteiras")
        resumo = resumir(dados["registros"])
        self.assertEqual(resumo["valor"], 4000)
        self.assertEqual(resumo["pendentes"], 3)
        self.assertEqual(resumo["regular"], 1)
        self.assertEqual(resumo["sem_custo"], 1)
        pdf = gerar_pdf_paleteiras(base_paleteiras(), 7)
        self.assertIn(b"QR004", texto_pdf(pdf))

    def test_relatorios_vazios(self):
        for funcao, resultado in [
            (gerar_pdf_checklist, base_checklist()),
            (gerar_pdf_paleteiras, base_paleteiras()),
        ]:
            pdf = funcao(resultado.iloc[:0], 10)
            self.assertTrue(pdf.startswith(b"%PDF-"))
            self.assertIn(b"Nenhum equipamento neste recorte", texto_pdf(pdf))

    def test_mensagem_preserva_html_e_anexa_pdf_integro(self):
        pdf = gerar_pdf_checklist(base_checklist(), 7, "FROTAS")
        mensagem = criar_mensagem(
            ["teste@example.com"], "Teste", "<p>Olá</p>", ("Resumo_FROTAS.pdf", pdf)
        )
        recebida = BytesParser(policy=policy.default).parsebytes(mensagem.as_bytes())
        self.assertEqual(recebida.get_content_type(), "multipart/mixed")
        self.assertIn("Olá", recebida.get_body(preferencelist=("html",)).get_content())
        anexos = list(recebida.iter_attachments())
        self.assertEqual(len(anexos), 1)
        self.assertEqual(anexos[0].get_filename(), "Resumo_FROTAS.pdf")
        self.assertEqual(anexos[0].get_content_type(), "application/pdf")
        self.assertEqual(anexos[0].get_payload(decode=True), pdf)

    def test_email_checklist_anexa_todos_da_filial_sem_misturar_outra(self):
        with patch.object(
            checklist,
            "DESTINATARIOS",
            {"FROTAS": ["teste@example.com"], "CON": ["con@example.com"]},
        ), patch.object(checklist, "enviar_email") as enviar, patch("builtins.print"):
            checklist.enviar_emails(base_checklist())
        enviar.assert_called_once()
        nome, pdf = enviar.call_args.kwargs["anexo_pdf"]
        self.assertEqual(nome, "Resumo_Checklist_FROTAS.pdf")
        conteudo = texto_pdf(pdf)
        self.assertIn(b"REGULAR_FROTAS", conteudo)
        self.assertNotIn(b"OUTRA_FILIAL_CON", conteudo)

    def test_email_paleteiras_anexa_pdf(self):
        with patch.object(
            paleteiras, "DESTINATARIOS", {"CON": ["teste@example.com"]}
        ), patch.object(paleteiras, "enviar_email") as enviar, patch("builtins.print"):
            paleteiras.enviar_emails(base_paleteiras())
        enviar.assert_called_once()
        nome, pdf = enviar.call_args.kwargs["anexo_pdf"]
        self.assertEqual(nome, "Resumo_Paleteiras_CON.pdf")
        self.assertIn(b"QR001", texto_pdf(pdf))

    def test_salvar_preserva_excel_e_gera_pdf(self):
        for modulo, base, funcao, quantidade in [
            (checklist, base_checklist(), "salvar_pdf_checklist", 3),
            (paleteiras, base_paleteiras(), "salvar_pdf_paleteiras", 2),
        ]:
            with patch.object(Path, "mkdir"), patch.object(
                pd.DataFrame, "to_excel"
            ) as excel, patch.object(modulo, funcao) as pdf, patch("builtins.print"):
                modulo.salvar_resultados(base, modulo.gerar_resumo(base))
            self.assertEqual(excel.call_count, quantidade)
            pdf.assert_called_once_with(base, modulo.SAIDA_DIR, modulo.PRAZO_DIAS)


if __name__ == "__main__":
    unittest.main()
