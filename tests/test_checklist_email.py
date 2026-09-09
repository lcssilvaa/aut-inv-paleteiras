import contextlib
import io
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pandas as pd

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import checklist


class EmailChecklistTest(unittest.TestCase):
    def setUp(self):
        self.enterContext(contextlib.redirect_stdout(io.StringIO()))
        self.enterContext(patch.object(checklist, "PRAZO_DIAS", 7))

    def processar(self, codigo_resposta="2", ativo="Sim", com_resposta=True):
        base = pd.DataFrame(
            [
                ["TESTE_BACKUP", 1, "BACKUP", "Bateria", ativo],
                ["CADASTRO_CON", 2, "CON", "Bateria", "Não"],
            ],
            columns=[
                "Placa",
                "Código Filial",
                "Filial",
                "Tipo de Equipamento",
                "Ativo",
            ],
        )
        with patch.object(pd, "read_excel", return_value=base):
            equipamentos = checklist.carregar_equipamentos()
        historico = pd.DataFrame(
            [
                {
                    "PLACA_NORMALIZADA": "TESTE_BACKUP" if com_resposta else "OUTRA",
                    "CODIGOFILIALRESPOSTA": codigo_resposta,
                    "CODIGO_FILIAL_RESPOSTA_NORMALIZADO": codigo_resposta,
                    "NOMECHECKLIST": "Automatic",
                    "TIPO_REALIZADO_NORMALIZADO": "AUTOMATIC",
                    "DATARESPOSTA_CONVERTIDA": pd.Timestamp.now()
                    - pd.Timedelta(days=dias),
                    "RESPONDIDOPOR": pessoa,
                    "RESPOSTALOG": "",
                }
                for dias, pessoa in [(10, "Ana <Silva> & Souza"), (20, "Antigo")]
            ]
        )
        return checklist.processar(equipamentos, historico)

    def test_siglas_e_ultimo_responsavel_nas_duas_tabelas(self):
        resultado = self.processar()
        row = resultado.iloc[0]
        self.assertEqual(row["FILIAL"], "BACKUP")
        self.assertEqual(row["FILIAL_ESPERADA"], "BHZ")
        # CON continua identificável mesmo com seu equipamento desabilitado.
        self.assertEqual(row["FILIAL_RESPOSTA"], "CON")
        self.assertEqual(len(resultado), 1)
        self.assertTrue(row["INCONSISTENCIA_FILIAL"])
        corpo = checklist.gerar_html_email("BACKUP", resultado)
        self.assertIn("Filial esperada: BHZ | Filial informada: CON", corpo)
        self.assertNotIn("Código informado", corpo)
        self.assertNotIn("Cód. filial", corpo)
        self.assertEqual(corpo.count("Realizado por"), 2)
        self.assertEqual(corpo.count("Ana &lt;Silva&gt; &amp; Souza"), 2)
        self.assertNotIn("Antigo", corpo)

    def test_backup_resposta_bhz_e_envio_separado(self):
        resultado = self.processar(codigo_resposta="1")
        row = resultado.iloc[0]
        self.assertEqual(row["FILIAL_RESPOSTA"], "BHZ")
        self.assertFalse(row["INCONSISTENCIA_FILIAL"])
        with (
            patch.object(
                checklist,
                "DESTINATARIOS",
                {"BACKUP": ["backup@example.com"], "BHZ": ["bhz@example.com"]},
            ),
            patch.object(checklist, "enviar_email") as enviar,
        ):
            checklist.enviar_emails(resultado)
        enviar.assert_called_once()
        self.assertEqual(enviar.call_args.args[0], ["backup@example.com"])
        self.assertIn("[CHECKLIST] BACKUP", enviar.call_args.args[1])

    def test_filial_desconhecida_e_responsavel_ausente(self):
        resultado = self.processar(codigo_resposta="999999")
        resultado["ULTIMO_RESPONSAVEL"] = None
        corpo = checklist.gerar_html_email("BACKUP", resultado)
        self.assertIn("Não identificada na base", corpo)
        self.assertNotIn("999999", corpo)
        self.assertNotIn("None", corpo)

    def test_sem_checklist_e_todos_desabilitados(self):
        resultado = self.processar(com_resposta=False)
        self.assertEqual(resultado.iloc[0]["FILIAL_RESPOSTA"], "-")
        corpo = checklist.gerar_html_email("BACKUP", resultado)
        self.assertIn("Nunca realizado", corpo)
        self.assertNotIn("nan", corpo)
        vazio = self.processar(ativo="Não")
        self.assertTrue(vazio.empty)
        self.assertTrue(checklist.gerar_resumo(vazio).empty)


if __name__ == "__main__":
    unittest.main()
