import contextlib
import io
import os
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

    def processar(
        self,
        codigo_resposta="1",
        ativo="Sim",
        com_resposta=True,
        grupo="BACKUP",
        codigo_esperado=2,
    ):
        base = pd.DataFrame(
            [
                [f"TESTE_{grupo}", codigo_esperado, grupo, "Bateria", ativo],
                ["CADASTRO_CON", 2, "CON", "Bateria", "Não"],
                ["CADASTRO_BHZ", 1, "BHZ", "Bateria", "Não"],
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
                    "PLACA_NORMALIZADA": f"TESTE_{grupo}" if com_resposta else "OUTRA",
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
        self.assertEqual(row["FILIAL_ESPERADA"], "CON")
        # BHZ continua identificável mesmo com seu equipamento desabilitado.
        self.assertEqual(row["FILIAL_RESPOSTA"], "BHZ")
        self.assertEqual(len(resultado), 1)
        self.assertTrue(row["INCONSISTENCIA_FILIAL"])
        corpo = checklist.gerar_html_email("BACKUP", resultado)
        self.assertIn("Filial esperada: CON | Filial informada: BHZ", corpo)
        self.assertNotIn("Código informado", corpo)
        self.assertNotIn("Cód. filial", corpo)
        self.assertEqual(corpo.count("Realizado por"), 2)
        self.assertEqual(corpo.count("Ana &lt;Silva&gt; &amp; Souza"), 2)
        self.assertNotIn("Antigo", corpo)

    def test_backup_resposta_con_e_envio_separado(self):
        resultado = self.processar(codigo_resposta="2")
        row = resultado.iloc[0]
        self.assertEqual(row["FILIAL_RESPOSTA"], "CON")
        self.assertFalse(row["INCONSISTENCIA_FILIAL"])
        with (
            patch.object(
                checklist,
                "DESTINATARIOS",
                {"BACKUP": ["backup@example.com"], "CON": ["con@example.com"]},
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

    def test_frotas_aceita_con_e_sinaliza_outras_filiais(self):
        for codigo, inconsistente, sigla in [("2", False, "CON"), ("1", True, "BHZ")]:
            with self.subTest(codigo=codigo):
                resultado = self.processar(
                    codigo_resposta=codigo, grupo="FROTAS", codigo_esperado=2
                )
                row = resultado.iloc[0]
                self.assertEqual(row["FILIAL"], "FROTAS")
                self.assertEqual(row["FILIAL_ESPERADA"], "CON")
                self.assertEqual(row["FILIAL_RESPOSTA"], sigla)
                self.assertEqual(row["INCONSISTENCIA_FILIAL"], inconsistente)

    def test_frotas_e_con_recebem_mensagens_separadas(self):
        frotas = self.processar(grupo="FROTAS", codigo_esperado=2)
        con = self.processar(grupo="CON", codigo_esperado=2)
        resultado = pd.concat([frotas, con], ignore_index=True)
        resumo = checklist.gerar_resumo(resultado)
        self.assertEqual(set(resumo["FILIAL"]), {"FROTAS", "CON"})
        with patch.dict(
            os.environ,
            {"FROTAS_EMAIL_1": "frotas@example.com", "CON_EMAIL_1": "con@example.com"},
            clear=True,
        ):
            destinatarios = checklist.carregar_destinatarios()
        with (
            patch.object(checklist, "DESTINATARIOS", destinatarios),
            patch.object(checklist, "enviar_email") as enviar,
        ):
            checklist.enviar_emails(resultado)
        self.assertEqual(enviar.call_count, 2)
        mensagens = {
            chamada.args[0][0]: chamada.args for chamada in enviar.call_args_list
        }
        for grupo, outro in [("FROTAS", "CON"), ("CON", "FROTAS")]:
            mensagem = mensagens[f"{grupo.lower()}@example.com"]
            self.assertIn(f"[CHECKLIST] {grupo}", mensagem[1])
            self.assertIn(f"TESTE_{grupo}", mensagem[2])
            self.assertNotIn(f"TESTE_{outro}", mensagem[2])

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
