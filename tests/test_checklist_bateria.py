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
from automacao.relatorios_pdf import dados_relatorio, gerar_pdf_checklist, resumir


class BateriaChecklistTest(unittest.TestCase):
    def setUp(self):
        self.enterContext(contextlib.redirect_stdout(io.StringIO()))
        self.enterContext(patch.object(checklist, "PRAZO_DIAS", 7))

    def simular(self, respostas, tipos=None):
        tipos = tipos or ["Bateria", "Troca inicial - Bateria", "Troca Final - Bateria"]
        base = pd.DataFrame(
            [["BAT001", 2, "CON", tipo] for tipo in tipos],
            columns=["Placa", "Código Filial", "Filial", "Tipo de Equipamento"],
        )
        with patch.object(pd, "read_excel", return_value=base):
            equipamentos = checklist.carregar_equipamentos()
        historico = pd.DataFrame(
            [
                {
                    "PLACA_NORMALIZADA": "BAT001" if tipo is not None else "OUTRA",
                    "CODIGOFILIALRESPOSTA": 2,
                    "CODIGO_FILIAL_RESPOSTA_NORMALIZADO": "2",
                    "NOMECHECKLIST": tipo or "Bateria",
                    "TIPO_REALIZADO_NORMALIZADO": (tipo or "Bateria").upper(),
                    "DATARESPOSTA_CONVERTIDA": pd.Timestamp.now()
                    - pd.Timedelta(days=dias),
                    "RESPONDIDOPOR": "Pessoa teste",
                    "RESPOSTALOG": "",
                }
                for tipo, dias in respostas
            ]
        )
        return checklist.processar(equipamentos, historico)

    def test_trocas_recentes_nao_quitam_pendencia_mesmo_com_bateria_anterior(self):
        for tipo in ["Troca inicial - Bateria", "Troca Final - Bateria", "Automatic"]:
            with self.subTest(tipo=tipo):
                resultado = self.simular([(tipo, 1), ("Bateria", 2)])
                row = resultado.iloc[0]
                self.assertEqual(row["TIPOS_ACEITOS"], ["Bateria"])
                self.assertEqual(row["CHECKLIST_REALIZADO"], tipo)
                self.assertEqual(row["STATUS_PRAZO"], "OK")
                self.assertEqual(row["STATUS_CHECKLIST"], "PENDENTE")
                self.assertTrue(row["PENDENCIA_BATERIA"])
                self.assertTrue(row["TEM_PROBLEMA"])
                self.assertIn("Checklist Bateria obrigatório", row["MOTIVO_PENDENCIA"])
                resumo = checklist.gerar_resumo(resultado).iloc[0]
                self.assertEqual(resumo["PENDENTES"], 1)
                self.assertEqual(resumo["OK"], 0)

    def test_bateria_mais_recente_regulariza(self):
        resultado = self.simular([("Troca Final - Bateria", 2), ("Bateria", 1)])
        row = resultado.iloc[0]
        self.assertEqual(row["STATUS_CHECKLIST"], "OK")
        self.assertFalse(row["PENDENCIA_BATERIA"])
        self.assertFalse(row["TEM_PROBLEMA"])

    def test_cadastro_antigo_com_apenas_troca_exige_bateria(self):
        resultado = self.simular(
            [("Troca inicial - Bateria", 1)], tipos=["Troca inicial - Bateria"]
        )
        self.assertEqual(resultado.iloc[0]["TIPOS_ACEITOS"], ["Bateria"])
        self.assertEqual(resultado.iloc[0]["STATUS_CHECKLIST"], "PENDENTE")

    def test_pdf_conta_pendencia_sem_confundir_com_vencimento(self):
        for respostas, situacao in [
            ([("Troca Final - Bateria", 1)], "tipo_obrigatorio"),
            ([("Troca Final - Bateria", 10)], "vencido"),
            ([(None, 1)], "sem"),
        ]:
            with self.subTest(situacao=situacao):
                resultado = self.simular(respostas)
                dados = dados_relatorio(resultado, 7, "checklist")
                self.assertEqual(dados["registros"][0]["situacao"], situacao)
                resumo = resumir(dados["registros"])
                self.assertEqual(resumo["pendentes"], 1)
                self.assertEqual(resumo["regular"], 0)
                self.assertEqual(
                    sum(resumo[k] for k in ["sem", "vencido", "tipo_obrigatorio"]), 1
                )
                self.assertTrue(gerar_pdf_checklist(resultado, 7).startswith(b"%PDF-"))

    def test_email_conta_troca_recente_como_pendente_e_mantem_tabela(self):
        resultado = self.simular([("Troca Final - Bateria", 1)])
        with (
            patch.object(checklist, "DESTINATARIOS", {"CON": ["teste@example.com"]}),
            patch.object(checklist, "enviar_email") as enviar,
        ):
            checklist.enviar_emails(resultado)
        enviar.assert_called_once()
        assunto, corpo = enviar.call_args.args[1:3]
        self.assertIn("1 pendência(s)", assunto)
        self.assertIn("Checklist Bateria obrigatório", corpo)
        self.assertIn("<table", corpo)
        self.assertIn("Pessoa teste", corpo)

    def test_outros_equipamentos_preservam_regra_de_tipo(self):
        resultado = self.simular([("Bateria", 1)], tipos=["Automatic"])
        row = resultado.iloc[0]
        self.assertEqual(row["STATUS_CHECKLIST"], "OK")
        self.assertFalse(row["PENDENCIA_BATERIA"])
        self.assertTrue(row["INCONSISTENCIA_TIPO"])


if __name__ == "__main__":
    unittest.main()
