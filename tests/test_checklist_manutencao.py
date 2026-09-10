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


class ManutencaoChecklistTest(unittest.TestCase):
    def setUp(self):
        self.enterContext(contextlib.redirect_stdout(io.StringIO()))
        self.enterContext(patch.object(checklist, "PRAZO_DIAS", 7))

    def simular(self, respostas_corretas):
        locais = [
            ("EST0037", 7, "VIX"),
            ("EST EL007", 20, "SSA"),
            ("EST EL016", 2, "CON"),
            ("EST EL019", 2, "CON"),
        ]
        linhas = [
            [placa, codigo, "MANUTENCAO", "Esteira Elétrica", "Sim"]
            for placa, codigo, _ in locais
        ]
        # O cadastro operacional identifica as siglas, mesmo desabilitado.
        linhas += [
            [f"BASE_{sigla}", codigo, sigla, "Esteira Elétrica", "Não"]
            for _, codigo, sigla in locais
        ]
        base = pd.DataFrame(
            linhas,
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
                    "PLACA_NORMALIZADA": checklist.normalizar_placa(placa),
                    "CODIGOFILIALRESPOSTA": (
                        codigo if respostas_corretas else (2 if codigo != 2 else 7)
                    ),
                    "CODIGO_FILIAL_RESPOSTA_NORMALIZADO": str(
                        codigo if respostas_corretas else (2 if codigo != 2 else 7)
                    ),
                    "NOMECHECKLIST": "Esteira Elétrica",
                    "TIPO_REALIZADO_NORMALIZADO": checklist.normalizar_tipo(
                        "Esteira Elétrica"
                    ),
                    "DATARESPOSTA_CONVERTIDA": pd.Timestamp.now()
                    - pd.Timedelta(days=1),
                    "RESPONDIDOPOR": "Teste manutenção",
                    "RESPOSTALOG": "",
                }
                for placa, codigo, _ in locais
            ]
        )
        return equipamentos, checklist.processar(equipamentos, historico)

    def test_filial_esperada_manutencao_com_codigo_individual(self):
        _, resultado = self.simular(True)
        self.assertEqual(resultado["FILIAL_ESPERADA"].tolist(), ["MANUTENCAO"] * 4)
        self.assertEqual(
            resultado["FILIAL_RESPOSTA"].tolist(), ["VIX", "SSA", "CON", "CON"]
        )
        self.assertFalse(resultado["INCONSISTENCIA_FILIAL"].any())
        self.assertFalse(resultado["TEM_PROBLEMA"].any())

    def test_codigo_de_outro_local_continua_inconsistente(self):
        _, resultado = self.simular(False)
        self.assertTrue(resultado["INCONSISTENCIA_FILIAL"].all())
        self.assertTrue(
            resultado["DESCRICAO_INCONSISTENCIA"]
            .str.contains("Filial esperada: MANUTENCAO", regex=False)
            .all()
        )
        html = checklist.gerar_html_email("MANUTENCAO", resultado)
        self.assertEqual(html.count("Filial esperada: MANUTENCAO"), 4)
        self.assertNotIn("Filial esperada: CON", html)

    def test_siglas_operacionais_nao_recebem_nome_do_grupo(self):
        equipamentos, _ = self.simular(True)
        mapa = equipamentos.attrs["SIGLAS_FILIAIS"]
        self.assertEqual(
            {k: mapa[k] for k in ["2", "7", "20"]},
            {"2": "CON", "7": "VIX", "20": "SSA"},
        )

    def test_cobranca_so_para_manutencao(self):
        _, resultado = self.simular(False)
        destinatarios = {
            sigla: [f"{sigla.lower()}@example.com"]
            for sigla in ["MANUTENCAO", "CON", "VIX", "SSA"]
        }
        with (
            patch.object(checklist, "DESTINATARIOS", destinatarios),
            patch.object(checklist, "enviar_email") as enviar,
        ):
            checklist.enviar_emails(resultado)
        enviar.assert_called_once()
        self.assertEqual(enviar.call_args.args[0], ["manutencao@example.com"])
        self.assertIn("[CHECKLIST] MANUTENCAO", enviar.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
