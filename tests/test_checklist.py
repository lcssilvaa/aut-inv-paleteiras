import contextlib
import io
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd

# Permite executar este arquivo diretamente, inclusive de dentro de tests.
if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import checklist


class EquipamentosAtivosTest(unittest.TestCase):
    def setUp(self):
        temporario = TemporaryDirectory()
        self.addCleanup(temporario.cleanup)
        self.pasta = Path(temporario.name)
        self.base = self.pasta / "Base Checklist.xlsx"
        self.enterContext(patch.object(checklist, "ARQUIVO_EQUIPAMENTOS", self.base))
        self.enterContext(patch.object(checklist, "SAIDA_DIR", self.pasta / "saida"))
        self.enterContext(contextlib.redirect_stdout(io.StringIO()))
        self.enviar = self.enterContext(patch.object(checklist, "enviar_email"))
        self.enterContext(
            patch.object(checklist, "DESTINATARIOS", {"BACKUP": ["teste@example.com"]})
        )

    def carregar(self, linhas, com_ativo=True):
        colunas = ["Placa", "Código Filial", "Filial", "Tipo de Equipamento"]
        if com_ativo:
            colunas.append("Ativo")
        pd.DataFrame(linhas, columns=colunas).to_excel(self.base, index=False)
        return checklist.carregar_equipamentos()

    def processar(self, equipamentos):
        # Ambos os equipamentos têm um checklist recente, mas de tipo incorreto.
        historico = pd.DataFrame(
            [
                {
                    "PLACA_NORMALIZADA": placa,
                    "CODIGOFILIALRESPOSTA": 1,
                    "CODIGO_FILIAL_RESPOSTA_NORMALIZADO": "1",
                    "NOMECHECKLIST": "Automatic",
                    "TIPO_REALIZADO_NORMALIZADO": "AUTOMATIC",
                    "DATARESPOSTA_CONVERTIDA": pd.Timestamp.now(),
                    "RESPONDIDOPOR": "Teste",
                    "RESPOSTALOG": "",
                }
                for placa in ["ATIVA", "MANUTENCAO"]
            ]
        )
        resultado = checklist.processar(equipamentos, historico)
        resumo = checklist.gerar_resumo(resultado)
        checklist.salvar_resultados(resultado, resumo)
        checklist.enviar_emails(resultado)
        return resultado, resumo

    def test_coluna_ausente_ou_vazia_preserva_cadastro(self):
        linha = ["ATIVA", 1, "BACKUP", "Bateria"]
        antigo = self.carregar([linha], com_ativo=False)
        vazio = self.carregar([linha + [None]])
        pd.testing.assert_frame_equal(antigo, vazio)

    def test_valores_aceitos_e_erro_de_digitacao(self):
        for valor in [None, "", "  ", "Sim", " sIM ", True, 1, 1.0]:
            with self.subTest(valor=valor):
                self.assertTrue(checklist.interpretar_ativo(valor))
        for valor in ["Não", " nao ", False, 0, 0.0]:
            with self.subTest(valor=valor):
                self.assertFalse(checklist.interpretar_ativo(valor))
        with self.assertRaisesRegex(ValueError, "coluna Ativo"):
            self.carregar([["ATIVA", 1, "BACKUP", "Bateria", "Talvez"]])

    def test_uma_linha_desabilita_todos_os_tipos_da_placa(self):
        linhas = [
            ["MANUTENCAO", 1, "BACKUP", "Bateria", "Sim"],
            [" manutencao ", 1, "BACKUP", "Troca inicial - Bateria", "Não"],
            ["MANUTENCAO", 1, "BACKUP", "Troca Final - Bateria", None],
            ["ATIVA", 1, "BACKUP", "Bateria", "Sim"],
        ]
        equipamentos = self.carregar(linhas)
        self.assertEqual(equipamentos["Placa"].tolist(), ["ATIVA"])
        linhas[1][-1] = "Sim"
        reativados = self.carregar(linhas).set_index("PLACA_NORMALIZADA")
        self.assertEqual(len(reativados), 2)
        self.assertEqual(len(reativados.loc["MANUTENCAO", "TIPOS_ACEITOS"]), 3)

    def test_desabilitados_nao_entram_em_relatorios_resumo_ou_email(self):
        equipamentos = self.carregar(
            [
                ["ATIVA", 1, "BACKUP", "Bateria", "Sim"],
                ["MANUTENCAO", 1, "BACKUP", "Bateria", "Não"],
            ]
        )
        resultado, resumo = self.processar(equipamentos)
        self.assertEqual(resultado["Placa"].tolist(), ["ATIVA"])
        self.assertEqual(resumo["TOTAL"].sum(), 1)
        self.assertEqual(resumo["FILIAL"].tolist(), ["BACKUP"])
        for nome in ["Resultado_Checklist.xlsx", "Problemas_Checklist.xlsx"]:
            exportado = pd.read_excel(checklist.SAIDA_DIR / nome)
            self.assertEqual(exportado["Placa"].tolist(), ["ATIVA"])
        self.enviar.assert_called_once()
        corpo = self.enviar.call_args.args[2]
        self.assertIn("ATIVA", corpo)
        self.assertNotIn("MANUTENCAO", corpo)

    def test_todos_desabilitados_geram_relatorios_vazios_sem_email(self):
        linhas = [["ATIVA", 1, "BACKUP", "Bateria", "Sim"]]
        self.processar(self.carregar(linhas))
        self.enviar.reset_mock()
        linhas[0][-1] = "Não"
        resultado, resumo = self.processar(self.carregar(linhas))
        self.assertTrue(resultado.empty)
        self.assertTrue(resumo.empty)
        for nome in [
            "Resultado_Checklist.xlsx",
            "Problemas_Checklist.xlsx",
            "Resumo_Checklist.xlsx",
        ]:
            exportado = pd.read_excel(checklist.SAIDA_DIR / nome)
            self.assertTrue(exportado.empty)
            self.assertGreater(len(exportado.columns), 0)
        self.enviar.assert_not_called()


if __name__ == "__main__":
    unittest.main()
