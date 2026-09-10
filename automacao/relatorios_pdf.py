"""Resumos visuais em PDF para checklists e inventário de paleteiras."""

from collections import defaultdict
from functools import lru_cache
from io import BytesIO
import os
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import getAscentDescent, registerFont, stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import LongTable, Paragraph, TableStyle

LARGURA, ALTURA = 1280, 900
LOGO = Path(__file__).resolve().parents[1] / "logo" / "logo.png"
VERMELHO, VERDE, LARANJA = "#C90838", "#168756", "#EF922D"
CINZA, ROXO, TEXTO, SUAVE, FUNDO = "#AAA5B1", "#7961B2", "#24232B", "#787681", "#F5F3F5"


@lru_cache(maxsize=1)
def fonte_cabecalho():
    arquivo = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / "arialbd.ttf"
    if arquivo.is_file():
        registerFont(TTFont("CabecalhoPDF", str(arquivo)))
        return "CabecalhoPDF"
    return "Helvetica-Bold"


def texto(valor):
    return "" if pd.isna(valor) else str(valor)


def numero(valor):
    return None if pd.isna(valor) else float(valor)


def data(valor):
    return "-" if pd.isna(valor) else pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")


def decimal(valor, casas=1):
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def dinheiro(valor):
    return f"R$ {decimal(valor, 2)}"


def dados_relatorio(resultado, prazo_dias, tipo, filial=None):
    # Filtra também chamadas diretas para impedir anexos com outra filial.
    if filial is not None:
        resultado = resultado[
            resultado["FILIAL"].astype(str).str.strip().eq(filial)
        ].copy()
    registros = []
    for _, row in resultado.iterrows():
        checklist = tipo == "checklist"
        ultima = row["ULTIMA_RESPOSTA" if checklist else "ULTIMA_LEITURA"]
        no_prazo = row["STATUS_PRAZO"] == "OK" if checklist else row["STATUS"] == "LIDO"
        situacao = "sem" if pd.isna(ultima) else "no_prazo" if no_prazo else "vencido"
        if checklist and no_prazo and row.get("PENDENCIA_BATERIA", False):
            situacao = "tipo_obrigatorio"
        registros.append(
            {
                "id": texto(row["Placa"] if checklist else row["QR Code"]),
                "disk": "" if checklist else texto(row["NR_DISK"]),
                "filial": texto(row["FILIAL"]).strip(),
                "categoria": texto(
                    row["TIPOS_ACEITOS_TEXTO"] if checklist else row["MODELO"]
                ),
                "ultima": data(ultima),
                "responsavel": texto(
                    row["ULTIMO_RESPONSAVEL"]
                    if checklist
                    else row["ULTIMO_INVENTARIADOR"]
                ),
                "dias": numero(
                    row["DIAS_SEM_CHECKLIST"] if checklist else row["DIAS_SEM_LEITURA"]
                ),
                "situacao": situacao,
                "tipo_erro": bool(row["INCONSISTENCIA_TIPO"]) if checklist else False,
                "filial_erro": (
                    bool(row["INCONSISTENCIA_FILIAL"]) if checklist else False
                ),
                "problema": (
                    bool(row["TEM_PROBLEMA"]) if checklist else situacao != "no_prazo"
                ),
                "detalhe": texto(row["DESCRICAO_INCONSISTENCIA"]) if checklist else "",
                "custo": None if checklist else numero(row["CUSTO_NUMERICO"]),
            }
        )
    atualizacao = (
        resultado["DATA_ANALISE"].max() if not resultado.empty else pd.Timestamp.now()
    )
    return {
        "tipo": tipo,
        "titulo": "Checklists" if tipo == "checklist" else "Paleteiras",
        "filial": filial,
        "prazo": int(prazo_dias),
        "atualizacao": data(atualizacao),
        "registros": registros,
    }


def resumir(registros):
    pendentes = [r for r in registros if r["situacao"] != "no_prazo"]
    return {
        "total": len(registros),
        "regular": sum(not r["problema"] for r in registros),
        "no_prazo": sum(r["situacao"] == "no_prazo" for r in registros),
        "vencido": sum(r["situacao"] == "vencido" for r in registros),
        "sem": sum(r["situacao"] == "sem" for r in registros),
        "tipo_obrigatorio": sum(r["situacao"] == "tipo_obrigatorio" for r in registros),
        "pendentes": len(pendentes),
        "tipo": sum(r["tipo_erro"] for r in registros),
        "filial": sum(r["filial_erro"] for r in registros),
        "inconsistencias": sum(r["tipo_erro"] or r["filial_erro"] for r in registros),
        "valor": sum(r["custo"] or 0 for r in pendentes),
        "sem_custo": sum(r["custo"] is None for r in pendentes),
    }


def agrupar(registros, chave):
    grupos = defaultdict(list)
    for registro in registros:
        grupos[registro[chave] or "Não informado"].append(registro)
    return sorted(
        [{"nome": nome, **resumir(itens)} for nome, itens in grupos.items()],
        key=lambda g: (-g["total"], g["nome"]),
    )


class PaginaVisual:
    """Desenha em coordenadas a partir do topo da página."""

    def __init__(self, canvas, dados):
        self.canvas, self.dados = canvas, dados

    def retangulo(self, x, y, largura, altura, cor, raio=0):
        self.canvas.setFillColor(colors.HexColor(cor))
        self.canvas.roundRect(
            x, ALTURA - y - altura, largura, altura, raio, fill=1, stroke=0
        )

    def escrever(
        self, x, y, valor, tamanho=10, cor=TEXTO, negrito=False, limite=None, fonte=None
    ):
        fonte = fonte or ("Helvetica-Bold" if negrito else "Helvetica")
        valor = str(valor)
        if limite is not None:
            while valor and stringWidth(valor, fonte, tamanho) > limite:
                valor = valor[:-4] + "..." if len(valor) > 4 else ""
        self.canvas.setFillColor(colors.HexColor(cor))
        self.canvas.setFont(fonte, tamanho)
        self.canvas.drawString(x, ALTURA - y, valor)

    def escrever_centralizado(self, cx, cy, valor, tamanho=25):
        fonte = "Helvetica-Bold"
        largura = stringWidth(str(valor), fonte, tamanho)
        ascendente, descendente = getAscentDescent(fonte, tamanho)
        self.escrever(
            cx - largura / 2,
            cy + (ascendente + descendente) / 2,
            valor,
            tamanho,
            negrito=True,
        )

    def cabecalho(self, subtitulo):
        self.retangulo(0, 0, LARGURA, ALTURA, FUNDO)
        self.retangulo(0, 0, LARGURA, 82, VERMELHO)
        self.canvas.drawImage(
            str(LOGO),
            32,
            ALTURA - 64,
            width=125,
            height=46,
            preserveAspectRatio=True,
            mask="auto",
        )
        self.retangulo(184, 20, 3, 42, "#FFFFFF", 1.5)
        fonte = fonte_cabecalho()
        titulo = "CHECKLISTS" if self.dados["tipo"] == "checklist" else "INVENTÁRIO"
        self.escrever(214, 51, titulo, 28, "#FFFFFF", fonte=fonte)
        self.escrever(32, 119, subtitulo, 20, fonte=fonte)
        grupo = self.dados["filial"] or "Todas as filiais"
        self.escrever(
            32,
            137,
            f"Grupo: {grupo}   |   Prazo: {self.dados['prazo']} dias",
            10,
            SUAVE,
        )
        self.escrever(943, 119, f"Atualizado: {self.dados['atualizacao']}", 10, SUAVE)

    def cartao(self, x, y, largura, titulo, subtitulo):
        self.retangulo(x, y, largura, 257, "#FFFFFF", 11)
        self.escrever(x + 18, y + 28, titulo.upper(), 10, VERMELHO, True, largura - 36)
        self.escrever(x + 18, y + 47, subtitulo, 8.5, SUAVE, limite=largura - 36)

    def barras(self, x, y, largura, itens, percentual=False):
        # Cada item contém rótulo, segmentos (valor/cor) e texto à direita.
        if not itens:
            self.escrever(x + 18, y + 125, "Nenhum registro neste recorte.", 10, SUAVE)
            return
        maior = max(1, max(sum(v for v, _ in partes) for _, partes, _ in itens))
        rotulo_largura = 112
        barra_largura = largura - rotulo_largura - 76
        for indice, (rotulo, partes, valor) in enumerate(itens[:10]):
            topo = y + 68 + indice * 16
            self.escrever(
                x + 18, topo + 9, rotulo, 8, negrito=True, limite=rotulo_largura - 8
            )
            self.retangulo(
                x + 18 + rotulo_largura, topo, barra_largura, 11, "#F1EDF1", 2
            )
            total = sum(v for v, _ in partes)
            escala = total if percentual and total else maior
            inicio = x + 18 + rotulo_largura
            for valor_parte, cor in partes:
                comprimento = max(0, valor_parte / escala * barra_largura)
                if comprimento:
                    self.retangulo(inicio, topo, comprimento, 11, cor)
                if percentual and comprimento >= 34:
                    rotulo_percentual = f"{decimal(valor_parte / total * 100, 0)}%"
                    deslocamento = (
                        comprimento
                        - stringWidth(rotulo_percentual, "Helvetica-Bold", 7)
                    ) / 2
                    self.escrever(
                        inicio + deslocamento,
                        topo + 8.5,
                        rotulo_percentual,
                        7,
                        "#FFFFFF",
                        True,
                    )
                inicio += comprimento
            self.escrever(x + largura - 44, topo + 9, valor, 8, limite=40)
        if len(itens) > 10:
            self.escrever(
                x + 18,
                y + 243,
                f"Maiores 10 de {len(itens)}. Detalhes nas próximas páginas.",
                7.5,
                SUAVE,
            )


def desenhar_visao(canvas, dados):
    p = PaginaVisual(canvas, dados)
    p.cabecalho("Visão geral da operação")
    registros = dados["registros"]
    s = resumir(registros)
    checklist = dados["tipo"] == "checklist"
    indicadores = [
        ("EQUIPAMENTOS", str(s["total"]), "Total de registros no recorte", VERMELHO),
        (
            "EM DIA" if checklist else "LEITURAS NO PRAZO",
            str(s["regular"]),
            (
                "Prazo, tipo e filial corretos"
                if checklist
                else "Leituras dentro do prazo"
            ),
            VERDE,
        ),
        (
            "PENDENTES",
            str(s["pendentes"]),
            f"{s['vencido']} vencidos | {s['sem']} sem registro"
            + (
                f" | {s['tipo_obrigatorio']} tipo obrigatório"
                if s["tipo_obrigatorio"]
                else ""
            ),
            VERMELHO,
        ),
        (
            "COM INCONSISTÊNCIAS" if checklist else "VALOR DOS EQUIPAMENTOS PENDENTES",
            str(s["inconsistencias"]) if checklist else dinheiro(s["valor"]),
            (
                f"{s['tipo']} de tipo | {s['filial']} de filial"
                if checklist
                else "Soma dos custos cadastrados"
            ),
            LARANJA,
        ),
    ]
    for indice, (titulo, valor, nota, cor) in enumerate(indicadores):
        x = 32 + indice * 310
        p.retangulo(x, 160, 286, 88, "#FFFFFF", 9)
        p.retangulo(x, 173, 4, 62, cor, 2)
        p.escrever(x + 17, 181, titulo, 8.5, SUAVE, True, 255)
        p.escrever(
            x + 17, 214, valor, 29 if len(valor) < 10 else 22, negrito=True, limite=253
        )
        p.escrever(x + 17, 234, nota, 8.5, SUAVE, limite=252)
    largura, primeira, segunda = 389, 273, 552
    xs = [32, 445, 858]
    chave = "categoria" if dados["filial"] else "filial"
    eixo = ("tipo" if checklist else "modelo") if dados["filial"] else "filial"
    grupos = agrupar(registros, chave)
    p.cartao(
        xs[0],
        primeira,
        largura,
        f"Equipamentos por {eixo}",
        "Distribuição dos registros no cadastro",
    )
    p.barras(
        xs[0],
        primeira,
        largura,
        [(g["nome"], [(g["total"], LARANJA)], str(g["total"])) for g in grupos],
    )
    p.cartao(
        xs[1],
        primeira,
        largura,
        f"Situação dos registros por {eixo}",
        "Verde: no prazo | Vermelho: vencido | Cinza: sem registro",
    )
    p.barras(
        xs[1],
        primeira,
        largura,
        [
            (
                g["nome"],
                [
                    (g["no_prazo"], VERDE),
                    (g["vencido"], VERMELHO),
                    (g["sem"], CINZA),
                    (g["tipo_obrigatorio"], LARANJA),
                ],
                str(g["total"]),
            )
            for g in grupos
        ],
        percentual=True,
    )
    if s["tipo_obrigatorio"]:
        p.escrever(
            xs[1] + 18,
            primeira + 59,
            "Laranja: pendente por checklist Bateria obrigatório",
            7.5,
            SUAVE,
        )
    if checklist:
        p.cartao(
            xs[2],
            primeira,
            largura,
            f"Inconsistências por {eixo}",
            "Laranja: tipo | Roxo: filial (podem ocorrer juntas)",
        )
        grupos_falha = sorted(
            [g for g in grupos if g["inconsistencias"]],
            key=lambda g: -(g["tipo"] + g["filial"]),
        )
        p.barras(
            xs[2],
            primeira,
            largura,
            [
                (
                    g["nome"],
                    [(g["tipo"], LARANJA), (g["filial"], ROXO)],
                    str(g["tipo"] + g["filial"]),
                )
                for g in grupos_falha
            ],
        )
    else:
        p.cartao(
            xs[2],
            primeira,
            largura,
            f"Valor pendente por {eixo}",
            "Custos cadastrados em R$ mil; não representa perda",
        )
        custos = sorted(
            [g for g in grupos if g["pendentes"]], key=lambda g: -g["valor"]
        )
        p.barras(
            xs[2],
            primeira,
            largura,
            [
                (g["nome"], [(g["valor"], LARANJA)], decimal(g["valor"] / 1000))
                for g in custos
            ],
        )
    p.cartao(
        xs[0],
        segunda,
        largura,
        "Há mais tempo sem registro",
        "Equipamentos vencidos, com data conhecida (dias)",
    )
    vencidos = sorted(
        [r for r in registros if r["situacao"] == "vencido" and r["dias"] is not None],
        key=lambda r: -r["dias"],
    )
    p.barras(
        xs[0],
        segunda,
        largura,
        [(r["id"], [(r["dias"], VERMELHO)], decimal(r["dias"])) for r in vencidos],
    )
    p.cartao(
        xs[1],
        segunda,
        largura,
        "Situação dos equipamentos" if checklist else "Cobertura do inventário",
        (
            "Cumprimento do prazo, tipo e filial"
            if checklist
            else "Proporção de equipamentos com leitura no prazo"
        ),
    )
    proporcao = s["regular"] / s["total"] if s["total"] else 0
    cx, cy = xs[1] + 113, ALTURA - (segunda + 151)
    canvas.setFillColor(colors.HexColor(VERMELHO if s["total"] else CINZA))
    canvas.circle(cx, cy, 68, stroke=0, fill=1)
    if proporcao:
        canvas.setFillColor(colors.HexColor(VERDE))
        canvas.wedge(
            cx - 68, cy - 68, cx + 68, cy + 68, 90, -360 * proporcao, stroke=0, fill=1
        )
    canvas.setFillColor(colors.white)
    canvas.circle(cx, cy, 53, stroke=0, fill=1)
    p.escrever_centralizado(
        cx,
        segunda + 151,
        f"{decimal(proporcao * 100)}%" if s["total"] else "-",
        25,
    )
    p.escrever(
        xs[1] + 211,
        segunda + 107,
        "Em dia" if checklist else "No prazo",
        10,
        SUAVE,
    )
    p.escrever(xs[1] + 211, segunda + 135, s["regular"], 25, VERDE, True)
    p.escrever(
        xs[1] + 211,
        segunda + 177,
        "Com pendência ou divergência" if checklist else "Pendentes",
        9,
        SUAVE,
    )
    p.escrever(
        xs[1] + 211, segunda + 204, s["total"] - s["regular"], 25, VERMELHO, True
    )
    p.cartao(
        xs[2],
        segunda,
        largura,
        "Equipamentos por situação",
        f"Prazo vencido ao completar {dados['prazo']} dias",
    )
    valores = [
        ("No prazo", s["no_prazo"], VERDE),
        ("Vencido", s["vencido"], VERMELHO),
        ("Sem checklist" if checklist else "Nunca lido", s["sem"], CINZA),
    ]
    if checklist:
        valores.append(("Tipo obrigatório", s["tipo_obrigatorio"], LARANJA))
    maior = max(1, max(valor for _, valor, _ in valores))
    passo = (largura - 50) / len(valores)
    for i, (nome, valor, cor) in enumerate(valores):
        centro = xs[2] + 25 + passo * (i + 0.5)
        altura = valor / maior * 125
        p.retangulo(
            centro - passo * 0.28,
            segunda + 211 - altura,
            passo * 0.56,
            max(1, altura),
            cor,
            3,
        )
        p.escrever(
            centro - stringWidth(str(valor), "Helvetica-Bold", 11) / 2,
            segunda + 203 - altura,
            valor,
            11,
            negrito=True,
        )
        p.escrever(
            centro - stringWidth(nome, "Helvetica", 8) / 2,
            segunda + 230,
            nome,
            8,
            limite=passo,
        )
    canvas.showPage()


def desenhar_tabela(canvas, dados, titulo, cabecalhos, linhas, larguras):
    estilo = ParagraphStyle(
        "celula",
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=colors.HexColor(TEXTO),
        wordWrap="CJK",
    )
    estilo_cabecalho = ParagraphStyle(
        "cabecalho",
        parent=estilo,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor(VERMELHO),
    )
    conteudo = [[Paragraph(escape(str(c)), estilo_cabecalho) for c in cabecalhos]]
    conteudo.extend(
        [
            [
                Paragraph(escape(str(c if c is not None and c != "" else "-")), estilo)
                for c in linha
            ]
            for linha in linhas
        ]
    )
    if not linhas:
        conteudo.append(
            [Paragraph("Nenhum equipamento neste recorte.", estilo)]
            + [""] * (len(cabecalhos) - 1)
        )
    tabela = LongTable(conteudo, colWidths=larguras, repeatRows=1, hAlign="LEFT")
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F9EAF0")),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.white, colors.HexColor("#FAF8FA")],
                ),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("LINEBELOW", (0, 0), (-1, -1), 0.3, colors.HexColor("#E5E0E6")),
            ]
        )
    )
    pendentes = [tabela]
    while pendentes:
        atual = pendentes.pop(0)
        _, altura = atual.wrap(1216, 680)
        if altura > 680:
            partes = atual.split(1216, 680)
            if len(partes) < 2:
                raise ValueError(
                    "Uma linha do relatório excede o espaço disponível na página."
                )
            atual = partes[0]
            pendentes = partes[1:] + pendentes
            _, altura = atual.wrap(1216, 680)
        pagina = PaginaVisual(canvas, dados)
        pagina.cabecalho(titulo)
        atual.drawOn(canvas, 32, ALTURA - 158 - altura)
        canvas.showPage()


def gerar_pdf(dados):
    arquivo = BytesIO()
    canvas = Canvas(arquivo, pagesize=(LARGURA, ALTURA), pageCompression=1)
    canvas.setTitle(f"{dados['titulo']} - {dados['filial'] or 'Todas as filiais'}")
    canvas.setAuthor("Patrus Transportes")
    desenhar_visao(canvas, dados)
    checklist = dados["tipo"] == "checklist"
    if dados["filial"] is None:
        grupos = agrupar(dados["registros"], "filial")
        cabecalhos = [
            "Filial / grupo",
            "Total",
            "No prazo",
            "Vencidos",
            "Sem registro",
            "Em dia" if checklist else "Leituras no prazo",
            "Tipo incorreto" if checklist else "Valor pendente",
            "Filial incorreta" if checklist else "Sem custo",
        ]
        linhas = [
            [
                g["nome"],
                g["total"],
                g["no_prazo"],
                g["vencido"],
                g["sem"],
                g["regular"],
                g["tipo"] if checklist else dinheiro(g["valor"]),
                g["filial"] if checklist else g["sem_custo"],
            ]
            for g in grupos
        ]
        larguras = [200, 110, 140, 140, 140, 160, 180, 146]
        if checklist:
            cabecalhos.append("Pendente por tipo obrigatório")
            for linha, grupo in zip(linhas, grupos):
                linha.append(grupo["tipo_obrigatorio"])
            larguras = [160, 80, 120, 110, 120, 120, 170, 170, 166]
        desenhar_tabela(
            canvas, dados, "Resumo de todas as filiais", cabecalhos, linhas, larguras
        )
    nomes_status = {
        "no_prazo": "No prazo",
        "vencido": "Prazo vencido",
        "sem": "Sem checklist" if checklist else "Nunca lido",
        "tipo_obrigatorio": "Pendente: Bateria obrigatório",
    }
    registros = sorted(
        dados["registros"],
        key=lambda r: (
            not r["problema"],
            r["filial"],
            -(r["dias"] if r["dias"] is not None else float("inf")),
            r["id"],
        ),
    )
    cabecalhos = [
        "Placa" if checklist else "QR Code / Disk",
        "Filial / grupo",
        "Tipos aceitos" if checklist else "Modelo",
        "Último registro",
        "Realizado por",
        "Dias",
        "Situação",
        "Inconsistências" if checklist else "Custo cadastrado",
    ]
    linhas = [
        [
            r["id"] if checklist else f"{r['id']} / {r['disk']}",
            r["filial"],
            r["categoria"],
            r["ultima"],
            r["responsavel"],
            "-" if r["dias"] is None else decimal(r["dias"]),
            nomes_status[r["situacao"]],
            (
                r["detalhe"]
                if checklist
                else "-" if r["custo"] is None else dinheiro(r["custo"])
            ),
        ]
        for r in registros
    ]
    desenhar_tabela(
        canvas,
        dados,
        "Detalhamento dos equipamentos",
        cabecalhos,
        linhas,
        [130, 85, 180, 116, 140, 48, 103, 414],
    )
    canvas.save()
    return arquivo.getvalue()


def gerar_pdf_checklist(resultado, prazo_dias, filial=None):
    return gerar_pdf(dados_relatorio(resultado, prazo_dias, "checklist", filial))


def gerar_pdf_paleteiras(resultado, prazo_dias, filial=None):
    return gerar_pdf(dados_relatorio(resultado, prazo_dias, "paleteiras", filial))


def salvar_pdf_checklist(resultado, saida_dir, prazo_dias):
    destino = Path(saida_dir) / "Resumo_Visual_Checklist.pdf"
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(gerar_pdf_checklist(resultado, prazo_dias))
    return destino


def salvar_pdf_paleteiras(resultado, saida_dir, prazo_dias):
    destino = Path(saida_dir) / "Resumo_Visual_Paleteiras.pdf"
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(gerar_pdf_paleteiras(resultado, prazo_dias))
    return destino
