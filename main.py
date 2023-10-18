from PIL.Image import Image
from openpyxl import Workbook
from datetime import date

from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, PatternFill,fills
from openpyxl.chart import LineChart, Reference
from classes import Leitor, PropriedadesGrafico, GerenciadorPlanilha
#acao = input("qual o codigo da acao? ") # BIDI4
try:
    acao = "BIDI4"

    leitor = Leitor(caminho_arq='./dados/')
    leitor.ler_arquivo(acao)

    gerenciador_planilha = GerenciadorPlanilha()
    planilha_dados = gerenciador_planilha.criar_planilha(nome_planilha="dados")

    gerenciador_planilha.adiciona_linha(["DATA", "COTACAO", "BANDA INFERIOR", "BANDA SUPERIOR"])

    indice = 2

    for linha in leitor.dados:
        # DATA
        ano_mes_dia = linha[0].split(" ")[0]
        data = date(year=int(ano_mes_dia.split("-")[0]), month=int(ano_mes_dia.split("-")[1]),
                    day=int(ano_mes_dia.split("-")[2]))

        # COTACAO
        cotacao = float(linha[1])

        gerenciador_planilha.atualiza_cell(cell=f'A{indice}', valor=data)
        gerenciador_planilha.atualiza_cell(cell=f'B{indice}', valor=cotacao)

        # BANDA INFERIOR
        # media movel (20p) - 2 * desvio padrao (20p)
        gerenciador_planilha.atualiza_cell(cell=f'C{indice}',
                                           valor=f'=AVERAGE(B{indice}:B{indice + 19})-2*STDEV(B{indice}:B{indice + 19})')

        # BANDA SUPERIOR
        # media movel (20p) + 2 * desvio padrao (20p)
        gerenciador_planilha.atualiza_cell(cell=f'D{indice}',
                                           valor=f'=AVERAGE(B{indice}:B{indice + 19})+2*STDEV(B{indice}:B{indice + 19})')
        indice += 1

    # criando a planilha grafica
    planilha_grafica = gerenciador_planilha.criar_planilha(nome_planilha="grafico")

    # criando o cabeçalho do gráfico
    gerenciador_planilha.mescla_cells(cell_ini="A1", cell_fim="T2")

    gerenciador_planilha.aplica_estilo(
        celula="A1",
        estilos=[
            ("font", Font(b=True, sz=18, color="FFFFFF")),
            ("alignment", Alignment(horizontal="center", vertical="center")),
            ("fill", PatternFill(start_color="07838f", end_color="07838f", fill_type="solid"))
        ]
    )

    gerenciador_planilha.atualiza_cell(cell="A1", valor=f"historico de cotacoes")

    referencia_datas = Reference(planilha_dados, min_col=1, min_row=2, max_col=1, max_row=indice)
    referencia_cotacoes = Reference(planilha_dados, min_col=2, min_row=2, max_col=4, max_row=indice)

    # criando o gráfico
    gerenciador_planilha.adiciona_grafico(
        cell='A3',
        comprimento=33.87,
        altura=14.82,
        titulo=f"Cotacoes - {acao}",
        titu_x="Data",
        titu_y="valor",
        referencia_x=referencia_cotacoes,
        referencia_y=referencia_datas,
        propriedades_graf=[
            PropriedadesGrafico(grossura=0, cor="0a55ab"),
            PropriedadesGrafico(grossura=0, cor="a61508"),
            PropriedadesGrafico(grossura=0, cor="12a154")
        ]
    )

    gerenciador_planilha.salva_planilha(nome_arquivo=f"{acao}")

except FileNotFoundError:
    print("Arquivo não encontrado")

except ValueError:
    print("Arquivo com formato incorreto")

except Exception as e:
    print(f"Erro desconhecido: {e}")
