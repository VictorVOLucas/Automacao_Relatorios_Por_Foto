import subprocess
import pyautogui
import time
from utils import Utils

class RoboV4:
    software_path = r'C:\Program Files (x86)\GraphOn\GO-Global\Client\gg-client.exe'
        # Argumentos para o executável
    arguments = [
        software_path,
        '-c',
        '-h', 'erpcloud.korp.com.br',
        '-mm', '1',
        '-u', 'SUPPER_15',
        '-p', 'p#&jHPF8yvJ',
        '-a', 'SUPPER',
        '-ac', 'all'
    ]

    nomeArquivo_EstoqueSupper = 'EstoqueSupper'
    nomeArquivo_FaturamentoPeso = 'FaturamentoPeso'
    nomeArquivo_PedidoCompra = 'PedidoCompra'
    nomeArquivo_Req_Almoxarifado = 'Req_Almoxarifado'
    nomeArquivo_Mov_Saida_ODF = 'Mov_Saida_ODF'
    nomeArquivo_MRP_Filial_2 = 'CarteiraProducao_Filial_2'
    nomeArquivo_MRP_JMS_1 = 'CarteiraProducao_JMS_1'
    nomeArquivo_MRP_JM_3 = 'CarteiraProducao_JM_3'
    nomeArquivo_Parametros = 'Parametros'+Utils.current_date.strftime('%d%m%Y')

    def abrir_software():
        try:
            subprocess.Popen(RoboV4.arguments, creationflags=subprocess.CREATE_NO_WINDOW)
            print(f"O software {RoboV4.software_path} foi aberto com sucesso com os argumentos fornecidos.")
            time.sleep(15)  # Aguardar o software abrir completamente
        except Exception as e:
            print(f"Ocorreu um erro ao abrir o software: {e}")
            raise

    def realizar_login():
        try:
            Utils.clicar_elemento('img/Logar_Sistema/Usuario.png', "Campo 'Usuário'")
            pyautogui.write('BOT_Relatorio_Supper')
            time.sleep(2)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Logar_Sistema/Senha.png', "Campo 'Senha'")
            pyautogui.write('210624')
            time.sleep(2)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Logar_Sistema/BotaoEntrar.png', "Botão 'Entrar'")
            time.sleep(10)  # Aguardar a resposta do login

        except Exception as e:
            print(f"Ocorreu um erro ao realizar login: {e}")
            raise

    def abrir_cubo_de_decisao():
        try:
            Utils.clicar_elemento('img/Menu_Relatorio/Relatorios.png', "Palavra 'Relatórios'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Menu_Relatorio/Cubo.png', "Palavra 'Cubo de decisão'")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"Ocorreu um erro ao abrir o Cubo de Decisão: {e}")
            raise

    def relatorio_faturamento_peso():
        try:
            Utils.clicar_elemento('img/Rel_FaturamentoPeso/FaturamentoPeso.png', "Icone 'Faturamento Peso'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.aplicar_data_mes_atual('img/Rel_FaturamentoPeso/DataInicial.png', 'img/Rel_FaturamentoPeso/DataFinal.png')
            Utils.apicar_modelo_de_relatorio('img/Rel_FaturamentoPeso/FaturamentoPeso_Lista.png')
            Utils.salvar_arquivo(RoboV4.nomeArquivo_FaturamentoPeso)

        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório de Faturamento Peso: {e}")
            raise

    def relatorio_estoque_supper():
        try:
            Utils.clicar_elemento('img/Rel_EstoqueSupper/EstoqueSupper.png', "Icone 'Estoque Supper'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_EstoqueSupper/EstoqueSupper_Lista.png')
            Utils.salvar_arquivo(RoboV4.nomeArquivo_EstoqueSupper)
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório de Estoque Supper: {e}")
            raise

    def relatorio_pedido_compra():
        try:
            Utils.clicar_elemento('img/Rel_PedidoCompra/PedidoCompra.png', "Icone 'Pedido Compra'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_PedidoCompra/PedidoCompra_Lista.png')
            Utils.salvar_arquivo(RoboV4.nomeArquivo_PedidoCompra)
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório de Pedido Compra: {e}")
            raise

    def abrir_mrp():
        try:
            Utils.clicar_elemento('img/Menu_Logistica/Logistica.png', "Palavra 'Logística'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Menu_Logistica/MRP.png', "Palavra 'MRP'")
            time.sleep(10)  # Ajuste conforme necessário


            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta'")
            time.sleep(3)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Supper_Filial_2.png', "Botão 'Empresa Supper Filial 2'", 'img/Rel_MRP/Trocar_Empresa/Supper_Filial_2_Alternativa.png')
            time.sleep(10)  # Ajuste conforme necessário

            RoboV4.relatorio_mrp(RoboV4.nomeArquivo_MRP_Filial_2)

            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta'")
            time.sleep(3)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/JMS_1.png', "Botão 'JMS 1'")
            time.sleep(10)  # Ajuste conforme necessário

            RoboV4.relatorio_mrp(RoboV4.nomeArquivo_MRP_JMS_1)

            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta'")
            time.sleep(3)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/JM_3.png', "Botão 'JM 3'")
            time.sleep(10)  # Ajuste conforme necessário

            RoboV4.relatorio_mrp(RoboV4.nomeArquivo_MRP_JM_3)
        except Exception as e:
            print(f"Ocorreu um erro ao abrir o MRP: {e}")
            raise

    def relatorio_mrp(nomearquivo_MRP):
        try:
            # clicar_elemento('img/Rel_MRP/Carteira_de_Producao.png', "Botão 'Carteira de Produção'")
            # time.sleep(5)  # Ajuste conforme necessário

            Utils.salvar_arquivo(nomearquivo_MRP, 'img/Salvar_Arquivo/ExportarExcel_v2.png')
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório MRP: {e}")
            raise

    def relatorio_Req_Almoxarifado():
        try:
            Utils.clicar_elemento('img/Rel_Req_Almoxarifado/Req_Almoxarifado.png', "Icone 'Req. Almoxarifado'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_Req_Almoxarifado/Req_Almoxarifado_Lista.png')

            Utils.clicar_elemento('img/Rel_Req_Almoxarifado/Produto.png', "Botão 'Produto'")

            Utils.expand_All_Notes('img/Rel_Req_Almoxarifado/Expand_Notes/Ano_Entrega.png', 'img/Rel_Req_Almoxarifado/Expand_Notes/Mes_Entrega.png')

            Utils.salvar_arquivo(RoboV4.nomeArquivo_Req_Almoxarifado)
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório de Estoque Supper: {e}")
            raise

    def relatorio_Mov_Saida_ODF():
        try:
            Utils.clicar_elemento('img/Rel_Mov_Saida_ODF/Mov_Saida_ODF.png', "Icone 'Mov. Saída ODF'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.aplicar_data_mes_atual('img/Rel_FaturamentoPeso/DataInicial.png', 'img/Rel_FaturamentoPeso/DataFinal.png')

            Utils.apicar_modelo_de_relatorio('img/Rel_Mov_Saida_ODF/Mov_Saida_ODF_Lista.png')

            Utils.clicar_elemento('img/Rel_Mov_Saida_ODF/Desc_Prod.png', "Botão 'Desc_Prod'")

            Utils.expand_All_Notes('img/Rel_Mov_Saida_ODF/Expand_Notes/Ano.png', 'img/Rel_Mov_Saida_ODF/Expand_Notes/Mes.png', 'img/Rel_Mov_Saida_ODF/Expand_Notes/Cod_Prod.png')

            Utils.salvar_arquivo(RoboV4.nomeArquivo_Mov_Saida_ODF)

        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório de Movimento Saida de ODF: {e}")
            raise

    def abrir_Parametros():
        try:
            Utils.clicar_elemento('img/Menu_Configuracoes/Configuracoes.png', "Palavra 'Configurações'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Menu_Configuracoes/Parametros.png', "Palavra 'Parâmetros'")
            time.sleep(5)  # Ajuste conforme necessário

            RoboV4.relatorio_Parametros()
        except Exception as e:
            print(f"Ocorreu um erro ao abrir os Parâmetros: {e}")
            raise

    def relatorio_Parametros():
        try:
            # Utils.clicar_elemento('img/Rel_Parametros/Parametros.png', "Icone 'Parametros'")
            # time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/Filtro.png', "Botão 'Filtro'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/Excel.png', "Botão 'Excel'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/BotaoOK.png', "Botão 'Excel'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Rel_Parametros/SetaDupla.png', "Botão 'OK'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.salvar_arquivo_parametros(RoboV4.nomeArquivo_Parametros, 'img/Rel_Parametros/Excel.png')

        except Exception as e:
            print(f"Ocorreu um erro ao gerar o relatório de Parametros: {e}")
            raise