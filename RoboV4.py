import subprocess
import pyautogui
import time
from utils import Utils
import tkinter as tk
from tkinter import messagebox
import os
from datetime import datetime

class RoboV4:

    now = datetime.now()

    nomeArquivo_EstoqueSupper = 'EstoqueSupper'
    nomeArquivo_FaturamentoPeso = 'FaturamentoPeso'
    nomeArquivo_PedidoCompra = 'PedidoCompra'+now.strftime('%Y-%m')
    nomeArquivo_Req_Almoxarifado = 'Req_Almoxarifado'
    nomeArquivo_Mov_Saida_ODF = 'Mov_Saida_ODF'+now.strftime('%Y-%m')
    nomeArquivo_MRP_Filial_2 = 'CarteiraProducao_Filial_2'
    nomeArquivo_MRP_JMS_1 = 'CarteiraProducao_JMS_1'
    nomeArquivo_MRP_JM_3 = 'CarteiraProducao_JM_3'
    nomeArquivo_Parametros = 'Parametros'+Utils.data_atual.strftime('%d%m%Y')

    @staticmethod
    def Widgets_RoboV4(parent):
        # Conteúdo da aba Configurações
        tk.Label(parent, text="CONFIGURAÇÕES DA APLICAÇÃO").pack(pady=5)

        # Entrada para o caminho do software
        tk.Label(parent, text="Informe o Caminho do Sistema:").pack(pady=5)
        parent.software_path_entry = tk.Entry(parent, width=50)
        parent.software_path_entry.pack(pady=5)

        # Entrada para os argumentos
        tk.Label(parent, text="Argumentos para o executável:").pack(pady=5)
        parent.arguments_entry = tk.Entry(parent, width=50)
        parent.arguments_entry.pack(pady=5)

        # Botão para salvar o caminho do software
        save_button = tk.Button(parent, text="Salvar Caminho e Argumentos", command=lambda: (RoboV4.salvar_caminho_e_argumentos(parent), Utils.salvar_caminho_relatorios(parent)))
        save_button.pack(pady=10)

        try:
            with open('configuracoes_software.txt', 'r') as f:
                lines = f.readlines()
                if len(lines) >= 2:
                    parent.software_path_entry.insert(0, lines[0].strip())
                    parent.arguments_entry.insert(0, lines[1].strip())
        except FileNotFoundError:
            messagebox.showwarning("Aviso", "Arquivo de configurações não encontrado.")

    @staticmethod
    def salvar_caminho_e_argumentos(parent):
        software_path = parent.software_path_entry.get()
        arguments = parent.arguments_entry.get()

        # Validar se o caminho informado é válido
        if not os.path.exists(software_path):
            messagebox.showerror("Erro", "O caminho especificado não é válido.")
            return

        # Salvar o caminho e os argumentos em um arquivo de texto
        try:
            with open('configuracoes_software.txt', 'w') as f:
                f.write(f"{software_path}\n")
                f.write(arguments)
            messagebox.showinfo("Sucesso", "Caminho do software e argumentos salvos com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Erro ao salvar as configurações: {str(e)}")

    @staticmethod
    def abrir_software():
        # Ler o caminho e os argumentos do arquivo de texto
        try:
            with open('configuracoes_software.txt', 'r') as f:
                lines = f.readlines()
                if len(lines) >= 2:
                    software_path = lines[0].strip()
                    arguments = lines[1].strip().split(',')  # Supondo que os argumentos estejam separados por vírgula ou outro separador
                else:
                    messagebox.showerror("Erro", "Arquivo de configurações incompleto.")
                    return

            # Argumentos para o executável
            full_arguments = [software_path] + arguments

            subprocess.Popen(full_arguments, creationflags=subprocess.CREATE_NO_WINDOW)
            print(f"O software {software_path} foi aberto com sucesso com os argumentos fornecidos: {arguments}")
            time.sleep(25)  # Aguardar o software abrir completamente
        except Exception as e:
            print(f"Ocorreu um erro ao abrir o software: {e}")
            messagebox.showerror("Erro", f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao abrir o software: {e}")

    @staticmethod
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
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao realizar login: {e}")
            raise

    @staticmethod
    def abrir_cubo_de_decisao():
        try:
            Utils.clicar_elemento('img/Menu_Relatorio/Relatorios.png', "Palavra 'Relatórios'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Menu_Relatorio/Cubo.png', "Palavra 'Cubo de decisão'")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao abrir o Cubo de Decisão: {e}")
            raise

    @staticmethod
    def relatorio_faturamento_peso():
        try:
            Utils.clicar_elemento('img/Rel_FaturamentoPeso/FaturamentoPeso.png', "Icone 'Faturamento Peso'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.aplicar_data_mes_atual('img/Rel_FaturamentoPeso/DataInicial.png', 'img/Rel_FaturamentoPeso/DataFinal.png')
            Utils.apicar_modelo_de_relatorio('img/Rel_FaturamentoPeso/FaturamentoPeso_Lista.png')
            Utils.salvar_arquivo(RoboV4.nomeArquivo_FaturamentoPeso)

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Faturamento Peso: {e}")
            raise

    @staticmethod
    def relatorio_estoque_supper():
        try:
            Utils.clicar_elemento('img/Rel_EstoqueSupper/EstoqueSupper.png', "Icone 'Estoque Supper'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_EstoqueSupper/EstoqueSupper_Lista.png')
            Utils.salvar_arquivo(RoboV4.nomeArquivo_EstoqueSupper)
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Estoque Supper: {e}")
            raise

    @staticmethod
    def relatorio_pedido_compra():
        try:
            Utils.clicar_elemento('img/Rel_PedidoCompra/PedidoCompra.png', "Icone 'Pedido Compra'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_PedidoCompra/PedidoCompra_Lista.png')
            Utils.salvar_arquivo_pedido_compra(RoboV4.nomeArquivo_PedidoCompra)
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Pedido Compra: {e}")
            raise

    @staticmethod
    def abrir_mrp():
        try:
            Utils.clicar_elemento('img/Menu_Logistica/Logistica.png', "Palavra 'Logística'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Menu_Logistica/MRP.png', "Palavra 'MRP'")
            time.sleep(10)  # Ajuste conforme necessário


            # Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta'")
            # time.sleep(3)  # Ajuste conforme necessário
            # Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Supper_Filial_2.png', "Botão 'Empresa Supper Filial 2'")
            # time.sleep(10)  # Ajuste conforme necessário

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

            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Seta.png', "Botão 'Seta'")
            time.sleep(3)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Rel_MRP/Trocar_Empresa/Supper_Filial_2.png', "Botão 'Empresa Supper Filial 2'")
            time.sleep(3)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao abrir o MRP: {e}")
            raise

    @staticmethod
    def relatorio_mrp(nomearquivo_MRP):
        try:
            Utils.salvar_arquivo_mrp(nomearquivo_MRP)
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório MRP: {e}")
            raise

    @staticmethod
    def relatorio_Req_Almoxarifado():
        try:
            Utils.clicar_elemento('img/Rel_Req_Almoxarifado/Req_Almoxarifado.png', "Icone 'Req. Almoxarifado'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.apicar_modelo_de_relatorio('img/Rel_Req_Almoxarifado/Req_Almoxarifado_Lista.png')

            Utils.clicar_elemento('img/Rel_Req_Almoxarifado/Produto.png', "Botão 'Produto'")
            time.sleep(2)
            # Utils.expand_All_Notes('img/Rel_Req_Almoxarifado/Expand_Notes/Ano_Entrega.png', 'img/Rel_Req_Almoxarifado/Expand_Notes/Mes_Entrega.png')

            Utils.salvar_arquivo(RoboV4.nomeArquivo_Req_Almoxarifado)
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Estoque Supper: {e}")
            raise

    @staticmethod
    def relatorio_Mov_Saida_ODF():
        try:
            Utils.clicar_elemento('img/Rel_Mov_Saida_ODF/Mov_Saida_ODF.png', "Icone 'Mov. Saída ODF'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.aplicar_data_mes_atual('img/Rel_FaturamentoPeso/DataInicial.png', 'img/Rel_FaturamentoPeso/DataFinal.png')

            Utils.apicar_modelo_de_relatorio('img/Rel_Mov_Saida_ODF/Mov_Saida_ODF_Lista.png')

            Utils.clicar_elemento('img/Rel_Mov_Saida_ODF/Desc_Prod.png', "Botão 'Desc_Prod'")
            time.sleep(2)
            Utils.expand_All_Notes('img/Rel_Mov_Saida_ODF/Expand_Notes/Cod_Prod.png')

            Utils.salvar_arquivo_pedido_compra(RoboV4.nomeArquivo_Mov_Saida_ODF)
            time.sleep(5)

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Movimento Saida de ODF: {e}")
            raise

    @staticmethod
    def abrir_Parametros():
        try:
            Utils.clicar_elemento('img/Menu_Configuracoes/Configuracoes.png', "Palavra 'Configurações'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Menu_Configuracoes/Parametros.png', "Palavra 'Parâmetros'")
            time.sleep(5)  # Ajuste conforme necessário

            RoboV4.relatorio_Parametros()
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao abrir os Parâmetros: {e}")
            raise

    @staticmethod
    def relatorio_Parametros():
        try:

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
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Parametros: {e}")
            raise

    @staticmethod
    def fechar_sistema():
        try:
            time.sleep(5)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Fechar_Sistema/X.png', "Botão 'X - Fechar'")
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao gerar o relatório de Parametros: {e}")
            raise