import pyautogui
import time
from datetime import datetime, timedelta
import calendar
import tkinter as tk
from tkinter import messagebox
import sys

class Utils:

    #achar o dia do mês atual
    data_atual = datetime.now()
    primeiro_dia_mes_atual = data_atual.replace(day=1).strftime('%d/%m/%Y')
    ultimo_dia_mes_atual = data_atual.replace(day=calendar.monthrange(data_atual.year, data_atual.month)[1]).strftime('%d/%m/%Y')

    #achar o dia do mês anterior
    primeiro_dia_mes_atual_data = data_atual.replace(day=1)
    ultimo_dia_mes_anterior_data = primeiro_dia_mes_atual_data - timedelta(days=1)
    primeiro_dia_mes_anterior_data = ultimo_dia_mes_anterior_data.replace(day=1)
    
    primeiro_dia_mes_anterior = primeiro_dia_mes_anterior_data.strftime('%d/%m/%Y')
    ultimo_dia_mes_anterior = ultimo_dia_mes_anterior_data.strftime('%d/%m/%Y')

    @staticmethod
    def Widgets_Utils(parent):

        # Entrada para o caminho do software
        tk.Label(parent, text="informe o caminho para salvar os relatorios:").pack(pady=5)
        parent.caminho_entry = tk.Entry(parent, width=50)
        parent.caminho_entry.pack(pady=5)

        # Entrada para os argumentos
        tk.Label(parent, text="Informe o caminho para salvar os parametros:").pack(pady=5)
        parent.caminho_parametros_entry = tk.Entry(parent, width=50)
        parent.caminho_parametros_entry.pack(pady=5)

         # Entrada para os argumentos
        tk.Label(parent, text="Informe o caminho para salvar o Relatorio Pedido de Compra:").pack(pady=5)
        parent.caminho_pedido_compra_entry = tk.Entry(parent, width=50)
        parent.caminho_pedido_compra_entry.pack(pady=5)

         # Entrada para os argumentos
        tk.Label(parent, text="Informe o caminho para salvar o Relatorio Movimento Saida de ODF:").pack(pady=5)
        parent.caminho_mov_odf_entry = tk.Entry(parent, width=50)
        parent.caminho_mov_odf_entry.pack(pady=5)

        

        try:
            with open('configuracoes_caminhos.txt', 'r') as f:
                lines = f.readlines()
                if len(lines) >= 4:
                    parent.caminho_entry.insert(0, lines[0].strip())
                    parent.caminho_parametros_entry.insert(0, lines[1].strip())
                    parent.caminho_pedido_compra_entry.insert(0, lines[2].strip())
                    parent.caminho_mov_odf_entry.insert(0, lines[3].strip())

        except FileNotFoundError:
            messagebox.showwarning("Aviso", "Arquivo de configurações não encontrado.")

    @staticmethod
    def salvar_caminho_relatorios(parent):
        caminho = parent.caminho_entry.get()
        caminho_parametros = parent.caminho_parametros_entry.get()
        caminho_pedido_compra = parent.caminho_pedido_compra_entry.get()
        caminho_mov_odf = parent.caminho_mov_odf_entry.get()

        # Salvar o caminho e os argumentos em um arquivo de texto
        try:
            with open('configuracoes_caminhos.txt', 'w') as f:
                f.write(f"{caminho}\n")
                f.write(f"{caminho_parametros}\n")
                f.write(f"{caminho_pedido_compra}\n")
                f.write(caminho_mov_odf)
            messagebox.showinfo("Sucesso", "Caminho do software, argumentos e relatorios salvos com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Erro ao salvar as configurações: {str(e)}")

    @staticmethod
    def clicar_elemento(imagem, descricao, tentativas=3, confidence=0.9):
        for tentativa in range(tentativas):
            try:
                location = pyautogui.locateCenterOnScreen(imagem, confidence=confidence)
                
                if location:
                    pyautogui.click(location)
                    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {descricao} encontrado(a) e clicado(a) com sucesso.")
                    return
                else:
                    raise Exception(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {descricao} não encontrado(a) na tela. Tentativa {tentativa + 1} de {tentativas}")
            
            except Exception as e:
                print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao clicar em {descricao}: {e}")
                if tentativa == tentativas - 1:  # Se for a última tentativa, encerrar o processo
                    time.sleep(2)  # Esperar antes da próxima tentativa
                    raise

    @staticmethod
    def apicar_modelo_de_relatorio(caminho_imagem_lista):
        try:
            Utils.clicar_elemento('img/Modelo_Relatorio/SetaModelos.png', "Icone 'Seta Modelos'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Modelo_Relatorio/SetaLista.png', "Icone 'Seta Lista'")
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento(caminho_imagem_lista, "Modelo de Relatório, caminho imagem lista")
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao aplicar o modelo de relatório: {e}")
            raise
    
    @staticmethod
    def salvar_arquivo(nome_arquivo):
        try:
            # Ler o caminho do arquivo de configuração
            with open('configuracoes_caminhos.txt', 'r') as f:
                caminho = f.readline().strip()  # Lê apenas a primeira linha do arquivo

            # Simulação de operações de salvar arquivo usando pyautogui (adapte conforme necessário)
            # pyautogui.click('img/Salvar_Arquivo/ExportarExcel.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/ExportarExcel.png', "Icone 'Exportar Excel'")
            time.sleep(15)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho.png', "Icone 'Area de Trabalho'")
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador_Pesquisa.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)
            pyautogui.write(caminho)
            pyautogui.press('enter')
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NomeDocumento.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NomeDocumento.png', "Elemento 'Nome do Documento'")
            time.sleep(2)
            pyautogui.write(nome_arquivo)
            time.sleep(2)

            # pyautogui.click('img/Salvar_Arquivo/BotaoSalvar.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoSalvar.png', "Icone 'Botão Salvar'")
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NaoAbrirArquivo.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NaoAbrirArquivo.png', "Botão 'Não/Ok'")
            time.sleep(5)

        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de configurações não encontrado.")
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao salvar o arquivo: {e}")
    
    @staticmethod
    def salvar_arquivo_mrp(nome_arquivo):
        try:
            # Ler o caminho do arquivo de configuração
            with open('configuracoes_caminhos.txt', 'r') as f:
                caminho = f.readline().strip()  # Lê apenas a primeira linha do arquivo

            # Simulação de operações de salvar arquivo usando pyautogui (adapte conforme necessário)
            # pyautogui.click('img/Salvar_Arquivo/ExportarExcel.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/ExportarExcel_v2.png', "Icone 'Exportar Excel'")
            time.sleep(10)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho.png', "Icone 'Este Computador'")
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador_Pesquisa.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)
            pyautogui.write(caminho)
            pyautogui.press('enter')
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NomeDocumento.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NomeDocumento.png', "Elemento 'Nome do Documento'")
            time.sleep(2)
            pyautogui.write(nome_arquivo)
            time.sleep(2)

            # pyautogui.click('img/Salvar_Arquivo/BotaoSalvar.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoSalvar_v2.png', "Icone 'Botão Salvar'")
            time.sleep(15)

            # pyautogui.click('img/Salvar_Arquivo/NaoAbrirArquivo.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoOK_v2.png', "Botão 'Não/Ok'")
            time.sleep(5)

        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de configurações não encontrado.")
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao salvar o arquivo: {e}")

    # # @staticmethod
    # # def procurar_imagem_alternativa(imagem_alternativa):
    #     try:
    #         Utils.clicar_elemento(imagem_alternativa, "Imagem Alternativa para dar OK ao salvar o arquivo")
    #         time.sleep(5)  # Ajuste conforme necessário
    #     except Exception as e:
    #         print(f"Ocorreu um erro ao procurar a imagem alternativa: {e}")
    #         sys.exit(1)

    @staticmethod
    def aplicar_data_mes_atual(data_inicial_img, data_final_img):
        try:
            Utils.clicar_elemento(data_inicial_img, "Campo 'Data Inicial'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(Utils.primeiro_dia_mes_atual)

            Utils.clicar_elemento(data_final_img, "Campo 'Data Final'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(Utils.ultimo_dia_mes_atual)

            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoOk.png', "Botão 'Ok'")
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao aplicar a data do mês atual: {e}")
            raise

    @staticmethod
    def expand_All_Notes(coluna1, coluna2=None, coluna3=None):
        try:
            Utils.clicar_elemento(coluna1, "Primeira Coluna")
            pyautogui.rightClick()
            time.sleep(5)  # Ajuste conforme necessário
            Utils.clicar_elemento('img/Expand_All_Notes/Expand_All_Notes.png', "Botão 'Expand All Notes'")
            time.sleep(5)  # Ajuste conforme necessário

            if coluna2:
                Utils.clicar_elemento(coluna2, "Segunda Coluna")
                pyautogui.rightClick()
                time.sleep(5)  # Ajuste conforme necessário
                Utils.clicar_elemento('img/Expand_All_Notes/Expand_All_Notes.png', "Botão 'Expand All Notes'")
                time.sleep(5)  # Ajuste conforme necessário

            if coluna3:
                Utils.clicar_elemento(coluna3, "Terceira Coluna")
                pyautogui.rightClick()
                time.sleep(5)  # Ajuste conforme necessário
                Utils.clicar_elemento('img/Expand_All_Notes/Expand_All_Notes.png', "Botão 'Expand All Notes'")
                time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao expandir todas as notas: {e}")
            raise

    @staticmethod
    def salvar_arquivo_parametros(nome_arquivo, exportar_excel_img):
        try:

            with open('configuracoes_caminhos.txt', 'r') as f:
                segunda_linha = f.readlines()
                if len(segunda_linha) >= 2:
                    caminho_parametros = segunda_linha[1].strip() 
                else:
                    # Caso o arquivo tenha menos de duas linhas
                    print("O arquivo não tem pelo menos duas linhas.")

            Utils.clicar_elemento(exportar_excel_img, "Icone 'Exportar Excel'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador.png', "Icone 'Este Computador'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(caminho_parametros)
            pyautogui.press('enter')
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Salvar_Arquivo/NomeDocumento.png', "Elemento 'Nome do Documento'")
            time.sleep(2)  # Ajuste conforme necessário

            pyautogui.write(nome_arquivo)
            time.sleep(2)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Salvar_Arquivo/SalvarParametros.png', "Icone 'Botão Salvar'")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao salvar o arquivo: {e}")
            raise

    @staticmethod
    def salvar_arquivo_pedido_compra(nome_arquivo):
    
        try:
            # Ler o caminho do arquivo de configuração
            with open('configuracoes_caminhos.txt', 'r') as f:
                terceira_linha = f.readlines()
                if len(terceira_linha) >= 3:
                    caminho_pedido_compra = terceira_linha[2].strip() 
                else:
                    # Caso o arquivo tenha menos de duas linhas
                    print("O arquivo não tem pelo menos tres linhas.")

            # Simulação de operações de salvar arquivo usando pyautogui (adapte conforme necessário)
            # pyautogui.click('img/Salvar_Arquivo/ExportarExcel.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/ExportarExcel.png', "Icone 'Exportar Excel'")
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho.png', "Icone 'Este Computador'")
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador_Pesquisa.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)
            pyautogui.write(caminho_pedido_compra)
            pyautogui.press('enter')
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NomeDocumento.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NomeDocumento.png', "Elemento 'Nome do Documento'")
            time.sleep(2)
            pyautogui.write(nome_arquivo)
            time.sleep(2)

            # pyautogui.click('img/Salvar_Arquivo/BotaoSalvar.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoSalvar.png', "Icone 'Botão Salvar'")
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NaoAbrirArquivo.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NaoAbrirArquivo.png', "Botão 'Não/Ok'")
            time.sleep(5)

        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de configurações não encontrado.")
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao salvar o arquivo: {e}")

    @staticmethod
    def salvar_arquivo_mov_saida_odf(nome_arquivo):
        try:
            # Ler o caminho do arquivo de configuração
            with open('configuracoes_caminhos.txt', 'r') as f:
                quarta_linha = f.readlines()
                if len(quarta_linha) >= 4:
                    caminho_mov_saida_odf = quarta_linha[3].strip() 
                else:
                    # Caso o arquivo tenha menos de duas linhas
                    print("O arquivo não tem pelo menos tres linhas.")

            # Simulação de operações de salvar arquivo usando pyautogui (adapte conforme necessário)
            # pyautogui.click('img/Salvar_Arquivo/ExportarExcel.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/ExportarExcel.png', "Icone 'Exportar Excel'")
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho.png', "Icone 'Este Computador'")
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador_Pesquisa.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/AreaDeTrabalho_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)
            pyautogui.write(caminho_mov_saida_odf)
            pyautogui.press('enter')
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NomeDocumento.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NomeDocumento.png', "Elemento 'Nome do Documento'")
            time.sleep(2)
            pyautogui.write(nome_arquivo)
            time.sleep(2)

            # pyautogui.click('img/Salvar_Arquivo/BotaoSalvar.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoSalvar.png', "Icone 'Botão Salvar'")
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NaoAbrirArquivo.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NaoAbrirArquivo.png', "Botão 'Não/Ok'")
            time.sleep(5)

        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de configurações não encontrado.")
        except Exception as e:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Ocorreu um erro ao salvar o arquivo: {e}")

