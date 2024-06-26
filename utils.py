import pyautogui
import time
from datetime import datetime
import calendar
import tkinter as tk
from tkinter import messagebox

class Utils:

    current_date = datetime.now()
    first_day_of_month = current_date.replace(day=1).strftime('%d/%m/%Y')
    last_day_of_month = current_date.replace(day=calendar.monthrange(current_date.year, current_date.month)[1]).strftime('%d/%m/%Y')

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

        try:
            with open('configuracoes_caminhos.txt', 'r') as f:
                lines = f.readlines()
                if len(lines) >= 2:
                    parent.caminho_entry.insert(0, lines[0].strip())
                    parent.caminho_parametros_entry.insert(0, lines[1].strip())
        except FileNotFoundError:
            messagebox.showwarning("Aviso", "Arquivo de configurações não encontrado.")

    @staticmethod
    def salvar_caminho_e_parametros(parent):
        caminho = parent.caminho_entry.get()
        caminho_parametros = parent.caminho_parametros_entry.get()

        # Salvar o caminho e os argumentos em um arquivo de texto
        try:
            with open('configuracoes_caminhos.txt', 'w') as f:
                f.write(f"{caminho}\n")
                f.write(caminho_parametros)
            messagebox.showinfo("Sucesso", "Caminho do software e argumentos salvos com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar as configurações: {str(e)}")

    @staticmethod
    def clicar_elemento(imagem, descricao, imagem_alternativa=None, tentativas=3, confidence=0.9):
        for tentativa in range(tentativas):
            try:
                location = pyautogui.locateCenterOnScreen(imagem, confidence=confidence)
                
                if location:
                    pyautogui.click(location)
                    print(f"{descricao} encontrado(a) e clicado(a) com sucesso.")
                    return
                else:
                    raise Exception(f"{descricao} não encontrado(a) na tela. Tentativa {tentativa + 1} de {tentativas}")
            except Exception as e:
                print(f"Ocorreu um erro ao clicar em {descricao}: {e}")
            time.sleep(2)  # Esperar antes da próxima tentativa
        Utils.procurar_imagem_alternativa(imagem_alternativa)

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
            print(f"Ocorreu um erro ao aplicar o modelo de relatório: {e}")
            raise
    
    @staticmethod
    def salvar_arquivo(nome_arquivo):
        try:
            # Ler o caminho do arquivo de configuração
            with open('configuracoes_caminhos.txt', 'r') as f:
                caminho = f.readline().strip()  # Lê apenas a primeira linha do arquivo

            # Simulação de operações de salvar arquivo usando pyautogui (adapte conforme necessário)
            # pyautogui.click('img/Salvar_Arquivo/ExportarExcel.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/ExportarExcel.png', "Icone 'Exportar Excel'", 'img/Salvar_Arquivo/ExportarExcel_v2.png')
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador.png', "Icone 'Este Computador'", 'img/Salvar_Arquivo/AreaDeTrabalho.png')
            time.sleep(5)
            
            # pyautogui.click('img/Salvar_Arquivo/EsteComputador_Pesquisa.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador_Pesquisa.png', "Icone 'Este Computador Pesquisa'", 'img/Salvar_Arquivo/AreaDeTrabalho_Pesquisa.png')
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
            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoSalvar.png', "Icone 'Botão Salvar'", 'img/Salvar_Arquivo/BotaoSalvar_v2.png')
            time.sleep(5)

            # pyautogui.click('img/Salvar_Arquivo/NaoAbrirArquivo.png')
            Utils.clicar_elemento('img/Salvar_Arquivo/NaoAbrirArquivo.png', "Botão 'Não/Ok'", 'img/Salvar_Arquivo/BotaoOK_v2.png' )
            time.sleep(5)

        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de configurações não encontrado.")
        except Exception as e:
            print(f"Ocorreu um erro ao salvar o arquivo: {e}")

    @staticmethod
    def procurar_imagem_alternativa(imagem_alternativa):
        try:
            Utils.clicar_elemento(imagem_alternativa, "Imagem Alternativa para dar OK ao salvar o arquivo")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"Ocorreu um erro ao procurar a imagem alternativa: {e}")
            raise

    @staticmethod
    def aplicar_data_mes_atual(data_inicial_img, data_final_img):
        try:
            Utils.clicar_elemento(data_inicial_img, "Campo 'Data Inicial'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(Utils.first_day_of_month)

            Utils.clicar_elemento(data_final_img, "Campo 'Data Final'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(Utils.last_day_of_month)

            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoOk.png', "Botão 'Ok'")
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao aplicar a data do mês atual: {e}")
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

            Utils.clicar_elemento('img/Salvar_Arquivo/SalvarParametros.png', "Icone 'Botão Salvar'", 'img/Salvar_Arquivo/BotaoSalvar_v2.png')
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"Ocorreu um erro ao salvar o arquivo: {e}")
            Utils.procurar_imagem_alternativa("img/Salvar_Arquivo/Imagem_Alternativa_OK.png")
