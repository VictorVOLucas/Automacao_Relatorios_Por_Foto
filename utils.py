import pyautogui
import time
from datetime import datetime
import calendar

class Utils:
    caminho = r'K:\Users\victor.vinicius\Desktop\Automacao\Testes'

    caminho_parametros = r'K:\Users\victor.vinicius\Documents\GoogleDrive\Parametros'

    current_date = datetime.now()
    first_day_of_month = current_date.replace(day=1).strftime('%d/%m/%Y')
    last_day_of_month = current_date.replace(day=calendar.monthrange(current_date.year, current_date.month)[1]).strftime('%d/%m/%Y')

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

    def salvar_arquivo(nome_arquivo, exportar_excel_img='img/Salvar_Arquivo/ExportarExcel.png'):
        try:
            Utils.clicar_elemento(exportar_excel_img, "Icone 'Exportar Excel'", 'img/Salvar_Arquivo/ExportarExcel_v2.png')
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador.png', "Icone 'Este Computador'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(Utils.caminho)
            pyautogui.press('enter')
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Salvar_Arquivo/NomeDocumento.png', "Elemento 'Nome do Documento'")
            time.sleep(2)  # Ajuste conforme necessário

            pyautogui.write(nome_arquivo)
            time.sleep(2)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Salvar_Arquivo/BotaoSalvar.png', "Icone 'Botão Salvar'", 'img/Salvar_Arquivo/BotaoSalvar_v2.png')
            time.sleep(5)  # Ajuste conforme necessário

            Utils.clicar_elemento('img/Salvar_Arquivo/NaoAbrirArquivo.png', "Botão 'Não/Ok'", 'img/Salvar_Arquivo/Imagem_Alternativa_OK.png')
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao salvar o arquivo: {e}")
            Utils.procurar_imagem_alternativa("img/Salvar_Arquivo/Imagem_Alternativa_OK.png")

    def procurar_imagem_alternativa(imagem_alternativa):
        try:
            Utils.clicar_elemento(imagem_alternativa, "Imagem Alternativa para dar OK ao salvar o arquivo")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"Ocorreu um erro ao procurar a imagem alternativa: {e}")
            raise

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

    def salvar_arquivo_parametros(nome_arquivo, exportar_excel_img):
        try:
            Utils.clicar_elemento(exportar_excel_img, "Icone 'Exportar Excel'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador.png', "Icone 'Este Computador'")
            time.sleep(5)  # Ajuste conforme necessário
            
            Utils.clicar_elemento('img/Salvar_Arquivo/EsteComputador_Pesquisa.png', "Icone 'Este Computador Pesquisa'")
            time.sleep(2)  # Ajuste conforme necessário
            pyautogui.write(Utils.caminho_parametros)
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
