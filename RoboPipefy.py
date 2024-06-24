import webbrowser
import urllib.parse
import pyautogui
import time
import os
import shutil
import datetime

class RoboPipefy:
    @staticmethod
    def abrir_navegador_e_pesquisar(consulta):
        # Codificar a consulta de pesquisa para uma URL
        url = f"https://app.pipefy.com{consulta}"
        # Abrir o navegador padrão e navegar para a URL de pesquisa
        webbrowser.open(url)
        print(f"Navegador aberto e pesquisando por: {consulta}")
        time.sleep(5)  # Ajuste conforme necessário

    @staticmethod
    def clicar_elemento(imagem, descricao, imagem_alternativa=None):
        try:
            # Tenta encontrar a localização do elemento na tela
            location = pyautogui.locateCenterOnScreen(imagem, confidence=0.9)
            if not location and imagem_alternativa:
                location = pyautogui.locateCenterOnScreen(imagem_alternativa, confidence=0.9)
            
            if location:
                pyautogui.click(location)
                print(f"{descricao} encontrado(a) e clicado(a) com sucesso.")
            else:
                raise Exception(f"{descricao} não encontrado(a) na tela.")
        except Exception as e:
            print(f"Ocorreu um erro ao clicar em {descricao}: {e}")
            raise

    @staticmethod
    def exportar_relatorio():
        try:
            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Exportar.png', "Icone 'Exportar Excel'")
            time.sleep(5)  # Ajuste conforme necessário

            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Download.png', "Icone 'Download Excel'")
            time.sleep(10)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Ocorreu um erro ao exportar o relatório: {e}")
            raise

    @staticmethod
    def mover_arquivo_com_data(nome_base, pasta_origem, pasta_destino):
        try:
            # Obter a data atual no formato desejado
            data_atual = datetime.datetime.now().strftime('%d-%m-%Y')
            # Construir o nome completo do arquivo
            nome_arquivo = f"{nome_base}{data_atual}.xlsx"
            
            # Caminho completo do arquivo de origem
            caminho_arquivo_origem = os.path.join(pasta_origem, nome_arquivo)
            
            # Caminho completo do arquivo de destino
            caminho_arquivo_destino = os.path.join(pasta_destino, nome_arquivo)
            
            # Verifica se o arquivo de origem existe
            if os.path.exists(caminho_arquivo_origem):
                # Move o arquivo para o destino
                shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)
                print(f"Arquivo '{nome_arquivo}' movido com sucesso para: {caminho_arquivo_destino}")
            else:
                print(f"O arquivo '{nome_arquivo}' não existe em '{pasta_origem}'.")
        except Exception as e:
            print(f"Ocorreu um erro ao mover o arquivo: {e}")

    @staticmethod
    def logoff_pipefy():
        try:
            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Sair_Pipefy/Relatorios.png', "Icone 'Relatorios'")
            time.sleep(5)  # Ajuste conforme necessário

            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Sair_Pipefy/Menu_Sair.png', "Icone 'Perfil'")
            time.sleep(5)  # Ajuste conforme necessário

            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Sair_Pipefy/Botao_Sair.png', "Botão 'Sair'")
            time.sleep(5)  # Ajuste conforme necessário

        except Exception as e:
            print(f"Erro no logoff Robo Pipefy: {e}")
            raise

    @staticmethod
    def login_pipefy(email_pipefy, senha_pipefy):
        try:
            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Login_Pipefy/Email.png', "Campo de email")
            pyautogui.write(email_pipefy)
            time.sleep(5)  # Ajuste conforme necessário

            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Login_Pipefy/Senha.png', "Campo de senha")
            pyautogui.write(senha_pipefy)
            time.sleep(5)  # Ajuste conforme necessário

            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Login_Pipefy/LogIn.png', "Botão de login")
            time.sleep(5)  # Ajuste conforme necessário
        except Exception as e:
            print(f"Erro no login Robo Pipefy: {e}")
            raise

    @staticmethod
    def main():
        try:
            consulta = "Consulta Aqui"
            RoboPipefy.abrir_navegador_e_pesquisar(consulta)
            email_pipefy = ('victor.vinicius@supper.com.br')  # Certifique-se de configurar essa variável de ambiente
            senha_pipefy = ('Coxinha110103')  # Certifique-se de configurar essa variável de ambiente
            RoboPipefy.login_pipefy(email_pipefy, senha_pipefy)
            RoboPipefy.exportar_relatorio()
            pasta_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
            nome_base_arquivo = 'sequenciamento_'
            pasta_destino = r'C:\Users\victor.vinicius\Desktop\Automacao\Testes\sequenciamentoPipefy'
            RoboPipefy.mover_arquivo_com_data(nome_base_arquivo, pasta_downloads, pasta_destino)
            RoboPipefy.logoff_pipefy()
        except Exception as e:
            print(f"Erro no main Robo Pipefy: {e}")
            raise

if __name__ == "__main__":
    RoboPipefy.main()
