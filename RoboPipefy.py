import webbrowser
import pyautogui
import time
import os
import shutil
import datetime
import psutil
import sys

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
    def fechar_todos_navegadores():
        navegadores_comuns = ["chrome", "firefox", "msedge", "opera", "safari", "brave", "vivaldi"]
        
        try:
            for processo in psutil.process_iter(['pid', 'name']):
                for navegador_nome in navegadores_comuns:
                    if navegador_nome.lower() in processo.info['name'].lower():
                        p = psutil.Process(processo.info['pid'])
                        p.terminate()  # ou p.kill() para forçar o encerramento
                        print(f"Navegador {processo.info['name']} encerrado com sucesso.")
            
            print("Todos os navegadores conhecidos foram encerrados.")
        except Exception as e:
            print(f"Ocorreu um erro ao tentar fechar os navegadores: {e}")
            sys.exit(1)

    @staticmethod
    def clicar_elemento(imagem, descricao, tentativas=3, confidence=0.9):
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
                if tentativa == tentativas - 1:  # Se for a última tentativa, encerrar o processo
                    time.sleep(2)  # Esperar antes da próxima tentativa
                    raise

    @staticmethod
    def exportar_relatorio():
        try:
            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Exportar.png', "Icone 'Exportar Excel'")
            time.sleep(5)  # Ajuste conforme necessário

            RoboPipefy.clicar_elemento('img/Rel_Pipefy/Download.png', "Icone 'Download Excel'")
            time.sleep(30)  # Ajuste conforme necessário

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
            time.sleep(10)  # Ajuste conforme necessário

            RoboPipefy.fechar_todos_navegadores()

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
            consulta = "/pipes/790423/reports_v2/300276569"
            RoboPipefy.abrir_navegador_e_pesquisar(consulta)
            email_pipefy = ('bot@supper.com.br')  # Certifique-se de configurar essa variável de ambiente
            senha_pipefy = ('2115FZ7y')  # Certifique-se de configurar essa variável de ambiente
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
