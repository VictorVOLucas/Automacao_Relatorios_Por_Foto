import tkinter as tk  # Módulo para criar interfaces gráficas
import tkinter.ttk as ttk  # Módulo para widgets estilizados
from tkinter import messagebox  # Módulo para exibir caixas de mensagem
import schedule  # Biblioteca para agendamento de tarefas
import time  # Módulo para manipulação de tempo
from threading import Thread  # Biblioteca para threads (execução paralela)
from datetime import datetime  # Biblioteca para manipulação de datas e horários
from RoboV4 import RoboV4  # Importa funcionalidades específicas de RoboV4
from utils import Utils  # Importa utilitários
#from RoboPipefy import RoboPipefy  # Importa funcionalidades específicas de RoboPipefy
import os  # Módulo para interagir com o sistema operacional
import sys  # Módulo para interagir com o interpretador Python
#import telegram  # Biblioteca para interagir com a API do Telegram
from Enviar_Log import EnviarLogs  # Importa a classe EnviarLogs do arquivo Logs.py

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Extração de Relatórios")  # Define o título da janela principal
        self.root.geometry("850x700")  # Define o tamanho da janela

        self.create_widgets()  # Cria os widgets da interface

        self.is_running = False  # Flag para indicar se o agendamento está ativo
        self.load_schedule()  # Carrega agendamentos salvos

        self.stdout_redirector = StdoutRedirector(self.log_text)  # Inicializa o redirecionador de stdout

    def create_widgets(self):
        # Criação das abas
        self.notebook = ttk.Notebook(self.root)  # Cria um notebook para abas
        self.notebook.pack(fill=tk.BOTH, expand=True)  # Preenche a área disponível

        # Aba Inicio
        self.tab_inicio = tk.Frame(self.notebook)  # Cria um frame para a aba Inicio
        self.notebook.add(self.tab_inicio, text="Inicio")  # Adiciona a aba ao notebook

        # Conteúdo da aba Inicio
        tk.Label(self.tab_inicio, text="Horários para extração de relatórios (HH:MM)").pack(pady=5)  # Rótulo de instrução

        self.time_entries = []  # Lista para armazenar os campos de entrada de horários
        for i in range(4):  # Cria quatro campos de entrada para horários
            entry = tk.Entry(self.tab_inicio)
            entry.pack(pady=5)
            self.time_entries.append(entry)

        self.save_button = tk.Button(self.tab_inicio, text="Salvar Agendamentos", command=self.save_schedule)  # Botão para salvar os horários agendados
        self.save_button.pack(pady=20)

        self.stop_button = tk.Button(self.tab_inicio, text="Encerrar Sistema", command=self.stop_system)  # Botão para encerrar o sistema
        self.stop_button.pack(pady=10)

        self.status_label = tk.Label(self.tab_inicio, text="Status: Inativo", fg="red")  # Rótulo para exibir o status do sistema
        self.status_label.pack(pady=20)

        self.log_text = tk.Text(self.tab_inicio, height=30, width=100)  # Campo de texto para exibir o log de eventos
        self.log_text.pack(pady=10)

        # Aba Configurações
        self.tab_config = tk.Frame(self.notebook)  # Cria um frame para a aba Configurações
        self.notebook.add(self.tab_config, text="Configurações")  # Adiciona a aba ao notebook

        RoboV4.Widgets_RoboV4(self.tab_config)  # Adiciona widgets específicos de RoboV4 na aba Configurações
        Utils.Widgets_Utils(self.tab_config)  # Adiciona widgets utilitários na aba Configurações

    def load_schedule(self):
        try:
            with open("agendamentos.txt", "r") as file:
                times = file.readlines()  # Lê os horários agendados do arquivo
                times = [t.strip() for t in times if self.validate_time_format(t.strip())]  # Filtra e limpa os horários válidos
                for i, time_entry in enumerate(self.time_entries):
                    if i < len(times):
                        time_entry.insert(tk.END, times[i])  # Insere os horários nos campos de entrada
        except FileNotFoundError:
            pass  # Se o arquivo não existe, não faz nada

    def save_schedule(self):
        times = [entry.get() for entry in self.time_entries]  # Obtém os horários dos campos de entrada
        if all(self.validate_time_format(t) for t in times):  # Verifica se todos os horários são válidos
            with open("agendamentos.txt", "w") as file:
                file.write("\n".join(times))  # Salva os horários no arquivo
            self.start_schedule(times)  # Inicia o agendamento
        else:
            messagebox.showerror("Erro", "Por favor, insira horários válidos no formato HH:MM.")  # Exibe mensagem de erro

    def validate_time_format(self, time_str):
        try:
            time.strptime(time_str, "%H:%M")  # Verifica se o horário está no formato HH:MM
            return True
        except ValueError:
            return False

    def start_schedule(self, times):
        if not self.is_running:
            self.is_running = True  # Define o flag como True
            self.status_label.config(text="Status: Ativo", fg="green")  # Atualiza o status para ativo
            self.log("Agendamento iniciado.")  # Loga a mensagem de início do agendamento
            Thread(target=self.run_schedule, args=(times,)).start()  # Inicia a execução do agendamento em uma thread separada
        else:
            messagebox.showinfo("Info", "O agendamento já está em execução.")  # Informa que o agendamento já está em execução

    def run_schedule(self, times):
        try:
            schedule.clear()  # Limpa agendamentos anteriores

            for time_str in times:
                schedule.every().day.at(time_str).do(self.extrair_relatorios)  # Agenda a extração dos relatórios
            
            while self.is_running:
                schedule.run_pending()  # Executa tarefas agendadas
                time.sleep(1)
        except Exception as e:
            self.log(f"Ocorreu um erro no main: {e}")  # Loga qualquer erro ocorrido durante a execução

    def extrair_relatorios(self):
        try:
            self.log("Iniciando o sistema...")
            RoboV4.abrir_software()  # Abre o software RoboV4
            self.log("Realizando Login...")
            RoboV4.realizar_login()  # Realiza o login

            #Sequência de chamadas para extrair diferentes relatórios
            self.log("Abrindo Cubo de Decisão...")
            RoboV4.abrir_cubo_de_decisao()
            self.log("Retirando Relatorio Estoque SUPPER...")
            RoboV4.relatorio_estoque_supper()
            self.log("Retirando Relatorio Faturamento Peso...")
            RoboV4.relatorio_faturamento_peso()
            self.log("Retirando Relatorio Pedido Compra...")
            RoboV4.relatorio_pedido_compra()
            self.log("Retirando Relatorio Movimento de Saida de ODF...")
            RoboV4.relatorio_Mov_Saida_ODF()

            if datetime.now().weekday() == 0:  # Se for segunda-feira
                self.log("Iniciando extração dos Parametros...")
                RoboV4.abrir_Parametros()  # Extrai parâmetros específicos
            else:
                self.log("Não é segunda-feira, não serão extraídos os parâmetros...")

            self.log("Retirando Relatorio MRP...")
            RoboV4.abrir_mrp()
            self.log("Relatórios extraídos com sucesso.")  # Loga a mensagem de sucesso

            self.log("Fechando o Sistema...")
            RoboV4.fechar_sistema()  # Certifica-se de que o sistema será fechado

            #self.log("Abrindo Pipefy e Retirando o Relatorio...")
            #RoboPipefy.main()
        except Exception as e:
            self.log(f"Ocorreu um erro no main principal(): {e}")
            self.salvar_log_em_arquivo()  # Salva o log em um arquivo
            EnviarLogs.enviar_logs()  # Envia o log para o Telegram
   
    def log(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")  # Adiciona mensagem ao log
        self.log_text.see(tk.END)  # Rola o log para a última linha

    def stop_schedule(self):
        self.is_running = False  # Define o flag como False
        self.status_label.config(text="Status: Inativo", fg="red")  # Atualiza o status para inativo
        self.log("Agendamento parado.")  # Loga a mensagem de parada do agendamento

    def stop_system(self):
        self.is_running = False
        self.stop_schedule()  # Para o agendamento se estiver em execução
        self.root.destroy()  # Fecha a janela
        os._exit(0)  # Força a saída do processo

    def salvar_log_em_arquivo(self):
        log_content = self.log_text.get("1.0", tk.END)
        folder_path = "Logs"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        filename = f"{folder_path}/error_log_{datetime.now().strftime('%Y_%m_%d_%H_%M')}.txt"
        with open(filename, "w") as file:
            file.write(log_content)
        self.log(f"Log salvo em {filename}")  # Loga o nome do arquivo salvo
        return filename

class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_space = text_widget
        self.original_stdout = sys.stdout
        sys.stdout = self

    def __del__(self):
        sys.stdout = self.original_stdout

    def write(self, message):
        self.text_space.insert(tk.END, message)
        self.text_space.see(tk.END)

    def flush(self):
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.protocol("WM_DELETE_WINDOW", app.stop_system)  # Garante que o sistema pare ao fechar a janela
    root.mainloop()
