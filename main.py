import schedule
import time
from datetime import datetime
from RoboV4 import RoboV4
from RoboPipefy import RoboPipefy

# class Main:
#     def main():
#         try:
#             RoboV4.abrir_software()
#             RoboV4.realizar_login()
#             RoboV4.abrir_cubo_de_decisao()
#             RoboV4.relatorio_estoque_supper()
#             RoboV4.relatorio_faturamento_peso()
#             RoboV4.relatorio_pedido_compra()
#             RoboV4.relatorio_Mov_Saida_ODF()
#             RoboV4.relatorio_Req_Almoxarifado()
#             RoboV4.abrir_Parametros()
#             RoboV4.abrir_mrp()
#             RoboPipefy.main()
#         except Exception as e:
#             print(f"Ocorreu um erro no main: {e}")

# if __name__ == "__main__":
#     Main.main()

class Main:
    @staticmethod
    def main():
        try:
            # Agendar a extração dos relatórios nos horários específicos
            schedule.every().day.at("06:45").do(Main.extrair_relatorios)
            schedule.every().day.at("10:45").do(Main.extrair_relatorios)
            schedule.every().day.at("17:00").do(Main.extrair_relatorios)
            
            # Loop principal para aguardar os agendamentos
            while True:
                schedule.run_pending()
                time.sleep(1)

        except Exception as e:
            print(f"Ocorreu um erro no main: {e}")

    @staticmethod
    def extrair_relatorios():
        try:
            RoboV4.abrir_software()
            RoboV4.realizar_login()
            RoboV4.abrir_cubo_de_decisao()
            RoboV4.relatorio_estoque_supper()
            RoboV4.relatorio_faturamento_peso()
            RoboV4.relatorio_pedido_compra()
            RoboV4.relatorio_Mov_Saida_ODF()

            # Verificar se é segunda-feira para extrair relatório de parâmetros
            if datetime.now().weekday() == 0:  # 0 corresponde a segunda-feira
                RoboV4.abrir_Parametros()

        except Exception as e:
            print(f"Ocorreu um erro ao extrair os relatórios: {e}")

if __name__ == "__main__":
    Main.main()