# Enviar_Log.py
from telegram import Bot
import asyncio
import os

class EnviarLogs:

    def __init__(self):
        # Substitua 'YOUR_TOKEN' pelo token do seu bot do Telegram
        self.TOKEN = '7460964482:AAEOWjVUfAhsA-8WBxAitXErNN4FL3TJZvQ'
        
        # Substitua 'CHAT_ID' pelo ID do chat onde você quer enviar o arquivo
        self.CHAT_ID = '-4271575797'
        
        # Caminho para a pasta que contém os arquivos TXT
        self.FOLDER_PATH = 'Logs'
        
        # Cria uma instância do bot
        self.bot = Bot(token=self.TOKEN)

    # PEGAR O ULTIMO ARQUIVO TXT GERADO NA PASTA
    def get_latest_file(self):
        txt_files = [f for f in os.listdir(self.FOLDER_PATH) if f.endswith('.txt')]
        if not txt_files:
            return None
        full_paths = [os.path.join(self.FOLDER_PATH, f) for f in txt_files]
        latest_file = max(full_paths, key=os.path.getmtime)
        return latest_file

    # ENVIAR O ARQUIVO PARA O TELEGRAM
    async def send_file(self, file_path):
        try:
            with open(file_path, 'rb') as file:
                await self.bot.send_document(chat_id=self.CHAT_ID, document=file)
            print(f"Arquivo {os.path.basename(file_path)} enviado com sucesso!")
        except Exception as e:
            print(f"Erro ao enviar o arquivo: {e}")

    # FUNÇÃO PRINCIPAL
    async def main(self):
        latest_file = self.get_latest_file()
        if latest_file:
            await self.send_file(latest_file)
        else:
            print("Nenhum arquivo TXT encontrado na pasta.")

    # Função para executar o envio de logs
    def enviar_logs():
        enviar_logs_instance = EnviarLogs()
        asyncio.run(enviar_logs_instance.main())
