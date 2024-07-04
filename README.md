# Sistema de Extração de Relatórios

Este projeto é um sistema para extração de relatórios utilizando Tkinter para a interface gráfica. O sistema permite agendar e executar tarefas de extração de relatórios de forma automática e manual.

## Estrutura do Projeto

- `main.py`: Arquivo principal que inicializa a aplicação Tkinter.
- `RoboV4.py`: Contém a lógica principal do robô de extração de relatórios.
- `Enviar_Log.py`: Script responsável pelo envio de logs.
- `utils.py`: Contém funções utilitárias usadas em vários pontos do projeto.

## Funcionalidades

- **Agendamento de Tarefas**: Permite agendar a extração de relatórios para horários específicos.
- **Execução Manual**: Possibilidade de executar a extração de relatórios manualmente.
- **Logs**: Geração e envio de logs para monitoramento e debug.

## Instalação

### Pré-requisitos

- Python 3.x
- Bibliotecas necessárias listadas no `requirements.txt`

### Passos para Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/sistema-extracao-relatorios.git
Navegue até o diretório do projeto:

bash
Copiar código
cd sistema-extracao-relatorios
Crie um ambiente virtual e ative-o:

bash
Copiar código
python -m venv venv
source venv/bin/activate  # Para Linux/Mac
.\venv\Scripts\activate  # Para Windows
Instale as dependências:

bash
Copiar código
pip install -r requirements.txt
Execute a aplicação:

bash
Copiar código
python main.py
Uso
Interface Gráfica
A aplicação possui uma interface gráfica feita com Tkinter, onde você pode:

Agendar novas tarefas de extração.
Executar a extração manualmente.
Visualizar logs e status das execuções.
Alterar Ícone da Janela
Para alterar o ícone da janela, certifique-se de ter um arquivo de ícone (.ico ou .png) e ajuste o caminho no main.py:

python
Copiar código
# No arquivo main.py
icon = ImageTk.PhotoImage(file='caminho/para/seu/icone.png')
self.root.iconphoto(False, icon)
Criar Executável
Para criar um executável do seu projeto, utilize o PyInstaller:

Instale o PyInstaller:

bash
Copiar código
pip install pyinstaller
Crie o executável:

bash
Copiar código
pyinstaller --onefile --windowed --icon=caminho/para/seu/icone.ico main.py
O executável será gerado no diretório dist.

Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e pull requests.

Licença
Este projeto está licenciado sob a MIT License. Veja o arquivo LICENSE para mais detalhes.

markdown
Copiar código

### Observações

1. **Certifique-se de ajustar os caminhos dos ícones** no `main.py` conforme necessário.
2. **Inclua um arquivo `requirements.txt`** com todas as dependências necessárias.
3. **Verifique se há um arquivo de licença** e inclua o tipo de licença no README.
