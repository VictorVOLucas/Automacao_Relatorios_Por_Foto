# Automação de Extração de Relatórios

Este projeto visa automatizar a extração de relatórios de um software específico usando Python e automação de interface gráfica.

## Pré-requisitos

- Python 3.12.4 instalado
- Bibliotecas Python necessárias (instaláveis via `pip install <nome_da_biblioteca>`):
  - pyautogui
  - schedule

## Configuração

1. Clone o repositório para o seu ambiente local:
   ```bash
   git clone https://github.com/VictorVOLucas/Robo_Relatorio_KORP.git
   cd Robo_Relatorio_KORP

# Funcionamento do Código

## Estrutura de Arquivos

- Main.py: Arquivo principal que inicia o processo de automação.
- RoboV4.py: Contém a classe RoboV4 com métodos para interagir com o software específico.
- RoboPipefy.py: Integração com o Pipefy, seção não detalhada neste exemplo.
- Utils.py: Funções utilitárias para interagir com a interface gráfica usando pyautogui.
- img/: Diretório contendo imagens utilizadas para localizar elementos na interface gráfica.

## Agendamento de Extração

O script utiliza a biblioteca schedule para agendar a extração dos relatórios em três horários diferentes durante o dia (9h, 12h e 15h). Para isso, a função extrair_relatorios() é chamada nos horários especificados todos os dias. Além disso, o relatório de parâmetros é extraído toda segunda-feira.

## Notas Importantes
Certifique-se de que as imagens utilizadas (*.png) em Utils.py estão atualizadas e correspondem aos elementos visuais da interface do software.
Ajuste os caminhos de arquivos e outras configurações específicas conforme necessário para o seu ambiente de automação.
