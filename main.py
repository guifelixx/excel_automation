import openpyxl  # Importa a biblioteca openpyxl para trabalhar com planilhas Excel
import pyperclip  # Importa a biblioteca pyperclip para manipular a área de transferência do sistema
import pyautogui  # Importa a biblioteca pyautogui para automatizar processos
from openpyxl.utils import FORMULAE  # Importa a constante FORMULAE da biblioteca openpyxl.utils

# Carrega os arquivos das planilhas Excel
resina_fatima = openpyxl.load_workbook('FATIMA 30-01.xlsx', data_only=True)
fechamento = openpyxl.load_workbook('FECHAMENTO FATIMA.xlsx')
sheet_resina_fatima = resina_fatima['Plan1']
sheet_fechamento = fechamento['JANEIRO 03-01']

# Inicializa a variável para armazenar o valor da célula
valor_celula = None

# Itera sobre as células da faixa I29:J30 da planilha FATIMA 30-01.xlsx para encontrar um valor válido
for row in sheet_resina_fatima['I29:J30']:
    for cell in row:
        # Verifica se o valor da célula não é None e não é uma fórmula
        if cell.value is not None and cell.data_type not in FORMULAE:
            valor_celula = cell.value  # Armazena o valor da célula se for válido
            break  # Interrompe a iteração se um valor válido for encontrado
    if valor_celula is not None:  # Se um valor válido foi encontrado, interrompe a iteração externa
        break

# Verifica se o valor da célula não é None e pode ser convertido para um número float
# Tenta converter o valor da célula para float
# Substitui o ponto pela vírgula na representação do float
# Copia o valor formatado com vírgula para a área de transferência
if valor_celula is not None:
    try:
        valor_float = float(valor_celula)
        valor_com_virgula = str(valor_float).replace('.', ',')
        pyperclip.copy(valor_com_virgula)
    except ValueError:
        print("O valor da célula não pôde ser convertido para um número float.")

# Clica em uma posição na tela para ativar a janela da planilha de destino e cola o valor copiado
def rodar():
    pyautogui.click(163, 302, duration=1)
    pyautogui.hotkey('ctrl', 'v')
