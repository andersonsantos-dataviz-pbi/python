# Importando bibliotecas do pywin32
import win32gui
import win32con
# Importando a biblioteca para trabalhar com tempo
import time
# Importando biblioteca para tratar subprocessos
import subprocess
# Importanto biblioteca para controlar teclado
import pyautogui as py
# Importando biblioteca de data
from datetime import date

# Pausa entre etapas de execução
py.PAUSE = 0.5

# Variáveis de ambiente
link = 'https://www.nfse.gov.br/EmissorNacional'
cnpj_empresa = '47.350.412/0001-15'
senha = 'Ams@1979'
hoje = date.today().strftime("%d/%m/%Y")
tomador = '21.269.365/0001-96'
municipio_ocorrencia = 'Belo Horizonte'
tributacao_nacional = '010701'
"""
# Acessando o Microsoft Edge em perfil especifico, maximizado utilizando subprocess

subprocess.Popen([r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe','--start-maximized', '--profile-directory=NFSe', 'https://www.nfse.gov.br/EmissorNacional'])
# Aguardando o navegador abrir
time.sleep(5)
"""
# Acessando o Google Chrome maximizado utilizando subprocess
chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

subprocess.Popen([
    chrome_path,
    '--new-window',
    '--start-maximized',
    '--disable-session-crashed-bubble',
    '--profile-directory=NFSe',
    'https://www.nfse.gov.br/EmissorNacional'
])

# Aguarda o Chrome criar a janela
time.sleep(2)

# Abrindo o navegador do Google Chrome Maximizado
def maximize_chrome_window():
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if 'Chrome' in title:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

    win32gui.EnumWindows(enum_handler, None)

maximize_chrome_window()
time.sleep(5)

# Fazendo o login no sistema
py.click(x=459, y=335) # Clicando no campo do CNPJ
py.write(cnpj_empresa) # Digitando o CNPJ no campo
py.press('tab') # Navegando para o campo de senha
py.write(senha) # Digitando a senha
py.press('tab') # Navegando para o botão de entrar
py.press('enter') # Pressionando enter para logar no sistema

# Posicionando o cursor do mouse sobre o botão de emissão de NFSe
time.sleep(5) # Aguardando o sistema carregar
py.click(x=1018, y=559) # Clicando no botão de emissão de NFSe
time.sleep(1)

# Preenchendo a data de competência
py.click(x=440, y=402) # Clicando no campo da data de competência
py.write(hoje) # Digitando a data de hoje no campo de competência
py.press('tab') # Validando preenchimento da data de competência
time.sleep(1) # Aguardando o sistema carregar
py.scroll(-5000) # Rolando a tela para cima

# Selecionando o tipo de tributação
py.click(x=207, y=164) # Clicando no campo de tipo de tributação
time.sleep(1) # Aguardando o sistema carregar
py.click(x=180, y=260) # Selecionando o tipo de tributação "Regime de apuração dos tributos federais e municipal pelo Simples Nacional"
time.sleep(1) # Aguardando o sistema carregar

# Preenchendo o campo de tomador de serviço
py.click(x=130, y=429) # Clicando no campo do tomador
py.scroll(-5000) # Rolando a tela para cima
py.click(x=138, y=223) # Clicando no campo do CNPJ do tomador
py.write(tomador) # Digitando o CNPJ do tomador
py.press('tab')
time.sleep(1) # Aguardando o sistema carregar
py.scroll(-5000) # Rolando a tela para cima
time.sleep(1) # Aguardando o sistema carregar
"""


py.click(x=1727, y=747) # Clicando no botão para avançar para a tela seguinte.
"""
"""
# Navegando para o campo de tipo de tributação
for _ in range(5):
    py.press('tab')
    time.sleep(0.1)


# Preenchendo as informações da tela de pessoas
# Posicionando o cursor do mouse no campo de município de ocorrência
for _ in range(1):
    py.press('tab')
    time.sleep(0.1)
py.click(x=836, y=426) # Clicando no listbox de município de ocorrência
py.write(municipio_ocorrencia) # Digitando o município de ocorrência
time.sleep(1) # Aguardando o sistema carregar as opções
# Selecionando o município de ocorrência e navegando para o próximo campo
for _ in range(1):
    py.press('tab')
    time.sleep(0.1)
py.click(x=142, y=596) # Clicando no listbox de tributação nacional
py.click(x=142, y=644) # Posicionando o cursor do mouse no campo de tributação nacional
py.write(tributacao_nacional) # Digitando o código de tributação nacional
time.sleep(1) # Aguardando o sistema carregar as opções
# Selecionando a tributação nacional e navegando para o próximo campo
for _ in range(1):
    py.press('tab')
    time.sleep(0.1)
"""





