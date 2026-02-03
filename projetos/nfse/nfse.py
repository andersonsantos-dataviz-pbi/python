# Importando bibliotecas do pywin32
import win32gui
import win32con
# Importando a biblioteca para trabalhar com tempo
import time
# Importando biblioteca para tratar subprocessos
import subprocess
# Importanto biblioteca para controlar teclado
import pyautogui
# Importando biblioteca de data
from datetime import date

# Pausa entre etapas de execução
pyautogui.PAUSE = 0.5

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
pyautogui.click(x=459, y=335) # Clicando no campo do CNPJ
pyautogui.write(cnpj_empresa) # Digitando o CNPJ no campo
pyautogui.press('tab') # Navegando para o campo de senha
pyautogui.write(senha) # Digitando a senha
pyautogui.press('tab') # Navegando para o botão de entrar
pyautogui.press('enter') # Pressionando enter para logar no sistema

# Posicionando o cursor do mouse sobre o botão de emissão de NFSe
time.sleep(5) # Aguardando o sistema carregar
pyautogui.click(x=1021, y=504) # Clicando no botão de emissão de NFSe
# Navegando até a data de competência
for _ in range(21):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.write(hoje) # Digitando a data de hoje no campo de competência
pyautogui.press('tab') # Validando preenchimento da data de competência
time.sleep(5) # Aguardando o sistema carregar
# Navegando para o campo de tipo de tributação
for _ in range(5):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.click(x=158, y=503) # Clicando no campo de tipo de tributação
time.sleep(5) # Aguardando o sistema carregar
pyautogui.click(x=183, y=598) # Selecionando o tipo de tributação "Regime de apuração dos tributos federais e municipal pelo Simples Nacional"
pyautogui.scroll(-5000) # Rolando a tela para cima
pyautogui.click(x=131, y=424) # Clicando no campo do tomador
pyautogui.scroll(-5000) # Rolando a tela para cima
pyautogui.click(x=138, y=223) # Clicando no campo do CNPJ do tomador
pyautogui.write(tomador) # Digitando o CNPJ do tomador
pyautogui.press('tab')
pyautogui.scroll(-5000) # Rolando a tela para cima
pyautogui.click(x=1727, y=747) # Clicando no botão para avançar para a tela seguinte

# Preenchendo as informações da tela de pessoas
# Posicionando o cursor do mouse no campo de município de ocorrência
for _ in range(1):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.click(x=836, y=426) # Clicando no listbox de município de ocorrência
pyautogui.write(municipio_ocorrencia) # Digitando o município de ocorrência
time.sleep(1) # Aguardando o sistema carregar as opções
# Selecionando o município de ocorrência e navegando para o próximo campo
for _ in range(1):
    pyautogui.press('tab')
    time.sleep(0.1)
pyautogui.click(x=142, y=596) # Clicando no listbox de tributação nacional
pyautogui.click(x=142, y=644) # Posicionando o cursor do mouse no campo de tributação nacional
pyautogui.write(tributacao_nacional) # Digitando o código de tributação nacional
time.sleep(1) # Aguardando o sistema carregar as opções
# Selecionando a tributação nacional e navegando para o próximo campo
for _ in range(1):
    pyautogui.press('tab')
    time.sleep(0.1)






