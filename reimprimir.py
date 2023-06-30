import openpyxl as xl
import pygetwindow as gw
import pyautogui
import unidecode as decode
import date, time, datetime
import os


def preencher(linha, ra):
    print("Pegando tela do Previewer")
    while True:
        try:
            impresso = gw.getWindowsWithTitle(titlepreview)[0]
            impresso.activate()
            break
        except IndexError:
            pass

    time.sleep(1)
    botaoprint = pyautogui.locateOnScreen(r'imagens\print.png')
    pyautogui.click(botaoprint)

    while True:
        try:
            confirmacao = gw.getWindowsWithTitle('Imprimir')[0]
            confirmacao.activate()
            pyautogui.press('enter')
            break
        except IndexError:
            pass

    print('pegando tela do doro pdf')
    while True:
        try:
            doro = gw.getWindowsWithTitle('Doro PDF Writer [1.82]')[0]
            doro.activate()
            break
        except IndexError:
            pass

    time.sleep(2)
    header = ['RA ',
              str(ra),
              ' ',
              str(sheet['A' + str(linha)].value),
              ' ',
              str(sheet['B' + str(linha)].value),
              ' OS',
              str(sheet['D' + str(linha)].value),
              ' FAZENDA ',
              str(sheet['G' + str(linha)].value),
              ' ',
              decode.unidecode(sheet['H' + str(linha)].value),
              ' SECAO',
              str(sheet['I' + str(linha)].value),
              ' - ',
              decode.unidecode(sheet['N' + str(linha)].value),
              ]
    pyautogui.write(''.join(header), interval=0.05)
    pyautogui.press('enter')

    impresso.close()


titleformulario = 'Instância - [RAF00450 - Receituário Agronômico]'
titlepreview = 'RAR00450: Previewer '
titleerro = 'SOL for Windows 95'
titlereport = 'Reports Background Engine'

while True:
    try:
        wb = xl.load_workbook(filename=r'Controle RA.xlsm', data_only=True, keep_vba=True)
        cadastro = wb['ControleART']
        sheet = wb['ControleRA']
        break
    except PermissionError:
        print('Planilha de controle aberta. Por favor feche a pasta de trabalho e continue.')
        os.system("PAUSE")

# Variavel armazena informações da tela de receituario
receituario = None

# 'SOL - [Usina de Acucar Santa Terezinha Ltda. - em Recuperação Judicial]' nome da janela inicial do sol
# 'Instância - [RAF00450 - Receituário Agronômico]' nome da jenela de receituario do sol.

agronomos = []
i = 2

usa = sheet["P1"].value
i = 3  # Loop para percorrer a planilha de receituario
retorno = 0

try:
    while True:
        tela = gw.getWindowsWithTitle('Instância - [RAF01010 - Reimpressão do Receituário Agronômico]')[0]
        if sheet["P" + str(i)].value == '2022-04-27 00:00:00':
            cpf = pyautogui.locateOnScreen(r'imagens\cpf.png')
            pyautogui.click(cpf)
            pyautogui.write(sheet["W" + str(i)].value)
            escolher = pyautogui.locateOnScreen(r'imagens\escolher.png')
            pyautogui.click(escolher)
            pyautogui.write(sheet["R" + str(i)].value)
            print2 = pyautogui.locateOnScreen(r'imagens\print2.png')
            preencher(i, sheet["R" + str(i)].value)
            tela.activate()
            pyautogui.press('alt')
            pyautogui.press('r')
            pyautogui.press('m')
            pyautogui.press('alt')
            pyautogui.press('a')
            pyautogui.press('down')
            pyautogui.press('r')
            i = i+1
        elif sheet["P" + str(i)].value is None:
            i = i+1
finally:
    pass
