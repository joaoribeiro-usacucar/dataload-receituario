import pyautogui as pag
import pygetwindow as gw
import openpyxl as xl
from troca_instancia import trocainstancia
import time

titleformulario = 'Instância - [RAF00450 - Receituário Agronômico]'
titlepreview = 'RAR00450: Previewer '
titleerro = 'SOL for Windows 95'
titlereport = 'Reports Background Engine'
receituario = None


def arredondar(numero):
    quociente, resto = numero // 5, numero % 5
    resultado = (5 * (quociente + (resto != 0)))
    if resultado == 0:
        resultado = 1
    return resultado


while receituario is None:
    try:  # Tenta atribuir as informações da tela de receituario até conseguir
        receituario = gw.getWindowsWithTitle(titleformulario)[0]
    except IndexError:
        pag.alert('Alerta',
                  'Sol não está aberto, por favor faça login no SOL e abra o formulário de Receituário Agronômico.')

wb = xl.load_workbook(filename=r'para_corrigir.xlsx', data_only=True, keep_vba=True)
sheet = wb['Planilha1']
i = 2
usa = 'USA2'
while True:
    if sheet['V' + str(i)].value == 'Verificar Area' and sheet['U' + str(i)].value == 'NÃO':
        receituario.activate()
        time.sleep(5)

        if sheet['B' + str(i)].value != usa:
            usa = trocainstancia(usa, sheet['B' + str(i)].value)

        pag.press('alt')
        pag.press('a')
        pag.press('down')
        pag.press('f')
        pag.press('enter')
        pag.press('TAB')

        time.sleep(2)
        pag.write(str(sheet['S' + str(i)].value))  # Escreve numero do receituário

        pag.press('TAB')
        pag.press('TAB')

        time.sleep(2)
        pag.write(str(sheet['O' + str(i)].value))  # Escreve o CPF do responsável

        time.sleep(2)
        pag.press('alt')
        pag.press('a')
        pag.press('down')
        pag.press('x')

        # 5  Tab
        for x in range(1, 6):
            pag.press('TAB')

        pag.write(str(sheet['I' + str(i)].value))  # Escreve area correta

        # 7  Tab
        for x in range(1, 9):
            pag.press('TAB')

        pag.write(str(arredondar(int(sheet["X" + str(i)].value))))  # Escreve quantidade correta

        pag.press('F10')
        print('Obtendo tela do previewer')
        while True:
            try:
                impresso = gw.getWindowsWithTitle(titlepreview)[0]
                impresso.activate()
                break
            except IndexError:
                pass

        time.sleep(1)
        botaoprint = pag.locateOnScreen(r'imagens\print.png')
        pag.click(botaoprint)

        while True:
            try:
                confirmacao = gw.getWindowsWithTitle('Imprimir')[0]
                confirmacao.activate()
                pag.press('enter')
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

        pag.write(str(sheet['W' + str(i)].value), interval=0.05)
        pag.press('enter')

        while True:
            try:
                doro = gw.getWindowsWithTitle('Doro PDF Writer')[0]
                doro.activate()
                pag.press('left')
                pag.press('enter')
                break
            except IndexError:
                break

        print('saiu do loop de impressao')
        time.sleep(1)
        impresso.close()
        i = i + 1
    elif sheet['U' + str(i)].value is not None:
        i = i + 1
    else:
        break
