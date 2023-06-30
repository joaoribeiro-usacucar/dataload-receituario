import pyautogui
import time
import pygetwindow as gw
import unidecode as decode


titleformulario = 'Instância - [RAF00450 - Receituário Agronômico]'
titlepreview = 'RAR00450: Previewer '
titleerro = 'SOL for Windows 95'
titlereport = 'Reports Background Engine'

pyautogui.FAILSAFE = True


def escreve(dados):
    try:
        pyautogui.write(str(dados), interval=0.1)
        pyautogui.press('TAB')
        try:  # Tenta localizar tela de erro, se não conseguir retorna 0, se conseguir fecha e tenta escrever novamente
            sol95 = gw.getWindowsWithTitle(titleerro)[0]
            sol95.close()
            pyautogui.press('backspace')
            pyautogui.write(str(dados), interval=0.1)
            pyautogui.press('TAB')
            try:
                sol95 = gw.getWindowsWithTitle(titleerro)[0]
                sol95.close()
                return 1
            except IndexError:
                return 0

        except IndexError:
            return 0
    except pyautogui.FailSafeException:
        return 2


def preencher(dadosra, tela, ra):
    controle = 3
    try:
        print(*dadosra, sep=',')
        tela.activate()
        time.sleep(1)
        controle = 0
        counter = 1
        for campo in range(0, 20):
            result = escreve(dadosra[campo])
            if result == 1:
                pyautogui.press('alt')
                pyautogui.press('r')
                pyautogui.press('m')
                controle = 1
                break
            elif result == 0:
                if counter == 10:
                    pyautogui.press('TAB')
                counter = counter + 1
            else:
                break

        if counter == 18:
            controle = 0

        if controle == 0:
            pyautogui.press('f10')
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
                      str(ra + 1),
                      ' ',
                      str(dadosra[20]),
                      ' ',
                      str(dadosra[21]),
                      ' OS',
                      str(dadosra[22]),
                      ' FAZENDA ',
                      str(dadosra[3]),
                      ' ',
                      decode.unidecode(dadosra[23]),
                      ' SECAO',
                      str(dadosra[4]),
                      ' - ',
                      decode.unidecode(dadosra[24]),
                      ]
            pyautogui.write(''.join(header), interval=0.05)
            pyautogui.press('enter')

            impresso.close()
            return 0
        else:
            return 1
    except pyautogui.FailSafeException:
        if controle == 0:
            return 2
        if controle == 3:
            return 3


if __name__ == '__main__':
    pass
