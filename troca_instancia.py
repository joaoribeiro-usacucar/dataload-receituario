import time
import pyautogui as pag
import pygetwindow as gw
import pandas as pd

maringa = r'imagens\maringa.png'
paranacity = r'imagens\paranacity.png'
terrarica = r'imagens\terrarica.png'
receituario = r'imagens\receituario.png'
processos = r'imagens\processos.png'
receituarioagr = r'imagens\receituarioagronomico.png'

titlereceituario = 'Instância - [RAF00450 - Receituário Agronômico]'

unidades = pd.DataFrame(
    [
        ['USA1', r'imagens\maringa.png'],
        ['USA2', r'imagens\paranacity.png'],
        ['USA13', r'imagens\terrarica.png'],
        ['USA15', r'imagens\rondon.png'],
        ['USA16', r'imagens\gaucha.png'],
        ['USA4', r'imagens\ivate.png'],
        ['USA3', r'imagens\tapejara.png']
    ], columns=['INSTANCIA', 'PATH']
)


def to_formulario():
    path_form = pag.locateOnScreen(receituarioagr, confidence=0.85)
    pag.doubleClick(path_form)


def troca_instancia(destino):
    try:
        tela = gw.getWindowsWithTitle(titlereceituario)[0]
        tela.activate()
        pag.press('alt')
        pag.press('r')
        pag.press('m')
        pag.press('alt')
        pag.press('a')
        pag.press('down')
        pag.press('r')
        time.sleep(1)
    except IndexError:
        pass

    instancia = get_instancia()

    imagem = pag.locateOnScreen(unidades.loc[unidades['INSTANCIA'] == instancia].values[0][1], confidence=0.85)
    pag.click(imagem)
    pag.moveTo(540, 300)
    time.sleep(1)
    imagem = pag.locateOnScreen(unidades.loc[unidades['INSTANCIA'] == destino].values[0][1], confidence=0.85)
    pag.click(imagem)
    time.sleep(0.5)

    instancia = get_instancia()
    to_formulario()
    return instancia


def get_instancia():
    ins = None
    # self.sol.activate()
    for _, instancia in unidades.iterrows():
        botao = pag.locateOnScreen(instancia[1], confidence=0.85)
        if botao is not None:
            ins = instancia[0]
            break
    return ins


if __name__ == '__main__':
    troca_instancia('USA3')
