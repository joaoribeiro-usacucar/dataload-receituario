import openpyxl as xl
import pygetwindow as gw
import pyautogui
import os
import time
import unidecode as decode
from datetime import datetime
from troca_instancia import troca_instancia
from ger_arquivos import *

# from tkinter.filedialog import askdirectory

pyautogui.FAILSAFE = True
impressora()
cria_pasta()


class Agronomo:
    def __init__(self, instancia, nomeagr, cpf, art, ultimo, datamin, datamax, ramax):
        self.instancia = instancia
        self.nome = nomeagr
        self.cpf = cpf
        self.art = art
        self.ultimo = ultimo
        self.datamin = datamin
        self.datamax = datamax
        self.ramax = ramax

    def __str__(self) -> str:
        return 'Nome: %s CPF: %s' % (self.nome, self.cpf)


def arredondar(x):
    quociente, resto = x // 5, x % 5
    resultado = (5 * (quociente + (resto != 0)))
    if resultado == 0:
        resultado = 1
    return resultado


def logdeerro(linha, campo):
    erro = ['Erro na linha: ', str(linha), ", tentando digitar:", str(campo), ", ",
            datetime.today().strftime("%d/%m/%y %H:%M"), "\n"]
    conteudo.append(''.join(erro))


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
            sol95 = gw.getWindowsWithTitle(titleerro)[0]
            sol95.close()
            return 1
        except IndexError:
            return 0
    except pyautogui.FailSafeException:
        return 2


# data_emi, cpf, fazenda, secao, area
# cultura, alvo, diag, qt_aplic, produto, dose, quant_ad, unidade, mod, intervalo, precau, manejo, epi, outros


def preencher(dadosra, linha, tela, art, ra):
    controle = 3
    try:
        print(*dadosra, sep=',')
        tela.activate()
        time.sleep(1)
        controle = 0
        counter = 1
        for digitando in dadosra:
            resultado = escreve(digitando)
            if resultado == 1:
                logdeerro(linha, digitando)
                pyautogui.press('alt')
                pyautogui.press('r')
                pyautogui.press('m')
                controle = 1
                break
            elif resultado == 2:
                controle = 99
                break
            else:
                if counter == 2:
                    pyautogui.write(str(art), interval=0.1)
                    pyautogui.press('TAB')
                elif counter == 10:
                    pyautogui.press('TAB')
            counter = counter + 1

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
                    try:
                        pyautogui.write('')
                    except pyautogui.FailSafeException:
                        return 2

            time.sleep(1)
            botaoprint = pyautogui.locateOnScreen(r'imagens\print.png', confidence=0.80)
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
                    try:
                        pyautogui.write('')
                    except pyautogui.FailSafeException:
                        return 2

            time.sleep(2)
            header = ['RA ',
                      str(ra + 1),
                      ' ',
                      str(sheet['C' + str(linha)].value),
                      ' ',
                      str(sheet['D' + str(linha)].value),
                      ' OS',
                      str(sheet['F' + str(linha)].value),
                      ' FAZENDA ',
                      str(sheet['I' + str(linha)].value),
                      ' ',
                      decode.unidecode(sheet['J' + str(linha)].value),
                      ' SECAO',
                      str(sheet['K' + str(linha)].value),
                      ' - ',
                      decode.unidecode(sheet['R' + str(linha)].value),
                      ]
            pyautogui.write(''.join(header), interval=0.05)
            pyautogui.press('enter')

            impresso.close()
            return 0
        else:
            if controle == 99:
                return 2
            else:
                return 1
    except pyautogui.FailSafeException:
        if controle == 0:
            return 2
        if controle == 3:
            return 3


titleformulario = 'Instância - [RAF00450 - Receituário Agronômico]'
titlepreview = 'RAR00450: Previewer '
titleerro = 'SOL for Windows 95'
titlereport = 'Reports Background Engine'

if pyautogui.confirm('Deseja desligar o pc apos a conclusão?', buttons=['Sim', 'Não']) == 'Sim':
    print('O computador irá desligar ao fim do processo!')
    desligar = 1
else:
    desligar = 0

try:
    logs = open('log de erros.txt', 'r')
    conteudo = logs.readlines()
    logs = open('log de erros.txt', 'w')
except FileNotFoundError:
    conteudo = []
    logs = open('log de erros.txt', 'w')

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

while receituario is None:
    try:  # Tenta atribuir as informações da tela de receituario até conseguir
        receituario = gw.getWindowsWithTitle(titleformulario)[0]
    except IndexError:
        print("Sol não está aberto, por favor faça login no SOL e abra o formulário de Receituário Agronômico.")
        os.system("PAUSE")

# 'SOL - [Usina de Acucar Santa Terezinha Ltda. - em Recuperação Judicial]' nome da janela inicial do sol
# 'Instância - [RAF00450 - Receituário Agronômico]' nome da jenela de receituario do sol.

saida = 0
agronomos = []
i = 2
while True:
    if cadastro['J' + str(i)].value is not None:
        agronomos.append(
            Agronomo(
                cadastro['J' + str(i)].value,  # Instancia
                cadastro['K' + str(i)].value,  # Nome
                cadastro['L' + str(i)].value,  # CPF
                cadastro['M' + str(i)].value,  # ART
                cadastro['N' + str(i)].value,  # Ultima RA
                str(cadastro['O' + str(i)].value.strftime("%d/%m/%Y")),  # Data Min
                str(cadastro['P' + str(i)].value.strftime("%d/%m/%Y")),  # Data Max
                cadastro['Q' + str(i)].value  # RA Max
            )
        )
        i = i + 1
    else:
        break

usa = sheet["T1"].value
i = 3  # Loop para percorrer a planilha de receituario
retorno = 0

while True:
    if sheet["AN" + str(i)].value == "OK" and saida == 0:
        campos = [
            str(sheet["V" + str(i)].value.strftime("%d/%m/%Y")),  # data_emi
            sheet["W" + str(i)].value,  # cpf
            sheet["X" + str(i)].value,  # fazenda
            sheet["Y" + str(i)].value,  # secao
            sheet["Z" + str(i)].value,  # area
            sheet["AA" + str(i)].value,  # cultura
            sheet["AB" + str(i)].value,  # alvo
            sheet["AC" + str(i)].value,  # diag
            1,  # qt_aplic
            sheet["AD" + str(i)].value,  # produto
            sheet["AE" + str(i)].value,  # dose
            arredondar(int(sheet["AF" + str(i)].value)),  # quant_ad arredondada em multiplo de 5
            sheet["AG" + str(i)].value,  # unidade
            sheet["AH" + str(i)].value,  # mod
            sheet["AI" + str(i)].value,  # intervalo
            sheet["AJ" + str(i)].value,  # precau
            sheet["AK" + str(i)].value,  # manejo
            sheet["AL" + str(i)].value,  # epi
            sheet["AM" + str(i)].value]  # outros

        for obj in agronomos:
            if obj.cpf == campos[1]:
                if time.strptime(obj.datamin, "%d/%m/%Y") <= time.strptime(campos[0],
                                                                           "%d/%m/%Y") <= time.strptime(obj.datamax,
                                                                                                        "%d/%m/%Y"):
                    if obj.ramax > obj.ultimo:
                        campos[0] = sheet["V" + str(i)].value.strftime("%d%m%Y")

                        if sheet["C" + str(i)].value != usa:
                            print(i)
                            usa = troca_instancia(sheet["C" + str(i)].value)

                        retorno = preencher(campos, i, receituario, obj.art, obj.ultimo)
                        if retorno == 0:
                            campos.clear()
                            obj.ultimo = obj.ultimo + 1
                            sheet['O' + str(i)] = obj.ultimo
                            sheet['P' + str(i)] = 'SIM'
                        elif retorno == 2:
                            sheet['P' + str(i)] = 'RA Preenchida Verificar se foi Salva'
                            saida = 1
                            break
                        elif retorno == 3:
                            break
                        else:
                            campos.clear()
                            sheet['O' + str(i)] = ''
                    else:
                        sheet['O' + str(i)] = 'ART atingiu limite de RA'
                else:
                    sheet['O' + str(i)] = 'Data fora do periodo da ART'
            if not campos:
                break
        i = i + 1
        if 1 < retorno > 4:
            break
    elif sheet["AN" + str(i)].value == "-":
        i = i + 1
    else:
        break

fim = ['Fim da execução em: ', datetime.today().strftime("%d/%m/%y - %H:%M"), "\n"]

conteudo.append(''.join(fim))
logs.writelines(conteudo)
logs.close()

try:
    wb.save(filename=r'Controle RA.xlsm')
except PermissionError:
    wb.save(filename='output.xlsm')

wb.close()

os.system(r'mover.bat')

if desligar == 1:
    os.system("shutdown /s /t 1")
else:
    mensagem = ['Concluido as: ', datetime.today().strftime("%d/%m/%y - %H:%M")]
    pyautogui.alert(''.join(mensagem), 'Operação Finalizada')
