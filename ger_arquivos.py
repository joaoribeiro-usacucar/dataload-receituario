import os


def cria_pasta():
    os.system(r'mkdir C:\RA')


def impressora():
    comando_path = r'REG ADD HKCU\SOFTWARE\CompSoft\Doro\ /v Path /t REG_SZ /d C:\RA /f'
    comando_flag = r'REG ADD HKCU\SOFTWARE\CompSoft\Doro\ /v Flags /t REG_DWORD /d 0 /f'
    comando_impressora = r'RUNDLL32 PRINTUI.DLL, PrintUIEntry /y /n "Doro PDF Writer"'

    os.system(comando_path)
    os.system(comando_flag)
    os.system(comando_impressora)


def criar_arquivo():
    texto = [
        '@CHCP 1252 >NUL',
        "\n",
        r'move "C:\RA\*USA1 TRATOS*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Iguatemi\Tratos Culturais"',
        "\n",
        r'move "C:\RA\*USA2 TRATOS*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Paranacity\Tratos Culturais"',
        "\n",
        r'move "C:\RA\*USA13 TRATOS*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Terra Rica\Tratos Culturais"',
        "\n",
        r'move "C:\RA\*USA1 APLICACAO AEREA*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Iguatemi\Aplicação Aérea"',
        "\n",
        r'move "C:\RA\*USA2 APLICACAO AEREA*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Paranacity\Aplicação Aérea"',
        "\n",
        r'move "C:\RA\*USA13 APLICACAO AEREA*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Terra Rica\Aplicação Aérea"',
        "\n",
        r'move "C:\RA\*USA1 FORMACAO*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Iguatemi\Formação"',
        "\n",
        r'move "C:\RA\*USA2 FORMACAO*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Paranacity\Formação"',
        "\n",
        r'move "C:\RA\*USA13 FORMACAO*.PDF" "\\USA1FS1\Vol3\USER\RA - Receituário Agronômico\RAs 2022\Terra Rica\Formação"'
        "\n"
    ]
    bat = open('mover.bat', 'w')
    bat.writelines(texto)
    bat.close()


if __name__ == '__main__':
    cria_pasta()
