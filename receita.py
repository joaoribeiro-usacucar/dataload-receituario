class Receituario:
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


if __name__ == '__main__':
    pass
