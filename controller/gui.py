import numpy as np
import subprocess
import sys

from numpy import linalg 
from openpyxl import * 
from tkinter import *
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import ttk
from datetime import datetime
from math import *

sys.path.insert(0, "/home/ubuntu/Downloads/freela01/controller")
from functions import *

Nome_aquivo = ''
Carno = 0
nno = 0
aberto = 0
Titulo = 'Nenhum arquivo aberto'


class Tela_About:
    def __init__(self, Raiz, Original):
        self.TelAbout = Raiz
        self.Teloriginal = Original
        self.TelAbout.title('Sobre')
        self.TelAbout.geometry("400x250+400+150")
        self.TelAbout.resizable(0, 0)
        self.canvas = Canvas(self.TelAbout, width=380, height=230, bg='#EEE8AA')
        self.canvas.place(x=10, y=10)
        self.lb1 = Label(self.TelAbout, text="PUC Minas - Poços de Caldas",
                         font=('Times', '13', 'bold'), bg='#EEE8AA')
        self.lb1.place(x=80, y=20)
        self.lb1 = Label(self.TelAbout, text="ESCREVER SOBRE O METODO UTILIZADO",
                         font=('Times', '11'),bg='#EEE8AA')
        self.lb1.place(x=50, y=80)
        self.lb1 = Label(self.TelAbout, text="Data : 01/02/2022",
                         font=('Times', '11'), bg='#EEE8AA')
        self.lb1.place(x=50, y=150)
        self.lb1 = Label(self.TelAbout, text="Versão: V1.0",
                         font=('Times', '11'), bg='#EEE8AA')
        self.lb1.place(x=50, y=130)
        self.lb1 = Label(self.TelAbout, text="Programadores:  -\n\t\t-",
                         font=('Times', '11'), bg='#EEE8AA')
        self.lb1.place(x=50, y=200)

class Tela_Principal:
    def __init__(self, Raiz):
        self.TelPrin = Raiz
        self.agora = datetime.now()
        self.TelPrin.title(Titulo)
        self.TelPrin.geometry("1000x545+150+10") 
        self.TelPrin.resizable(0, 0)
        self.canvas = Canvas(self.TelPrin, width=200, height=30, bg='#DCDCDC')
        self.canvas.place(x=0, y=510)
        self.canvas = Canvas(self.TelPrin, width=800, height=30, bg='Silver')
        self.canvas.place(x=200, y=510)
        self.Relogio = Label(self.TelPrin, text=self.agora.strftime('%d/%m/%Y       %H:%M:%S'),
                                           font=('Times', '10', 'bold'),
                                           bg='#DCDCDC',
                                           relief = "sunken")
        self.Relogio.place(x=5, y=515)
        self.ajud = Label(self.TelPrin, text="ESCREVER SOBRE O OBJETIVO DO PROJETO - Versão : V1.0                          " ,
                                        font=('Times', '10', 'bold'),
                                        bg='Silver',
                                        relief = "sunken")
        self.ajud.place(x=210, y=517)


        # ////Cria meunu superior  -  Principal
        self.Menusup = Menu(self.TelPrin, tearoff=0)  # Cria um menu superior
        self.TelPrin.configure(menu=self.Menusup)  # Configura a tela inicial para o menu

        # ////Cria meunu superior  -  Arquivo
        self.arquivo = Menu(self.Menusup, tearoff=0)  # Cria um menu dentro do menu superior Arquivo
        self.arquivo.add_command(label="Abrir", command=self.Sub_Abrir)  # Adiciona Abrir ao menu
        self.arquivo.add_command(label="Limpar Área", command=self.Sub_Limpa)  # Adiciona Abrir ao menu
        self.arquivo.add_command(label="Sair", command=self.Sub_Sair)  # Adiciona sair no sub menu
        self.Menusup.add_cascade(label="Arquivo", menu=self.arquivo)  # Adiciona  cascata

        # //  Cria Menu entrada de dados
        self.Menusup.add_command(label="Entrada Dados", command=self.Entrada_Dados)

       # //  Cria Menu para calcular os esforços
        self.Menusup.add_command(label="Dimensionar", command=self.Sub_Dimensionar)

        # //// Cria  Menu superior - Saída de dados
        self.Menusup.add_command(label="Resultados", command=self.Sub_Resultados)

        # //// Cria  Menu superior - About
        self.Menusup.add_command(label="Sobre", command=self.Sub_About)

        # //// Cria  Menu superior - Sair
        self.Menusup.add_command(label="Sair", command=self.Sub_Sair)

        self.areatexto = scrolledtext.ScrolledText(self.TelPrin,
                                                   width=120,
                                                   height=26,
                                                   font=('Consolas', '12'))

        self.areatexto.pack()
        #  atualização de tela
        self.alteracao()

        self.TelPrin.protocol("WM_DELETE_WINDOW", self.on_closing_Principal)

    # ************************************************Sub rotinas *******************************************
    # Atualização do Relógio
    def alteracao(self):
        self.agora = datetime.now()
        self.Relogio['text'] = self.agora.strftime('%d/%m/%Y       %H:%M:%S')
        self.TelPrin.after(1000, self.alteracao)

    def Sub_About(self):
        self.Telabout = Toplevel(self.TelPrin)
        Tela_About(self.Telabout, self.TelPrin)
        self.Telabout.protocol("WM_DELETE_WINDOW", self.on_close_about)

    def on_close_about(self):
        self.Telabout.destroy()

    def on_closing_Principal(self):
     self.Sub_Sair()

    def Sub_Sair(self):
        self.escolha = messagebox.askquestion('CONFIRMAÇÃO!!!', 'Tem certeza que deseja Sair?')
        if self.escolha == 'yes':
            self.TelPrin.destroy()

    def Sub_Abrir(self):
        global  Nome_aquivo , Titulo

        entrada_arquivo = filedialog.askopenfile(initialdir="/", title="Abrir arquivo",
                                               filetypes=(("Entrada", "*.XLSX"), ("all files", "*.*")))
        if entrada_arquivo != None:
            Nome_aquivo = entrada_arquivo.name
            Titulo = Nome_aquivo
            self.TelPrin.title(Titulo)
            self.areatexto.delete('1.0', END)
            self.areatexto.insert(END, str(Titulo)+'\n')
            self.areatexto.update()

    def Sub_Limpa(self):
         self.areatexto.delete('1.0',END)

    def Entrada_Dados(self):
      global Matno, Matbar, Carno, Carbar, nno, nbar, Nome_aquivo , aberto
      if Nome_aquivo == '' :
          messagebox.showinfo('ABOUT', 'Abra um arquivo Primeiro')
      else:
        # Abre o arquivo do excel para entrada de dados
        t = subprocess.Popen(Nome_aquivo, shell=True)
        t.wait()

        nno = numero_nos(Nome_aquivo)  # Determina o número de nós
        nbar = numero_barras(Nome_aquivo)  # Determina o número de barras

        Matno = np.zeros((nno, 10))
        Matbar = np.zeros((nbar, 11))
        Carno = np.zeros((nno, 4))
        Carbar = np.zeros((nbar, 7))

        # faz a leitura da pasta "Entrada de nós

        arqexcel = load_workbook(Nome_aquivo)
        Pastatraba = arqexcel["Entrada Nós"]
        for i in range(int(nno)):

            Matno[i, 0] = Pastatraba.cell(row=6 + i, column=1).value
            Matno[i, 1] = Pastatraba.cell(row=6 + i, column=2).value
            Matno[i, 2] = Pastatraba.cell(row=6 + i, column=3).value
            if str(Pastatraba.cell(row=6 + i, column=4).value) == 'Sim ':
                Matno[i, 3] = 1
            else:
                Matno[i, 3] = 0
            if str(Pastatraba.cell(row=6 + i, column=5).value) == 'Sim ':
                Matno[i, 4] = 1
            else:
                Matno[i, 4] = 0
            if str(Pastatraba.cell(row=6 + i, column=6).value) == 'Sim ':
                Matno[i, 5] = 1
            else:
                Matno[i, 5] = 0
            if str(Pastatraba.cell(row=6 + i, column=7).value) == 'Sim ':
                Matno[i, 6] = 1
            else:
                Matno[i, 6] = 0

            Matno[i, 7] = Pastatraba.cell(row=6 + i, column=8).value
            Matno[i, 8] = Pastatraba.cell(row=6 + i, column=9).value
            Matno[i, 9] = Pastatraba.cell(row=6 + i, column=10).value

        # Leitura da pasta 'Entrada Barras

        Pastatraba = arqexcel["Entrada Barras"]
        for i in range(int(nbar)):

            Matbar[i, 0] = Pastatraba.cell(row=6 + i, column=1).value

            if str(Pastatraba.cell(row=6 + i, column=2).value) == 'Engaste-Engaste':
                Matbar[i, 1] = 11
            if str(Pastatraba.cell(row=6 + i, column=2).value) == 'Engaste-Apoio':
                Matbar[i, 1] = 10
            if str(Pastatraba.cell(row=6 + i, column=2).value) == 'Apoio-Engasste':
                Matbar[i, 1] = 1
            if str(Pastatraba.cell(row=6 + i, column=2).value) == 'Apoio-Apoio':
                Matbar[i, 1] = 0

            Matbar[i, 2] = Pastatraba.cell(row=6 + i, column=3).value
            Matbar[i, 3] = Pastatraba.cell(row=6 + i, column=4).value
            Matbar[i, 4] = Pastatraba.cell(row=6 + i, column=5).value
            Matbar[i, 5] = Pastatraba.cell(row=6 + i, column=6).value
            Matbar[i, 6] = Pastatraba.cell(row=6 + i, column=7).value
            Matbar[i, 7] = Pastatraba.cell(row=6 + i, column=8).value
            Matbar[i, 8] = Pastatraba.cell(row=6 + i, column=9).value
            Matbar[i, 9] = Pastatraba.cell(row=6 + i, column=10).value
            Matbar[i, 10] = Pastatraba.cell(row=6 + i, column=11).value

        # Leitura da pasta 'Carregamentos nos Nós'

        Pastatraba = arqexcel["Car nos"]
        for i in range(int(nno)):
            Carno[i, 0] = Pastatraba.cell(row=6 + i, column=1).value
            Carno[i, 1] = Pastatraba.cell(row=6 + i, column=2).value
            Carno[i, 2] = Pastatraba.cell(row=6 + i, column=3).value
            Carno[i, 3] = Pastatraba.cell(row=6 + i, column=4).value

        # Leitura da pasta 'Carregamentos Barras'

        Pastatraba = arqexcel["Car Barras"]
        for i in range(int(nbar)):
            Carbar[i, 0] = Pastatraba.cell(row=6 + i, column=1).value
            Carbar[i, 1] = Pastatraba.cell(row=6 + i, column=2).value
            Carbar[i, 2] = Pastatraba.cell(row=6 + i, column=3).value
            Carbar[i, 3] = Pastatraba.cell(row=6 + i, column=4).value
            Carbar[i, 4] = Pastatraba.cell(row=6 + i, column=5).value
            Carbar[i, 5] = Pastatraba.cell(row=6 + i, column=6).value
            Carbar[i, 6] = Pastatraba.cell(row=6 + i, column=7).value

        aberto = 1
        messagebox.showinfo('ABOUT', 'Arquivo carregado com sucesso')
        self.areatexto.insert(END, 'Arquivo carregado com sucesso \n')
        self.areatexto.update()

    def Sub_Dimensionar(self):

     global Nome_aquivo , aberto , Matno, Matbar, Carno, Carbar, nno, nbar,aux, desl , delta

     escolha = messagebox.askquestion('CONFIRMAÇÃO!!!', 'Tem Certeza Que Deseja Realizar o Processamento?')
     if escolha == 'yes' and aberto == 1 :

        #Zerar o contador de saida
        aux = 0

        self.areatexto.insert(END, 'Iniciando o processamento ...\n')
        self.areatexto.update()

        self.areatexto.insert(END, 'Limpando arquivos ...\n')
        self.areatexto.update()

        # limpa a planilha
        limpa(Nome_aquivo)

        self.areatexto.insert(END, 'Alocando Matrizes ...\n')
        self.areatexto.update()

        # alocação de Matrizes
        Matglo = np.zeros((4 * nno, 4 * nno))  # Matriz global
        Fo = np.zeros((8, 1))  # Matriz de forças nodais
        delta = np.zeros((8, 1))  # Matriz de deslocamentos na barra
        Matfor = np.zeros((4 * nno, 1))  # Matriz de forças nodais ( Foreq + Fo )

        # monta-se a matriz de força e de rigidez de cada barra
        for bar in range(int(nbar)):

            self.areatexto.insert(END, 'Matriz de Rigidez e Força da barra - ' + str(bar + 1) + '\n')
            self.areatexto.update()

            # Tipo de Calculo
            tipo = int(Matbar[bar, 1])

            # nó inicial e final da barra
            noi = int(Matbar[bar, 2])
            nof = int(Matbar[bar, 3])

            # molas nos nós

            khi = Matno[noi - 1, 7]
            kvi = Matno[noi - 1, 8]
            kgi = Matno[noi - 1, 9]

            khf = 0
            kvf = 0
            kgf = 0

            if bar+1 == nbar:
                khf = Matno[nof - 1, 7]
                kvf = Matno[nof - 1, 8]
                kgf = Matno[nof - 1, 9]
                print('ok')

            # Coordenadas iniciais e finais da barra
            xi = Matno[noi - 1, 1]
            yi = Matno[noi - 1, 2]
            xf = Matno[nof - 1, 1]
            yf = Matno[nof - 1, 2]


            # Propriedades das Barras
            A = Matbar[bar, 4]  # área da barra
            I = Matbar[bar, 5]  # inércia flexional
            It = Matbar[bar, 6]  # inércia torcional
            E = Matbar[bar, 7]  # Módulo de Elasticidade
            poisson = Matbar[bar, 8]  # poisson
            kvd = Matbar[bar, 9]  # Coeficiente de mola na barra
            ptor = Matbar[bar, 10]  # porcentagem da rigidez a torção
            P = 0  # Valor para futuro cálculo de flambagem
            G = E / (2 * (1 + poisson))

            # Carregamentos na Barra
            qix = Carbar[bar, 1]  # carga inicial paralelo a barra
            qfx = Carbar[bar, 2]  # carga final paralelo a barra
            qiy = Carbar[bar, 3]  # carga inicial perpendicular a barra
            qfy = Carbar[bar, 4]  # carga final perpendicular a barra
            qig = Carbar[bar, 5]  # carga de torcão inicial na barra
            qfg = Carbar[bar, 6]  # carga de torçao final na barra


            # Carregamentos nos nós
            Fo[0, 0] = 0
            Fo[1, 0] = Carno[noi - 1, 1]
            Fo[2, 0] = Carno[noi - 1, 2]
            Fo[3, 0] = Carno[noi - 1, 3]
            Fo[4, 0] = 0
            Fo[5, 0] = 0
            Fo[6, 0] = 0
            Fo[7, 0] = 0

            if bar+1 == nbar:
                Fo[5, 0] = Carno[nof - 1, 1]
                Fo[6, 0] = Carno[nof - 1, 2]
                Fo[7, 0] = Carno[nof - 1, 3]


            L = comprimento(xi, xf, yi, yf)  # comprimento da Barra

            vsen, vcos = angulo(xi, xf, yi, yf)  # Angulo da Barra


            mge = Smge(tipo, G, It * ptor / 100, E, I, A, L, kvd, P)  # Matriz de Rigidez local

            Matcin = matriz_cinematica(vsen, vcos)  # natriz cinemática : B

            Matcint = Matcin.transpose()  # Transposta de B

            mge = Matcint.dot(mge.dot(Matcin))  # BT * re * B

            # Soma as molas nodais a matriz de Rigidez
            mge[1, 1] = mge[1, 1] + khi
            mge[2, 2] = mge[2, 2] + kvi
            mge[3, 3] = mge[3, 3] + kgi
            mge[5, 5] = mge[5, 5] + khf
            mge[6, 6] = mge[6, 6] + kvf
            mge[7, 7] = mge[7, 7] + kgf

            Mfeq = forcaequivale(tipo, L, qix, qfx, qiy, qfy, qig, qfg)  # Matriz de força equivalente

            Mfeq = Matcint.dot(Mfeq)  # Matriz de força equivalente rotacionada para o eixo global

            Matfor = forcanodal(Matfor, Fo, noi, nof, Mfeq)  # Matriz de Força nodal

            Matglo = matrizglobal(Matglo, mge, noi, nof)  # monta a matriz global

            Sub_said(mge, 1, Nome_aquivo,bar)  # imprime a matriz de rigidez local

            Sub_said(Mfeq, 2, Nome_aquivo,bar)  # imprime o vetor de carga nodal local

            aux += 9  # incrementa linha para a próxima barra

        Sub_said(Matglo, 5, Nome_aquivo,bar)  # imprime a Matriz de rigidez global
        Sub_said(Matfor, 6, Nome_aquivo,bar)  # imprime o vetor de carga globas

        self.areatexto.insert(END, 'Condições de Contorno ...\n')
        self.areatexto.update()

        #  Impoe as condições de contorno
        for nos in range(nno):
            rx = int(Matno[nos, 3])
            ry = int(Matno[nos, 4])
            rz = int(Matno[nos, 5])
            rg = int(Matno[nos, 6])
            Matglo, Matfor = vinculacao(Matglo, Matfor, nos + 1, rx, ry, rz, rg)

        Sub_said(Matglo, 3, Nome_aquivo,bar)  # imprime a Matriz de rigidez global
        Sub_said(Matfor, 4, Nome_aquivo,bar)  # imprime o vetor de carga globas

        self.areatexto.insert(END, 'Resolvendo o sistema...\n')
        self.areatexto.update()

        #  Resolução do sistema K * U = F
        desl = linalg.solve(Matglo, Matfor)

        Sub_said(desl, 7, Nome_aquivo,bar)  # imprime os deslocamentos

        self.areatexto.insert(END, 'Cálculo dos Esforços Solicitantes...\n')
        self.areatexto.update()

        # Calcula os Esforços Solicigantes
        Esforcos_solicitantes(Nome_aquivo)

        self.areatexto.insert(END, 'Fim do processamento...\n' )
        self.areatexto.update()

     else:
         messagebox.showinfo('ABOUT', 'Sem Entrada de Dados...\n')

    def Sub_Resultados(self):
        global Nome_aquivo
        if Nome_aquivo == '':
            messagebox.showinfo('ABOUT', 'Abra um arquivo Primeiro...\n')
        else:
            t = subprocess.Popen(Nome_aquivo, shell=True)
            t.wait()

if __name__ == '__main__':
    Principal = Tk()
    Tela_Principal(Principal)
    Principal.mainloop()