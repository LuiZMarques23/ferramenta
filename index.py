
###### bibliotecas ########
from tkinter import *
from tkinter import ttk
from tkinter import messagebox## avisos mensagens
import os
import pandas as pd
from datetime import date
import numpy as np
import tkinter as tk

###########################################################################################################


# tela do sistema 
tela = Tk() ## incio da tela 
tela.geometry('980x558+200+40')#tamanho da tela
tela ['bg'] = 'dimgray'# cor da tela
tela.title('Ferramenta controle Qualidade ')# nome do tidolo
##############################################################################################################
stilo = ttk.Style()
stilo.theme_use('alt')
stilo.configure('.',font=('Arial 20'), rowheight=31)

treeviewDados = ttk.Treeview(tela,column=(1,2,3,4), show='headings')


treeviewDados.column('1', width=300 , anchor=CENTER)
treeviewDados.heading('1', text='Nome')


treeviewDados.column('2' ,anchor=CENTER)
treeviewDados.heading('2', text='Vencimento')


treeviewDados.column('3', width=200 ,anchor=CENTER)
treeviewDados.heading('3', text='Preço')

treeviewDados.column('4', anchor=CENTER)
treeviewDados.heading('4',text='Quantidade')



## grid = divite a tela em partes
##  row ´= linhas
#treeviewDados.grid(row=5, column=0,columnspan=8,sticky='NSEW' )

treeviewDados.place(x=40,y=258,height=300)


##########################################################################################################



# chamando pasta execel
controle = pd.read_excel('controle.xlsx')

#converte a coluna Nascimento em texto
controle['Vencimento'] = controle['Vencimento'].astype(str)

#criando a coluna Ano
controle['Ano'] = controle['Vencimento'].str[:4]

#criando a coluna Mes
controle['Mes'] = controle['Vencimento'].str[5:7]

#criando coluna Dia
controle['Dia'] = controle['Vencimento'].str[-2:]

########### criando data atual de hoje ################
controle['Data Atual'] = date.today()


#criandouma coluna de Data Atual para texto
controle['Data Atual'] = controle['Data Atual'].astype(str)

 #criando coluna do Ano Atual
controle['Ano Atual'] = controle['Data Atual'].str[:4]

#criando coluna Mes Atual
controle['Mes Atual'] = controle['Data Atual'].str[5:7]

#criando coluna Dia Atual
controle['Dia Atual'] = controle['Data Atual'].str[-2:]

#comparando Mes e dia 
controle['Vencimento'] = np.where((controle['Mes'] == controle['Mes Atual']) &
                                      (controle['Dia'] == controle['Dia Atual']),'vencido','')

# loc ele localiza e filtra 
controle = controle.loc[controle['Vencimento'] != '', ['Nome','Vencimento','preco','Quantidade']]


print(controle)

for linha in range(len(controle)):
    

# poupulando os itens na Treeview com os dados vencimento
    treeviewDados.insert('', 'end',
                         values=(str(controle.iloc[linha, 0] ),#nome
                                 str(controle.iloc[linha, 1] ),#vencimento
                                 str(controle.iloc[linha, 2] ),
                                 str(controle.iloc[linha, 3] )))#preço
                                 
######################################################################################### funçoes def ####################################################
# adicionar um novo itens
def deletarItemTreeview():

    itens = treeviewDados.selection()

    for item in itens:

        treeviewDados.delete(item)
        
      #### mensagens de aviso ########
        messagebox.showinfo("Aviso", "Item Deletado com Sucesso!")

        contaNumeroLinhas()

##### comando tecla sair####
def deletar():
    tela.destroy()

    

### comando mensagem  Ajuda ####
def sobre():
      messagebox.showinfo("Sobre",  'Olá! Passo 1: Aprenda navegar no sistema Antes de começar a usar o sistema, para a navegar primeiro colacar arquivo excel_controle na mesma pasta do sistema .\n\n passo 2: As informações são editadas no próprio Excel_controle que estão junto na pasta do sistema \n\n passo 3: após de cadastrar item no excel_controle aperta para salvar no excel \n\n depois sair do excel entra novamente no sistema \n\n Fim! \n\n Desenvolvido por Luiz Henrique Marques')


barramenu = tk.Menu(tela)
menu_func = tk.Menu(barramenu)# menu funcionamento 
menu_ajuda = tk.Menu(barramenu)# menu sair

tela.config(menu=barramenu)



barramenu.add_cascade(label = 'Informaçao', menu=menu_ajuda)
menu_ajuda.add_command(label = 'Sobre', command=sobre)
###########################################################################################################################################

#### criando contagen de vencimento
labeNumeroLinhas = Label(text= 'Linhas:', font='Georgia 13',fg='white', bg='purple')
labeNumeroLinhas.place(x=40, y=535, width= 145, height=25)

def contaNumeriLinha(item=''):
    numero = 0
    linha = treeviewDados.get_children(item)
    for linha in linha:
        numero += 1

    labeNumeroLinhas.config(text='Vencido:' + str(numero))

contaNumeriLinha()
#################################################################################botoes###############################################################

### criando botão deleta 
btdeleta = Button(text= 'Deletar',
                   bg = 'purple',foreground='white',font= 'Ariel 20',overrelief=RIDGE ,command= deletarItemTreeview)
btdeleta.place(x = 40, y =200, width = 140)


### botao de sair 
btsair = Button(text = 'Sair',
                 bg = 'purple',foreground='red',font='Ariel 20',overrelief=RIDGE, command= deletar)
btsair.place(x=780, y=20, width=190)



####################################################################################################################################################




tela.mainloop() # fim da tela 


