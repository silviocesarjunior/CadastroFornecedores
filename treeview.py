#TreeView
from tkinter import *
from tkinter import messagebox
from tkinter import ttk

janela = Tk()

janela.geometry("750x350")

#theme_use - alt, default, classic
stiloDaTreeView= ttk.Style()
stiloDaTreeView.theme_use("alt")
stiloDaTreeView.configure(".", font = "Arial 10")

#column - criando as colunas
treeViewDados = ttk.Treeview(janela, column=(1,2,3,4,5,6,7,8), show="headings" )

treeViewDados.column("1", anchor=CENTER)
treeViewDados.heading("1", text="Nome Completo")

treeViewDados.column("2", anchor=CENTER)
treeViewDados.heading("2", text="Fornecedor")

treeViewDados.column("3", anchor=CENTER)
treeViewDados.heading("3", text="Placa")

treeViewDados.column("4", anchor=CENTER)
treeViewDados.heading("4", text="Local de descarga")

treeViewDados.column("5", anchor=CENTER)
treeViewDados.heading("5", text="Nota fiscal")

treeViewDados.column("6", anchor=CENTER)
treeViewDados.heading("6", text="Hora inicial")

treeViewDados.column("7", anchor=CENTER)
treeViewDados.heading("7", text="Hora final")

treeViewDados.column("8", anchor=CENTER)
treeViewDados.heading("8", text="Observações")

#inserindo dados na treeview
treeViewDados.insert("", "end", text="1", values=("Allan",29, "ERU-1010", "Masculino"))
treeViewDados.insert("", "end", text="2", values=("Ana",19, "AGU-3030", 19, "Feminino"))
treeViewDados.insert("", "end", text="3", values=("Roger", "TET-2200", 40, "Masculino"))


treeViewDados.pack()
janela.mainloop()