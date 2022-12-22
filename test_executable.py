from tkinter import *
from tkinter import ttk

root= Tk()


class Functions():
    # def on_click(text):

class Application(Functions):
    def __init__(self):
        self.root = root
        self.tela()
        self.frames()
        self.btn_1()
        self.btn_2()
        root.mainloop()
    def tela(self):
        self.root.title("Agiw Sistemas - INFORMATEC")
        self.root.configure(background= 'lightblue' ) # cor do fundo
        self.root.geometry("400x300") # tamanho da tela
        self.root.resizable(True, True) # tela responsiva
        self.root.minsize(width=400, height=300) # tamanho maximo da tela
        self.root.maxsize(width=800, height=600) # tamanho minimo da tela
    def frames(self):
        self.frame_1 = Frame(self.root, bd=4, bg= "white", highlightbackground="black", highlightthickness="3")
        self.frame_1.place(relx= 0.05, rely= 0.1, relwidth=0.9, relheight=0.5)
        #label
        self.lb_titulo = Label(self.frame_1, text="Escolha o robo", width=1, bg="light blue",
        fg='black',font=('Comic',11,"bold italic"))
        self.lb_titulo.pack()
        self.lb_titulo.place(relx= 0.08, rely= 0.01, relwidth= 0.35)
    def btn_1(self):
        self.btn1= Button(self.frame_1, text="1 - Criar_Nota_Fiscal", bd=3, bg='#8FBC8F',
                   font=('Comic',11))         
        self.btn1.place(relx= 0.2, rely=0.2, relheight=0.18) 
    def btn_2(self):
        self.btn_teste= Button(self.frame_1, text="2 - teste_robo", bd=3, bg='#8FBC8F',font=('Comic',11))
        self.btn_teste.place(relx= 0.2, rely=0.5, relheight=0.18)

Application()