import os
from tkinter import *
from tkinter import filedialog, Label
import FillExcel

root = Tk()


class Funcs:
    def open_files(self):
        files = filedialog.askopenfilenames(parent=root, title='Selecione os jogadores',
                                            filetypes=[("HTML Files", "*.html")], initialdir=os.getcwd())
        FillExcel.create_sheets(self.input_template.get(), files)

    def open_excel(self):
        file = filedialog.askopenfilenames(parent=root, title='Selecione a planilha',
                                           filetypes=[("Excel File", "*.xlsx")], initialdir=os.getcwd())
        self.input_template.delete(0, "end")
        self.input_template.insert(0, file)

    def check_template_same_folder(self):
        files = os.listdir(os.getcwd())
        for file in files:
            if file == "Analise_de_Atributos_-_TaticaMente.xlsx":
                self.input_template.insert(0, os.path.abspath(os.path.join(os.getcwd(), file)))


class Application(Funcs):
    def __init__(self):
        self.lb_text = None
        self.bt_selecionar = None
        self.bt_abrir = None
        self.input_template = None
        self.lb_template = None
        self.frame = None
        self.root = root
        self.screen()
        self.screen_frames()
        self.widgets_frame()
        root.mainloop()

    def screen(self):
        self.root.title("Cadastro de Clientes")
        self.root.configure(background='#2e0d46')
        self.root.geometry("700x250")
        self.root.resizable(False, False)

    def screen_frames(self):
        self.frame = Frame(self.root, bd=4, bg='#dfe3ee', highlightbackground='#6e56dd', highlightthickness=3)
        self.frame.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)

    def widgets_frame(self):
        # label e input do template
        self.lb_template = Label(self.frame, text="Template Excel", bg='#dfe3ee', fg='#1f0930')
        self.lb_template.place(relx=0.05, rely=0.05)
        self.input_template = Entry(self.frame)
        self.input_template.place(relx=0.05, rely=0.15, relwidth=0.7)
        self.check_template_same_folder()

        # botao abrir
        self.bt_abrir = Button(self.frame, command=self.open_excel, text="Abrir", bd=2, bg='#6e56dd', fg='white',
                               font=('verdana', 8, 'bold'))
        self.bt_abrir.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        # botao para selecionar os arquivos html e gerar os correspondentes arquivos excel
        self.bt_selecionar = Button(self.frame, command=self.open_files, text="Selecionar jogadores", bd=2,
                                    bg='#6e56dd', fg='white', font=('verdana', 8, 'bold'))
        self.bt_selecionar.place(relx=0.35, rely=0.3, relwidth=0.25, relheight=0.15)

        # label com texto
        self.lb_text = Label(self.frame, text='Utilize o botão "abrir" para selecionar o local da planilha excel se '
                                              'necessário, então utilize o botão \n"selecionar jogadores" para escolher '
                                              'um ou mais arquivos html e gerar as planilhas correspondentes.',
                             bg='#dfe3ee', fg='#1f0930')
        self.lb_text.place(relx=0.05, rely=0.6)


Application()
