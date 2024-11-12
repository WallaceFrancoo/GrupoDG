import os
import main
import BancoDeDados
from tkinter import *
from PIL import Image
import customtkinter
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename

def abrir_janela_depara():
    def cadastrar():
        DE = entry_de.get()
        PARA = entry_para.get()
        BANCO = combo_opcoes.get()

        if DE and PARA and BANCO:
            confirmacao = messagebox.askokcancel(
                "Confirmação", f"Você digitou o historico {DE} para ir para a conta {PARA} no banco de dados da empresa {BANCO}. Deseja confirmar?"
            )
            if confirmacao:
                if BANCO == '840':
                    BancoDeDados.adicionar_DadosChempack(DE, PARA)
                elif BANCO == '841':
                    BancoDeDados.adicionar_DadosCertificadora(DE, PARA)
                elif BANCO == '842':
                    BancoDeDados.adicionar_DadosDGCompliance(DE, PARA)
                elif BANCO == '843':
                    BancoDeDados.adicionar_DadosDGServicos(DE, PARA)
                elif BANCO == "TODAS":
                    BancoDeDados.adicionar_DadosGeral(DE, PARA)
        else:
            messagebox.showwarning("Atenção", "Preencha todos os campos!")

    janela_depara = customtkinter.CTkToplevel()
    janela_depara.title("Cadastrar DePara")
    janela_depara.geometry("500x200")

    janela_depara.focus_force()
    janela_depara.grab_set()

    label_de = customtkinter.CTkLabel(janela_depara, text="DE:")
    label_de.grid(row=0, column=0, padx=20, pady=10)
    entry_de = customtkinter.CTkEntry(janela_depara,width=200)
    entry_de.grid(row=0, column=1, padx=20, pady=10)

    label_para = customtkinter.CTkLabel(janela_depara, text="PARA:")
    label_para.grid(row=1, column=0, padx=20, pady=10)
    entry_para = customtkinter.CTkEntry(janela_depara, width=200)
    entry_para.grid(row=1, column=1, padx=20, pady=10)

    # Adiciona a caixa de seleção
    label_opcoes = customtkinter.CTkLabel(janela_depara, text="Cadastrar no banco de dados da: ")
    label_opcoes.grid(row=2, column=0, padx=20, pady=10)
    combo_opcoes = customtkinter.CTkComboBox(janela_depara,
                                             values=["840", "841", "842", "843", "TODAS"])
    combo_opcoes.grid(row=2, column=1, padx=20, pady=10)
    combo_opcoes.set("843")  # Define uma opção padrão
    combo_opcoes.configure(state="readonly")

    botao_cadastrar = customtkinter.CTkButton(janela_depara, text="Cadastrar", command=cadastrar)
    botao_cadastrar.grid(row=3, column=0, columnspan=2, pady=20)

self = customtkinter.CTk()
self.title("Conversão Excel Bradesco com identificação - Grupo Concepta!")
self.geometry("600x450")

self.grid_rowconfigure(0, weight=1)
self.grid_columnconfigure(1, weight=1)

image_path =  "Z:\Diversos\Projetos T.I\GrupoDG\imagens"
logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "Transparente2.png")), size=(26,26))
large_test_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "Simbolo.png")), size=(150, 100))
image_icon_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "pdf.png")), size=(20, 20))
image_icon_image2 = customtkinter.CTkImage(Image.open(os.path.join(image_path, "excel.png")), size=(20, 20))

# create navigation frame = column 0
navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
navigation_frame.grid(row=0, column=0, sticky="nsew")
navigation_frame.grid_rowconfigure(6, weight=1)
navigation_frame_label = customtkinter.CTkLabel(navigation_frame, text="  Sergecont Contabilidade", image=logo_image,
                                                compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

textbox = customtkinter.CTkTextbox(navigation_frame,
                                   corner_radius=10,
                                   width=100,
                                   fg_color="transparent",
                                   wrap="word")
textbox.grid(row=1, column=0, rowspan=6, padx=(10, 10), pady=(10, 10), sticky="nsew")
textbox.insert("0.0", "Segue aplicativo para conversão do extrato bancario Bradesco para TXT\n\n"
                      "1. Clique no botão Selecionar Planilhas!\n\n"
                      "2. Digite o mês necessario\n\n"
                      "3. Selecione o arquivo do extrato bancario primeiro e após seleciona a planilha de identificações enviadas pelo cliente\n\n"
                      "4. Aguarde até aparecer a mensagem ( - Os documentos estão prontos para serem gerados! - )\n\n"
                      "5. Após isso clique em gerar Arquivo TXT ele irá gerar 5 sendo eles:"
                      "* Extrato Bancario da empresa\n* Pagamentos entre as empresas!\nCada arquivo terá que ser importado em sua respectiva empresa!\n\n"
                      "6. O relatorio de erros são as contas que estarão em acertos assim que importado a "
                      "O Programa foi parametrizado para ler o texto que contem no banco de dados!\n\n"
                      "Qualquer duvida entrar em contato com o T.I\n\n"
                      "At., Wallace"

               )



home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
home_frame.grid_columnconfigure(0, weight=1)
home_frame_large_image_label = customtkinter.CTkLabel(home_frame, text="", image=large_test_image)
home_frame_large_image_label.grid(row=0, column=0, columnspan=2, padx=10, pady=15)

combo_opcoes = customtkinter.CTkComboBox(home_frame,
                                             values=["840", "841", "842", "843"])
combo_opcoes.grid(row=2, column=0, padx=20, pady=10)
combo_opcoes.configure(state="readonly")


botao_planilha = customtkinter.CTkButton(home_frame, text="Selecionar Planilhas", image=image_icon_image2, compound="right",
                                    command=lambda: main.buscarArquivo(combo_opcoes.get()))
botao_planilha.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)
botao_planilha.grid(row=3, column=0, pady=50, columnspan=2)





botao_gerar = customtkinter.CTkButton(home_frame, text="Gerar Arquivo TXT", compound="right",
                                      command=lambda: main.gerarArquivo(combo_opcoes.get()))
botao_gerar.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER)
botao_gerar.grid(row=4, column=0, pady=0, columnspan=2)

frame_cadastro = customtkinter.CTkFrame(home_frame,fg_color="transparent")
frame_cadastro.grid(row=5, column=0, pady=10, columnspan=2)

botao_cadastrodePara = customtkinter.CTkButton(frame_cadastro, text="Cadastrar dePara", compound="left",command=abrir_janela_depara)
botao_cadastrodePara.grid(row=0, column=0, padx=10, pady=10)

home_frame.grid(row=0, column=1, sticky="nsew")

self.mainloop()