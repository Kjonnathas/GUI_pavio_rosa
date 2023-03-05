### importação das bibliotecas que serão utilizadas para o projeto ###

import customtkinter as ctk
from tkinter.messagebox import showinfo, showwarning, showerror, askyesno
from tkinter.filedialog import askdirectory
from tkinter import PhotoImage
import win32com.client as win32
import pandas as pd
import numpy as np
import webbrowser
import pyodbc
from datetime import datetime
from time import sleep
from warnings import filterwarnings
from os import environ, getcwd
from re import compile, fullmatch, sub

# --------------------------------------

### ignorando os avisos para não poluir o terminal ###

filterwarnings("ignore")

### criando a janela principal da aplicação ###

# aplicando as configurações da tela

ctk.set_appearance_mode("light")

janela_principal = ctk.CTk(fg_color="#F7ADAF")
janela_principal.title("Pavio Rosa - Sistema de Cadastro")
janela_principal.geometry("360x400+500+150")
janela_principal.iconbitmap("icone.ico")
janela_principal.resizable(width=False, height=False)

def fn_tela_esqueci_minha_senha():

     # configurando a tela de recuperação de senha

     janela_recuperar_senha = ctk.CTk(fg_color="#F7ADAF")
     janela_recuperar_senha.title("Pavio Rosa - Sistema de Cadastro")
     janela_recuperar_senha.geometry("340x180+500+150")
     janela_recuperar_senha.iconbitmap("icone.ico")
     janela_recuperar_senha.resizable(width=False, height=False)

     # função para validar a recuperação de senha

     def fn_recuperar_senha():

          with open(file="credenciais.txt", mode="r", encoding="utf-8") as arq_credenciais:
               credenciais = arq_credenciais.readline().split(",")
               senha = credenciais[1].strip()
               arq_credenciais.close()

          if "gmail" in en_email.get():
               url = "www.gmail.com"
          elif "outlook" in en_email.get():
               url = "www.outlook.com.br"
          else:
               url = "www.yahoo.com"

          try:
               outlook = win32.Dispatch("outlook.application")
               email = outlook.CreateItem(0)
               email.To = en_email.get()
               email.Subject = "Pavio Rosa - Recuperação de senha"
               email.HTMLBody = f'''
                                   <p> Prezada Juliana, </p>
                                   <p> Segue a sua senha, conforme solicitado: <b> {senha} </b> </p>
                              '''
               email.Send()
               showinfo(title="Pavio Rosa - Atenção", message="E-mail enviado com sucesso!\n\nVocê será redirecionada para a página do seu e-mail.")
               sleep(1)
               showinfo(title="Pavio Rosa - Atenção", message="O programa fechará dentro de alguns instantes!\n\nFavor, abrir o programa novamente e realizar o login.")
               sleep(3)
               janela_recuperar_senha.destroy()
               webbrowser.open(url=url)
          except:
               showerror(title="Pavio Rosa - Atenção", message="Houve um erro durante o envio do e-mail de recuperação de senha!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos")

     # criação dos widgets da tela de recuperação de senha

     en_email = ctk.StringVar()

     lb_1 = ctk.CTkLabel(master=janela_recuperar_senha, 
                         text="Recuperação de senha", 
                         text_color="#e35d6a", 
                         font=("Helvetica", 18,'bold'))

     lb_2 = ctk.CTkLabel(master=janela_recuperar_senha, 
                         text="Informe o seu e-mail", 
                         text_color="#e35d6a", 
                         font=("Helvetica", 14,'bold'))

     en_1 = ctk.CTkEntry(master=janela_recuperar_senha, 
                         width=200,
                         border_width=1,
                         textvariable=en_email)

     en_2 = ctk.CTkButton(master=janela_recuperar_senha, 
                         font=("Helvetica", 12, "bold"), 
                         text="RECUPERAR SENHA", 
                         width=150, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_recuperar_senha)

     # posicionando os widgtes da tela de alteração de senha

     lb_1.place(x=70, y=10)
     lb_2.place(x=95, y=40)
     en_1.place(x=70, y=70)
     en_2.place(x=90, y=120)

     janela_recuperar_senha.mainloop()

### criando a janela para alteração de senha ###

def fn_tela_alterar_senha():

     # configurando a tela de alterar senha

     janela_alterar_senha = ctk.CTk(fg_color="#F7ADAF")
     janela_alterar_senha.title("Pavio Rosa - Sistema de Cadastro")
     janela_alterar_senha.geometry("360x280+500+150")
     janela_alterar_senha.iconbitmap("icone.ico")
     janela_alterar_senha.resizable(width=False, height=False)

     # função para validar a alteração de senha

     def fn_alterar_senha():
          senha_atual = en_senha_atual.get()
          senha_nova = en_senha_nova.get()
          with open(file="credenciais.txt", mode="r", encoding="utf-8") as arq_credenciais:
               credenciais = arq_credenciais.readline().split(",")
               senha_anterior = credenciais[1].strip()
               lembrar_login = ",0" if credenciais[2] == 0 else ",1"
               arq_credenciais.close()
               if senha_atual == "" and senha_nova == "":
                    showwarning(title="Pavio Rosa - Atenção", message="Os dois campos precisam estar preenchidos!")
               elif senha_atual == "":
                    showwarning(title="Pavio Rosa - Atenção", message="O campo da senha atual não foi preenchido!")
               elif senha_nova == "":
                    showwarning(title="Pavio Rosa - Atenção", message="O campo da senha nova não foi preenchido!")
               elif senha_anterior == senha_atual:
                    with open(file="credenciais.txt", mode="r", encoding="utf-8") as arq_credenciais:
                         arq_credenciais.write("juliana.pavio_rosa," + senha_nova + lembrar_login)
                         arq_credenciais.close()
                         janela_alterar_senha.destroy()
                         fn_tela_menu()
               else:
                    showwarning(title="Pavio Rosa - Atenção", message="A senha atual está incorreta!")
                    en_senha_atual.set("")
                    en_senha_nova.set("")

     # criação dos widgtes da tela de alteração de senha

     img_logo = PhotoImage(file="logo.png")

     en_senha_atual = ctk.StringVar()

     en_senha_nova = ctk.StringVar()

     lb_1 = ctk.CTkLabel(master=janela_alterar_senha, 
                         text="Alteração de senha", 
                         text_color="#e35d6a", 
                         font=("Helvetica", 18,'bold'))

     lb_2 = ctk.CTkLabel(master=janela_alterar_senha,
                         text="Senha atual",
                         text_color="#e35d6a", 
                         font=("Helvetica", 12, "bold"))

     lb_3 = ctk.CTkLabel(master=janela_alterar_senha,
                         text="Senha nova",
                         text_color="#e35d6a", 
                         font=("Helvetica", 12, "bold"))

     lb_4 = ctk.CTkLabel(master=janela_alterar_senha, 
                         text="", 
                         image=img_logo)

     en_1 = ctk.CTkEntry(master=janela_alterar_senha, 
                         width=200,
                         border_width=1,
                         show="*",
                         textvariable=en_senha_atual)

     en_2 = ctk.CTkEntry(master=janela_alterar_senha, 
                         width=200,
                         border_width=1, 
                         show="*",
                         textvariable=en_senha_nova)

     bt_1 = ctk.CTkButton(master=janela_alterar_senha, 
                         font=("Helvetica", 12, "bold"), 
                         text="ALTERAR SENHA", 
                         width=150, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_alterar_senha)

     # posicionando os widgets na tela de alteração de senha

     lb_1.place(x=90, y=10)
     lb_2.place(x=70, y=40)
     lb_3.place(x=70, y=95)
     lb_4.place(x=80, y=230)
     en_1.place(x=70, y=65)
     en_2.place(x=70, y=120)
     bt_1.place(x=90, y=170)

     janela_alterar_senha.mainloop()

# criando a tela de menu

def fn_tela_menu():

     janela_menu = ctk.CTk(fg_color="#F7ADAF")
     janela_menu.title("Pavio Rosa - Sistema de Cadastro")
     janela_menu.geometry("360x220+500+150")
     janela_menu.iconbitmap("icone.ico")
     janela_menu.resizable(width=False, height=False)

     # função para prosseguir para a tela de cadastro de clientes

     def fn_tela_clientes():
          
          janela_menu.destroy()
          fn_tela_cadastro_clientes()

     # função para prosseguir para a tela de cadastro de produtos

     def fn_tela_produtos():

          janela_menu.destroy()
          fn_tela_cadastro_produtos()

     # função para prosseguir para a tela de cadastro de transações

     def fn_tela_transacao():

          janela_menu.destroy()
          fn_tela_transacoes()

     # criação dos widgets da tela de menu

     lb_1 = ctk.CTkLabel(master=janela_menu,
                         width=100,
                         text="MENU DE CADASTRO",
                         text_color="#e35d6a",
                         font=("Helvetica", 18, "bold"),
                         anchor="center")

     lb_2 = ctk.CTkLabel(master=janela_menu,
                         width=100,
                         text="Selecione o cadastro a realizar",
                         text_color="#e35d6a",
                         font=("Helvetica", 14, "bold"),
                         anchor="center")

     bt_1 = ctk.CTkButton(master=janela_menu,
                         font=("Helvetica", 12, "bold"), 
                         text="CLIENTES", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_tela_clientes)

     bt_2 = ctk.CTkButton(master=janela_menu,
                         font=("Helvetica", 12, "bold"), 
                         text="PRODUTOS", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_tela_produtos)

     bt_3 = ctk.CTkButton(master=janela_menu,
                         font=("Helvetica", 12, "bold"), 
                         text="TRANSAÇÃO", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_tela_transacao)

     # posicionando os widgtes da tela de menu

     lb_1.place(x=85, y=20)
     lb_2.place(x=70, y=60)
     bt_1.place(x=110, y=100)
     bt_2.place(x=110, y=140)
     bt_3.place(x=110, y=180)

     janela_menu.mainloop()

# criação da tela de cadastro de clientes

def fn_tela_cadastro_clientes():

     janela_clientes = ctk.CTk(fg_color="pink")
     janela_clientes.title("Pavio Rosa - Sistema de Cadastro")
     janela_clientes.geometry("740x340+300+150")
     janela_clientes.iconbitmap("icone.ico")
     janela_clientes.resizable(width=False, height=False)

     # função para tratamento dos dados de cadastro fornecidos pelo usuário

     def fn_tratar_respostas():

          global stop_email
          global stop_data_nascimento
          global stop_telefone
          global stop_cep

          var_email = en_email.get().strip()
          padrao_email = "[.|_]*[a-zA-Z0-9]+[.|_]*[\w]+@[\w]+\.[a-z]+"
          checar_email = fullmatch(compile(padrao_email), var_email)
          if var_email == "":
               pass
               stop_email = False
          elif checar_email == None:
               showwarning(title="Pavio Rosa - Atenção", message="O e-mail fornecido é inválido!")
               en_email.set("")
               stop_email = True        
          else:
               en_email.set(var_email)
               stop_email = False

          var_data = en_data_nascimento.get().strip()
          padrao_data_nascimento = "\d{2}/\d{2}/\d{4}"
          checar_data_nascimento = fullmatch(compile(padrao_data_nascimento), var_data)
          try:
               data_convertida = datetime.strptime(var_data, "%d/%m/%Y")
          except:
               pass
          if var_data == "":
               pass
               stop_data_nascimento = False
          elif data_convertida > datetime.today():
               showwarning(title="Pavio Rosa - Atenção", message="A data de nascimento não pode ser maior que a data de hoje!")
               en_data_nascimento.set("")
               stop_data_nascimento = True
          elif checar_data_nascimento == None:
               showwarning(title="Pavio Rosa - Atenção", message="A data de nascimento fornecida é inválida!")
               en_data_nascimento.set("")
               stop_data_nascimento = True
          else:
               en_data_nascimento.set(var_data)
               stop_data_nascimento = False

          var_telefone = en_telefone.get().strip()
          if var_telefone == "":
               pass
               stop_telefone = False
          else:
               var_tel = ",".join([num for num in var_telefone if num.isnumeric()]).replace(",", "")
               if len(var_tel) == 11:
                    var_tel = "(" + str(var_tel[:2]) + ")" + " " + str(var_tel[2:3]) + " " + str(var_tel[3:7]) + "-" + str(var_tel[-4:])
                    en_telefone.set(var_tel)
                    stop_telefone = False
               else:
                    showwarning(title="Pavio Rosa - Atenção", message="O número de telefone informado é inválido!")
                    en_telefone.set("")
                    stop_telefone = True

          var_cep = en_cep.get().strip()
          if var_cep == "":
               pass
               stop_cep = False
          else:
               variavel_cep = ",".join([num for num in var_cep if num.isnumeric()]).replace(",", "")
               if len(variavel_cep) != 8:
                    showwarning(title="Pavio Rosa - Atenção", message="O CEP informado é inválido!")
                    en_cep.set("")
                    stop_cep = True
               else:
                    variavel_cep = variavel_cep[:5] + "-" + variavel_cep[-3:]
                    en_cep.set(variavel_cep)
                    stop_cep = False

     # função para o cadastramento dos dados dos clientes no banco de dados

     def fn_cadastrar_cliente():

          # chamando a função de tratamento das respostas

          fn_tratar_respostas()

          # validando se os dados estão de acordo para não cadastrar errado

          if stop_email != True and stop_data_nascimento != True and stop_telefone != True and stop_cep != True:

               var_nome = en_nome.get().strip().title()
               var_sobrenome = en_sobrenome.get().strip().title()
               var_data_nascimento = en_data_nascimento.get().strip()
               var_genero = en_genero.get().strip().capitalize()
               var_email = en_email.get().strip()
               var_telefone = en_telefone.get().strip()
               var_rua = en_rua.get().strip().title()
               var_numero = en_numero.get().strip()
               var_bairro = en_bairro.get().strip().title()
               var_cidade = en_cidade.get().strip().title()
               var_estado = en_estado.get().strip()
               var_cep = en_cep.get().strip()

               try:
                    dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
                    )
                    conn = pyodbc.connect(dados_conexao)
                    cursor = conn.cursor()
                    cursor.execute(f'''
                                        INSERT INTO dClientes
                                             (Data_cadastro, 
                                             Nome, 
                                             Sobrenome, 
                                             Data_nascimento, 
                                             Genero, Email, 
                                             Telefone, Rua, 
                                             Numero, 
                                             Bairro, 
                                             Cidade, 
                                             Estado, 
                                             CEP)
                                        VALUES
                                             (FORMAT(GETDATE(), 'dd/MM/yyyy HH:mm:ss'), 
                                             '{var_nome}', 
                                             '{var_sobrenome}', 
                                             '{var_data_nascimento}', 
                                             '{var_genero}', 
                                             '{var_email}', 
                                             '{var_telefone}', 
                                             '{var_rua}', 
                                             '{var_numero}', 
                                             '{var_bairro}', 
                                             '{var_cidade}', 
                                             '{var_estado}', 
                                             '{var_cep}')
                    ''')
                    cursor.commit()
                    cursor.close()
                    conn.close()
                    showinfo(title="Pavio Rosa - Atenção", message="Cliente cadastrado com sucesso!")

                    # chama a função de limpar os campos que foram preenchidos
                    fn_limpar_campos()
               except:
                    showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante o cadastrado do cliente!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")
          else:
               showwarning(title="Pavio Rosa - Atenção", message="Não foi possível cadastrar o cliente!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     # função para exportar relatório em excel dos dados de clientes cadastrados na base de dados

     def fn_exportar_tabela():

          try:
               dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
               )
               conn = pyodbc.connect(dados_conexao)
               df = pd.read_sql_query(
                                        sql="SELECT * FROM dClientes", 
                                        con=conn
                                   )
               conn.close()
               try:
                    df['Data_nascimento'] = df['Data_nascimento'].astype(np.datetime64)
                    df['Data_nascimento'] = df['Data_nascimento'].dt.strftime("%d/%m/%Y")
                    df['Data_cadastro'] = df['Data_cadastro'].dt.strftime("%d/%m/%Y %H:%M:%S")
                    df['Data_atualizacao'] = df['Data_atualizacao'].dt.strftime("%d/%m/%Y %H:%M:%S")
               except:
                    pass
               diretorio = askdirectory(title="Selecione o local para salvar o arquivo",
                                        initialdir=fr"C:\Users\{environ['USERNAME']}\Downloads",
                                        mustexist=True)
               if diretorio == "" or diretorio == None:
                    sleep(1)
                    var_confirmar = askyesno(title="Pavio Rosa - Atenção", message="Você não selecionou nenhum local para salvar o arquivo. Deseja mesmo cancelar a operação?", default="no")
                    if var_confirmar == True:
                         showinfo(title="Pavio Rosa - Atenção", message="Operação cancelada com sucesso!")
                    else:     
                         diretorio = getcwd()
                         df.to_excel(diretorio + "\Pavio_Rosa_clientes.xlsx", index=False, sheet_name="Clientes")
                         showinfo(title="Pavio Rosa - Atenção", message=f"O arquivo foi salvo em:\n\n{diretorio}")
               else:
                    df.to_excel(diretorio + "\Pavio_Rosa_clientes.xlsx", index=False, sheet_name="Clientes")
                    showinfo(title="Pavio Rosa - Atenção", message="Relatório de clientes exportado com sucesso!")
          except:
               showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante o processo de exportação dos dados da base de clientes!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     # função para realizar a atualização dos dados dos clientes no banco de dados

     def fn_atualizar_dados():

          fn_tratar_respostas()

          var_nome = en_nome.get().strip().title()
          var_sobrenome = en_sobrenome.get().strip().title()
          var_data_nascimento = en_data_nascimento.get().strip()
          var_genero = en_genero.get().strip().capitalize()
          var_email = en_email.get().strip()
          var_telefone = en_telefone.get().strip()
          var_rua = en_rua.get().strip().title()
          var_numero = en_numero.get().strip()
          var_bairro = en_bairro.get().strip().title()
          var_cidade = en_cidade.get().strip().title()
          var_estado = en_estado.get().strip()
          var_cep = en_cep.get().strip()

          if en_id_cliente.get().strip() == "":
                    showwarning(title="Pavio Rosa - Atenção", message="O id cliente não foi informado!")
          else:
               var_id_cliente = int(en_id_cliente.get().strip())
               try:
                    dados_conexao = (
                         "Driver={SQL Server};"
                         "Server=DESKTOP-OJGFM82;"
                         "Database=db_pavio_rosa"
                    )
                    
                    conn = pyodbc.connect(dados_conexao)
                    cursor = conn.cursor()
                    cursor.execute("SELECT * FROM dProdutos")
                    valores = cursor.fetchall()
                    lista_id_clientes = [valor[0] for valor in valores]
                    if en_id_cliente.get().strip() in lista_id_clientes:
                         pass
                    else:
                         cursor.close()
                         conn.close()
                         showwarning(title="Pavio Rosa - Atenção", message="O id cliente informado não está cadastrado na base de dados!")
                    if var_nome != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Nome = '{var_nome}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_sobrenome != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Sobrenome = '{var_sobrenome}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_data_nascimento != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Data_nascimento = '{var_data_nascimento}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_genero != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Genero = '{var_genero}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_email != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Email = '{var_email}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_telefone != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Telefone = '{var_telefone}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_rua != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Rua = '{var_rua}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_numero != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Numero = '{var_numero}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_bairro != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Bairro = '{var_bairro}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_cidade != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Cidade = '{var_cidade}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_estado != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET Estado = '{var_estado}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    if var_cep != "":
                         cursor.execute(f'''
                              UPDATE dClientes
                              SET CEP = '{var_cep}'
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    cursor.execute(f'''
                              UPDATE dClientes
                              SET Data_atualizacao = FORMAT(GETDATE(), 'dd/MM/yyyy HH:mm:ss')
                              WHERE ID_cliente = '{var_id_cliente}'
                         ''')
                    cursor.commit()
                    cursor.close()
                    conn.close()
                    showinfo(title="Pavio Rosa - Atenção", message="Dados atualizados com sucesso!")
                    fn_limpar_campos()
               except:
                    showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante a atualização dos dados da base de clientes!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")
                    fn_limpar_campos()

     # função para limpar os campos que foram preenchidos

     def fn_limpar_campos():

          en_nome.set("")
          en_sobrenome.set("")
          en_data_nascimento.set("")
          en_genero.set("")
          en_email.set("")
          en_telefone.set("")
          en_rua.set("")
          en_numero.set("")
          en_bairro.set("")
          en_cidade.set("")
          en_estado.set("")
          en_cep.set("")
          en_id_cliente.set("")

     # função para retornar ao menu

     def fn_voltar_menu():

          janela_clientes.destroy()
          fn_tela_menu()

     # criando dos widgets da tela de cadastro de clientes

     en_nome = ctk.StringVar()

     en_sobrenome = ctk.StringVar()

     en_data_nascimento = ctk.StringVar()

     en_genero = ctk.StringVar()

     en_email = ctk.StringVar()

     en_telefone = ctk.StringVar()

     en_rua = ctk.StringVar()  
     
     en_numero = ctk.StringVar() 
     
     en_bairro = ctk.StringVar()

     en_cidade = ctk.StringVar() 
     
     en_estado = ctk.StringVar() 
     
     en_cep = ctk.StringVar() 
     
     en_id_cliente = ctk.StringVar()

     lb_1 = ctk.CTkLabel(master=janela_clientes, 
                         text="CADASTRO DE CLIENTES",
                         text_color="#e35d6a",
                         font=('Helvetica', 26, "bold"))

     lb_2 = ctk.CTkLabel(master=janela_clientes, 
                         text="Nome",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_3 = ctk.CTkLabel(master=janela_clientes, 
                         text="Sobrenome",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_4 = ctk.CTkLabel(master=janela_clientes, 
                         text="Data de Nascimento",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_5 = ctk.CTkLabel(master=janela_clientes, 
                         text="Gênero",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_6 = ctk.CTkLabel(master=janela_clientes, 
                         text="E-mail",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_7 = ctk.CTkLabel(master=janela_clientes, 
                         text="Telefone",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_8 = ctk.CTkLabel(master=janela_clientes, 
                         text="Rua",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_9 = ctk.CTkLabel(master=janela_clientes, 
                         text="Número",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_10 = ctk.CTkLabel(master=janela_clientes, 
                         text="Bairro",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_11 = ctk.CTkLabel(master=janela_clientes, 
                         text="Cidade",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_12 = ctk.CTkLabel(master=janela_clientes, 
                         text="Estado",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_13 = ctk.CTkLabel(master=janela_clientes, 
                         text="CEP",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_14 = ctk.CTkLabel(master=janela_clientes, 
                         text="ID Cliente",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_15 = ctk.CTkLabel(master=janela_clientes, 
                         text="Só informar o ID Cliente em caso de atualização*",
                         text_color="#e35d6a",
                         font=('Helvetica', 10, "bold"))

     en_1 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_nome)

     en_2 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_sobrenome)

     en_3 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_data_nascimento)

     en_4 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_email)

     en_5 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_telefone)

     en_6 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_rua)

     en_7 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_numero)

     en_8 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_bairro)

     en_9 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_cidade)
     
     en_10 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_cep)

     en_11 = ctk.CTkEntry(master=janela_clientes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_id_cliente)   

     cbb_1 = ctk.CTkComboBox(master=janela_clientes,
                              width=150,
                              height=20,
                              border_width=1,
                              values=["Feminino", "Masculino", "Transgênero", "Neutro", "Não-Binário", "Agênero", "Pangênero"],
                              button_color="#e35d6a",
                              dropdown_fg_color="pink",
                              dropdown_text_color="#e35d6a",
                              dropdown_hover_color="#c98276",
                              font=('Helvetica', 12),
                              justify="left",
                              variable=en_genero)

     cbb_2 = ctk.CTkComboBox(master=janela_clientes,
                              width=150,
                              height=20,
                              border_width=1,
                              values=["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"],
                              
                              button_color="#e35d6a",
                              dropdown_fg_color="pink",
                              dropdown_text_color="#e35d6a",
                              dropdown_hover_color="#c98276",
                              font=('Helvetica', 12),
                              justify="left",
                              variable=en_estado)

     bt_1 = ctk.CTkButton(master=janela_clientes,
                         font=("Helvetica", 12, "bold"), 
                         text="CADASTRAR", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_cadastrar_cliente)

     bt_2 = ctk.CTkButton(master=janela_clientes,
                         font=("Helvetica", 12, "bold"), 
                         text="ATUALIZAR DADOS", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_atualizar_dados)

     bt_3 = ctk.CTkButton(master=janela_clientes,
                         font=("Helvetica", 12, "bold"), 
                         text="GERAR RELATÓRIO", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_exportar_tabela)

     bt_4 = ctk.CTkButton(master=janela_clientes,
                         font=("Helvetica", 12, "bold"), 
                         text="LIMPAR CAMPOS", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_limpar_campos)

     bt_5 = ctk.CTkButton(master=janela_clientes,
                         font=("Helvetica", 12, "bold"), 
                         text="MENU", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_voltar_menu)

     # posicionando os widgets da tela de cadastro de clientes

     lb_1.place(x=200, y=20)
     lb_2.place(x=10, y=78)
     lb_3.place(x=220, y=78)
     lb_4.place(x=455, y=78)
     lb_5.place(x=10, y=118)
     lb_6.place(x=220, y=118)
     lb_7.place(x=455, y=118)
     lb_8.place(x=10, y=158)
     lb_9.place(x=220, y=158)
     lb_10.place(x=455, y=158)
     lb_11.place(x=10, y=198)
     lb_12.place(x=220, y=198)
     lb_13.place(x=455, y=198)
     lb_14.place(x=10, y=238)
     lb_15.place(x=10, y=260)
     en_1.place(x=60, y=80)
     en_2.place(x=295, y=80)
     en_3.place(x=580, y=80)
     en_4.place(x=295, y=120)
     en_5.place(x=580, y=120)
     en_6.place(x=60, y=160)
     en_7.place(x=295, y=160)
     en_8.place(x=580, y=160)
     en_9.place(x=60, y=200)
     en_10.place(x=580, y=200)
     en_11.place(x=80, y=240)
     cbb_1.place(x=60, y=120)
     cbb_2.place(x=295, y=200)
     bt_1.place(x=10, y=300)
     bt_2.place(x=140, y=300)
     bt_3.place(x=280, y=300)
     bt_4.place(x=420, y=300)
     bt_5.place(x=550, y=300)

     janela_clientes.mainloop()

# criando a tela de cadastro de produtos

def fn_tela_cadastro_produtos():

     janela_produtos = ctk.CTk(fg_color="pink")
     janela_produtos.title("Pavio Rosa - Sistema de Cadastro")
     janela_produtos.geometry("670x220+300+150")
     janela_produtos.iconbitmap("icone.ico")
     janela_produtos.resizable(width=False, height=False)

     def fn_tratar_respostas():

          global stop_preco
          global stop_cod_produto
          global stop_produto

          var_preco = en_preco.get().strip()
          padrao_preco = "[\d]+[\.|,]?[\d]+"
          checar_preco = fullmatch(compile(padrao_preco), var_preco)
          if var_preco == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo de preço!")
               stop_preco = True
          elif "R$" in var_preco:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de preço só aceita números.\n\nSe desejar informar os centavos, utilizar a vírgula ou o ponto.")
               stop_preco = True
          elif "-" in var_preco:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de preço não aceita valores negativos!")
               stop_preco = True
          elif checar_preco == None:
               showwarning(title="Pavio Rosa - Atenção", message="O formato de preço passado é inválido!")
               stop_preco = True
          else:
               en_preco.set(var_preco)
               stop_preco = False

          var_produto = en_produto.get().strip()
          if var_produto == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo produto!")
               stop_produto = True
          else:
               en_produto.set(var_produto)
               stop_produto = False

          var_cod_produto = en_cod_produto.get().strip().upper()
          if var_cod_produto == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo de código do produto")
               stop_cod_produto = True
          else:
               en_cod_produto.set(var_cod_produto)
               stop_cod_produto = False

     # realiza a conexão com o banco e tenta efetuar o cadastramento dos produtos na tabela de produtos

     def fn_cadastrar_produto():

          fn_tratar_respostas()

          if stop_preco != True and stop_cod_produto != True and stop_produto != True:

               var_cod_produto = en_cod_produto.get().strip().upper()
               var_produto = en_produto.get().strip().title()
               var_preco = en_preco.get().strip()
               result = sub(",", ".", var_preco)
               var_preco = round(float(result), 2)

               try:
                    dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
                    )
                    conn = pyodbc.connect(dados_conexao)
                    cursor = conn.cursor()
                    cursor.execute("SELECT * FROM dProdutos")
                    valores = cursor.fetchall()
                    lista_cod_produto = [valor[0] for valor in valores]
                    if en_cod_produto.get().strip() not in lista_cod_produto:
                         pass
                    else:
                         cursor.close()
                         conn.close()
                         showwarning(title="Pavio Rosa - Atenção", message="O produto informado já está cadastrado na base de produtos!")
                    cursor.execute(f'''
                                        INSERT INTO dProdutos
                                        (Cod_produto, 
                                        Data_cadastro, 
                                        Produto, 
                                        Preco)
                                        VALUES
                                        ('{var_cod_produto}', 
                                        FORMAT(GETDATE(), 'dd/MM/yyyy HH:mm:ss'), 
                                        '{var_produto}', 
                                        '{var_preco}')
                                   ''')
                    cursor.commit()
                    cursor.close()
                    conn.close()
                    showinfo(title="Pavio Rosa - Atenção", message="Produto cadastrado com sucesso!")
                    fn_limpar_campos()
               except:
                    showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante o cadastrado do produto!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")
          else:
               showwarning(title="Pavio Rosa - Atenção", message="Não foi possível cadastrar o produto!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     # função para exportar os dados referentes a tabela de produtos

     def fn_exportar_tabela():

          try:
               dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
               )
               conn = pyodbc.connect(dados_conexao)
               df = pd.read_sql_query(
                                        sql="SELECT * FROM dProdutos", 
                                        con=conn
                                   )
               try:
                    df['Data_cadastro'] = df['Data_cadastro'].dt.strftime("%d/%m/%Y %H:%M:%S")
                    df['Data_atualizacao'] = df['Data_atualizacao'].dt.strftime("%d/%m/%Y %H:%M:%S")
               except:
                    pass
               diretorio = askdirectory(title="Selecione o local para salvar o arquivo",
                                        initialdir=fr"C:\Users\{environ['USERNAME']}\Downloads",
                                        mustexist=True)
               if diretorio == "" or diretorio == None:
                    sleep(1)
                    var_confirmar = askyesno(title="Pavio Rosa - Atenção", message="Você não selecionou nenhum local para salvar o arquivo. Deseja mesmo cancelar a operação?", default="no")
                    if var_confirmar == True:
                         showinfo(title="Pavio Rosa - Atenção", message="Operação cancelada com sucesso!")
                    else:     
                         diretorio = getcwd()
                         df.to_excel(diretorio + "\Pavio_Rosa_produtos.xlsx", index=False, sheet_name="Produtos")
                         showinfo(title="Pavio Rosa - Atenção", message=f"O arquivo foi salvo em:\n\n{diretorio}")
               else:
                    df.to_excel(diretorio + "\Pavio_Rosa_produtos.xlsx", index=False, sheet_name="Clientes")
                    showinfo(title="Pavio Rosa - Atenção", message="Relatório de produtos exportado com sucesso!")
          except:
               showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante o processo de exportação dos dados da base de produtos!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     def fn_atualizar_dados():

          var_cod_produto = en_cod_produto.get().strip()
          if var_cod_produto == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo de código do produto!")
               stop_cod_produto = True
          else:
               en_cod_produto.set(var_cod_produto)
               stop_cod_produto = False

          var_preco = en_preco.get().strip()
          padrao_preco = "[\d]+[\.|,]?[\d]+"
          checar_preco = fullmatch(compile(padrao_preco), var_preco)
          if var_preco == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo de preço!")
               stop_preco = True
          elif "R$" in var_preco:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de preço só aceita números.\n\nSe desejar informar os centavos, utilizar a vírgula ou o ponto.")
               stop_preco = True
          elif "-" in var_preco:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de preço não aceita valores negativos!")
               stop_preco = True
          elif checar_preco == None:
               showwarning(title="Pavio Rosa - Atenção", message="O formato de preço passado é inválido!")
               stop_preco = True
          else:
               en_preco.set(var_preco)
               stop_preco = False
          
          var_novo_preco = en_preco.get().strip()
          result = sub(",", ".", var_novo_preco)
          var_novo_preco = round(float(result), 2)

          var_cod_produto = en_cod_produto.get().strip().upper()

          if stop_cod_produto != True and stop_preco != True:

               try:
                    dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
                    )
                    conn = pyodbc.connect(dados_conexao)
                    cursor = conn.cursor()
                    cursor.execute("SELECT * FROM dProdutos")
                    valores = cursor.fetchall()
                    cod_produtos = [valor[0] for valor in valores]
                    if en_cod_produto.get() not in cod_produtos:
                         showwarning(title="Pavio Rosa - Atenção", message="O código do produto informado não foi cadastrado no banco de dados!\n\n Favor, verificar se o código do produto está realmente correto.")
                         fn_limpar_campos()
                    else:
                         cursor.execute(f'''
                              UPDATE dProdutos
                              SET Preco = '{var_novo_preco}', Data_atualizacao = FORMAT(GETDATE(), 'dd/MM/yyyy HH:mm:ss')
                              WHERE Cod_produto = '{var_cod_produto}'
                         ''')
                         cursor.commit()
                         cursor.close()
                         conn.close()
                         showinfo(title="Pavio Rosa - Atenção", message=f"Atualização do produto {en_cod_produto.get()} feita com sucesso!")
                         fn_limpar_campos()
               except:
                    showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante a atualização dos dados!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")
          else:
               showwarning(title="Pavio Rosa - Atenção", message="Não foi possível cadastrar o produto\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     # função para limpar os campos que foram preenchidos

     def fn_limpar_campos():

          en_cod_produto.set("")
          en_produto.set("")
          en_preco.set("")

     def fn_voltar_menu():

          janela_produtos.destroy()
          fn_tela_menu()

     # criando os widgets da tela de cadastro de produtos

     en_produto = ctk.StringVar()

     en_preco = ctk.StringVar()

     en_cod_produto = ctk.StringVar()

     lb_1 = ctk.CTkLabel(master=janela_produtos, 
                         text="CADASTRO DE PRODUTOS",
                         text_color="#e35d6a",
                         font=('Helvetica', 26, "bold"))

     lb_2 = ctk.CTkLabel(master=janela_produtos, 
                         text="Código do Produto",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_3 = ctk.CTkLabel(master=janela_produtos, 
                         text="Produto",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_4 = ctk.CTkLabel(master=janela_produtos, 
                         text="Preço",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     en_1 = ctk.CTkEntry(master=janela_produtos,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_cod_produto)

     en_2 = ctk.CTkEntry(master=janela_produtos,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_produto)

     en_3 = ctk.CTkEntry(master=janela_produtos,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_preco)

     bt_1 = ctk.CTkButton(master=janela_produtos,
                         font=("Helvetica", 12, "bold"), 
                         text="CADASTRAR", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_cadastrar_produto)

     bt_2 = ctk.CTkButton(master=janela_produtos,
                         font=("Helvetica", 12, "bold"), 
                         text="ATUALIZAR", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_atualizar_dados)

     bt_3 = ctk.CTkButton(master=janela_produtos,
                    font=("Helvetica", 12, "bold"), 
                    text="GERAR RELATÓRIO", 
                    width=120, 
                    height=30, 
                    corner_radius=10, 
                    fg_color="#e35d6a", 
                    hover=True, 
                    hover_color="#c98276",
                    command=fn_exportar_tabela)

     bt_4 = ctk.CTkButton(master=janela_produtos,
                         font=("Helvetica", 12, "bold"), 
                         text="LIMPAR CAMPOS", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_limpar_campos)

     bt_5 = ctk.CTkButton(master=janela_produtos,
                         font=("Helvetica", 12, "bold"), 
                         text="MENU", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_voltar_menu)

     # posicionando os widgtes da tela de cadastro de produtos

     lb_1.place(x=10, y=20)
     lb_2.place(x=10, y=78)
     lb_3.place(x=10, y=108)
     lb_4.place(x=10, y=138)
     en_1.place(x=120, y=80)
     en_2.place(x=120, y=110)
     en_3.place(x=120, y=140)
     bt_1.place(x=10, y=180)
     bt_2.place(x=140, y=180)
     bt_3.place(x=270, y=180)
     bt_4.place(x=410, y=180)
     bt_5.place(x=540, y=180)

     janela_produtos.mainloop()

# criando a tela de cadastro de transações

def fn_tela_transacoes():

     janela_transacoes = ctk.CTk(fg_color="pink")
     janela_transacoes.title("Pavio Rosa - Sistema de Cadastro")
     janela_transacoes.geometry("720x220+300+150")
     janela_transacoes.iconbitmap("icone.ico")
     janela_transacoes.resizable(width=False, height=False)

     def fn_tratar_respostas():

          global stop_data_venda
          global stop_cod_produto
          global stop_id_cliente
          global stop_quantidade
          global stop_preco

          var_data_venda = en_data_venda.get().strip()
          padrao_data_venda = "\d{2}/\d{2}/\d{4}"
          checar_data_venda = fullmatch(compile(padrao_data_venda), var_data_venda)
          try:
               data_convertida = datetime.strptime(var_data_venda, "%d/%m/%Y")
          except:
               pass
          if var_data_venda == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório informar a data da venda!")
               stop_data_venda = True
          elif data_convertida > datetime.today():
               showwarning(title="Pavio Rosa - Atenção", message="A data da venda não pode ser maior do que a data de hoje!")
               stop_data_venda = True
          elif checar_data_venda == None:
               showwarning(title="Pavio Rosa - Atenção", message="A data da venda fornecida não é válida ou não está no padrão correto!")
               stop_data_venda = True
          else:
               en_data_venda.set(var_data_venda)
               stop_data_venda = False

          var_quantidade = en_quantidade.get().strip()
          if var_quantidade == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório informar a quantidade vendida!")
               stop_quantidade = True
          elif "," in var_quantidade or "." in var_quantidade:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de quantidade só aceita valores inteiros!")
               stop_quantidade = True
          else:
               en_quantidade.set(var_quantidade)
               stop_quantidade = False

          var_preco = en_preco.get().strip()
          padrao_preco = "[\d]+[\.|,]?[\d]+"
          checar_preco = fullmatch(compile(padrao_preco), var_preco)
          if var_preco == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo de preço!")
               stop_preco = True
          elif "R$" in var_preco:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de preço só aceita números.\n\nSe desejar informar os centavos, utilizar a vírgula ou o ponto.")
               stop_preco = True
          elif "-" in var_preco:
               showwarning(title="Pavio Rosa - Atenção", message="O campo de preço não aceita valores negativos!")
               stop_preco = True
          elif checar_preco == None:
               showwarning(title="Pavio Rosa - Atenção", message="O formato de preço passado é inválido!")
               stop_preco = True
          else:
               en_preco.set(var_preco)
               stop_preco = False

          var_cod_produto = en_cod_produto.get().strip()
          if var_cod_produto == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório preencher o campo de código do produto!")
               stop_cod_produto = True
          else:
               en_cod_produto.set(var_cod_produto)
               stop_cod_produto = False

          var_id_cliente = en_id_cliente.get().strip()
          if var_id_cliente == "":
               showwarning(title="Pavio Rosa - Atenção", message="É obrigatório informar o id do cliente!")
               stop_id_cliente = True
          else:
               en_id_cliente.set(var_id_cliente)
               stop_id_cliente = False

     # realiza a conexão com o banco e tenta efetuar o cadastramento de venda na tabela de transações

     def fn_cadastrar_venda():

          fn_tratar_respostas()

          if stop_data_venda != True and stop_quantidade != True and stop_preco != True and stop_cod_produto != True and stop_id_cliente != True:

               var_data_venda = en_data_venda.get().strip()
               var_quantidade = int(en_quantidade.get().strip())
               var_cod_produto = en_cod_produto.get().strip()
               var_id_cliente = int(en_id_cliente.get().strip())
               var_preco = en_preco.get().strip()
               result = sub(",", ".", var_preco)
               var_preco = round(float(result), 2)
               var_valor_total = round(var_quantidade * var_preco, 2)

               try:
                    dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
                    )
                    conn = pyodbc.connect(dados_conexao)
                    cursor = conn.cursor()
                    tab_produtos = cursor.execute("SELECT * FROM dProdutos")
                    lista_produtos = tab_produtos.fetchall()
                    cod_produtos = [produto[0] for produto in lista_produtos]
                    tab_clientes = cursor.execute("SELECT * FROM dClientes")
                    lista_clientes = tab_clientes.fetchall()
                    id_clientes = [cliente[0] for cliente in lista_clientes]
                    if en_cod_produto.get() not in cod_produtos:
                         showwarning(title="Pavio Rosa - Atenção", message="O código do produto informado não foi cadastrado no banco de dados!\n\n Favor, verificar se o código do produto está realmente correto.")
                         en_cod_produto.set("")
                    elif int(en_id_cliente.get()) not in id_clientes:
                         showwarning(title="Pavio Rosa - Atenção", message="O código do cliente informado não foi cadastrado no banco de dados!\n\n Favor, verificar se o código do cliente está realmente correto.")
                         en_id_cliente.set("")
                    else:
                         cursor.execute(f'''
                                        INSERT INTO fTransacoes
                                        (Data_venda, 
                                        Cod_produto, 
                                        ID_cliente,
                                        Quantidade, 
                                        Preco, 
                                        Valor_total)
                                        VALUES
                                        ('{var_data_venda}', 
                                        '{var_cod_produto}', 
                                        '{var_id_cliente}', 
                                        '{var_quantidade}', 
                                        '{var_preco}', 
                                        '{var_valor_total}')
                                   ''')
                         cursor.commit()
                         cursor.close()
                         conn.close()
                         showinfo(title="Pavio Rosa - Atenção", message="Venda cadastrada com sucesso!")
                         fn_limpar_campos()
               except:
                    showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante o cadastrado da venda!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")
          else:
               showwarning(title="Pavio Rosa - Atenção", message="Não foi possível cadastrar a venda!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     # função para exportar os dados referentes as transações da loja

     def fn_exportar_tabela():

          try:
               dados_conexao = (
                    "Driver={SQL Server};"
                    "Server=DESKTOP-OJGFM82;"
                    "Database=db_pavio_rosa"
               )
               conn = pyodbc.connect(dados_conexao)
               df = pd.read_sql_query(
                                        sql="SELECT * FROM fTransacoes", 
                                        con=conn
                                   )
               diretorio = askdirectory(title="Selecione o local para salvar o arquivo",
                                        initialdir=fr"C:\Users\{environ['USERNAME']}\Downloads",
                                        mustexist=True)
               try:
                    df['Data_venda'] = df['Data_venda'].dt.strftime("%d/%m/%Y")
               except:
                    pass
               if diretorio == "" or diretorio == None:
                    sleep(1)
                    var_confirmar = askyesno(title="Pavio Rosa - Atenção", message="Você não selecionou nenhum local para salvar o arquivo. Deseja mesmo cancelar a operação?", default="no")
                    if var_confirmar == True:
                         showinfo(title="Pavio Rosa - Atenção", message="Operação cancelada com sucesso!")
                    else:     
                         diretorio = getcwd()
                         df.to_excel(diretorio + "\Pavio_Rosa_transacoes.xlsx", index=False, sheet_name="Transacoes")
                         showinfo(title="Pavio Rosa - Atenção", message=f"O arquivo foi salvo em:\n\n{diretorio}")
               else:
                    df.to_excel(diretorio + "\Pavio_Rosa_transacoes.xlsx", index=False, sheet_name="Clientes")
                    showinfo(title="Pavio Rosa - Atenção", message="Relatório de transações exportado com sucesso!")
          except:
               showerror(title="Pavio Rosa - Atenção", message="Algo inesperado ocorreu durante o processo de exportação dos dados da base de transações!\n\nTente novamente e caso o erro persista, entre em contato com o desenvolvedor do sistema para maiores esclarecimentos.")

     # função para limpar os campos que foram preenchidos

     def fn_limpar_campos():

          en_data_venda.set("")
          en_cod_produto.set("")
          en_id_cliente.set("")
          en_quantidade.set("")
          en_preco.set("")

     def fn_voltar_menu():

          janela_transacoes.destroy()
          fn_tela_menu()

     # criando os widgets da tela de cadastro de transações

     en_data_venda = ctk.StringVar()

     en_cod_produto = ctk.StringVar()

     en_id_cliente = ctk.StringVar()

     en_quantidade = ctk.StringVar()

     en_preco = ctk.StringVar()

     lb_1 = ctk.CTkLabel(master=janela_transacoes, 
                         text="CADASTRO DE VENDAS",
                         text_color="#e35d6a",
                         font=('Helvetica', 26, "bold"))

     lb_2 = ctk.CTkLabel(master=janela_transacoes, 
                         text="Data da Venda",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_3 = ctk.CTkLabel(master=janela_transacoes, 
                         text="Cod Produto",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_4 = ctk.CTkLabel(master=janela_transacoes, 
                         text="ID Cliente",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_5 = ctk.CTkLabel(master=janela_transacoes, 
                         text="Quantidade",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     lb_6 = ctk.CTkLabel(master=janela_transacoes, 
                         text="Preço",
                         text_color="#e35d6a",
                         font=('Helvetica', 12, "bold"))

     en_1 = ctk.CTkEntry(master=janela_transacoes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_data_venda)

     en_2 = ctk.CTkEntry(master=janela_transacoes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_cod_produto)

     en_3 = ctk.CTkEntry(master=janela_transacoes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_id_cliente)

     en_4 = ctk.CTkEntry(master=janela_transacoes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_quantidade)

     en_5 = ctk.CTkEntry(master=janela_transacoes,
                         width=150,
                         height=20,
                         border_width=1,
                         textvariable=en_preco)

     bt_1 = ctk.CTkButton(master=janela_transacoes,
                         font=("Helvetica", 12, "bold"), 
                         text="CADASTRAR", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_cadastrar_venda)

     bt_2 = ctk.CTkButton(master=janela_transacoes,
                         font=("Helvetica", 12, "bold"), 
                         text="GERAR RELATÓRIO", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_exportar_tabela)

     bt_3 = ctk.CTkButton(master=janela_transacoes,
                         font=("Helvetica", 12, "bold"), 
                         text="LIMPAR CAMPOS", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_limpar_campos)

     bt_4 = ctk.CTkButton(master=janela_transacoes,
                         font=("Helvetica", 12, "bold"), 
                         text="MENU", 
                         width=120, 
                         height=30, 
                         corner_radius=10, 
                         fg_color="#e35d6a", 
                         hover=True, 
                         hover_color="#c98276",
                         command=fn_voltar_menu)

     # posicionando os widgets da tela de cadastro de transações

     lb_1.place(x=200, y=20)
     lb_2.place(x=10, y=78)
     lb_3.place(x=260, y=78)
     lb_4.place(x=500, y=78)
     lb_5.place(x=10, y=118)
     lb_6.place(x=260, y=118)
     en_1.place(x=100, y=80)
     en_2.place(x=340, y=80)
     en_3.place(x=560, y=80)
     en_4.place(x=100, y=120)
     en_5.place(x=340, y=120)
     bt_1.place(x=10, y=180)
     bt_2.place(x=140, y=180)
     bt_3.place(x=280, y=180)
     bt_4.place(x=410, y=180)

     janela_transacoes.mainloop()

# criando as funções da tela principal

# função para gravar no arquivo credenciais.txt se o usuário deseja ou não ser lembrado

def fn_gravar_usuario():

     with open(file="credenciais.txt", mode="w", encoding="utf-8") as arq_credenciais:
          arq_credenciais.write("juliana.pavio_rosa,pavio_rosa.2023," + str(cb_checar_login.get()))
          arq_credenciais.close()

# função para lembrar o usuário se o mesmo desejar

def fn_lembrar_usuario():

     with open(file="credenciais.txt", mode="r", encoding="utf-8") as arq_credenciais:
          credenciais = arq_credenciais.readline().split(",")
          usuario = credenciais[0].strip()
          senha = credenciais[1].strip()
          lembrar_login = credenciais[2].strip()
          if lembrar_login == str(1):
               cb_checar_login.set(1)
               en_usuario.set(usuario)
          else:
               cb_checar_login.set(0)
          arq_credenciais.close()

# função para logar caso as credenciais estejam corretas

def fn_logar():

     with open(file="credenciais.txt", mode="r", encoding="utf-8") as arq_credenciais:
          credenciais = arq_credenciais.read().split(",")
          usuario = credenciais[0].strip()
          senha = credenciais[1].strip()
          if usuario == en_usuario.get() and senha == en_senha.get():
               aguardar_ok = showinfo(title="Pavio Rosa", message="Seja bem-vinda, Juliana!")
               if aguardar_ok == "ok":
                    janela_principal.destroy()
                    fn_tela_menu()
          elif usuario != en_usuario.get() and senha != en_senha.get():
               showerror(title="Pavio Rosa - Atenção", message="Usuário e senha inválidos!")
          elif usuario!= en_usuario.get():
               showerror(title="Pavio Rosa - Atenção", message="Usuário inválido!")
          else:
               showerror(title="Pavio Rosa - Atenção", message="Senha inválida!")

# função para realizar alteração de senha

def fn_alterar_senha():

     janela_principal.destroy()
     fn_tela_alterar_senha()

# função para recuperar a senha caso o usuário esqueça

def fn_recuperar_senha():

     janela_principal.destroy()
     fn_tela_esqueci_minha_senha()

# criando os widgets da janela

img_logo = PhotoImage(file="logo.png")

en_usuario = ctk.StringVar()

en_senha = ctk.StringVar()

cb_checar_login = ctk.IntVar()

lb_1 = ctk.CTkLabel(master=janela_principal, 
                    text="Faça o seu login", 
                    text_color="#e35d6a", 
                    font=("Helvetica", 18,'bold'))

lb_2 = ctk.CTkLabel(master=janela_principal,
                    text="Usuário",
                    text_color="#e35d6a", 
                    font=("Helvetica", 12, "bold"))

lb_3 = ctk.CTkLabel(master=janela_principal,
                    text="Senha",
                    text_color="#e35d6a", 
                    font=("Helvetica", 12, "bold"))

en_1 = ctk.CTkEntry(master=janela_principal, 
                    width=200,
                    border_width=1,
                    textvariable=en_usuario)

en_2 = ctk.CTkEntry(master=janela_principal, 
                    width=200,
                    border_width=1, 
                    show="*",
                    textvariable=en_senha)

en_3 = ctk.CTkLabel(master=janela_principal, 
                    text=None, 
                    image=img_logo)

cb = ctk.CTkCheckBox(master=janela_principal, 
                    text="Lembrar meu login", 
                    text_color="#e35d6a", 
                    checkbox_width=20, 
                    checkbox_height=20,
                    border_width=3, 
                    font=("Helvetica", 12, "bold"),
                    hover_color="#c98276", 
                    fg_color="#c98276",
                    variable=cb_checar_login,
                    command=fn_gravar_usuario)

bt_1 = ctk.CTkButton(master=janela_principal, 
                    font=("Helvetica", 12, "bold"), 
                    text="LOGAR", 
                    width=150, 
                    height=30, 
                    corner_radius=10, 
                    fg_color="#e35d6a", 
                    hover=True, 
                    hover_color="#c98276", 
                    command=fn_logar)

bt_2 = ctk.CTkButton(master=janela_principal, 
                    font=("Helvetica", 12, "bold"), 
                    text="ALTERAR SENHA", 
                    width=150, 
                    height=30, 
                    corner_radius=10, 
                    fg_color="#e35d6a", 
                    hover=True, 
                    hover_color="#c98276",
                    command=fn_alterar_senha)

bt_3 = ctk.CTkButton(master=janela_principal, 
                    font=("Helvetica", 12, "bold"), 
                    text="ESQUECI MINHA SENHA", 
                    width=150, 
                    height=30, 
                    corner_radius=10, 
                    fg_color="#e35d6a", 
                    hover=True, 
                    hover_color="#c98276",
                    command=fn_recuperar_senha)

# executando a função lembrar_usuario logo após iniciar a aplicação

fn_lembrar_usuario()

# posicionando os widgtes na tela

lb_1.place(x=100, y=10)
lb_2.place(x=70, y=40)
lb_3.place(x=70, y=95)
en_1.place(x=70, y=65)
en_2.place(x=70, y=120)
en_3.place(x=80, y=350)
cb.place(x=70, y=160)
bt_1.place(x=90, y=200)
bt_2.place(x=90, y=250)
bt_3.place(x=90, y=300)

janela_principal.mainloop()