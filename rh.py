from logging import RootLogger
import pandas as pd
import smtplib
from datetime import datetime, timedelta
from PIL import Image, ImageDraw, ImageFont
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Variáveis globais para a GUI
logs_text = ""
executando_pela_gui = False  # Variável para verificar se o programa está sendo executado pela GUI
root = None  # Adicione uma declaração global para a variável root
bg_image_path = r"C:\Users\rafael.bezerra\Desktop\programaRH\imagens\bg5.png"  # Caminho da imagem de fundo
status_label = None  # Adicione uma declaração global para a variável status_label

# Variáveis globais para o caminho da planilha e fotos
planilha_path = r'C:\Users\pedro.guilherme\Desktop\programaRH\dados_colaboradores.xlsx'
caminho_foto = ""

# Dicionário para armazenar os aniversariantes do dia e se o e-mail já foi enviado
aniversariantes_enviados = {}
aniversariantes_tempo_empresa_enviados = {}

# Função para criar cartão de aniversário
def criar_cartao_aniversario(nome, data_nascimento, imagem_path):
    hoje = datetime.now().date()

    # Carregar o mockup
    mockup_path = r"\\192.168.15.126\C$\Users\pedro.guilherme\Desktop\programaRH\imagens\mockup.jpeg"
    output_path = f"\\\\192.168.15.126\\C$\\Users\\pedro.guilherme\\Desktop\\programaRH\\imagens\\imagensMontadas\\{nome}.png"

    with Image.open(mockup_path) as img:
        # Carregar a foto do colaborador
        with Image.open(imagem_path) as foto:
            # Redimensionar a foto para um círculo
            tamanho = (600, 600)
            mascara = Image.new("L", tamanho, 0)
            draw = ImageDraw.Draw(mascara)
            draw.ellipse((0, 0) + tamanho, fill=255)
            foto = foto.resize(tamanho)
            foto = Image.composite(foto, Image.new("RGBA", tamanho, 0), mascara)

            # Ajustar as coordenadas para posicionar a foto onde desejar
            posicao_x = 270
            posicao_y = 620

            # Adicionar a foto ao mockup
            img.paste(foto, (posicao_x, posicao_y), mask=foto)

            # Adicionar o nome do colaborador
            draw = ImageDraw.Draw(img)
            font = ImageFont.load_default()
            font_size = 56
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except IOError:
                pass  # Se não conseguir carregar a fonte, use a fonte padrão
            draw.text((250, 500), nome, fill="white", font=font)

            # Salvar a imagem resultante
            img.save(output_path)

    return output_path

# Função para criar cartão de tempo de empresa
def criar_cartao_tempo_empresa(nome, dias_trabalhados, imagem_path):
    # Implemente conforme necessário
    # Similar à função criar_cartao_aniversario, mas para tempo de empresa
    pass

# Função para enviar e-mail com várias imagens anexadas
def enviar_email(nome_aniversariantes, caminhos_imagens, assunto):
    remetente = "samuel.labjt@gmail.com"
    senha = ""
    destinatario = "samuel.labjt@gmail.com"

    # Configurações do servidor SMTP
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    # Criar a mensagem
    mensagem = MIMEMultipart()
    mensagem.attach(MIMEText(f"Parabéns, {', '.join(nome_aniversariantes)}!\n\nDesejamos um dia incrível para vocês.", "plain"))

    # Anexar as imagens ao e-mail
    for nome, caminho_imagem in zip(nome_aniversariantes, caminhos_imagens):
        with open(caminho_imagem, "rb") as fp:
            anexo = MIMEImage(fp.read())
        anexo.add_header("Content-Disposition", f"attachment; filename={nome}.png")
        mensagem.attach(anexo)

    # Configurar conexão com o servidor SMTP
    servidor = smtplib.SMTP(smtp_server, smtp_port)
    servidor.starttls()
    servidor.login(remetente, senha)

    # Enviar e-mail
    servidor.sendmail(remetente, destinatario, mensagem.as_string())

    # Fechar a conexão com o servidor SMTP
    servidor.quit()

   
def enviar_email_tempo_empresa(nome_aniversariantes, caminhos_imagens, assunto):
    remetente = "samuel.labjt@gmail.com"
    senha = ""
    destinatario = "samuel.labjt@gmail.com"

    # Configurações do servidor SMTP
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    # Criar a mensagem
    mensagem = MIMEMultipart()
    mensagem.attach(MIMEText(f"Parabéns, hoje é seu naiversário na Padim{', '.join(nome_aniversariantes)}!\n\nDesejamos um dia incrível para vocês.", "plain"))

    # Anexar as imagens ao e-mail
    for nome, caminho_imagem in zip(nome_aniversariantes, caminhos_imagens):
        with open(caminho_imagem, "rb") as fp:
            anexo = MIMEImage(fp.read())
        anexo.add_header("Content-Disposition", f"attachment; filename={nome}.png")
        mensagem.attach(anexo)

    # Configurar conexão com o servidor SMTP
    servidor = smtplib.SMTP(smtp_server, smtp_port)
    servidor.starttls()
    servidor.login(remetente, senha)

    # Enviar e-mail
    servidor.sendmail(remetente, destinatario, mensagem.as_string())

    # Fechar a conexão com o servidor SMTP
    servidor.quit()
    
    
# Função principal para processar aniversariantes e enviar e-mails
def processar_aniversariantes():
    global caminho_foto, logs_var, executando_pela_gui, aniversariantes_enviados, aniversariantes_tempo_empresa_enviados  # Adicionando global para caminho_foto, logs_var, executando_pela_gui e aniversariantes_enviados
    # Carregar dados da planilha
    dados = pd.read_excel(planilha_path, header=None, names=['Nome', 'DataNascimento', 'DataEntrada', 'CaminhoFoto'], skiprows=1)

    # Obter a data atual
    hoje = datetime.now().date()

    # Limpar a lista de aniversariantes do dia
    aniversariantes_do_dia_lista = []
    aniversariantes_do_dia_tempo_empresa_lista = []

    # Iterar sobre os colaboradores
    for index, row in dados.iterrows():
        nome = row['Nome']
        data_nascimento = row['DataNascimento']
        data_entrada = row['DataEntrada']
        caminho_foto = row['CaminhoFoto']

        # Convertendo a data de nascimento para um objeto datetime.date
        data_nascimento = datetime.strptime(data_nascimento, "%Y-%m-%d").date()

        # Calcular o tempo de empresa
        data_entrada = datetime.strptime(data_entrada, "%Y-%m-%d").date()
        tempo_empresa = hoje - data_entrada

        # Se for o aniversário do colaborador
        if hoje.month == data_nascimento.month and hoje.day == data_nascimento.day:
            logs_text = f"Aniversário: {nome}"
            output_path = criar_cartao_aniversario(nome, data_nascimento, caminho_foto)
            aniversariantes_do_dia_lista.append((nome, output_path))
        # Se for o aniversário de ENTRADA do colaborador
        if hoje.month == data_entrada.month and hoje.day == data_entrada.day:
            logs_text = f"Aniversário: {nome}"
            output_path = criar_cartao_aniversario(nome, tempo_empresa , caminho_foto)
            aniversariantes_do_dia_tempo_empresa_lista.append((nome, output_path))    

    # Verificar se há aniversariantes para enviar o e-mail
    if aniversariantes_do_dia_lista:
        nomes, caminhos = zip(*aniversariantes_do_dia_lista)
        logs_text = f"Enviando e-mail para: {', '.join(nomes)}"
        enviar_email(nomes, caminhos, "Feliz Aniversário!")

        # Atualizar os registros de envio
        for nome in nomes:
            aniversariantes_enviados[nome] = hoje

    else:
        logs_text = "Nenhum aniversariante do dia."

    # Atualizar os logs na GUI
    logs_var.set(logs_text)
    
    if aniversariantes_do_dia_tempo_empresa_lista:
        nomes, caminhos = zip(*aniversariantes_do_dia_tempo_empresa_lista)
        logs_text = f"Enviando e-mail para: {', '.join(nomes)}"
        enviar_email_tempo_empresa(nomes, caminhos, "Feliz Aniversário de Padim!")

        # Atualizar os registros de envio
        for nome in nomes:
            aniversariantes_tempo_empresa_enviados[nome] = hoje

    else:
        logs_text = "Nenhum aniversariante do dia."

    # Atualizar os logs na GUI
    logs_var.set(logs_text)

# Função para abrir a pasta da planilha
def abrir_pasta_planilha():
    planilha_folder = os.path.dirname(planilha_path)
    os.startfile(planilha_folder)

def abrir_pasta_fotos():
    fotos_folder = os.path.dirname(caminho_foto)
    os.startfile(fotos_folder)

# Função para limpar os logs
def limpar_logs():
    global logs_var
    logs_text = ""
    logs_var.set(logs_text)

# Função para atualizar os logs e iniciar o processo de envio de e-mails
def atualizar_logs_e_enviar_emails():
    global logs_var, status_label, executando_pela_gui  # Adicionando global para logs_var, status_label e executando_pela_gui
    logs_text = "Iniciando processamento..."
    logs_var.set(logs_text)

    if executando_pela_gui:
        try:
            processar_aniversariantes()
            logs_text = "E-mail enviado!"
            if status_label:
                status_label.config(text=logs_text, fg='green')
            else:
                print(logs_text)
        except Exception as e:
            logs_text = f"Erro ao enviar e-mail: {str(e)}"
            if status_label:
                status_label.config(text=logs_text, fg='red')
            else:
                print(logs_text)

    logs_var.set(logs_text)

# Criar a janela principal da GUI
root = tk.Tk()
root.title("Aniversariantes e-mails")
root.geometry("500x400")  
root.resizable(width=False, height=False)  # Tornar a janela não redimensionável


# Criar um Canvas para exibir a imagem de fundo
canvas = tk.Canvas(root, width=500, height=500, highlightthickness=3)
canvas.pack()

# Adicionar a imagem de fundo ao Canvas
bg_image = tk.PhotoImage(file=bg_image_path)
canvas.create_image(0, 0, anchor="nw", image=bg_image)

# Labels para exibir logs
logs_label = tk.Label(root, text="Logs:", bg='white')  # Adicionei um fundo branco para a legibilidade
logs_label.pack()

logs_var = tk.StringVar()
logs_var.set(logs_text)
logs_display = tk.Label(root, textvariable=logs_var, justify=tk.CENTER, wraplength=100, bg='white')  # Adicionei um fundo branco para a legibilidade
logs_display.pack()



#criando meu próprio estilo de botão 
style = ttk.Style()
style.configure('Rounded.TButton', borderwidth=5, relief='raised', foreground='black', background='#A9A9A9', font=('Arial', 12))

# Botão para atualizar os logs e iniciar o processo de envio de e-mails
atualizar_logs_button = ttk.Button(root, text="Enviar E-mails", command=atualizar_logs_e_enviar_emails, style='Rounded.TButton')
atualizar_logs_button.place(relx=0.4, rely=0.3) 

# Botão para limpar os logs
limpar_logs_button = ttk.Button(root, text="Limpar Logs", command=limpar_logs, style='Rounded.TButton')
limpar_logs_button.place(relx=0.405, rely=0.4) 

open_planilha_button = ttk.Button(root, text="Planilha", command=abrir_pasta_planilha, style='Rounded.TButton')
open_planilha_button.place(relx=0.403, rely=0.5)

open_fotos_button = ttk.Button(root, text="Fotos", command=abrir_pasta_fotos, style='Rounded.TButton')
open_fotos_button.place(relx=0.403, rely=0.6)  

# Iniciar o loop principal da GUI
executando_pela_gui = True
atualizar_logs_e_enviar_emails()
root.mainloop()
