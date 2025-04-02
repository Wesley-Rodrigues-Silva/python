import pandas as pd
import win32com.client
import os

# Definição dos caminhos para os arquivos
planilha_mensal = "C:\\Users\\trank\\Desktop\\Pasta1.xlsx"
historico_notificacoes = "C:\\Users\\trank\\Desktop\\historico.xlsx"

# Carregar o histórico de notificações, se existir
if os.path.exists(historico_notificacoes):
    historico_df = pd.read_excel(historico_notificacoes)
else:
    historico_df = pd.DataFrame(columns=["Nome", "Email", "Notificacoes", "Ultimas Datas"])

# Carregar os dados da planilha mensal
mensal_df = pd.read_excel(planilha_mensal)

# Converter a coluna de datas para formato brasileiro
mensal_df["Data"] = pd.to_datetime(mensal_df["Data"], errors='coerce').dt.strftime('%d/%m/%Y')

# Criar um dicionário para armazenar as notificações
notificacoes = {}
for _, row in mensal_df.iterrows():
    nome, email, data = row["Nome"], row["Email"], row["Data"]
    if email in notificacoes:
        notificacoes[email]["datas"].append(data)
    else:
        notificacoes[email] = {"nome": nome, "datas": [data]}

# Inicializar o Outlook para envio de e-mails
outlook = win32com.client.Dispatch("Outlook.Application")

# Iterar sobre os funcionários a serem notificados
for email, info in notificacoes.items():
    nome = info["nome"]
    datas = info["datas"]
    notificacoes_anteriores = historico_df.loc[historico_df["Email"] == email, "Notificacoes"].values
    
    # Verificar quantas notificações o funcionário já recebeu
    if len(notificacoes_anteriores) > 0:
        num_notificacoes = notificacoes_anteriores[0] + 1
    else:
        num_notificacoes = 1
    
    # Ajustar a mensagem dependendo da quantidade de datas
    if len(datas) == 1:
        texto_datas = f"você esqueceu de registrar na seguinte data: {datas[0]}"
    else:
        texto_datas = f"você esqueceu de registrar nas seguintes datas: {', '.join(datas)}"
    
    # Definir a mensagem personalizada conforme o número de notificações
    if num_notificacoes >= 3:
        mensagem = (f"Prezado(a) {nome},\n\n"
                    f"Identificamos que {texto_datas}.\n"
                    f"Esta é sua {num_notificacoes}ª notificação. Caso tenha outros esquecimentos nos meses seguintes, será descontado como falta.\n\n"
                    "Atenciosamente,\nRecursos Humanos")
    else:
        mensagem = (f"Prezado(a) {nome},\n\n"
                    f"Identificamos que {texto_datas}.\n"
                    f"Esta é sua {num_notificacoes}ª notificação de esquecimento.\n\n"
                    "Atenciosamente,\nRecursos Humanos")
    
    # Criar e enviar o e-mail via Outlook
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = "Notificação de Esquecimento de Ponto"
    mail.Body = mensagem
    mail.Send()
    
    # Atualizar o histórico de notificações
    historico_df = historico_df[historico_df["Email"] != email]
    historico_df = pd.concat([historico_df, pd.DataFrame([{
        "Nome": nome,
        "Email": email,
        "Notificacoes": num_notificacoes,
        "Ultimas Datas": ', '.join(datas)
    }])], ignore_index=True)

# Salvar o histórico atualizado em um arquivo Excel
historico_df.to_excel(historico_notificacoes, index=False)

print("Notificações enviadas com sucesso!")