import smtplib
from email.message import EmailMessage
from getpass import getpass
from openpyxl import Workbook

#================================= Cálculo de vendas ==========================================
print("\n============================== Preencher Dados das Vendas ===============================")
print("Digite os dados das vendas. Para encerrar, pressione Enter sem digitar o nome do produto.")

vendas = []

while True:
    nome_produto = input("\nDigite o nome do produto: ")
    if nome_produto == "":
        break

    try:
        quantidade = int(input("\nDigite a quantidade vendida: "))
        valor_unitario = float(input("\nDigite o valor unitário: ").replace(",", "."))
    except ValueError:
        print("\nValor inválido, digite apenas números. Essa venda não foi registrada.")
        continue

    vendas.append({"produto": nome_produto, "qtd": quantidade, "valor": valor_unitario})
    print("\nDados da venda registrados com sucesso!\n")

if not vendas:
    print("\nNenhuma venda registrada. Encerrando.")
    raise SystemExit

valor_geral = 0

try:
    for venda in vendas:
        nome = venda["produto"]
        qtd = venda["qtd"]
        valor = venda["valor"]
        total = qtd * valor
        valor_geral += total

        print(f"Produto: {nome}, Quantidade: {qtd}, Valor Unitário: R${valor:.2f}, Total: R${total:.2f}\n")

    print(f"\nValor geral das vendas: R${valor_geral:.2f}")
except Exception as e:
    print(f"\nErro ao calcular as vendas: {e}")

#=============================== Função para gerar planilha ==========================================
def gerar_planilha(vendas, nome_arquivo):
    wb = Workbook()
    planilha_vendas = wb.active
    planilha_vendas.title = "Vendas"

    # Adicionando cabeçalho
    planilha_vendas.append(["Produto", "Quantidade", "Valor Unitário", "Total"])

    for venda in vendas:
        nome = venda["produto"]
        qtd = venda["qtd"]
        valor = venda["valor"]
        total = qtd * valor
        planilha_vendas.append([nome, qtd, valor, total])

    wb.save(nome_arquivo)
    print(f"Planilha '{nome_arquivo}' gerada com sucesso!")

nome_arquivo = input("\nDigite o nome do arquivo da planilha: ")
if not nome_arquivo.endswith(".xlsx"):
    nome_arquivo += ".xlsx"

try:
    gerar_planilha(vendas, nome_arquivo)
except Exception as e:
    print(f"\nErro ao gerar a planilha: {e}")
    raise SystemExit

#================================ Função para enviar e-mail ==========================================
def enviar_email(remetente, senha, destinatario, assunto, corpo, nome_arquivo):
    msg = EmailMessage()
    msg.set_content(corpo)
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = assunto

    with open(nome_arquivo, "rb") as anexo:
        conteudo_anexo = anexo.read()

    msg.add_attachment(
        conteudo_anexo,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=nome_arquivo,
    )

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(remetente, senha)
        smtp.send_message(msg)

remetente = input("\nDigite o e-mail do remetente: ")
senha = getpass("\nDigite a senha de app do remetente: ")
destinatario = input("\nDigite o e-mail do destinatário: ")
assunto = input("\nDigite o assunto do e-mail: ")
corpo = input("\nDigite o corpo do e-mail: ")

try:
    enviar_email(remetente, senha, destinatario, assunto, corpo, nome_arquivo)
    print("\nE-mail enviado com sucesso!")
except smtplib.SMTPAuthenticationError:
    print("\nFalha na autenticação: verifique o e-mail e a senha de app.")
except Exception as e:
    print(f"\nErro ao enviar o e-mail: {e}")