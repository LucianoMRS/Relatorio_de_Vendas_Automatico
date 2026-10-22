# Registro de Vendas com Planilha Excel e Envio por E-mail

---

## 1. Bibliotecas importadas

| Importação | Para que serve |
|---|---|
| `smtplib` | Envio de e-mails pelo protocolo SMTP |
| `email.message.EmailMessage` | Montagem da mensagem (remetente, destinatário, assunto, corpo e anexo) |
| `getpass.getpass` | Leitura de senha sem mostrar o que é digitado |
| `openpyxl.Workbook` | Criação de planilhas `.xlsx` |

---

## 2. Funções nativas do Python

### `print()`
Exibe texto no terminal. No código, mostra mensagens ao usuário e o resumo das vendas.

```python
print("\nDados da venda registrados com sucesso!\n")
```

O `\n` dentro do texto pula uma linha.

### `input()`
Lê o que o usuário digita e **sempre devolve uma string**.

```python
nome_produto = input("\nDigite o nome do produto: ")
```

### `int()` e `float()`
Convertem texto em número inteiro e decimal, respectivamente. Se o texto não for numérico, lançam `ValueError`, que o código captura.

```python
quantidade = int(input("\nDigite a quantidade vendida: "))
valor_unitario = float(input("\nDigite o valor unitário: ").replace(",", "."))
```

### `open()`
Abre um arquivo. O modo `"rb"` significa *read binary* (leitura binária), necessário para ler um `.xlsx`, que não é um arquivo de texto.

```python
with open(nome_arquivo, "rb") as anexo:
    conteudo_anexo = anexo.read()
```

---

## 3. Métodos de string e lista

### `str.replace(antigo, novo)`
Troca trechos do texto. Aqui converte a vírgula decimal brasileira em ponto, para que `float()` aceite `80,50`.

```python
"80,50".replace(",", ".")  # "80.50"
```

### `str.endswith(sufixo)`
Retorna `True` se o texto termina com o sufixo informado. Evita gerar `relatorio.xlsx.xlsx`.

```python
if not nome_arquivo.endswith(".xlsx"):
    nome_arquivo += ".xlsx"
```

### `list.append(item)`
Adiciona um item ao final da lista. É usado para guardar cada venda (um dicionário) em `vendas` e, no Excel, para adicionar linhas à planilha.

```python
vendas.append({"produto": nome_produto, "qtd": quantidade, "valor": valor_unitario})
```

---

## 4. Estruturas de controle e tratamento de erros

### `while True` + `break`
Cria um loop infinito que só termina quando o usuário deixa o nome do produto em branco.

### `continue`
Pula para a próxima volta do loop. É usado quando a entrada é inválida, descartando aquela venda.

### `for ... in`
Percorre a lista de vendas para calcular totais e para escrever as linhas da planilha.

### `try` / `except`
Captura erros sem derrubar o programa.

| Bloco | Erro tratado | Motivo |
|---|---|---|
| `except ValueError` | Conversão de número | Usuário digitou letras |
| `except smtplib.SMTPAuthenticationError` | Login recusado | E-mail ou senha de app incorretos |
| `except Exception as e` | Qualquer outro erro | Mostra a descrição do problema em `e` |

### `raise SystemExit`
Encerra o programa imediatamente. É usado quando não há vendas ou quando a planilha não pôde ser gerada, evitando seguir adiante sem dados.

### `with`
Gerenciador de contexto: garante que o recurso (arquivo ou conexão SMTP) seja **fechado automaticamente**, mesmo se ocorrer um erro.

---

## 5. Formatação com f-strings

Strings iniciadas por `f` permitem inserir variáveis com `{}`. O sufixo `:.2f` formata o número com duas casas decimais.

```python
print(f"Produto: {nome}, Total: R${total:.2f}")
```

---

## 6. Biblioteca `openpyxl` (planilha)

| Recurso | O que faz |
|---|---|
| `Workbook()` | Cria um arquivo Excel novo em memória |
| `wb.active` | Retorna a planilha ativa (a primeira) |
| `planilha.title = "Vendas"` | Define o nome da aba |
| `planilha.append([...])` | Adiciona uma linha com os valores da lista |
| `wb.save(nome)` | Grava o arquivo no disco |

```python
wb = Workbook()
planilha_vendas = wb.active
planilha_vendas.title = "Vendas"
planilha_vendas.append(["Produto", "Quantidade", "Valor Unitário", "Total"])
wb.save("vendas.xlsx")
```

Se o arquivo estiver aberto no Excel, o `save` pode falhar com `PermissionError`.

---

## 7. Biblioteca `email` (montagem da mensagem)

### `EmailMessage()`
Cria o objeto que representa o e-mail.

### `msg["Campo"] = valor`
Define os cabeçalhos `From`, `To` e `Subject`, como se a mensagem fosse um dicionário.

### `msg.set_content(corpo)`
Define o texto principal do e-mail.

### `msg.add_attachment(dados, maintype, subtype, filename)`
Anexa um arquivo. Para `.xlsx`, o tipo MIME correto é:

```python
msg.add_attachment(
    conteudo_anexo,
    maintype="application",
    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    filename=nome_arquivo,
)
```

---

## 8. Biblioteca `smtplib` (envio)

| Método | O que faz |
|---|---|
| `smtplib.SMTP("smtp.gmail.com", 587)` | Abre a conexão com o servidor do Gmail na porta 587 |
| `smtp.starttls()` | Passa a conexão para modo criptografado (TLS) |
| `smtp.login(usuario, senha)` | Autentica no servidor |
| `smtp.send_message(msg)` | Envia o `EmailMessage` montado |

```python
with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
    smtp.starttls()
    smtp.login(remetente, senha)
    smtp.send_message(msg)
```

> **Atenção:** o Gmail exige uma **senha de app** (com verificação em duas etapas ativada), e não a senha normal da conta.

---

## 9. `getpass.getpass()`

Funciona como `input()`, mas **não exibe os caracteres digitados**, o que protege a senha de quem estiver olhando a tela.

```python
senha = getpass("\nDigite a senha de app do remetente: ")
```

---

## 10. Definição de funções (`def`)

O código organiza o trabalho em duas funções:

- **`gerar_planilha(vendas, nome_arquivo)`**: recebe a lista de vendas e o nome do arquivo, monta o Excel e salva.
- **`enviar_email(remetente, senha, destinatario, assunto, corpo, nome_arquivo)`**: monta a mensagem com anexo e envia pelo Gmail.

Passar `nome_arquivo` como parâmetro, em vez de reatribuí-lo dentro da função, evita o erro `UnboundLocalError`.

---

## Fluxo geral do programa

1. Coleta as vendas em um loop (`input`, `int`, `float`, `append`).
2. Calcula e exibe o total de cada item e o valor geral (`for`, f-strings).
3. Pergunta o nome do arquivo e garante a extensão `.xlsx` (`endswith`).
4. Gera a planilha (`Workbook`, `append`, `save`).
5. Coleta os dados do e-mail e envia com a planilha anexada (`EmailMessage`, `SMTP`).
