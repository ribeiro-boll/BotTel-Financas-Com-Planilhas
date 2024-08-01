# Portugues Brasileiro:
## Bot do Telegram para edição de arquivos `.xlsx`, focado para finanças.

Este repositório contém um bot de conversa para o Telegram que facilita a organização e registro de gastos monetários. O bot cria e edita arquivos `.xlsx` com o nome do usuário e o ano escolhido (até 6 anos atrás ou 3 anos no futuro). Ele automaticamente cria planilhas com todos os meses do ano selecionado, permitindo que você acompanhe seus gastos mês a mês.

## Funcionalidades

1. **Adicionar Gastos**:
   - O bot permite que você adicione novos gastos ao mês escolhido.
   - Você pode especificar o nome do gasto e o valor correspondente (Em Real Brasileiro, BRL).

2. **Listar Gastos**:
   - É possível listar todos os gastos, com seus respectivos valores, registrados no mês selecionado.

3. **Somar Gastos**:
   - O bot calcula automaticamente o total dos gastos no mês escolhido.

4. **Editar Gastos**:
   - Caso você cometa algum erro, pode editar tanto o nome quanto o valor de um gasto já registrado.

5. **Escolher o Ano**:
   - A seleção do ano é simples e intuitiva.
   - No entanto, se o ano do nome do arquivo exceder 6 anos da data atual, o bot nao poderá executar edições, pois ele irá ignorar tais arquivos .

## Language Availability

O bot, está disponivel apenas em pt-BR no momento.

## Como Usar
1. Certifique-se de que "pyTelegramBotAPI" e "openpyxl" estejam instalados, executando `pip install pyTelegramBotAPI` e `pip install openpyxl` no diretorio do bot
2. Clone este repositório para o seu ambiente local.
3. Edite a key API do bot no arquivo `BotTelegram.py` para usar em seu próprio bot de conversa no Telegram.
4. Execute o bot e mande uma mensagem contendo "/start"

## Observações

- Os arquivos `.xlsx` serão criados no diretório onde os arquivos `BotTelegram.py` e `integracaoExcel.py` estão localizados.
- Certifique-se de ajustar as permissões de escrita no diretório para que o bot possa criar e editar os arquivos.



# English:
## Telegram Bot for editing `.xlsx` files, focused in managing expenses.

This repository contains a Telegram chatbot designed to simplify the organization and tracking of monetary expenses. The bot creates and edits `.xlsx` files with the user's name and the chosen year (up to 6 years in the past or 3 years in the future). It automatically generates spreadsheets with all the months of the selected year, allowing users to monitor their monthly spending.

## Features

1. **Add Expenses**:
   - The bot allows users to add new expenses for the chosen month.
   - Users can specify the expense name and corresponding value (in Brazilian Real, BRL).

2. **List Expenses**:
   - Users can list all expenses, along with their respective values (in BRL), for the selected month.

3. **Sum Expenses**:
   - The bot automatically calculates the total expenses (in BRL) for the chosen month.

4. **Edit Expenses**:
   - If users make a mistake, they can edit both the name and value of a previously recorded expense.

5. **Choose the Year**:
   - Year selection is straightforward and intuitive.
   - However, if the year in the file name exceeds 6 years from the current date, the bot will not be able to create new edits, ignoring such files.

## Language Availability

The bot is currently available only in pt-BR for the moment.

## How to Use
1. Make sure that "pyTelegramBotAPI" and "openpyxl" is installed in your bot directory by entering `pip install pyTelegramBotAPI` and `pip install openpyxl`
2. Clone this repository to your local environment.
3. Edit the bot's key API in the `BotTelegram.py` file to use in your own Telegram chatbot.
4. Run the bot and send a message containing only "/start" 

## Notes

- The `.xlsx` files will be created in the directory where the `BotTelegram.py` and `integracaoExcel.py` files are located.
- Ensure that the directory has write permissions so that the bot can create and edit files.


