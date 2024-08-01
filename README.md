# Bot de Contabilidade no Telegram com integração com arquivos `.xlsx`

Este repositório contém um bot de conversa para o Telegram que facilita a organização e registro de gastos monetários. O bot cria e edita arquivos `.xlsx` com o nome do usuário e o ano escolhido (até 6 anos atrás ou 3 anos no futuro). Ele automaticamente cria planilhas com todos os meses do ano selecionado, permitindo que você acompanhe seus gastos mês a mês.

## Funcionalidades

1. **Adicionar Gastos**:
   - O bot permite que você adicione novos gastos ao mês escolhido.
   - Você pode especificar o nome do gasto e o valor correspondente.

2. **Listar Gastos**:
   - É possível listar todos os gastos, com seus respectivos valores, registrados no mês selecionado.

3. **Somar Gastos**:
   - O bot calcula automaticamente o total dos gastos no mês escolhido.

4. **Editar Gastos**:
   - Caso você cometa algum erro, pode editar tanto o nome quanto o valor de um gasto já registrado.

5. **Escolher o Ano**:
   - A seleção do ano é simples e intuitiva.
   - No entanto, se o ano do nome do arquivo exceder 6 anos da data atual, o bot não permitirá mais edições, pois ele irá ignorar tais arquivos .

## Como Usar

1. Clone este repositório para o seu ambiente local.
2. Edite a API do bot no arquivo `BotTelegram.py` para criar o seu próprio bot de conversa no Telegram.
3. Execute o bot e comece a registrar seus gastos!

## Observações

- Os arquivos `.xlsx` serão criados no diretório onde os arquivos `BotTelegram.py` e `integracaoExcel.py` estão localizados.
- Certifique-se de ajustar as permissões de escrita no diretório para que o bot possa criar e editar os arquivos.

