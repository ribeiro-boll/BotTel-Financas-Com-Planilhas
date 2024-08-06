import datetime
import telebot
from telebot import types
import dependencias
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from keys import keytel
bot = telebot.TeleBot(keytel())
planilha = dependencias.Planilha
gasto = dependencias.Gasto
id = dependencias.user
ano = []
arquivo = []
worksheet = []
gasto_taxi = []
mes_taxi = []
opcaoNumero_taxi = []
row_taxi = []
cell_taxi = []


@bot.message_handler(commands=["comecar"])
def inicio_bot(message):
    if arquivo == []:
        ano.append(datetime.datetime.now().year)
        worksheet.append(planilha(dependencias.nome_pessoa(message)))
        worksheet[0].definir_ano_inicial()
        arquivo.append(worksheet[0].arquivo)
        dependencias.criarArquivoAno(arquivo[0])

    opcao = types.ReplyKeyboardMarkup()
    a = types.KeyboardButton('/addGasto')
    b = types.KeyboardButton('/contasMes')
    c = types.KeyboardButton('/editarVer')
    d = types.KeyboardButton('/selecionarPlanilha')
    opcao.row(d)
    opcao.row(a, b, c)
    bot.send_message(id(message), "O limite de gastos por mês é de 50 gastos.")
    bot.send_message(id(message), f"Selecione sua opção: \n\n Para adicionar gasto em sua planilha \n-> /addGasto \n\n Caso queira ver o total de gastos no mês, toque em \n-> /contasMes \n\n Caso queira editar ou ver seus gastos, toque em \n-> /editarVer \n\n Para mudar a planilha, aperte em \n-> /selecionarPlanilha \n\n Sua planilha atual é do ano: {ano[0]}", reply_markup=opcao)


@bot.message_handler(commands=["selecionarPlanilha"])
def escolher_planilha1(message):
    try:
        bot.send_message(id(message), "Você pode escolher entre: \n\n 7 anos atrás ou 3 anos no futuro.")
        anos = dependencias.escolher_ano()
        anosEscolha = types.ReplyKeyboardMarkup()
        c = types.KeyboardButton(f'{anos[0]}')
        d = types.KeyboardButton(f'{anos[1]}')
        e = types.KeyboardButton(f'{anos[2]}')
        f = types.KeyboardButton(f'{anos[3]}')
        g = types.KeyboardButton(f'{anos[4]}')
        h = types.KeyboardButton(f'{anos[5]}')
        i = types.KeyboardButton(f'{anos[6]}')
        anosEscolha.row(f)
        anosEscolha.row(c, d, e)
        anosEscolha.row(g, h, i)
        anoParte1 = bot.send_message(id(message), "selecione o ano \n\n (obs: a primeira e maior opção, sera o ano atual)", reply_markup=anosEscolha)
        bot.register_next_step_handler(anoParte1, escolher_planilha2)

    except IndexError:
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/comecar')
        b = types.KeyboardButton('/selecionarPlanilha')
        voltar.row(a, b)
        bot.send_message(id(message), "Algo deu errado \n\n Para voltar para iniciar o programa, aperte \n-> /comecar\n\n ou para tentar novamente, aperte:\n->/selecionarPlanilha", reply_markup=voltar)


def escolher_planilha2(message):
    try:
        ano_escolhido = message.text
        arquivo.clear()
        ano.clear()
        worksheet[0].mudar_ano(ano_escolhido)
        dependencias.criarArquivoAno(worksheet[0].arquivo)
        arquivo.append(arquivo)
        ano.append(message.text)
        if 2010 < int(message.text) < 2099 and len(message.text) == 4:
            ano_escolhido = message.text
            arquivo.clear()
            ano.clear()
            worksheet[0].mudar_ano(ano_escolhido)
            dependencias.criarArquivoAno(worksheet[0].arquivo)
            arquivo.append(arquivo)
            ano.append(message.text)
            inicio = types.ReplyKeyboardMarkup()
            a = types.KeyboardButton('/comecar')
            inicio.row(a)
            bot.send_message(id(message), "Sua planilha foi criada/acessada com sucesso! \n\n Para começar a usar o programa, aperte em \n-> /comecar", reply_markup=inicio)
        else:
            voltar = types.ReplyKeyboardMarkup()
            a = types.KeyboardButton('/comecar')
            b = types.KeyboardButton('/selecionarPlanilha')
            voltar.row(a)
            voltar.row(b)
            bot.send_message(id(message), "Algo deu errado \n\n Para voltar para o início, aperte \n-> /comecar \n\n Para tentar novamente, aperte \n-> /selecionarPlanilha", reply_markup=voltar)
    except IndexError:
        voltar = types.ReplyKeyboardMarkup()
        b = types.KeyboardButton('/comecar')
        voltar.row(b)
        bot.send_message(id(message), "Algo deu errado. Por favor, toque em:\n-> /comecar \n\n pois atualmente nenhuma planilha está selecionada", reply_markup=voltar)


@bot.message_handler(commands=["addGasto"])
def selecionarMes(message):
    markup = types.ForceReply()
    valorNome = bot.send_message(id(message), "Digite o nome do gasto: ", reply_markup=markup)
    bot.register_next_step_handler(valorNome, editorExcel)


def editorExcel(message):
    novo_gasto = gasto(message.text)
    gasto_taxi.append(novo_gasto)
    markup = types.ForceReply()
    valorGasto = bot.send_message(id(message), "Digite o valor gasto: ", reply_markup=markup)
    bot.register_next_step_handler(valorGasto, seletorMes)


def seletorMes(message):
    i = 0
    mensagem = message.text
    listaMensagem = list(mensagem)
    if dependencias.floatcheck(message.text):
        while i < len(listaMensagem):
            if listaMensagem[i] == ",":
                listaMensagem[i] = "."
            i += 1
        mensagem = ''.join(listaMensagem)
        gasto_taxi[0].adicionar_valor(mensagem)
        mes = dependencias.funcao_opcao_mes(id(message), bot)
        bot.register_next_step_handler(mes, final)

    else:
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/addGasto')
        b = types.KeyboardButton('/comecar')
        voltar.row(a, b)
        bot.send_message(id(message), "Algo deu errado. Por favor, confirme se o valor do gasto é valido e tente novamente tocando em: \n-> /addGasto \n\n E tambem tenha certeza de que sua planilha está selecionada, tocando em: \n-> /comecar", reply_markup=voltar)


def final(message):
    try:
        valor = gasto_taxi[0].retornar_valor_gasto()
        a = gasto_taxi[0].adicionar_gasto(arquivo[0], message, gasto_taxi[0].gerar_lista(), valor)
        if a == "1":
            voltar = types.ReplyKeyboardMarkup()
            a1 = types.KeyboardButton('/addGasto')
            b1 = types.KeyboardButton('/comecar')
            voltar.row(a1, b1)
            gasto_taxi.clear()
            bot.send_message(id(message), "Gasto adicionado! \n\n Para adicionar outro gasto, toque em \n-> /addGasto \n\n Para iniciar novamente, toque em \n-> /comecar", reply_markup=voltar)
        if a == "0":
            gasto_taxi.clear()
            voltar = types.ReplyKeyboardMarkup()
            a1 = types.KeyboardButton('/editarVer')
            b2 = types.KeyboardButton('/addGasto')
            c3 = types.KeyboardButton('/comeco')
            voltar.row(a1, b2, c3)
            bot.send_message(id(message), f"Aparentemente, você estourou o limite de gastos (50) do mês {message.text} ou digitou algo errado. \n\n Para editar os gastos deste mês, toque em \n-> /editarVer e selecione o mês \n\n Para tentar adicionar novos gastos em outros meses, toque em \n-> /addGasto e selecione outro mês \n\n Para voltar ao início, toque em \n-> /comecar", reply_markup=voltar)

        else:
            gasto_taxi.clear()
            voltar = types.ReplyKeyboardMarkup()
            a = types.KeyboardButton('/addGasto')
            b = types.KeyboardButton('/comecar')
            voltar.row(a, b)
            bot.send_message(id(message), "Algo deu errado. Por favor, leia atentamente as mensagens que o bot enviar. \n Tente novamente com \n-> /addGasto \n\n Caso queira voltar para o início, toque em \n-> /comecar", reply_markup=voltar)

    except KeyError:
        gasto_taxi.clear()
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/addGasto')
        b = types.KeyboardButton('/comecar')
        voltar.row(a, b)
        bot.send_message(id(message), "Algo deu errado. Por favor, leia atentamente as mensagens que o bot enviar. \n Tente novamente com \n-> /addGasto \n\n Caso queira voltar para o início, toque em \n-> /comecar", reply_markup=voltar)

    except IndexError:
        gasto_taxi.clear()
        voltar = types.ReplyKeyboardMarkup()
        b = types.KeyboardButton('/comecar')
        voltar.row(b)
        bot.send_message(id(message), "Algo deu errado. Por favor, toque em:\n-> /comecar \n\n pois atualmente nenhuma planilha está selecionada", reply_markup=voltar)


@bot.message_handler(commands=['contasMes'])
def escolha_mes(message):
    mes = dependencias.funcao_opcao_mes(id(message), bot)
    bot.register_next_step_handler(mes, somatorio)


def somatorio(message):
    try:
        soma = dependencias.somatorio(worksheet[0].retornar_nome_arquivo_planilha_atual(), message)
        bot.send_message(id(message), f"O total gasto no mês {message.text} na planilha, {str(arquivo[0])}, foi de R${soma}")
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/contasMes')
        b = types.KeyboardButton('/comecar')
        voltar.row(a, b)
        bot.send_message(id(message), "Para ver o total de outro mês, toque novamente em \n-> /contasMes \n\n Para voltar ao início, toque em \n-> /comecar", reply_markup=voltar)

    except IndexError:
        gasto_taxi.clear()
        voltar = types.ReplyKeyboardMarkup()
        b = types.KeyboardButton('/comecar')
        voltar.row(b)
        bot.send_message(id(message), "Algo deu errado. Por favor, toque em:\n-> /comecar \n\n pois atualmente nenhuma planilha está selecionada", reply_markup=voltar)


@bot.message_handler(commands=["editarVer"])
def selecionar_mes(message):
    mes = dependencias.funcao_opcao_mes(id(message), bot)
    bot.register_next_step_handler(mes, listar_gastos)


def listar_gastos(message):
    try:
        mes_taxi.append(message.text)
        colunaTexto = dependencias.listar_gastos(arquivo[0], message)
        bot.send_message(id(message), colunaTexto)
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/comecar')
        b = types.KeyboardButton("/editar")
        voltar.row(a, b)
        escolha = bot.send_message(id(message), "Para editar algum valor, aperte em \n-> /editar \n\n Para voltar ao início, aperte \n-> /comecar", reply_markup=voltar)
        bot.register_next_step_handler(escolha, sair_ou_editar)
    except KeyError:
        chat_id = message.chat.id
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/comecar')
        b = types.KeyboardButton("/editarVer")
        voltar.row(a, b)
        bot.send_message(chat_id, "Algo deu errado \n\n Para voltar para o início, aperte \n-> /comecar \n\n Para tentar novamente, aperte \n-> /editarVer", reply_markup=voltar)
    except IndexError:
        gasto_taxi.clear()
        voltar = types.ReplyKeyboardMarkup()
        b = types.KeyboardButton('/comecar')
        voltar.row(b)
        bot.send_message(id(message), "Algo deu errado. Por favor, toque em:\n-> /comecar \n\n pois atualmente nenhuma planilha está selecionada", reply_markup=voltar)


def sair_ou_editar(message):
    if message.text == "/comecar":
        voltar = types.ReplyKeyboardMarkup()
        b = types.KeyboardButton('/comecar')
        voltar.row(b)
        bot.send_message(id(message), "Por favor, toque em:\n-> /comecar \n\n para confirmar sua escolha.", reply_markup=voltar)

    elif message.text == "/editar":
        try:
            lista = dependencias.listar_edicao(mes_taxi[0], arquivo[0])
            voltar = types.ReplyKeyboardRemove()
            nomesGastosEditar = bot.send_message(id(message), f'{lista}\ntoque no gasto que sera editado ou digite a sua numeração com uma barra na frente\n\nEx: Exemplo gasto nº /00001 - equivale ao gasto nº1', reply_markup=voltar)
            bot.register_next_step_handler(nomesGastosEditar, editar3)

        except KeyError and IndexError:
            chat_id = message.chat.id
            voltar = types.ReplyKeyboardMarkup()
            a = types.KeyboardButton('/comecar')
            b = types.KeyboardButton("/editarVer")
            voltar.row(a, b)
            bot.send_message(chat_id, "Algo deu errado \n\n Para voltar para o início, aperte \n-> /comecar \n\n Para tentar novamente, aperte \n-> /editarVer", reply_markup=voltar)


def editar3(message):
    try:
        escolha_editar_num = dependencias.funcao_editar_selecionar(message, arquivo[0], mes_taxi[0])
        row_taxi.append(escolha_editar_num[1])
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('Nome')
        b = types.KeyboardButton('Valor')
        voltar.row(a, b)
        opcao = bot.send_message(id(message), f"Para editar {escolha_editar_num[0]}, aperte em \n-> Nome. Para editar o valor, aperte em \n-> Valor", reply_markup=voltar)
        bot.register_next_step_handler(opcao, editar4)

    except ValueError:
        voltar = types.ReplyKeyboardMarkup()
        b = types.KeyboardButton('/comecar')
        voltar.row(b)
        bot.send_message(id(message), "Para voltar ao início, toque em \n-> /comecar", reply_markup=voltar)
    except KeyError:
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/comecar')
        voltar.row(a)
        bot.send_message(id(message), "Para voltar ao início, toque em \n-> /comecar", reply_markup=voltar)


def editar4(message):
    if message.text == "Nome":
        opcaoNumero_taxi.append("nome")
        markup = types.ForceReply()
        a = get_column_letter(1)
        cell = [a + row_taxi[0]]
        cell_taxi.append(cell)
        opcao = bot.send_message(id(message), "Digite o novo nome para o gasto", reply_markup=markup)
        bot.register_next_step_handler(opcao, editar5)
    elif message.text == "Valor":
        opcaoNumero_taxi.append("confirma")
        markup = types.ForceReply()
        b = get_column_letter(2)
        cell = [b + row_taxi[0]]
        cell_taxi.append(cell)
        opcao = bot.send_message(id(message), "Digite o novo valor para o gasto", reply_markup=markup)
        bot.register_next_step_handler(opcao, editar5)
    else:
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/editar')
        b = types.KeyboardButton('/comecar')
        voltar.row(a, b)
        bot.send_message(id(message), "Algo deu errado. Por favor, leia atentamente as mensagens que o bot enviar. \n\n Tente novamente com \n-> /editar \n\n Caso queira voltar para o início, toque em \n-> /comecar", reply_markup=voltar)


def editar5(message):
    wb = load_workbook(arquivo[0])
    ws = wb[mes_taxi[0]]
    mensagem1 = message.text
    listaMensagem = list(mensagem1)
    i = 0
    if opcaoNumero_taxi[0] != 'confirma':
        while i < len(listaMensagem):
            if listaMensagem[i] == ",":
                listaMensagem[i] = "."
            i += 1
        mensagem1 = ''.join(listaMensagem)
        ws[cell_taxi[0][0]].value = str(mensagem1)
        voltar = types.ReplyKeyboardMarkup()
        a = types.KeyboardButton('/editar')
        b = types.KeyboardButton('/comecar')
        voltar.row(a, b)
        bot.send_message(id(message), "Para editar mais alguma coisa, aperte \n-> /editar \n\n Para voltar ao início, aperte \n-> /comecar", reply_markup=voltar)
        wb.save(arquivo[0])
    else:
        if dependencias.floatcheck(message.text):
            while i < len(listaMensagem):
                if listaMensagem[i] == ",":
                    listaMensagem[i] = "."
                i += 1
            mensagem1 = ''.join(listaMensagem)
            ws[cell_taxi[0][0]].value = str(mensagem1)
            voltar = types.ReplyKeyboardMarkup()
            a = types.KeyboardButton('/editar')
            b = types.KeyboardButton('/comecar')
            voltar.row(a, b)
            bot.send_message(id(message), "Para editar mais alguma coisa, aperte \n-> /editar \n\n Para voltar ao início, aperte \n-> /comecar", reply_markup=voltar)
            wb.save(f"gastos de {dependencias.nome_pessoa(message)} {arquivo[0]}.xlsx")
        else:
            voltar = types.ReplyKeyboardMarkup()
            a = types.KeyboardButton('/editar')
            b = types.KeyboardButton('/comecar')
            voltar.row(a, b)
            bot.send_message(id(message), "Algo deu errado. Por favor, leia atentamente as mensagens que o bot enviar. \n\n Tente novamente com \n-> /editar \n\n Caso queira voltar para o início, toque em \n-> /comecar", reply_markup=voltar)
    gasto_taxi.clear()
    mes_taxi.clear()
    opcaoNumero_taxi.clear()
    row_taxi.clear()
    cell_taxi.clear()


@bot.message_handler(func=lambda m: True)
def echo_all(message):
    voltar = telebot.types.ReplyKeyboardMarkup()
    a = telebot.types.KeyboardButton('/comecar')
    voltar.row(a)
    bot.send_message(id(message), "Para iniciar, toque em /comecar", reply_markup=voltar)


bot.infinity_polling()
