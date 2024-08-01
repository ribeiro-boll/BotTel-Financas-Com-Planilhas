def floatcheck(a):
    b = list(str(a))
    c = 0
    for i in b:
        if i in "0123456789.,":
            c += 1
    if c == len(a):
        return True
    else:
        return False

def criarArquivoAno(nome, ano):
    from openpyxl import Workbook
    arquivo = f"gastos de {nome} {ano}.xlsx"
    try:
        open(arquivo, 'r')
        return arquivo
    except IOError:
        resumoAnual = Workbook()
        meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
        for mes in meses:
            resumoAnual.create_sheet(mes.capitalize())
        resumoAnual.save(arquivo)
        return arquivo


def opcoesAno():
    import datetime
    i = 0
    j = 0
    grupoAnos = []
    grupoAnos.clear()
    data = str(datetime.datetime.now())
    ano = int((data[0] + data[1] + data[2] + data[3]))
    ano1 = int((data[0] + data[1] + data[2] + data[3]))
    ano2 = int((data[0] + data[1] + data[2] + data[3]))
    k = ano2 - 6
    while j < 6:
        k += 1
        grupoAnos.append(k)
        j += 1
    grupoAnos.append(ano)
    while i < 4:
        ano1 += 1
        grupoAnos.append(ano1)
        i += 1
    return grupoAnos

def seuMes():
    import datetime
    a = str(datetime.date.today())[5] + str(datetime.date.today())[6]
    match int(a):
        case(1):
            return 'jan'
        case(2):
            return 'fev'
        case (3):
            return "mar"
        case (4):
            return "abr"
        case (5):
            return "mai"
        case (6):
            return "jun"
        case (7):
            return "jul"
        case (8):
            return "ago"
        case(9):
            return "set"
        case (10):
            return "out"
        case (11):
            return "nov"
        case (12):
            return "des"

