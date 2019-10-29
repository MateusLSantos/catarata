import os
import datetime
import subprocess
import pandas
import tkinter as tk
from tkinter import messagebox
#Inicialização de variáveis
currDir = os.getcwd()
patients = [];
errors = [];
#chrome_path = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
chrome_path = "firefox"
stop = False
save = False
forward = 0
data = pandas.read_excel("Output.xlsx")
erros = pandas.read_excel("Erros.xlsx")
xyProntuario = "500x250"
xyFichas = "500x350"

os.chdir("./Pacientes")
patients = sorted(os.listdir("."))

def ordenarArquivos(files):
    sortedFiles = sorted(files,key=lambda x:int(x.split("-")[0].rstrip()))
    return sortedFiles

#Retorna os nomes de todos os arquivos no diretório atual
def storeFiles():
    for root, directory, file in os.walk('.'):
        pass
        #files.append(file)
    files = ordenarArquivos(file)
    return files

#Inicializa um dataframe para o paciente
def createPatientDF(name):
    patientDF = pandas.DataFrame(columns = ['Nome', 'Registro', "Data da cirurgia", 'Idade', "Olho", "Dioptria", "Marca da Lente",
                                            "Modelo da Lente","Data Pré-op", "Esférico Pré-op", "Cilindro Pré-op",
                                            "Eixo Pré-op", 'Acuidade Pré-op', 'AR Pré-op', "Data Pós-op", "Esférico Pós-op", "Cilindro Pós-op",
                                            "Eixo Pós-op", 'Acuidade Pós-op', 'AR Pós-op'])
    patientDF.loc[0] = [name, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
                        None]
    return patientDF

#Função que abre cada prontuário cirúrgico na pasta atual
def getSurgeries(patientDF, files):
    global stop
    global forward
    global erros
    prontuarios = list(reversed([file for file in files if "PRONTU" in file]))
    utilizaveis = []
    ind = 0
    patientDF = patientDF
    while(ind<len(prontuarios) and not stop):
        subprocess.Popen([chrome_path, prontuarios[ind]])
        patientDF, pront = prontuario(patientDF, ind+1, len(prontuarios))
        if patientDF['Nome'][0] in errors:
            return patientDF, [];
        if stop:
            return createPatientDF('STOP'), [];
        if(forward == 1):
            ind+=1
            forward = 0
        elif(forward == -1):
            if(ind!=0):
                ind-=1
            forward = 0
        else:
            raise Exception("{} inválido, 1 ou -1 esperado.".format(forward))
        if pront > 0:
            utilizaveis.append(prontuarios[ind-1])
    if len(utilizaveis)>2 or len(utilizaveis) == 0:
        addToErrors(patientDF, "muitosOuNenhumUtilizavel")
    elif len(utilizaveis) == 2:
        if patientDF.loc[1]['Olho'] == patientDF.loc[2]['Olho']:
            addToErrors(patientDF, "olhoRepetido")
            return patientDF, []

    if patientDF.loc[0]['Olho'] == None:
        patientDF = patientDF.drop(index=0).reset_index().drop(columns=['index'])#ok

    #os.system('taskkill /f /im chrome.exe')

    return patientDF, utilizaveis


def addToErrors(patientDF, reason=-1):
    errors.append(patientDF['Nome'][0])  # global variable access
    if not isinstance(erros, pandas.DataFrame):
        raise Exception("{} não é do tipo esperado, esperava {}".format(erros.__class__, pandas.DataFrame.__class__))
    lastPos = erros.shape[0]
    erros.loc[lastPos, ["Nome", "Motivo"]] = [errors[-1], reason]  # global variable access




#Gambiarra, salva os dados no prontuário e avança para a próxima tela
def storeSave(instance,flag):
    global save
    save = True;
    flag[0] = True
    avancar(instance)

#Função que cria a interface e lida com os dados das fichas de atendimento
def prontuario(patientDF, currFile, totalFiles):

    pront = 0

    global save

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.geometry(xyProntuario)
    root.title(patientDF['Nome'][0] + ' - Prontuário ' + str(currFile) + ' de ' + str(totalFiles))
    tk.Label(root, text="Registro").grid(row=0)
    tk.Label(root, text="Data da Cirurgia ('DDMMAAAA')").grid(row=1)
    tk.Label(root, text="Idade (Anos Completos)").grid(row=2)
    tk.Label(root, text="Olho (OE ou OD)").grid(row=3)
    tk.Label(root, text="Dioptria").grid(row=4)
    tk.Label(root, text="Marca da Lente").grid(row=5)
    tk.Label(root, text="Modelo da Lente").grid(row=6)

    reg = tk.StringVar()
    data = tk.StringVar()
    idade = tk.StringVar()
    olho = tk.StringVar()
    dioptria = tk.StringVar()
    marca = tk.StringVar()
    modelo = tk.StringVar()

    regEntry = tk.Entry(root, textvariable=reg)
    dataEntry = tk.Entry(root, textvariable=data)
    idadeEntry = tk.Entry(root, textvariable=idade)
    olhoEntry = tk.Entry(root, textvariable=olho)
    dioptriaEntry = tk.Entry(root, textvariable=dioptria)
    marcaEntry = tk.Entry(root, textvariable=marca)
    modeloEntry = tk.Entry(root, textvariable=modelo)

    regEntry.grid(row=0, column=1)
    dataEntry.grid(row=1, column=1)
    idadeEntry.grid(row=2, column=1)
    olhoEntry.grid(row=3, column=1)
    dioptriaEntry.grid(row=4, column=1)
    marcaEntry.grid(row=5, column=1)
    modeloEntry.grid(row=6, column=1)

    tk.Button(root, text="Anterior", command=lambda: voltar(root)).grid(row=7, column=0, pady=4)
    tk.Button(root, text="Ok", command=lambda: storeSave(root,[None])).grid(row=7, column=1, pady=4)
    tk.Button(root, text="Próximo", command=lambda: avancar(root)).grid(row=7, column=2, pady=4)
    #tk.Button(root, text="Erro", command= lambda: salvarErro(patientDF,root)).grid(row=8, column=0, pady=4)
    tk.Button(root, text="Sair", command=lambda: sair(root)).grid(row=8, column=1)
    tk.Button(root, text="Checar data", command=lambda: descricoesCirurgicas()).grid(row=8, column=2)

    root.mainloop()

    if save:
        patientDF = salvarDadosCir(reg.get(), data.get(), idade.get(), olho.get().upper(), dioptria.get(), marca.get(),
                                                      modelo.get(), patientDF)
        save = False

    #Caso não esteja vazio, consideramos que o prontuário é válido
    if olho.get().upper() == 'OE' or olho.get().upper() == 'OD':
        pront = 1

    return patientDF, pront

#Função do botão
def descricoesCirurgicas():
    for file in os.listdir():
        if "DESCR" in file:
            subprocess.Popen([chrome_path, file])


#Salva os dados coletados na interface para o dataframe do paciente, criando uma entrada para cada cirurgia.
def salvarDadosCir(reg, data, idade, olho, dioptria, lente, modelo, patientDF):
    name = patientDF['Nome'][0]
    #Processando as variáveis
    try:
        data = processarData(data)
    except:
        pass

    try:
        dioptria = processarFloat(dioptria)
    except:
        pass

    try:
        if int(idade) > 120:
            print("Idade = " + str(idade))
            dataNasc = processarData(idade)
            print("Data de nascimento: " + str(dataNasc))
            idadeFinal = round((datetime.datetime.now() - datetime.datetime(dataNasc.year, dataNasc.month, dataNasc.day)).days/365)
            print("Idade final: " + str(idadeFinal))
            idade = idadeFinal
        else:
            idade = processarFloat(idade)
    except Exception as E:
        print(E)


    patientDF.loc[patientDF.shape[0]] = [name, reg, data, idade, olho, dioptria, lente, modelo, None, None, None, None, None, None,
                                         None, None, None, None, None, None]
    return patientDF

#Salva os dados coletados na interface para o dataframe do paciente
def salvarDadosFicha(data, esfE, esfD, cilE, cilD, eixoE, eixoD, acuE, acuD, ar, patientDF, pre):
    name = patientDF['Nome'][0]
    #Processando floats e data
    try:
        data = processarData(data)
    except:
        pass
    try:
        esfE = processarFloat(esfE)
    except:
        pass
    try:
        esfD = processarFloat(esfD)
    except:
        pass
    try:
        cilE = processarFloat(cilE)
    except:
        pass
    try:
        cilD = processarFloat(cilD)
    except:
        pass
    try:
        eixoE = processarFloat(eixoE)
    except:
        pass
    try:
        eixoD = processarFloat(eixoD)
    except:
        pass

    if ar:
        ar = True
    else:
        ar = False

    for ind in range(patientDF.shape[0]):
        oldRow = list(patientDF.loc[ind])
        newRow = oldRow.copy()
        if patientDF.loc[ind]['Olho'] == 'OE':
            if pre:
                newRow[8] = data
                newRow[9] = esfE
                newRow[10] = cilE
                newRow[11] = eixoE
                newRow[12] = acuE
                newRow[13] = ar
            else:
                newRow[14] = data
                newRow[15] = esfE
                newRow[16] = cilE
                newRow[17] = eixoE
                newRow[18] = acuE
                newRow[19] = ar
        else:
            if pre:
                newRow[8] = data
                newRow[9] = esfD
                newRow[10] = cilD
                newRow[11] = eixoD
                newRow[12] = acuD
                newRow[13] = ar
            else:
                newRow[14] = data
                newRow[15] = esfD
                newRow[16] = cilD
                newRow[17] = eixoD
                newRow[18] = acuD
                newRow[19] = ar
        patientDF.loc[ind] = newRow
    return patientDF

#Volta para a última ficha ou prontuário
def voltar(instance):
    global forward
    instance.destroy()
    forward = -1

#Vai para a próxima ficha ou prontuário
def avancar(instance):
    global forward
    instance.destroy()
    forward = 1

#Registra que há um erro neste paciente e avança para o próximo.
def salvarErro(patientDF,instance):
    global erros
    errors.append(patientDF['Nome'][0])
    erros.at[erros.shape[0],'Nome'] = errors[-1]
    instance.destroy()

#Para o programa
def sair(instance):
    global stop
    instance.destroy()
    stop = True;


#Transforma uma string do tipo "DDMMYYYY" em um datetime
def processarData(data):
    if '/' in data:
        data = data.replace('/', '')
    elif '-' in data:
        data = data.replace('-', '')
    if len(data) == 6:
        data = data = data[0:4] + '20' + data[4:]
    day = int(data[0:2])
    month = int(data[2:4])
    year = int(data[4:])
    return datetime.date(year, month, day)

#Transforma uma string com a dioptria em um float. Lida tanto com vírgula quanto com ponto.
def processarFloat(string):
    if ',' in string:
        string = string.replace(',','.')

    return float(string)

def ficha(patientDF, pre, currFile, totalFiles):

    global save
    flag = [False]
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.title(patientDF['Nome'][0] + ' - Ficha' + str(currFile) + ' de ' + str(totalFiles))
    root.geometry(xyFichas)

    tk.Label(root, text="Data").grid(row=0)
    tk.Label(root, text="Esférico OD").grid(row=1)
    tk.Label(root, text="Cilindro OD").grid(row=2)
    tk.Label(root, text="Eixo OD").grid(row=3)
    tk.Label(root, text="Acuidade OD").grid(row=4)
    tk.Label(root, text="Esférico OE").grid(row=5)
    tk.Label(root, text="Cilindro OE").grid(row=6)
    tk.Label(root, text="Eixo OE").grid(row=7)
    tk.Label(root, text="Acuidade OE").grid(row=8)
    tk.Label(root, text="Usei AR (Deixar vazio se não usou)").grid(row=9)
    tk.Label(root, text="Data da Cirurgia: " + str(patientDF['Data da cirurgia'][0].day) + '/' + str(patientDF['Data da cirurgia'][0].month) + '/' + str(patientDF['Data da cirurgia'][0].year) ).grid(row=12)

    data = tk.StringVar()
    esfE = tk.StringVar()
    esfD = tk.StringVar()
    cilE = tk.StringVar()
    cilD = tk.StringVar()
    eixoE = tk.StringVar()
    eixoD = tk.StringVar()
    acuE = tk.StringVar()
    acuD = tk.StringVar()
    ar = tk.StringVar()

    dataEntry = tk.Entry(root, textvariable=data)
    esfDEntry = tk.Entry(root, textvariable=esfD)
    cilDEntry = tk.Entry(root, textvariable=cilD)
    eixoDEntry = tk.Entry(root, textvariable=eixoD)
    acuDEntry = tk.Entry(root, textvariable=acuD)
    esfEEntry = tk.Entry(root, textvariable=esfE)
    cilEEntry = tk.Entry(root, textvariable=cilE)
    eixoEEntry = tk.Entry(root, textvariable=eixoE)
    acuEEntry = tk.Entry(root, textvariable=acuE)
    arEntry = tk.Entry(root, textvariable=ar)

    dataEntry.grid(row=0, column=1)
    esfDEntry.grid(row=1, column=1)
    cilDEntry.grid(row=2, column=1)
    eixoDEntry.grid(row=3, column=1)
    acuDEntry.grid(row=4, column=1)
    esfEEntry.grid(row=5, column=1)
    cilEEntry.grid(row=6, column=1)
    eixoEEntry.grid(row=7, column=1)
    acuEEntry.grid(row=8, column=1)
    arEntry.grid(row=9, column=1)


    tk.Button(root, text="Anterior", command= lambda: voltar(root)).grid(row=10, column=0, pady=4)
    tk.Button(root, text="Ok", command=       lambda: storeSave(root,flag)).grid(row=10, column=1, pady=4)
    tk.Button(root, text="Próximo", command=  lambda: avancar(root)).grid(row=10, column=2, pady=4)
    #tk.Button(root, text="Erro", command=     lambda: salvarErro(patientDF,root)).grid(row=11, column=0, pady=4)
    tk.Button(root, text="Sair", command=     lambda: sair(root)).grid(row=11, column=1)

    root.mainloop()

    if save:
        patientDF = salvarDadosFicha(data.get(), esfE.get(), esfD.get(), cilE.get(), cilD.get(),
                                                      eixoE.get(), eixoD.get(), acuE.get(), acuD.get(), ar.get(), patientDF, pre)
        save = False

    return patientDF,flag[0]

def quebrarLista(pronts, files):
    end = files[0:files.index(pronts[-1])], files[files.index(pronts[0])+1:]
    end = list(map(lambda x:list(filter(lambda y:"FICHA" in y,x)),end))
    return end

def getFichas(patientDF, fichas, boolean):
    global stop
    global forward
    ind = 0
    patientDF = patientDF
    while not stop:
        if patientDF['Nome'][0] in errors:
            return (createPatientDF(None).drop(index=0))
        subprocess.Popen([chrome_path, fichas[ind]])
        patientDF, flag = ficha(patientDF, boolean, ind+1, len(fichas))
        if stop:
            return createPatientDF("STOP");
        if patientDF['Nome'][0] in errors:
            break;
        if(forward == 1):
            if (ind != len(fichas) - 1):
                ind+=1
            forward = 0
        elif(forward == -1):
            if(ind != 0):
                ind-=1
            forward = 0
        else:
            raise Exception("{} inválido, 1 ou -1 esperado.".format(forward))
        if(flag):
            break
    #os.system('taskkill /f /im chrome.exe')
    return patientDF

#Iterando pelos pacientes

#ATENÇÃO! O CHROME PRECISA ESTAR ABERTO PARA FUNCIONAR

for patient in patients:
    if not stop and patient not in list(data['Nome']) and patient not in (list(erros['Nome'])):
        os.chdir(currDir + "/Pacientes/" + patient)
        files = storeFiles()
        patientDF = createPatientDF(patient)
        patientDF, pronts = getSurgeries(patientDF, files)
        if len(pronts) > 2 or len(pronts) == 0 or patientDF['Data da cirurgia'][0] == None:
            os.chdir(currDir)
            erros.to_excel('Erros.xlsx')
            continue
        else:
            listaPos, listaPre = quebrarLista(pronts, files)
            patientDF = getFichas(patientDF,listaPre,True)
            if patientDF.shape[0] == 0 :
                os.chdir(currDir)
                erros.to_excel('Erros.xlsx')
                continue
            patientDF = getFichas(patientDF,list(reversed(listaPos)),False)
            if patientDF.shape[0] == 0 :
                os.chdir(currDir)
                erros.to_excel('Erros.xlsx')
                continue
            #Checar se foi gerado algum erro
            if patientDF['Nome'][0] != 'STOP':
                data = pandas.concat([data,patientDF])
        os.chdir(currDir)
        data.to_excel('Output.xlsx')


print("Seu bloquinho acabou, pegue o próximo!!")
