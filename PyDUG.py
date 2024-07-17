from pyautocad import Autocad, APoint
import win32com.client
import pandas as pd
import PySimpleGUI as sg


def main():

    values = []

    # definição das perguntas 
    definicao(values)

    number = int(values[0])
    endereco = values[1]
    aba = values[2]

    par1 = []
    par2 = []
    par3 = []
    par4 = []

    leitura(number,endereco,aba,par1,par2,par3,par4)
    
    acad = Autocad()
    #for entity in acad.ActiveDocument.ModelSpace:
        #move(entity,number)

    acad = win32com.client.Dispatch("AutoCAD.Application")
    item = 0
    for entity in acad.ActiveDocument.ModelSpace:
        list = par1
        list1 = par2
        list2 = par3
        list3 = par4

        mano = list[item]
        bro = list1[item]
        broda = list2[item]
        circuit = list3[item]

        if broda == 'F+N':
            Npolos = '1P'
            polos = 'I'

        elif broda == '2F':
            Npolos = '2P'
            polos = 'II'

        elif broda == '3F':
            Npolos = '3P'
            polos = 'III'


        edit(entity,bro,mano,circuit,Npolos,polos)
        item += 1


def move(entity,circuito):
        add = APoint(13,0)
        for item in range(circuito-1):
            local = APoint(entity.InsertionPoint)
            new_local = APoint(local + add)
            copy = entity.Copy()
            copy.Move(APoint(local), APoint(new_local))
            item+=1
            add = add + [13,0,0]

def edit(entity,new,young,current,poles,number):
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                for attrib in entity.GetAttributes():
                    #print("  {}: {}".format(attrib.TagString, attrib.TextString))
                    if attrib.TagString == 'NOME':
                        if attrib.TextString == 'ZZZZ':
                            attrib.TextString =  new
                            attrib.Update()

                    elif attrib.TagString == 'CIRCUITO':
                        if attrib.TextString == 'YYYY':
                            attrib.TextString = young
                            attrib.Update()

                    elif attrib.TagString == 'IN_DJ_V':
                        if attrib.TextString == 'WWA':
                            attrib.TextString = current
                            attrib.Update()

                    elif attrib.TagString == 'POLOS_DJ':
                        if attrib.TextString == 'KKP':
                            attrib.TextString =  poles
                            attrib.Update() 

                    elif attrib.TagString == 'III_V':
                        if attrib.TextString == 'XX':
                            attrib.TextString =  number
                            attrib.Update()
            
def leitura(circuito,adress,aba,ctag,cname,csupply,ccurrent):
    data = pd.read_excel(adress, sheet_name=aba,skiprows=17)

    values = [4,5,6,14,circuito]
    for CB in range(circuito):
        row = data.values[CB]
        for info in range(4):
            cell = row[values[info]]
            if info == 0:
                ctag.append(cell)
            elif info == 1:
                cname.append(cell)
            elif info == 2:
                csupply.append(cell)
            elif info == 3:
                intcurrent = int(cell)
                strcurrent = str(intcurrent)
                ampher = strcurrent+'A'
                ccurrent.append(ampher)
            info += 1
    CB += 1

def definicao(info):
    # Definir textos e botões
    layout = [  [sg.Text('Número de disjuntores'),sg.InputText()],
            [sg.Text('Endereço do arquivo'),sg.InputText()],
            [sg.Text('Aba editada'),sg.InputText()],
            [sg.Button('OK'), sg.Button('Cancel')]]

    # Criar a janela
    window = sg.Window('Bem vindo ao PyDUG!', layout)

    #Loop para processar os eventos e fechar a janela
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
            break
        print('You entered ', values)

    #fecha a janela
        break

    window.close()
    info.append(values[0])
    info.append(values[1])
    info.append(values[2])      
    
main()