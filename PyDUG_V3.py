from pyautocad import Autocad, APoint
import win32com.client
import pandas as pd
import PySimpleGUI as sg


def main():

    

    values = []

    # definição das perguntas 
    definicao(values)

    quadro = int(values[0])
    endereco = values[1]
    soma = 0


    for counter in range(quadro):

        repeat = 0
        item = 0
        num = 0
        valores = []
        planilha(valores)

        number = int(valores[0])
        aba = valores[1]

        par1 = ['YYYY','YYYY']
        par2 = ['ZZZZ','ZZZZ']
        par3 = ['XX','XX']
        par4 = ['WWA','WWA']
        par5 = ['3ɸ-380/220V']
        par6 = ['QD-XX']
        par6.append(aba)


        leitura(number,endereco,aba,par1,par2,par3,par4,par5)
        acad = Autocad()
        n1 = 0


        for entity in acad.ActiveDocument.ModelSpace:
            if n1 == 0:
                move(entity,number+1,counter)
                n1 += 1

            elif n1 <= 5:
                move(entity,1,counter)
                n1 += 1
            

        acad = win32com.client.Dispatch("AutoCAD.Application")

        for entity in acad.ActiveDocument.ModelSpace:
            list = par1
            list1 = par2
            list2 = par3
            list3 = par4
                

            mano = list[item]
            bro = list1[item]
            broda = list2[item]
            circuit = list3[item]
                
            if '2F' in broda:
                Npolos = '2P'
                polos = 'II'

            elif '3F' in broda:
                Npolos = '3P'
                polos = 'III'
                
            else:
                Npolos = '1P'
                polos = 'I'

            if len(list) < number+4:
                if number <= 6:
                    list.insert(-1,'X2')

                elif number <= 12:
                    list.insert(-1,'X3')

                elif number <= 30:
                    list.insert(-1,'X4')

                elif number > 30:
                    reserva = number*0.15
                    reserva = int(reserva)
                    reserva = str(reserva)
                    reserva = 'X'+reserva
                    list.insert(-1,reserva)

                
                list1.insert(-1,'ESPAÇO RESERVA')
                list2.insert(-1,'XX')
                list3.insert(-1,'WWA')

            name = entity.EffectiveName
            if name == 'CIRCUITO' and repeat >= soma:
                edit1(entity,bro,mano,circuit,Npolos,polos)
                if item < number+3:
                    item+=1
                else:
                    soma = soma+item+3
            
            repeat+=1

        for entity in acad.ActiveDocument.ModelSpace:

            list4 = par5
            list5 = par6

            nome = list5[num]
            tensao = list4[num]

            name = entity.EffectiveName
            if name == 'E-TXT QD':
                edit2(entity,bro,mano,circuit,Npolos,polos,tensao,nome)
                if num < 1:
                    num+=1
                
                    


def move(entity,circuito,quadro):
        add = APoint(13,0)
        first_local = APoint(2000,100)-(APoint(400,0)*quadro)
        for item in range(circuito):

            local = APoint(entity.InsertionPoint)
            if item == 0:
                new_local = APoint(local+first_local)
            else:
                new_local = APoint(local+first_local + add)
                add = add + [13,0,0]
            

            copy = entity.Copy()
            copy.Move(APoint(local), APoint(new_local))
            

def edit1(entity,new,young,current,poles,number):

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

                    elif attrib.TagString == 'IN_DJ_H':
                        if attrib.TextString == 'WWA':
                            attrib.TextString =  current
                            attrib.Update()

                    elif attrib.TagString == 'POLOS_DJ_H':
                        if attrib.TextString == 'KKP':
                            attrib.TextString =  poles
                            attrib.Update()

                    elif attrib.TagString == 'III_H':
                        if attrib.TextString == 'XX':
                            attrib.TextString =  number
                            attrib.Update()
                    

def edit2(entity,tensao,nome):

            HasAttributes = entity.HasAttributes
            if HasAttributes:
                for attrib in entity.GetAttributes():           
                    if attrib.TagString == '3ɸ-YYY/XXXV':
                        if attrib.TextString == '3ɸ-380/220V':
                            attrib.TextString =  tensao
                            attrib.Update()

                    elif attrib.TagString == 'QD-XX':
                        if attrib.TextString == 'QD-XX':
                            attrib.TextString =  nome
                            attrib.Update()
                    
            
def leitura(circuito,adress,aba,ctag,cname,csupply,ccurrent,cvolt):

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

    volt = pd.read_excel(adress, sheet_name=aba)
    
    values = [6,8]
    for DB in range(2):
        row = volt.values[DB+2]
        for info in range(2):
            cell = row[values[info]]
            if info == 0 and DB==0:
                if cell == 220:
                    cell =  '3ɸ-220/127V'

                elif cell == 440:
                    cell = '3ɸ-440/380V'

                elif cell == 380:
                    cell = '3ɸ-380/220V'

                else:
                    cell = '3ɸ-380/220V'

                cvolt.append(cell)
            elif info ==1 and DB==0:
                ctag.append('YYYY')
                cname.append('ZZZZ')
                csupply.append(cell)
                
            elif info ==1 and DB==1:
                intcurrent = int(cell)
                strcurrent = str(intcurrent)
                ampher = strcurrent+'A'
                ccurrent.append(ampher)
            

def definicao(info):
    # Definir textos e botões
    layout = [  [sg.Text('Número de quadros'),sg.InputText()],
            [sg.Text('Nome do arquivo'),sg.InputText()],
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

def planilha(data):
    # Definir textos e botões
    layout = [  [sg.Text('Número de disjuntores'),sg.InputText()],
            [sg.Text('Aba editada'),sg.InputText()],
            [sg.Button('OK'), sg.Button('Cancel')]]

    # Criar a janela
    window = sg.Window('Defina as informações do painel', layout)

    #Loop para processar os eventos e fechar a janela
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
            break
        print('You entered ', values)

    #fecha a janela
        break

    window.close()
    data.append(values[0])
    data.append(values[1])  
    
main()