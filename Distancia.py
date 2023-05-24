import json
import time
import urllib.parse
import pandas as pd
import requests

from tkinter import filedialog, messagebox, Tk, Label, Entry, Button
#bingKey = "Av1PKSJKBf7UelqTaSCjNXjo0Qfo6HmPtoarsL6OB0wJKCrlJ6cWa6KEWARmmpRu"
global cidOri
global cidDes
global travelD
global menorCid
global menorEst
global menorDist


def processa():
    BingKeyFile = filedialog.askopenfilename(title = "Chave Bing", filetypes = (("Ini file", "*.ini"), ("all files", "*.*")))
    cidOrigem = filedialog.askopenfilename(title = "Arquivo de Cidades", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
    cidBase = filedialog.askopenfilename(title = "Arquivo de Base", filetypes = (("Excel files", "*.xlsx"), ("all files", "*.*")))
    pathDestino = filedialog.askdirectory()
    
    with open(BingKeyFile, "r") as arquivo:
        bingKey = arquivo.read()
        
   
                
    arqOri = pd.read_excel(cidOrigem)
    arqDes = pd.read_excel(cidBase) 
    
    menorDist = 20000

    lenOri = len(arqOri)
    lenDes = len(arqDes)
   
    i = 0
    while lenOri > i:
        menorDist = 20000
        d = 0
        while lenDes > d:
            cidOri = arqOri['Cidade'][i]
            estOri = arqOri['Estado'][i]
            cidDes = arqDes['Cidade'][d]
            estDes = arqDes['Estado'][d]

            origem = urllib.parse.quote(cidOri + "," + estOri)
            destino = urllib.parse.quote(cidDes + "," + estDes)

            route = "http://dev.virtualearth.net/REST/v1/Routes?wayPoint.1=" + \
                    origem + "&wayPoint.2=" + destino + "&key=" + bingKey

            r = requests.get(route)
            if r.status_code != 200:
                print(r.status_code)
                cont = 0
                while (r.status_code != 200) and (cont < 10):
                    cont = cont + 1
                    r = requests.get(route)
                    print(r.status_code)
            if r.status_code == 404:
                messagebox.showerror(title = "Erro de Rota", message="Cidade ou rota não encontrada  :" + cidOri + "-" + cidDes)
            if (r.status_code == 401) or (r.status_code==403):
                messagebox.showerror(title = "Erro de Autenticação", message="Erro de autenticação ou serviço indisponivel - Verifique sua chave BING")

            dt = json.loads(r.content)
            travelD = '%.2f' % (dt['resourceSets'][0]['resources'][0]['routeLegs'][0]['travelDistance'])

            travel = travelD.replace('.', ',')
            linha = cidOri + ";" + estOri + ";" + cidDes + ";" + estDes + ";" + travel + "\n"

            with open(pathDestino + "\distancias.csv", "a") as arquivo:
                arquivo.write(linha)
            print(linha)
            espaco = "                              "
            label1_1 = Label(Label(window, font=("Lucida Console", 12), text=espaco).place(x=150, y=80), )
            label1_1 = Label(Label(window, font=("Lucida Console", 12), text=cidOri).place(x=150, y=80), )
            label2_2 = Label(Label(window, font=("Lucida Console", 12), text=espaco).place(x=150, y=140), )
            label2_2 = Label(Label(window, font=("Lucida Console", 12), text=cidDes).place(x=150, y=140), )
            label2_2 = Label(Label(window, font=("Lucida Console", 12), text=espaco).place(x=150, y=200), )
            label2_2 = Label(Label(window, font=("Lucida Console", 12), text=travel).place(x=150, y=200), )
            label8 = Label(Label(window, font=("Lucida Console", 12), text=(str(d+1)+"/"+str(lenDes))).place(x=40, y=260), )

            label1_1.pack()
            label2_2.pack()
            label3_3.pack()
            label8.pack()


            window.update_idletasks()

            d = d + 1

            if (float(travelD) < menorDist):
                menorDist = float(travelD)
                menorCid = cidDes
                menorEst = estDes

        i = i + 1
        menDis = str(menorDist).replace('.', ',')
        linha2 = cidOri + ";" + estOri + ";" + menorCid + ";" + menorEst + ";" +str(menDis) + "\n"



        with open(pathDestino + "\menor_distancia.csv", "a") as arquivo:
            arquivo.write(linha2)
        print(linha2)

        espaco = "                              "
        label5_1 = Label(Label(window, font=("Lucida Console", 12), text=espaco).place(x=590, y=80), )
        label5_1 = Label(Label(window, font=("Lucida Console", 12), text=cidOri).place(x=590, y=80), )
        label6_2 = Label(Label(window, font=("Lucida Console", 12), text=espaco).place(x=590, y=140), )
        label6_2 = Label(Label(window, font=("Lucida Console", 12), text=menorCid).place(x=590, y=140), )
        label7_3 = Label(Label(window, font=("Lucida Console", 12), text=espaco).place(x=590, y=200), )
        label7_3 = Label(Label(window, font=("Lucida Console", 12), text=menDis).place(x=590, y=200), )
        label9 = Label(Label(window, font=("Lucida Console", 12), text=(str(i) + "/" + str(lenOri))).place(x=480, y=260), )

        label5_1.pack()
        label6_2.pack()
        label7_3.pack()
        label9.pack()

        window.update_idletasks()

    messagebox.showinfo(title="Calculator Sonda", message="Fim de Processamento")



if __name__ == '__main__':
    window = Tk()
    window.title("Sonda Distancia Calculator")
    window.geometry("900x450")


    # Labels
    label0 = Label(Label(window, font=("Lucida Console", 14), text="Distâncias entre Cidade e Base").place(x=40, y=30), )
    label1 = Label(Label(window, font=("Lucida Console", 12), text="Cidade    :").place(x=40, y=80), )
    label2 = Label(Label(window, font=("Lucida Console", 12), text="Base      :").place(x=40, y=140), )
    label3 = Label(Label(window, font=("Lucida Console", 12), text="Distância :").place(x=40, y=200), )
    label4 = Label(Label(window, font=("Lucida Console", 14), text="Menor Distância").place(x=480, y=30), )
    label5 = Label(Label(window, font=("Lucida Console", 12), text="Cidade    :").place(x=480,y=80),)
    label6 = Label(Label(window, font=("Lucida Console", 12), text="Base      :").place(x=480,y=140),)
    label7 = Label(Label(window, font=("Lucida Console", 12), text="Distância :").place(x=480, y=200), )
    label1_1 = Label(Label(window, font=("Lucida Console", 12), text="_______________________").place(x=150, y=80), )
    label2_2 = Label(Label(window, font=("Lucida Console", 12), text="_______________________").place(x=150, y=140), )
    label3_3 = Label(Label(window, font=("Lucida Console", 12), text="_______________________").place(x=150, y=200), )
    label5_1 = Label(Label(window, font=("Lucida Console", 12), text="_______________________").place(x=590, y=80), )
    label6_2 = Label(Label(window, font=("Lucida Console", 12), text="_______________________").place(x=590, y=140), )
    label7_3 = Label(Label(window, font=("Lucida Console", 12), text="_______________________").place(x=590, y=200), )
    #label8 = Label(Label(window, font=("Lucida Console", 12), text="XX/XX").place(x=40, y=260), )
    #label9 = Label(Label(window, font=("Lucida Console", 12), text="XX/XX").place(x=450, y=260), )


    # Entries
#    website_entry = Entry(width=35)
#    website_entry.grid(row=2, column=1, columnspan=2)
#    website_entry.focus()
    add_button = Button(text="Iniciar Processamento", width=36, command=processa, font=("Lucinda Console", 14))
    add_button.place(x=250, y=320)
#    root.update_idletasks()
    window.mainloop()
