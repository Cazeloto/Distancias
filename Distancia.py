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
            try:
                r = requests.get(route)
                if r.status_code != 200:
                    print(r.status_code)
                    cont = 0
                    while (r.status_code != 200) and (cont < 10):
                        cont = cont + 1
                        r = requests.get(route)
                        print(r.status_code)
                dt = json.loads(r.content)
                travelD = '%.2f' % (dt['resourceSets'][0]['resources'][0]['routeLegs'][0]['travelDistance'])
            except:
                print(r.content)
                quit()
                
            
            travel = travelD.replace('.', ',')
            linha = cidOri + ";" + estOri + ";" + cidDes + ";" + estDes + ";" + travel + "\n"

            with open(pathDestino + "\distancias.csv", "a") as arquivo:
                arquivo.write(linha)
            print(linha)
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

    messagebox.showinfo(title="Calculator Sonda", message="Fim de Processamento")



if __name__ == '__main__':
    window = Tk()
    window.title("Sonda Distancia Calculator")
    window.config(padx=30, pady=100)

    # Labels
#    website_label = Label(text="URL:")
#    website_label.grid(row=2, column=0)

    # Entries
#    website_entry = Entry(width=35)
#    website_entry.grid(row=2, column=1, columnspan=2)
#    website_entry.focus()
    add_button = Button(text="Iniciar Processamento", width=36, command=processa)
    add_button.grid(row=6, column=1, columnspan=2)
    window.mainloop()
