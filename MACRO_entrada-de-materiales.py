import sys
import pyautogui
import time
import os
from dotenv import load_dotenv


#btposY=0 # ajusta la posicion en Y del cursor
#btposX=0 # ajusta la posicion en X del cursor

#────────────────────────────────────
#  Global Vars
#────────────────────────────────────
# bspwm + vertical tabs  
# btposY=0
# btposX=0
btposY=0
btposX=0


# Botones de fieldmanager
btPlus=[253, 400] #[255,400] #boton de agregar consumible
btEnviar=[1055, 642] #enviar

# Var de delays (compu de mierda se llama) 
# delays
tiempoUsual=1.1
delayDeEscritura=(tiempoUsual/50)
tiempoPagina=tiempoUsual*5
pyautogui.PAUSE = (tiempoUsual/100) #as 0.01 secs

# cargar las variables y credenciales
load_dotenv("/env/env.sh", override=True)

# Ahora podés acceder a las variables
wfc_User = os.getenv("wfc_User")
wfc_Pass = os.getenv("wfc_Pass")
FM_User = os.getenv("FM_User")
FM_Pass =  os.getenv("FM_Pass")


#────────────────────────────────────

def listo():
    busqueda("Listo...",tiempoUsual*0.2,"Si")

def busqueda(busqueda, tiempo=tiempoUsual*0.2, listo="No"): #busqueda->str
    pyautogui.hotkey("ctrl","f")
    time.sleep(tiempo)
    pyautogui.typewrite(str(busqueda))
    if listo == "No":
        pyautogui.press("enter")
        time.sleep(tiempo)
        pyautogui.press("esc")

def ToasNotif(function):
    os.system("curl --silent -H 'Title: Macro Trabajo' -H 'Priority: High' -H 'Tags: heavy_check_mark,warning' -d 'Ya terminó el procedimiento del programa -> '"+function+" ntfy.sh/laconchadetumadrebobesponjaptm >> /dev/null")

# deberia de haber hecho un diccionario donde el valor sea
# una tupla o lista, el primer valor la forma de nombrar y el segundo la cantidad 
def materialesCantidades(material): 
    #---------------------
    #------ Equipos ------
    #---------------------
    if material == "decoAndroid":
        return "110110006"
    
    elif material == "decoFullIP":
        return "110100344"
    
    elif material == "ONTsagemcom":
        return "140110075"
    
    elif material == "term.volte":
        return "140111192"
    
    #---------------------
    #---- Consumibles ----
    #---------------------
    #     ↓      ↓
    #---[SAP, cantidad]---

    #only Frappe
    elif material == "70m":
        return ["drop", "1"]

    elif material == "80m":
        return ["80mt", "1"]

    elif material == "100mtRef":
        return ["20310032", "1"]

    elif material == "200m":
        return ["20310014", "2"]

    elif material == "bob500m":
        return ["20310004", "500"]
    
    elif material == "ctrlAndroid":
        return ["110100372", "5"]
    
    elif material == "pilAndroid":
        return ["320500065", "10"]
    
    elif material == "kitFullIP":
        return ["110100320", "1"]
    #---------------------
    elif material == "ret.p":
        return ["20710010", "3"]
    
    elif material == "conect":
        return ["20710003", "2"]
    
    elif material == "volte":
        return ["140111299", "3"]

    #elif material == "":
    #    return ["", "0"]


def openBrowser():
    os.system("xdotool key super+1")
    time.sleep(tiempoUsual)
    # pyautogui.keyDown('win')
    # pyautogui.press('1')
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.keyUp('win')


def corroboraSeriesFM(planilla="Nota"): 
    #revisar ot series recuperados, orden:
    #1. all-in-one
    #2. rsmt forms

    if planilla=="Series":
        cantSeries=input("Cuantos 'Series hay que verificar? -> ")
        if cantSeries == "":
            cantSeries=0
        else:
            cantSeries=int(cantSeries)

        openBrowser()
 
        if cantSeries > 0:

            # hago un for parecido al de los series normales
            iteracion = 1
            for i in range(cantSeries):
                pyautogui.hotkey("ctrl","home")
                time.sleep(tiempoUsual*0.4
                           )
                for x in range(i):
                    pyautogui.press("down")
                    time.sleep(tiempoUsual*0.4)

                # copiamos el serie
                time.sleep(tiempoUsual)                
                pyautogui.hotkey("ctrl", "c")
                time.sleep(tiempoUsual*0.5)

                pyautogui.hotkey("ctrl","pgdn")
                time.sleep(tiempoUsual)

                pyautogui.hotkey("ctrl", "h")
                time.sleep(tiempoUsual)
                pyautogui.hotkey("ctrl", "v")
                time.sleep(tiempoUsual*0.5)
                pyautogui.press("enter")
                time.sleep(tiempoPagina)

                pyautogui.press("esc")
                time.sleep(tiempoUsual*0.5)
                pyautogui.press("home")
                for i in range(2):
                    pyautogui.press("right")
                    time.sleep(tiempoUsual*0.3)
                
                pyautogui.hotkey("ctrl", "c")
                time.sleep(tiempoUsual*0.5)

                pyautogui.hotkey("ctrl","pgup")
                time.sleep(tiempoUsual)

                
                pyautogui.hotkey("ctrl","home")
                time.sleep(tiempoUsual*0.4)
                for x in range(i):
                    pyautogui.press("down")
                    time.sleep(tiempoUsual*0.4)
    
                for i in range(2):
                    pyautogui.hotkey("ctrl", "right")
                    time.sleep(tiempoUsual*0.5)
                pyautogui.press("right")
                time.sleep(tiempoUsual*0.5)
                pyautogui.hotkey("ctrl", "v")
                time.sleep(tiempoUsual*0.5)
                



    else:
        #Sheets (orden - Nota de entrega)
        cantInstalaciones=int(input("Cuantas 'Instalaciones hay que verificar? -> "))
        cantSeries=input("Cuantos 'Series hay que verificar? -> ")
        if cantSeries == "":
            cantSeries=0
        else:
            cantSeries=int(cantSeries)

        openBrowser()
        
        # macro de borrar los consumos
        pyautogui.hotkey("ctrl","shift","alt","8") 
        time.sleep(tiempoUsual)

        iteracion = 1
        for i in range(cantInstalaciones):
            # copiamos el serie
            time.sleep(tiempoUsual)
            pyautogui.hotkey("ctrl","home")
            time.sleep(tiempoUsual*0.4)
            busqueda("Corroborar")
            time.sleep(tiempoUsual*0.4)
            
            for i in range(iteracion):
                pyautogui.press("down")
                time.sleep(tiempoUsual*0.1)
            pyautogui.hotkey("ctrl","c")
            time.sleep(tiempoUsual*0.1)
            #---
            pyautogui.hotkey("ctrl","pgdn")
            time.sleep(tiempoUsual*0.5)
            pyautogui.hotkey("ctrl","pgdn")
            time.sleep(tiempoUsual*0.5)
            #---

            # fieldmanager
            # buscar OT
            busqueda("materiales des", 1)
            time.sleep(tiempoUsual)
            
            for i in range(4):
                pyautogui.press("tab")
                time.sleep(tiempoUsual*0.2)
                pyautogui.press("delete")
            
            time.sleep(tiempoUsual*0.5)
            pyautogui.hotkey("ctrl","v")
            time.sleep(tiempoUsual*0.2)

            time.sleep(tiempoUsual)
            pyautogui.moveTo(btposX+ 103, 326 +btposY)
            time.sleep(tiempoUsual*1.2)
            pyautogui.click()
            time.sleep(tiempoUsual*0.2)
            pyautogui.click()
            pyautogui.hotkey("ctrl","c")
            time.sleep(tiempoUsual)

            busqueda("materiales des", 1)
            pyautogui.press("tab")
            time.sleep(tiempoUsual*0.5)
            pyautogui.hotkey("ctrl","v")

            for i in range(3):
                pyautogui.press("tab")
                time.sleep(tiempoUsual*0.15)
                pyautogui.press("delete")

            # viendo OT con consumos
            time.sleep(tiempoUsual)
            busqueda("20710003",tiempoUsual)
            pyautogui.hotkey("ctrl","c")
            time.sleep(tiempoUsual)
            # si tiene decos
            busqueda("320500065",tiempoUsual)
            pyautogui.hotkey("ctrl","c")
            time.sleep(tiempoUsual)
            # Por las dudas si no se encuentra OT
            busqueda("Mostrando 1 - 100",tiempoUsual)
            pyautogui.hotkey("ctrl","c")
            time.sleep(tiempoUsual)

            pyautogui.hotkey("ctrl","pgup")
            time.sleep(tiempoUsual*0.4)
            pyautogui.hotkey("ctrl","pgup")
            time.sleep(tiempoUsual*0.4)

            pyautogui.hotkey("ctrl","home")
            time.sleep(tiempoUsual*0.4)
            busqueda("Consumos:")
            for i in range(iteracion):
                pyautogui.press("down")
                time.sleep(tiempoUsual*0.1)
            time.sleep(tiempoUsual*0.5)
            pyautogui.hotkey("ctrl","v")

            #end
            iteracion = iteracion + 1


        if cantSeries > 0:
        
            # introducir revision de todos los series
            pyautogui.hotkey("ctrl","home")
            time.sleep(tiempoUsual*0.4)
            busqueda("Consumos:")
            time.sleep(tiempoUsual*0.4)
            
            for i in range(cantInstalaciones+2):
                pyautogui.press("down")
                time.sleep(tiempoUsual*0.1)
            pyautogui.typewrite("PostRevision") # Super importante


            # hago un for parecido al de las instalaciones
            iteracion = 1
            for i in range(cantSeries):
                # copiamos el serie
                time.sleep(tiempoUsual)
                pyautogui.hotkey("ctrl","home")
                time.sleep(tiempoUsual*0.4)
                busqueda("Revisar")
                time.sleep(tiempoUsual*0.4)
                
                for i in range(iteracion):
                    pyautogui.press("down")
                    time.sleep(tiempoUsual*0.1)
                pyautogui.hotkey("ctrl","c")
                time.sleep(tiempoUsual*0.1)
                #---
                pyautogui.hotkey("ctrl","pgdn")
                time.sleep(tiempoUsual*0.5)
                pyautogui.hotkey("ctrl","pgdn")
                time.sleep(tiempoUsual*0.5)
                #---

                # fieldmanager
                # buscar OT
                busqueda("materiales des", 1)
                time.sleep(tiempoUsual)
                
                for i in range(4):
                    pyautogui.press("tab")
                    time.sleep(tiempoUsual*0.2)
                    pyautogui.press("delete")
                
                time.sleep(tiempoUsual*0.5)
                pyautogui.hotkey("ctrl","v")
                time.sleep(tiempoUsual*0.2)

                time.sleep(tiempoUsual)
                pyautogui.moveTo(btposX+ 103, 326 +btposY)
                pyautogui.click()
                time.sleep(tiempoUsual*0.2)
                pyautogui.click()
                pyautogui.hotkey("ctrl","c")
                time.sleep(tiempoUsual)

                pyautogui.hotkey("ctrl","pgup")
                time.sleep(tiempoUsual*0.4)
                pyautogui.hotkey("ctrl","pgup")
                time.sleep(tiempoUsual*0.4)

                pyautogui.hotkey("ctrl","home")
                time.sleep(tiempoUsual*0.4)
                busqueda("PostRevision")
                for i in range(iteracion):
                    pyautogui.press("down")
                    time.sleep(tiempoUsual*0.1)
                time.sleep(tiempoUsual*0.5)
                pyautogui.hotkey("ctrl","v")

                #end
                iteracion = iteracion + 1

        time.sleep(tiempoPagina)
        pyautogui.hotkey("ctrl","home")
        time.sleep(tiempoUsual*0.5)

    # insertar consumos en AppSheet
    # busqueda("cargado")
    # time.sleep(tiempoUsual)
    # pyautogui.press("down")
    # time.sleep(tiempoUsual)
    # pyautogui.hotkey("ctrl","c")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","home")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","pgdn")
    # time.sleep(tiempoUsual)
    # busqueda("Fecha")
    # [pyautogui.press("tab") for _ in range(14)]
    # time.sleep(tiempoUsual)
    # pyautogui.hotkey("ctrl","a")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","v")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","pgup")
    #
    # #Cargar instalaciones sin consumos
    # time.sleep(tiempoPagina)
    # busqueda("CopyPaste")
    # time.sleep(tiempoUsual)
    # pyautogui.press("down")
    # time.sleep(tiempoUsual)
    # pyautogui.hotkey("ctrl","shift","down")
    # time.sleep(tiempoUsual)
    # pyautogui.hotkey("ctrl","c")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","home")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","pgdn")
    # time.sleep(tiempoUsual)
    # busqueda("Fecha")
    # [pyautogui.press("tab") for _ in range(7)]
    # pyautogui.hotkey("ctrl","a")
    # time.sleep(tiempoUsual)
    # pyautogui.hotkey("ctrl","v")
    # time.sleep(tiempoUsual*0.5)
    # pyautogui.hotkey("ctrl","pgup")

    # listo()



def disclaimer():
    os.system("clear")
    print("__________Disclaimer___________")
    print("Estas macros pueden tener")
    print("errores y ocasionar acciones")
    print("no esperadas, utilizar esta")
    print("herramienta con sumo cuidado")
    print("")
    print("Estas avisado...")
    print("_______________________________")
    print("")
    print("")
    input("Continuar...  ")
#opcion = " "


def bulkOpen():
    #  En espacio nota de entrega
    # teniendo foco en la planilla "Orden" de la tabla
    # Nota de entrega y abajo abierto un formulario nuevo
    # para carga de instalaciones en AppSheet
    time.sleep(tiempoUsual*0.5)

    pyautogui.hotkey("win", "1")

    time.sleep(tiempoUsual)

    # URL Fotos para bulk open
    pyautogui.hotkey("ctrl", "shift", "alt", "8")
    time.sleep(tiempoPagina)
    pyautogui.hotkey("ctrl","c")
    time.sleep(tiempoUsual)

    pyautogui.hotkey("ctrl","t")
    pyautogui.hotkey("ctrl","shift","9")
    time.sleep(tiempoPagina*0.7)
    pyautogui.moveTo(1077, 201)
    time.sleep(tiempoUsual)
    pyautogui.click()
    # [pyautogui.press("tab") for _ in range(4)]
    pyautogui.hotkey("ctrl","a")
    pyautogui.hotkey("ctrl","v")
    pyautogui.hotkey("ctrl","end")
    pyautogui.press("left")
    pyautogui.press("delete")
    time.sleep(tiempoUsual*0.4)
    pyautogui.hotkey("ctrl","home")
    pyautogui.press("delete")
    pyautogui.press("tab")
    pyautogui.press("enter")
    pyautogui.press("esc")


def PRINTmateriales():
    #abrir browser
    pyautogui.keyDown('win')
    pyautogui.press('1')
    time.sleep(tiempoUsual*0.5)
    pyautogui.keyUp('win')

    pyautogui.hotkey('ctrl', 'end')
    pyautogui.hotkey('ctrl', 'home')
    pyautogui.hotkey('ctrl', 'shift', 'down')
    pyautogui.hotkey('shift', 'up')
    pyautogui.hotkey('shift', 'up')
    pyautogui.hotkey('shift', 'up')
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(tiempoPagina)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    #next
    for i in range(13):
        pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(tiempoUsual*2)
    #imprimir
    pyautogui.press('enter')
    time.sleep(tiempoPagina)
    #editConfirmation
    pyautogui.press('enter')
    time.sleep(tiempoUsual*2)
    #copiar movimiento
    pyautogui.hotkey('ctrl','end')
    pyautogui.hotkey('ctrl','home')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    #seleccionar 44 celdas
    # for i in range(44):
    pyautogui.hotkey('ctrl','shift','down')
    for i in range(4):
        pyautogui.hotkey('shift', 'up')
    for i in range(2):
        pyautogui.hotkey('shift','left')
    for i in range(3):
        pyautogui.hotkey('shift', 'down')

    pyautogui.hotkey('ctrl','c')
    time.sleep(tiempoUsual)
    pyautogui.press('esc')

    pyautogui.hotkey('ctrl','end')
    pyautogui.hotkey('ctrl','home')


def materiales(material, cantidad, custom):
    if custom == "S" or custom == "s":
        time.sleep(tiempoUsual)
        pyautogui.moveTo(btPlus[0], btPlus[1])
        pyautogui.click()
        time.sleep(tiempoUsual)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(material, interval=delayDeEscritura+0.4)
        time.sleep(tiempoUsual)
        time.sleep(tiempoUsual*0.5)

        pyautogui.moveTo(btposX+710, 410+btposY)
        pyautogui.click()

        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.typewrite(cantidad, interval=delayDeEscritura+0.4)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('enter')

    else:
        # cantidad=0
        #--------------
        #   CABLES
        #--------------
        
        #70m
        if material=="70m":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("70m")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(materialesCantidades("70m")[1], interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #80m
        if material=="80m":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("80m")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(materialesCantidades("80m")[1], interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #100mtRef
        if material=="100mtRef":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("100mtRef")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(materialesCantidades("100mtRef")[1], interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #conect.op
        if material == "conect.op" or material == "20710003":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')

            pyautogui.typewrite("conect.o", interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            pyautogui.typewrite("p")
            time.sleep(tiempoUsual*0.5)
            
            pyautogui.moveTo(btposX+725, 455+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(materialesCantidades("conect")[1], interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #ret.p    
        if material == "ret.p":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("ret.p")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(materialesCantidades("ret.p")[1], interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #sim volte
        if material == "volte":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("volte")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(materialesCantidades("volte")[1], interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #control android
        if material == "ctrlAndroid_Y_pilas":
            cantidadDecos = input ("cuantos decos son?\n -> ")

            pyautogui.keyDown('win')
            pyautogui.press('1')
            time.sleep(tiempoUsual*0.5)
            pyautogui.keyUp('win')
            
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')

            pyautogui.typewrite(materialesCantidades("ctrlAndroid")[0], interval=delayDeEscritura+0.4)
            time.sleep(tiempoUsual)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(str(cantidadDecos), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

            #pilas
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')

            pyautogui.typewrite(materialesCantidades("pilAndroid")[0], interval=delayDeEscritura+0.15)#320500065
            time.sleep(tiempoUsual)
            
            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(str(int(cantidadDecos)*2), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')


        if material == "ctrlAndroid":
            # cantidad = input ("cuantos decos son?\n -> ")

            pyautogui.keyDown('win')
            pyautogui.press('1')
            time.sleep(tiempoUsual*0.5)
            pyautogui.keyUp('win')
            
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')

            pyautogui.typewrite(materialesCantidades("ctrlAndroid")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(str(materialesCantidades("ctrlAndroid")[1]), interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')


        if material == "pilAndroid":
            #pilas
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')

            pyautogui.typewrite(materialesCantidades("pilAndroid")[0], interval=delayDeEscritura+0.2)#320500065
            time.sleep(tiempoUsual)
            
            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            if cantidad == 0:
                pyautogui.typewrite(str(int(materialesCantidades("ctrlAndroid")[1])*2), interval=delayDeEscritura+0.15)
            else:
                pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #kit
        if material == "kit":
            cantidad = input ("cuantos decos son?\n -> ")

            pyautogui.keyDown('win')
            pyautogui.press('1')
            time.sleep(tiempoUsual*0.5)
            pyautogui.keyUp('win')
            
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')

            pyautogui.typewrite(materialesCantidades("kitFullIP")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(str(cantidad), interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')

        #200m
        if material == "200m":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("200m")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("200m")[1], interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')


        #bob500m
        if material == "bob500m":
            time.sleep(tiempoUsual)
            pyautogui.moveTo(btPlus[0], btPlus[1])
            pyautogui.click()
            time.sleep(tiempoUsual)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("bob500m")[0], interval=delayDeEscritura+0.15)
            time.sleep(tiempoUsual)
            time.sleep(tiempoUsual*0.5)

            pyautogui.moveTo(btposX+710, 410+btposY)
            pyautogui.click()

            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.typewrite(materialesCantidades("bob500m")[1], interval=delayDeEscritura+0.15)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')




#MENU

def descvoLTE(mate):
    #abrir browser
    pyautogui.keyDown('win')
    pyautogui.press('1')
    time.sleep(tiempoUsual*0.5)
    pyautogui.keyUp('win')
    #for veces in range(1,5):
    #copiar solicitud
    pyautogui.hotkey('ctrl', 'c')
    #pegar solicitud
    pyautogui.hotkey('ctrl', 'pgdn')
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.typewrite("NRO. OT")
    pyautogui.press('esc')
    pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    #abrir solicitud
    pyautogui.moveTo(btposX+120, 381+btposY)
    pyautogui.click()
    pyautogui.moveTo(btposX+162, 659+btposY)
    pyautogui.click()
    time.sleep(2)
    #copiar serie
    pyautogui.moveTo(btposX+300, 507+btposY)
    pyautogui.click()
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(tiempoUsual)
    pyautogui.press('esc')
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'shift', 'alt', '2')
    time.sleep(2)
    pyautogui.press('up')
    pyautogui.press('left')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.press('down')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('ctrl', 'pgdn')
    time.sleep(tiempoUsual)

    if mate=="si":

        materiales("volte", 0, "N")
    
    else:
        pyautogui.moveTo(btposX+1097, 264+btposY)
        pyautogui.click()
    
    pyautogui.moveTo(btEnviar[0], btEnviar[1])
    time.sleep(15)
    
    pyautogui.hotkey('ctrl', 'pgup')
    pyautogui.press('down')


def loginFM():
    #abrir browser
    openBrowser()
    time.sleep(tiempoUsual)

    #abrir nueva pestaña
    pyautogui.hotkey('ctrl', 't')
    time.sleep(tiempoUsual*0.5)

    pyautogui.typewrite("fieldmanager.telecom.com.ar")
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('enter')

    time.sleep(tiempoPagina*1.3)

    #Login
    pyautogui.press('tab')
    pyautogui.typewrite(FM_User)
    pyautogui.press('tab')
    pyautogui.typewrite(FM_Pass)
    pyautogui.press('tab')
    pyautogui.press('enter')

    time.sleep(tiempoPagina)

    #por las dudas
    pyautogui.hotkey('ctrl', 'r')
    time.sleep(tiempoPagina)

def loginFrappe():
    #abrir browser
    openBrowser()
    time.sleep(tiempoUsual)

    #abrir nueva pestaña
    pyautogui.hotkey('ctrl', 't')
    time.sleep(tiempoUsual*0.5)

    #qps
    pyautogui.typewrite("qpsrl.zapto.org/desk") # 192.168.0.100") #qpsrl.zapto")
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('enter')
    time.sleep(tiempoPagina)
    time.sleep(tiempoPagina)

    #Login
    pyautogui.moveTo(btposX+571, 282+btposY)
    pyautogui.click()
    time.sleep(tiempoUsual*0.5)
    pyautogui.moveTo(btposX+665, 384+btposY)
    pyautogui.click()
    time.sleep(2)
    pyautogui.click()
    time.sleep(2)
    pyautogui.click()
    
    time.sleep(tiempoPagina)
    time.sleep(tiempoPagina)

def loadMaterialesFrappe(): 
    #abrir browser
    # openBrowser()

    # Desde Orden
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'home')
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(tiempoUsual*0.5)

    pyautogui.hotkey('ctrl', 't')
    pyautogui.hotkey('ctrl', 'v')
    # pyautogui.typewrite('https://docs.google.com/spreadsheets/d/e/2PACX-1vQFB5r6WpYQyHRqeNKBA_Z5N1iD5by823_xarzcQXNCwtOK1KOIfstvdvfe52geaC6CVKtw0a1pvqrT/pub?gid=1213551365&single=true&output=csv')
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('enter')
    time.sleep(tiempoPagina*1.2)
    pyautogui.hotkey('ctrl', 'w')
    
    #nueva nota de entrega
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'b')
    time.sleep(tiempoPagina)

    #cliente
    pyautogui.hotkey('ctrl','home')
    busqueda('entregar a')
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('esc')
    pyautogui.press('tab')
    time.sleep(tiempoUsual)
    pyautogui.press('tab')
    time.sleep(2.5)
    pyautogui.press('enter')
    time.sleep(tiempoUsual)

    # Fecha
    pyautogui.hotkey('ctrl', 'pgdn')
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey('ctrl', 'end')
    pyautogui.hotkey('ctrl', 'home')
    for i in range(4):
        pyautogui.press('left')
        time.sleep(tiempoUsual*0.2)
    pyautogui.press('down')
    time.sleep(tiempoUsual*0.2)
    pyautogui.press('down')
    pyautogui.hotkey('ctrl', 'c')
    #pegarFecha
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(tiempoUsual*0.5)
    busqueda('entregar a')
    time.sleep(tiempoUsual*0.5)
    for i in range(3):
        pyautogui.press('tab')
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(tiempoUsual)
    pyautogui.press('enter')
    time.sleep(tiempoUsual)

    #Tecnico
    busqueda('entregar a')
    time.sleep(tiempoUsual*0.5)
    for i in range(6):
        pyautogui.press('tab')
    #copiarTecnico
    pyautogui.hotkey('ctrl', 'pgdn')
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey('ctrl', 'end')
    pyautogui.hotkey('ctrl', 'home')
    for i in range(4):
        pyautogui.press('left')
        time.sleep(tiempoUsual*0.2)
    pyautogui.press('down')
    pyautogui.hotkey('ctrl', 'c')
    #pegarTecnico
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(tiempoUsual)
    pyautogui.press('enter')

    # pyautogui.press('tab')
    # pyautogui.press('tab')

    #subida csv
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(tiempoUsual)
    pyautogui.typewrite('subir')
    pyautogui.press('esc')
    pyautogui.press('enter')
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('enter')
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'shift', 'y')
    time.sleep(tiempoPagina)

    #yo te espero, carga tranqui
    pyautogui.moveTo(1086, 155)
    pyautogui.mouseDown()
    time.sleep(tiempoUsual*0.5)
    pyautogui.moveTo(452, 222, duration=1.5)
    time.sleep(tiempoPagina*0.5)
    openBrowser()
    time.sleep(tiempoUsual*0.5)
    pyautogui.mouseUp()
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'shift', 'y')
    time.sleep(tiempoUsual)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(tiempoUsual)
    
    #cerrar cuadro de subida
    busqueda('subir')
    for i in range(2):
        pyautogui.press("tab")
        time.sleep(tiempoUsual*0.5)
    pyautogui.press("enter")
    time.sleep(tiempoUsual)
    
    for i in range(2):
        busqueda('cerrar')
        pyautogui.press("tab")
        time.sleep(tiempoUsual*0.5)
        pyautogui.press("enter")
        time.sleep(tiempoUsual)

    time.sleep(tiempoUsual*0.4)


    #Save/guardar
    pyautogui.moveTo(1137, 141)
    pyautogui.click()

    #copiar movimientos
    pyautogui.hotkey('ctrl', 'pgdn')
    time.sleep(tiempoUsual)
    pyautogui.hotkey('ctrl', 'end')
    pyautogui.hotkey('ctrl', 'home')
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('down')
    pyautogui.keyDown('shift')
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('down')
    time.sleep(tiempoUsual*1.2)
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('down')
    pyautogui.hotkey('ctrl', 'up')
    

    for i in range(3):
        pyautogui.hotkey('ctrl', 'right')
        time.sleep(tiempoUsual*0.35)
    pyautogui.keyUp('shift')

    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'pgup')
    pyautogui.hotkey('ctrl', 's')
   
    time.sleep(tiempoUsual)

    busqueda("equipo de ventas")
    pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(tiempoUsual)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.hotkey("ctrl", "enter")
    time.sleep(tiempoUsual)
    busqueda("Detalles de la OC del Cliente")

    #Save/guardar
    pyautogui.moveTo(btposX+1205, 154+btposY)
    listo()





def serStoMov():
    # loginFrappe()

    #Serial
    pyautogui.moveTo(btposX+407, 216+btposY)
    pyautogui.click()
    time.sleep(tiempoPagina)
    pyautogui.moveTo(btposX+1071, 162+btposY)
    pyautogui.click()
    time.sleep(tiempoUsual*0.5)
    pyautogui.moveTo(btposX+1072, 428+btposY)
    pyautogui.click()

    time.sleep(2)
    pyautogui.moveTo(btposX+450, 264+btposY)
    pyautogui.click()
    pyautogui.moveTo(btposX+904, 143+btposY)
    pyautogui.click()
    
    #inicio
    time.sleep(tiempoUsual)
    pyautogui.moveTo(btposX+118, 106+btposY)
    pyautogui.click()
    time.sleep(tiempoPagina)


    #Stock
    pyautogui.moveTo(btposX+550, 213+btposY)
    pyautogui.click()
    time.sleep(tiempoPagina)
    pyautogui.moveTo(btposX+1141, 162+btposY)
    pyautogui.click()
    time.sleep(tiempoUsual*0.5)
    pyautogui.moveTo(btposX+1125, 343+btposY)
    pyautogui.click()

    time.sleep(2)
    pyautogui.moveTo(btposX+428, 264+btposY)
    pyautogui.click()
    pyautogui.moveTo(btposX+919, 145+btposY)
    pyautogui.click()
    
    #inicio
    time.sleep(tiempoUsual)
    pyautogui.moveTo(btposX+118, 106+btposY)
    pyautogui.click()
    time.sleep(tiempoPagina)

    #Movimientos
    pyautogui.moveTo(btposX+678, 213+btposY)
    pyautogui.click()
    time.sleep(tiempoPagina)
    pyautogui.moveTo(btposX+1138, 170+btposY)
    pyautogui.click()
    time.sleep(tiempoUsual*0.5)
    pyautogui.moveTo(btposX+1119, 339+btposY)
    pyautogui.click()

    time.sleep(2)
    pyautogui.moveTo(btposX+471, 259+btposY)
    pyautogui.click()
    pyautogui.moveTo(btposX+910, 148+btposY)
    pyautogui.click()
    
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'w')





def CompletarAlmacen_tieneFirefox():
    #Region
    busqueda("filtros")
    for i in range(2):
        pyautogui.press('tab')
        time.sleep(tiempoUsual*0.5)
    pyautogui.press('down')
    pyautogui.press("enter")
    time.sleep(tiempoUsual)

    #SubRegion
    for i in range(2):
        pyautogui.press('tab')
        time.sleep(tiempoUsual*0.5)
    pyautogui.press('down')
    pyautogui.press("enter")
    time.sleep(tiempoUsual)

    #centro
    for i in range(2):
        pyautogui.press('tab')
        time.sleep(tiempoUsual*0.5)
    pyautogui.press('down')
    pyautogui.press("enter")
    time.sleep(tiempoUsual)

    #Almacen
    for i in range(2):
        pyautogui.press('tab')
        time.sleep(tiempoUsual*0.5)
    for i in range(2):
        pyautogui.press('down')
        pyautogui.press("enter")
    time.sleep(tiempoUsual)


def CompletarAlmacen_tieneEdge():
    #Region
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(tiempoUsual)
    pyautogui.typewrite("non")
    time.sleep(tiempoUsual)
    pyautogui.press('esc')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('enter')
    pyautogui.press('tab')

    #Sub Region
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(tiempoUsual)
    pyautogui.typewrite("non")
    time.sleep(tiempoUsual)
    pyautogui.press('esc')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('enter')
    pyautogui.press('tab')

    #Centro Logístico
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(tiempoUsual)
    pyautogui.typewrite("non")
    time.sleep(tiempoUsual)
    pyautogui.press('esc')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('enter')
    pyautogui.press('tab')

    #Almacen
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(tiempoUsual)
    pyautogui.typewrite("non")
    time.sleep(tiempoUsual)
    pyautogui.press('esc')
    time.sleep(tiempoUsual)
    pyautogui.press('enter')
    for i in range(2):
        pyautogui.press('down')
        time.sleep(tiempoUsual*0.5)
        pyautogui.press('enter')


def materialesDescargados():
    loginFM()

    #Inicio
    # pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')

    time.sleep(tiempoPagina)
    
    # iniciar materiales descargados


    #Pone que navegador tiene
    CompletarAlmacen_tieneFirefox()
    for i in range(5):
        pyautogui.press('tab')

    # dateActual=str(str(date()[0])+"/"+str(date()[1])+"/"+str(date()[2]))
    # dateOld=str(str(date()[0])+"/"+str(date()[1]-2)+"/"+str(date()[2]))
    # #set fecha desde
    # pyautogui.typewrite(str(dateOld))
    # #fecha hasta
    # pyautogui.hotkey('ctrl', 'f')
    # pyautogui.typewrite('Fecha Ult Mod Hasta')
    # pyautogui.press('esc')
    # pyautogui.press('tab') 
    # pyautogui.typewrite(str(dateActual))
    # pyautogui.press('tab')
    #enviar
    pyautogui.press('enter')

def filtroGestionMat():
    tecnico=input("A cual tecnico queres filtrar?[none]")
    fecha=input("Que fecha queres filtrar?[none]")
    cerrada=input("Queres filtrar inst cerradas?[No]")
    
    tecnico=tecnico + " rsmt"

    openBrowser()

    busqueda("nro.")
    for i in range(2):
        pyautogui.press('tab')
        pyautogui.press('delete')
    if cerrada == "":
        pyautogui.typewrite('')
    else:
        pyautogui.typewrite('cer')

    pyautogui.press('tab')
    pyautogui.press('delete')
    pyautogui.typewrite(fecha)

    for i in range(2):
        pyautogui.press('tab')
        pyautogui.press('delete')
    pyautogui.typewrite(tecnico)
    # time.sleep(tiempoUsual)
    # pyautogui.moveTo(btposX+353, 230+btposY)
    # pyautogui.click()
    # time.sleep(tiempoUsual)
    # pyautogui.moveTo(btposX+353, 356+btposY)
    # pyautogui.click()
    # time.sleep(tiempoUsual)
    # pyautogui.moveTo(122,651)
    # pyautogui.click()
    


def macroDeGestion():
    loginFM()

    #Inicio
    # pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('enter')

    time.sleep(tiempoPagina)

   # usamos directamente la funcion y nos olvidamos 
    CompletarAlmacen_tieneFirefox()

    #enviar
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')


def comprobacionDeSerial():
    #abrir browser
    openBrowser()
    time.sleep(tiempoUsual)

    #seleccion del serial 1
    pyautogui.moveTo(btposX+269, 500+btposY)
    pyautogui.click()
    pyautogui.click()

    #copiar y buscar en sheets
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')

    #tiempo para verificacion
    time.sleep(3)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'alt', 'shift', '2')
    pyautogui.hotkey('ctrl', 'pgdn')
    
    #siguiente serial (2)
    pyautogui.press('down')
    pyautogui.press('down')

    #seleccionar
    pyautogui.press('home')
    pyautogui.hotkey('shift', 'end')

    #copiar y buscar en sheets
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')

    #tiempo para verificacion
    time.sleep(3)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'alt', 'shift', '2')
    pyautogui.hotkey('ctrl', 'pgdn')

    #siguiente serial (3)
    pyautogui.press('down')
    pyautogui.press('down')

    #seleccionar
    pyautogui.press('home')
    pyautogui.hotkey('shift', 'end')

    #copiar y buscar en sheets
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')

    #tiempo para verificacion
    time.sleep(3)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'alt', 'shift', '2')
    pyautogui.hotkey('ctrl', 'pgdn')

    #siguiente serial (4)
    pyautogui.press('down')
    pyautogui.press('down')

    #seleccionar
    pyautogui.press('home')
    pyautogui.hotkey('shift', 'end')

    #copiar y buscar en sheets
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'pgup')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')

    #tiempo para verificacion
    time.sleep(3)
    pyautogui.press('esc')
    pyautogui.hotkey('ctrl', 'alt', 'shift', '2')
    pyautogui.hotkey('ctrl', 'pgdn')

def activacionMaterialFrappe():
    cantidad=input("cuantas veces se itera?\n-> ")
    openBrowser()
    time.sleep(tiempoUsual)
    for i in range(int(cantidad)):
        pyautogui.press("1")
        time.sleep(tiempoUsual)
        pyautogui.press("left")
        pyautogui.press("del")
        time.sleep(tiempoUsual*1.2)
        pyautogui.press("enter")
        time.sleep(tiempoUsual)
        for i in range(8):
            pyautogui.press('tab')
            time.sleep(tiempoUsual*0.5)
        time.sleep(tiempoUsual)
        pyautogui.press("end")

def workforceRefresh():   
    time.sleep(tiempoUsual)
    pyautogui.hotkey("win","1")
    while True:
        time.sleep(tiempoPagina*1.5)
        pyautogui.hotkey("ctrl","c")
        for i in range(3):
            pyautogui.press('tab')
        pyautogui.press('enter')




def workforceSetup():
    #abrir browser
    openBrowser()
    time.sleep(tiempoUsual)

    #Ir a la pagina
    pyautogui.hotkey('ctrl', 't')
    time.sleep(tiempoUsual*0.5)

    pyautogui.typewrite("technet.teco.csgfsm.com")
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('down')
    time.sleep(tiempoUsual*0.5)
    pyautogui.press('enter')
    
    #Login
    time.sleep(tiempoPagina)
    pyautogui.typewrite(wfc_User)
    pyautogui.press('tab')
    pyautogui.typewrite(wfc_Pass)
    pyautogui.press('enter')
    
    #cronologia del tiempo
    time.sleep(tiempoPagina)
    pyautogui.press('tab')
    pyautogui.press('enter')
    for i in range(5):
        pyautogui.press('down')
        time.sleep(tiempoUsual*0.2)
    pyautogui.press('enter')
    
    #Buscar tecnicos
    time.sleep(tiempoPagina)
    for i in range(2):
        pyautogui.press('tab')
        time.sleep(tiempoUsual*0.2)
    pyautogui.typewrite("RSMT")
    pyautogui.press('enter')
    time.sleep(tiempoUsual*0.5)
    for i in range(2):
        pyautogui.press('tab')
        time.sleep(tiempoUsual*0.2)
    pyautogui.press('enter')
    pyautogui.press('tab')
    pyautogui.press('enter')
    
    #Mapa
    busqueda("cablevision")
    time.sleep(tiempoPagina)
    pyautogui.press('tab')
    pyautogui.press('enter')
    for i in range(4):
        pyautogui.press('down')
        time.sleep(tiempoUsual*0.2)
    pyautogui.press('enter')
    time.sleep(tiempoUsual*1.2)
    pyautogui.hotkey("ctrl", "f")
    time.sleep(tiempoUsual*0.5)
    pyautogui.typewrite("RSMT")


#----------------------------------------
#--------M-E-N-U-S-----------------------
#----------------------------------------

def menu1():

    opcion = " "
    while(opcion != "0"):
        os.system("clear")
        print("-------------------------------")
        print("             Menu              ")
        print("-------------------------------")
        print("01- Field Manager")
        print("02- Frappe")
        print("09- Otros...")
        print("00- chauchau")
        print()
        print("Que necesitas hacer?")
        opcion=input(" ")

        if opcion == "01" or opcion == "1":
            menuFieldManager()

        elif opcion == "02" or opcion == "2":
            menuFrappe()
            
        elif opcion == "09" or opcion == "9":
            otroMenu()

        elif opcion == "0":
            os.system("clear")
            print("morite puto")
            sys.exit()


def menuFrappe():
    opcion = " "
    while(opcion != "0"):
        os.system("clear")
        print("-------------------------------")
        print("            Frappe             ")
        print("-------------------------------")
        print("01- Bulk Open")
        print("02- FM login")
        print("03- Comprobacion de seriales")
        print("04- Cargar materiales")
        print("05- Comprueba y Carga materiales")
        print("09- Volver")
        print("00- chauchau")
        print()
        print("Que necesitas hacer?")
        opcion=input(" ")

        if opcion == "01" or opcion == "1":
            bulkOpen()

        elif opcion == "02" or opcion == "2":
            materialesDescargados()

        elif opcion == "03" or opcion == "3":
            corroboraSeriesFM()

        elif opcion == "04" or opcion == "4":
            openBrowser()
            loadMaterialesFrappe()

        elif opcion == "05" or opcion == "5":
            corroboraSeriesFM()
            time.sleep(tiempoPagina)
            loadMaterialesFrappe()

        elif opcion == "09" or opcion == "9":
            menu1()
        
        elif opcion == "0":
            os.system("clear")
            print("morite puto")
            sys.exit()

def menuFieldManager():

    opcion = " "
    while(opcion != "0"):
        os.system("clear")
        print("-------------------------------")
        print("   Sitios para revisar (FM)    ")
        print("-------------------------------")
        print("01- Gestion de materiales")
        print("02- Materiales descargados")
        print("03- Workforce (setup)")
        print("09- Volver")
        print("00- chauchau")
        print()
        print("Que necesitas hacer?")
        opcion=input(" ")
        
        if opcion == "01" or opcion == "1":
            menuFMGestion()

        elif opcion == "02" or opcion == "2":
            materialesDescargados()
            time.sleep(tiempoPagina)
            ToasNotif("materialesDescargados")

        elif opcion == "03" or opcion == "3":
            workforceSetup()

        elif opcion == "09" or opcion == "9":
            menu1()

        elif opcion == "0":
            os.system("clear")
            print("morite puto")
            sys.exit()


def menuFMGestion():

    def opcionalesMenu():
        opcion = " "
        while(opcion!="0"):
            os.system("clear")
            cantidad = ""
            print("-------------------------------")
            print("     Materiales Opcionales     ")
            print("-------------------------------")
            print()
            print("01 - Conectores (Max. 3)")
            print("02 - 40mt (020) (Max. 2)")
            print("022 - 40mt (002) (Max. 2)")
            print("03 - 40mt Ref. (004) (Max. 2)")
            print("04 - 70mt (Max. 2)")
            print("05 - 80mt (Max. 2)")
            print("~06 - 100mt (Max. 1)~")
            print("07 - 100mt Ref. (Max. 1)")
            print("08 - 150mt (Max. 1)")
            print("99 - ctrl 368 (Max. 10)")
            print("10 - ctrl 372 (Max. 10)")
            print("11 - ctrl 373 (Max. 10)")
            print("12 - pilas AAA (Max. 10)")
            print("13 - kit (320) (Max. 10)")
            print("14 - kit (009) (Max. 10)")
            print("15 - Cadenas (Max. 3)")
            print("16 - Patch UTP (Max. 3)")
            print()
            print("a-d Pagina/s")
            print("09- Volver")
            print("00- chauchau")
            print()
            print("Que necesitas hacer?")
            opcion=input(" ")
           
            if opcion == "9":
                menuFMGestion()
            if opcion == "0":
                os.system("clear")
                print("morite puto")
                sys.exit()

            # Paginas
            if opcion == "b":
                os.system("clear")
                print("-------------------------------")
                print("     Materiales Opcionales     ")
                print("-------------------------------")
                print(" pag. 2")
                print()
                print()
                print("a-d Paginas")
                print("09- Volver")
                print("00- chauchau")
                print()
                print("Que necesitas hacer?")
                opcion=input(" ")

                if opcion == "9":
                    menuFMGestion()
                if opcion == "0":
                    os.system("clear")
                    print("morite puto")
                    sys.exit()

                #Redireccion paginas
                if opcion == "a":
                    opcionalesMenu()
            
            # Here we go
            print("Que cantidad queres consumir?")
            cantidad = input(" ")
            print()
            
            openBrowser()
            if opcion == "1":
                materiales("conect.op", cantidad, "N")
            if opcion == "2":
                materiales("20310020", cantidad, "S")
            if opcion == "22":
                materiales("20300002", cantidad, "S")#40m nuevo
            if opcion == "3":
                materiales("20300004", cantidad, "S")
            if opcion == "4":
                materiales("20310031", cantidad, "S")
            if opcion == "5":
                materiales("20310019", cantidad, "S")
            # if opcion == "6":
                # materiales("20310013", cantidad, "S")
            if opcion == "7":
                materiales("20310032", cantidad, "S")
            if opcion == "8":
                materiales("20310012", cantidad, "S")
            if opcion == "99":
                materiales("110100368", cantidad, "S")
            if opcion == "10":
                materiales("110100372", cantidad, "S")
            if opcion == "11":
                materiales("110100373", cantidad, "S")
            if opcion == "12":
                materiales("320500065", cantidad, "S")
            if opcion == "13":
                materiales("110100320", cantidad, "S")
            if opcion == "14":
                materiales("110110009", cantidad, "S")
            if opcion == "15":
                materiales("211200138", cantidad, "S")
            if opcion == "16":
                materiales("140110033", cantidad, "S")

    opcion = " "
    while(opcion!="0"):
        os.system("clear")
        print("-------------------------------")
        print("     Gestion de materiales     ")
        print("-------------------------------")
        print("")
        print("01- 70m y conectores")
        print("02- 80m y conectores")
        print("03- 100m Reforzados")
        print("04- 40m (020) y conectores")
        print("05- 40m (002) y conectores")
        print("06- 40m Ref. (004) y conectores")
        print("07- controles y pilas")
        print("08- kit deco full ip")
        print("10- Otros Materiales")
        print("")
        print("-------------------------------")
        print("|    111 - Filtrar            |")
        print("|    666 - Para descargar     |")
        print("|    999 - Macro de gestion   |")
        print("-------------------------------")
        print("09- Volver")
        print("00- chauchau")
        print()
        print("Que necesitas hacer?")
        opcion=input(" ")
        
        if opcion == "01" or opcion == "1":
            openBrowser()
            materiales("70m", 0, "N")
            materiales("conect.op", 0, "N")
            
        elif opcion == "02" or opcion == "2":
            openBrowser()
            materiales("80m", 0, "N")
            materiales("conect.op", 0, "N")

        elif opcion == "03" or opcion == "3":
            openBrowser()
            materiales("100mtRef", 0, "N")
            materiales("conect.op", 0, "N")

        elif opcion == "04" or opcion == "4":
            openBrowser()
            materiales("20310020", "1", "S")
            materiales("conect.op", 0, "N")

        elif opcion == "05" or opcion == "5":
            openBrowser()
            materiales("20300002", "1", "S")#40m nuevo
            materiales("conect.op", 0, "N")

        elif opcion == "06" or opcion == "6":
            openBrowser()
            materiales("20300004", "1", "S")
            materiales("conect.op", 0, "N")

        elif opcion == "07" or opcion == "7":
            materiales("ctrlAndroid_Y_pilas", 0, "N")

        elif opcion == "08" or opcion == "8":
            materiales("kit", 0, "N")

        # elif opcion == "05" or opcion == "5":
        #     openBrowser()
        #     materiales("volte", 0, "N")

       
        elif opcion == "9" or opcion == "09":
            menuFieldManager()

        elif opcion == "10":
            opcionalesMenu()

        elif opcion == "0":
            os.system("clear")
            print("morite puto")
            sys.exit()

        elif opcion == "999":
            macroDeGestion()

        elif opcion == "666": #def descargaropcionales
            openBrowser()
            #--------
            materiales('20300002','1','S')
            materiales('20300004','1','S')
            materiales('20310020','1','S')
            materiales('20710003','3','N')
            materiales('140110033','3','S')
	
            #--------
            pyautogui.hotkey("ctrl","pgup")
            time.sleep(tiempoUsual)
            pyautogui.hotkey("ctrl","alt","shift","9")
            time.sleep(tiempoUsual)
            pyautogui.hotkey("ctrl","pgdn")
   

        #dejamos apoyado sobre el enviar 
        #por si se quiere enviar el formulario
        pyautogui.moveTo(btEnviar[0], btEnviar[1])
        if opcion == "111":
            filtroGestionMat()


def otroMenu():
    opcion = " "
    while(opcion != "0"):
        os.system("clear")
        print("-------------------------------")
        print("          Otras macros         ")
        print("-------------------------------")
        print("01- Comprobar series (Series)  ")
        # print("*desactivada*01- Activacion de material en la carga (FRAPPE)")
        # print("*desactivada*02- Imprimir Materiales")
        print("09- Volver")
        print("00- chauchau")
        print()
        print("Que necesitas hacer?")
        opcion=input(" ")

        if opcion == "09" or opcion == "9":
            menu1()

        elif opcion == "1":
            corroboraSeriesFM("Series")

        elif opcion == "0":
            os.system("clear")
            print("morite puto")
            sys.exit()

def preferencias_Velocidad():
    opcion = " "
    while(opcion != "0"):
        os.system("clear")
        print("-------------------------------")
        print("       Preferencias vel.       ")
        print("-------------------------------")
        print("01- tiempoUsual x2")
        print("02- tiempoUsual normal")
        print("09- Volver")
        print("00- chauchau")
        print()
        print("Que necesitas hacer?")
        opcion=input(" ")

        if opcion == "1" or opcion == "01":
            tiempoUsual=tiempoUsual*2
        elif opcion == "2" or opcion == "02":
            tiempoUsual=1.1 #secs
        if opcion == "09" or opcion == "9":
            menu1()

        elif opcion == "0":
            os.system("clear")
            print("morite puto")
            sys.exit()



disclaimer()
menu1()
