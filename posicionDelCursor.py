import sys
import pyautogui
import time
import os
tiempoUsual=1.1
tiempoPagina=tiempoUsual*5


def openBrowser():
    os.system("xdotool key super+1")
    time.sleep(tiempoUsual)

def busqueda(busqueda, tiempo=tiempoUsual*0.2, listo="No"): #busqueda->str
    pyautogui.hotkey("ctrl","f")
    time.sleep(tiempo)
    pyautogui.typewrite(str(busqueda))
    if listo == "No":
        pyautogui.press("enter")
        time.sleep(tiempo)
        pyautogui.press("esc")

def ppal():
    position=pyautogui.position()
    print(str(position[0])+", " +str(position[1]))

def otrasfunciones():
    cantInstalaciones=int(input("Cuantas 'Instalaciones hay que verificar? -> "))
    cantSeries=int(input("Cuantos 'Series hay que verificar? -> "))
    openBrowser()
    
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

    # insertar consumos en AppSheet
    time.sleep(tiempoPagina)
    pyautogui.hotkey("ctrl","home")
    time.sleep(tiempoUsual*0.5)
    busqueda("cargado")
    time.sleep(tiempoUsual)
    pyautogui.press("down")
    time.sleep(tiempoUsual)
    pyautogui.hotkey("ctrl","c")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","home")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","pgdn")
    time.sleep(tiempoUsual)
    busqueda("Fecha")
    [pyautogui.press("tab") for _ in range(14)]
    time.sleep(tiempoUsual)
    pyautogui.hotkey("ctrl","a")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","v")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","pgup")

    #Cargar instalaciones sin consumos
    time.sleep(tiempoPagina)
    busqueda("CopyPaste")
    time.sleep(tiempoUsual)
    pyautogui.press("down")
    time.sleep(tiempoUsual)
    pyautogui.hotkey("ctrl","shift","down")
    time.sleep(tiempoUsual)
    pyautogui.hotkey("ctrl","c")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","home")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","pgdn")
    time.sleep(tiempoUsual)
    busqueda("Fecha")
    [pyautogui.press("tab") for _ in range(7)]
    pyautogui.hotkey("ctrl","a")
    time.sleep(tiempoUsual)
    pyautogui.hotkey("ctrl","v")
    time.sleep(tiempoUsual*0.5)
    pyautogui.hotkey("ctrl","pgup")

    listo()



def main():
    print("Cual funcion queres?")
    print("1 - Principal(posicion del cursor)")
    print("2 - Otras")
    opcion=input(" -> ")

    if opcion == "1":
        ppal()
    elif opcion == "2":
        openBrowser()
        time.sleep(1)
        otrasfunciones()
    else:
        print("error...")
        main()

main()
