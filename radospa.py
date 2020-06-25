import pyautogui
from win32clipboard import *
import win32con
import win32api
import win32gui
import time
import sys
import os

def radospa(carpeta):
    """ Obtiene la posición de la leyenda "Almacen" """
    pos=pyautogui.locateOnScreen('Almacen.png')
    if pos==None: 
        """ Sale si no la encuentra """
        return
    x=pos.left+pos.width/2
    y=pos.top+pos.height+20
    pyautogui.click(x,y)
    pyautogui.press('Home')
    while True:
        pyautogui.scroll(-400)
        time.sleep(0.2)
        pos=pyautogui.locateOnScreen('almacen40.png')
        if pos!=None: 
            x=pos.left+pos.width/2
            y=pos.top+pos.height/2
            pyautogui.doubleClick(x,y)            
            break

    """ Busca Campo B """
    time.sleep(0.5)
    pos=pyautogui.locateOnScreen('campob.png',confidence=0.95)
    if pos==None:
        print("Campo B no encontrado, saliendo ...")
        return

    """ Se posiciona en el campo B y selecciona todo """
    x=pos.left+pos.width+20
    y=pos.top+pos.height/2
    pyautogui.click(x,y)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    """ Llena valor de B """
    pyautogui.write('TI')
    pyautogui.press('\t')
    """ Obtiene el valor calculado en C """
    pyautogui.keyDown('ctrl')
    pyautogui.press('c')
    pyautogui.keyUp('ctrl')
    time.sleep(.2)
    OpenClipboard()
    valorC=GetClipboardData(CF_UNICODETEXT)
    print('El valor de C es %s' % valorC)
    """ Pone xxxxxx en campo D para que abra el dialogo de file """
    pyautogui.press('\t')
    pyautogui.write('xxxxxx')
    pyautogui.press('\t')
    t0=time.time()
    while True:
        time.sleep(0.2)
        pos=pyautogui.locateOnScreen('filename.png',confidence=0.95)
        if pos!=None:
            break
        if time.time() - t0 > 10:
            break
    if pos==None:
        print("File name no encontrado, saliendo ...")
        return
    x=pos.left+pos.width+20
    y=pos.top+pos.height/2
    pyautogui.click(x,y)
        
    #pyautogui.press('Esc')    
    # f=open("ccccc","r")
    # for l in f:
    #     s=l.split('-')
    #     p=int(s[0])
    #     li=int(s[1])
    #     break
    # f.close()
    for a in os.listdir(carpeta):
        s=a.split('-')
        if (carpeta[-1]!='\\'):
            carpeta=carpeta+'\\'
        pyautogui.write(carpeta+a)
        pyautogui.press('Enter')
        pyautogui.keyDown('shift')
        pyautogui.press('tab')
        pyautogui.keyUp('shift')
        pyautogui.write(s[0])
        break
        
    

def main(argv):
    if len(argv)!=2:
        return
    radospa(argv[1])

if __name__ == '__main__':
    main(sys.argv)