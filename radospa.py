import pyautogui
from win32clipboard import *
import win32con
import win32api
import win32gui
import time
import sys
import os

def buscaImagen(img,tiempo=0):
    t0=time.time()
    while True:
        time.sleep(0.2)
        pos=pyautogui.locateOnScreen(img,confidence=0.95)
        if pos!=None:
            break
        if time.time() - t0 > tiempo:
            break
    return pos

def buscaImagenConScroll(img,tiempo=10):
    t0=time.time()
    while True:
        pyautogui.scroll(-400)
        time.sleep(0.2)
        pos=buscaImagen(img)
        if pos!=None:
            break
        if time.time() - t0 > tiempo:
            break
    return pos

def interaccion(carpeta,archivo):
    
    """ Obtiene la posición de la leyenda "Almacen" """
    pos=buscaImagen('Almacen.png',10)
    
    if pos==None: 
        """ Sale si no la encuentra """
        return
    
    """ Hace click en el 1er item visible """
    x=pos.left+pos.width/2
    y=pos.top+pos.height+20
    pyautogui.click(x,y)
    
    """ Vuelve al inicio de la lista """
    pyautogui.press('Home')

    """ Busca el item 40 """
    pos=buscaImagenConScroll('almacen40.png')
    if pos==None:
        """ Sale si no la encuentra """
        return

    """ Doble click en la imagen """
    x=pos.left+pos.width/2
    y=pos.top+pos.height/2
    pyautogui.doubleClick(x,y)        

    """ Espera un poco """
    time.sleep(0.5)

    """ Busca Campo B """
    pos=buscaImagen('campob.png')
    if pos==None:
        """ Sale si no la encuentra """
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

    """ Pasa al siguiente campo (C) """
    pyautogui.press('\t')

    """ Espera un poco """
    time.sleep(.5)

    """ Obtiene el valor calculado en C """
    pyautogui.keyDown('ctrl')
    pyautogui.press('c')
    pyautogui.keyUp('ctrl')
    OpenClipboard()
    valorC=GetClipboardData(CF_UNICODETEXT)
    CloseClipboard()

    """ Pasa al siguiente campo (D) """
    pyautogui.press('\t')

    """ Pone xxxxxx en campo D para que abra el dialogo de file """
    pyautogui.write('xxxxxx')

    """ Pasa al siguiente campo (E), como D=xxxx => abre el diálogo de file """
    pyautogui.press('\t')

    """ Espera que aparezca el díalogo """
    pos=buscaImagen('filename.png',10)
    if pos==None:
        """ Sale si no lo encuentra """
        return

    """ Click en el campo "File name:" """
    x=pos.left+pos.width+20
    y=pos.top+pos.height/2
    pyautogui.click(x,y)

    """ Escribe nombre del archivo (path completo) """
    pyautogui.write(carpeta+archivo)

    """ Presiona ENTER """
    pyautogui.press('Enter')
    
    """ Vuelve al campo anterior (D) """
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.keyUp('shift')

    """ Escribe los números hasta el primer guión """
    s=archivo.split('-')
    pyautogui.write(s[0])

    """ Busca el botón OK """
    pos=buscaImagen('OK.png')
 
    if pos!=None:
        """ Si lo encuentra lo presiona """
        x=pos.left+pos.width/2
        y=pos.top+pos.height/2
        pyautogui.click(x,y)

    return valorC

def radospa(carpeta):
    if (carpeta[-1]!='\\'):
        carpeta=carpeta+'\\'

    print('Carpeta: %s' % carpeta)
    for archivo in os.listdir(carpeta):
        print('Archivo: %s' % archivo)
        valorC=interaccion(carpeta,archivo)
        print('El valor de C es %s' % valorC)
        time.sleep(2)
            
def main(argv):
    if len(argv)!=2:
        return
    radospa(argv[1])

if __name__ == '__main__':
    main(sys.argv)