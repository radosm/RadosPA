import pyautogui
from win32clipboard import *
import win32con
import win32api
import win32gui
import time

def radospa():
    """ Obtiene la posición de la leyenda "Campo A:" """
    pos=pyautogui.locateOnScreen('campoa.png')
    if pos==None: 
        """ Sale si no la encuentra """
        return

    """ Se posiciona en el primer campo y selecciona todo """
    x=pos.left+pos.width+20
    y=pos.top+pos.height/2
    pyautogui.click(x,y)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    """ Llena valores A y B """
    pyautogui.write('algo ...')
    pyautogui.press('\t')
    pyautogui.write('TI')
    pyautogui.press('\t')
    """ Obtiene el valor calculado en C """
    pyautogui.keyDown('ctrl')
    pyautogui.press('c')
    pyautogui.keyUp('ctrl')
    OpenClipboard()
    valorC=GetClipboardData(CF_UNICODETEXT)
    print('El valor de C es %s' % valorC)
    """ Pone xxxxxx en campo D para que abra el dialogo de file"""
    pyautogui.press('\t')
    pyautogui.write('xxxxxx')
    pyautogui.press('\t')

def main():
    radospa()

if __name__ == '__main__':
    main()