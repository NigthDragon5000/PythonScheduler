# -*- coding: utf-8 -*-
"""
Created on Fri Feb 22 10:59:03 2019

@author: jcondori
"""


import pyodbc
import os
from sql_script import executeScriptsFromFile 
import schedule
import time
import win32com.client as win32
import PySimpleGUI as sg
#import unittest



def execute_macro(ruta,macro): 
    try:
        xlApp = win32.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser(ruta)
        print(xlsPath)
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        wb.RefreshAll()
        xlApp.Run(macro)             
        wb.Save()
        xlApp.Quit()
        print("Macro ran successfully!")
    except:
        print("Error found while running the excel macro!")
        #xlApp.Quit()


#ruta='D:\\previ.xlsm'
#macro='macrito'
#execute_macro(ruta,macro)

# Refresh all pivot tables



outlook = win32.Dispatch('outlook.application')

os.chdir("D:/")

#sql_con = pyodbc.connect('driver={SQL Server};SERVER=OF00SRVBDH;Trusted_Connection=True',autocommit=True)            

#sql_con.close()
#filename  = 'quericito.sql'
#filename  = 'Grupal.sql'
#filename  = 'GrupalEnero19.sql' # Verificado que funciona
#filename  = 'error.sql'
#filename  = 'GrupalFebrero19.sql' # Verificado que funciona
#filename  = 'pruebitas.sql' # Verificado que funciona
#filename  = 'detalle.sql' # Verificado que funciona


#filename = sg.PopupGetFile('Please enter a file name')
#hora = sg.PopupGetText('Hora:', 'Insertar la hora requerida de ejecucion')
#destinatarios=sg.PopupGetText('Destinatarios', 'Ingresar Destinatarios')

filename  = 'pruebitas.sql' # Verificado que funciona
hora = '18:02'
destinatarios='jcondori@compartamos.pe'

#
#layout = [[sg.PopupGetFile('Please enter a file name')],
#          [sg.PopupGetText('Title', 'Insertar la hora requerida de ejecucion')]]
#
#event, values = sg.Window('Enter a number example', layout).Read()
#
#sg.Popup(event, values[0])
#
#
#layout = [[sg.Text('Enter a Number')],
#          [sg.Input()],
#          [sg.OK()] ]
#
#event, values = sg.Window('Enter a number example', layout).Read()
#
#sg.Popup(event, values[0])

#ruta='D:\excel_macro.xlsm' 
#macro='macrito'

#execute_macro(ruta,macro)


#import pandas as pd

def ejecucion():
    sql_con = pyodbc.connect('driver={SQL Server};SERVER=OF00SRVBDH;Trusted_Connection=True',autocommit=True)
    e=executeScriptsFromFile(filename,sql_con)
    
    ''' Borrar'''
    #f1="select  * from ##pruebitass2"
    #base1  = pd.read_sql(f1, sql_con)
    #base1.to_excel("output.xlsx",
    #         sheet_name='Sheet1')
    
    ruta='D:\previ.xlsm'
    macro='macrito'
    execute_macro(ruta,macro)
    
    ''' '''
    
    #sql_con.close()
    mail = outlook.CreateItem(0)
    #mail.To = 'jcondori@compartamos.pe;cleto@compartamos.pe'
    mail.To = destinatarios
    mail.Subject = 'Proceso'
    if str(e) != 'None':
        mail.Body = 'Error: '+str(e)
    else:
        mail.Body = 'Sin errores'
        
    ''' Borrar '''
   # attachment  = "D:\output.xlsx"
   # mail.Attachments.Add(attachment)
    ''' '''
    
    mail.Send()
    return schedule.CancelJob


schedule.clear()
#schedule.every(1).minutes.do(ejecucion)
#schedule.every().hour.do(job)
schedule.every().day.at(hora).do(ejecucion)

#schedule.every(1).to(10).minutes.do(ejecucion)
#schedule.every().monday.do(job)

#schedule.every().wednesday.at("03:00").do(ejecucion)
#schedule.every().day.at("19:20").do(ejecucion())

while True:
    schedule.run_pending()
    time.sleep(1)


#executeScriptsFromFile('error.sql',sql_con)

#sql_con.close()
    
    
    
    
''' Otros '''

#import time
#
#def foo():
#  print(time.ctime())
#  
#
#
#while True:
#  foo()
#  time.sleep(10)
#  
#  
#  
#import threading, time, signal
#
#from datetime import timedelta
#
#WAIT_TIME_SECONDS = 1
#
#class ProgramKilled(Exception):
#    pass
#
#def foo():
#    print(time.ctime())
#    
#def signal_handler(signum, frame):
#    raise ProgramKilled
#    
#class Job(threading.Thread):
#    def __init__(self, interval, execute, *args, **kwargs):
#        threading.Thread.__init__(self)
#        self.daemon = False
#        self.stopped = threading.Event()
#        self.interval = interval
#        self.execute = execute
#        self.args = args
#        self.kwargs = kwargs
#        
#    def stop(self):
#                self.stopped.set()
#                self.join()
#    def run(self):
#            while not self.stopped.wait(self.interval.total_seconds()):
#                self.execute(*self.args, **self.kwargs)
#            
#if __name__ == "__main__":
#    signal.signal(signal.SIGTERM, signal_handler)
#    signal.signal(signal.SIGINT, signal_handler)
#    job = Job(interval=timedelta(seconds=WAIT_TIME_SECONDS), execute=foo)
#    job.start()
#    
#    while True:
#          try:
#              time.sleep(1)
#          except ProgramKilled:
#              print ("Program killed: running cleanup code")
#              job.stop()
#              break
