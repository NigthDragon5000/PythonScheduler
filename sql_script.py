# -*- coding: utf-8 -*-
"""
Created on Fri Feb 22 10:02:49 2019

@author: jcondori
"""


#import sqlite3
#from sqlite3 import OperationalError

#conn = sqlite3.connect('csc455_HW3.db')
#c = conn.cursor()



#import pyodbc
#import os

#os.chdir("D:/")

#sql_con = pyodbc.connect('driver={SQL Server};SERVER=OF00SRVBDH;Trusted_Connection=True')            

#c = sql_con.cursor()


import win32com.client as win32
#import schedule
outlook = win32.Dispatch('outlook.application')

def executeScriptsFromFile(filename,sql_con):
    
    if sql_con :

       c = sql_con.cursor()
       fd = open(filename, 'r')
       sqlFile = fd.read()
       fd.close()

    sqlCommands = sqlFile.split('GO')

    for command in sqlCommands:

        print(command)  
        try:
            c.execute(command)
            return ''
        except Exception as e:
           print(str(e))
           #mail = outlook.CreateItem(0)
           #mail.To = 'jcondori@compartamos.pe'
           #mail.Subject = 'Error'
           #mail.Body = 'Error en el query en: '+str(e)
           #attachment  = "D:\\example.xlsx"
           #mail.Attachments.Add(attachment)
           #mail.Send()
           sql_con.commit()
           return e
        

#try:
#    do_something()
#except BaseException as e:
#    logger.error('Failed to do something: ' + str(e))

          
#executeScriptsFromFile('quericito.sql')

