# -*- coding: utf-8 -*-
"""
Created on Tue Mar 17 20:20:02 2020

@author: jacques
"""

import numpy as np
import xlrd as xl      #pour importer les données de l'excel
import matplotlib.pyplot as plt

workbook = xl.open_workbook('Donnees_autoroute.xlsx')
SheetNameList = workbook.sheet_names()
for i in np.arange( len(SheetNameList) ):
#    print( SheetNameList[i] )


 """ SELECTIONNER UNE FEUILLE EXCEL """
worksheet = workbook.sheet_by_name(SheetNameList[0])
num_rows = worksheet.nrows 
num_cells = worksheet.ncols 
#print( 'num_rows, num_cells', num_rows, num_cells )

Liste_Vitesse_Moyenne = []
Liste_Heure_Journee = []
Liste_Debit = []


curr_row = 0
while curr_row < num_rows:
    row = worksheet.row(curr_row)
    curr_cell = 0
    while curr_cell < num_cells:
        # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
        cell_type = worksheet.cell_type(curr_row, curr_cell)
        cell_value = worksheet.cell_value(curr_row, curr_cell)
        
        if curr_cell == 0 : 
            Liste_Vitesse_Moyenne.append(cell_value)

        if curr_cell == 1:
            Liste_Heure_Journee.append(cell_value)
            
        if curr_cell == 3 :
            Liste_Debit.append(cell_value)
                        
        curr_cell += 1
    curr_row += 1

print(Liste_Vitesse_Moyenne,Liste_Heure_Journee)

#plt.plot(Liste_Heure_Journee,Liste_Vitesse_Moyenne)
#plt.plot(Liste_Heure_Journee,Liste_Debit)
#plt.grid()

plt.plot(Liste_Vitesse_Moyenne,Liste_Debit)


def Gain_temps(heure_début_etude,heure_fin_etude,Vimpose):
    """on prend l'heure de début et de fin en 1,9 """
    
    """on recherche l'indice correspondant aux heures """
    indice_depart = 0
    indice_fin = 0   
    for k in range (len(Liste_Vitesse_Moyenne)):
        if Liste_Heure_Journee[k] == heure_début_etude : 
            indice_depart = k
        if Liste_Heure_Journee[k] == heure_fin_etude : 
            indice_fin = k
    print(indice_depart)
    print(indice_fin)
    
    Gain = 0
    for i in range (indice_depart,indice_fin):
        Gain+= Liste_Debit[i]*(1/Liste_Vitesse_Moyenne[i]-1/Vimpose)
    
    return Gain

Vimpose = 110


t1 = 7
t2 = 10
Temps = Gain_temps(t1,t2,Vimpose)
print(Temps)
 



    
