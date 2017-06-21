#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import division
import sys
from xlrd import open_workbook
from xlwt import Workbook
import openpyxl
import xlsxwriter
import operator

## ****************** get data from xslx file ****************** 
def get_matrix_data(fileName, sheetNumber):
    values = []
    wb = open_workbook(fileName)
    first_sheet = wb.sheet_by_index(sheetNumber)
    nb_rows = first_sheet.nrows

    criteres = []
    for row in range (1, nb_rows):
        col_names = first_sheet.row(0)
        criteres.append(col_names)
        col_value = []
        for name, col in zip(col_names, range(first_sheet.ncols)):
            value = (first_sheet.cell(row,col).value)
            col_value.append(value)
        values.append(col_value)
    
    if (sheetNumber in [range(2, 5)]):
        values = values[0][1:]
   
    return values

def get_all(filename):
    performances = dict()
    Perf = get_matrix_data(filename, 0)
    nbActions = len(Perf)
    crit = ["competences_et_resultats",	"appreciation_sociale",	"penibilite_du_travail", "anciennete_dans_lentreprise",	"assiduite"]
            #[ 'g%d'%i for i in range(1, nbCriteres) ]
    for i in Perf:
        performances[i[0]] = dict(zip(crit, i[1:]))
    print("donnees avant normalisation : ")
    
    print_data(performances)
    
    weights = get_matrix_data(filename, 1)[0][1:]
    print("\n\n\n\n")
    print("Poids normalises: ")
    weights = dict(zip(crit, map(lambda p: p/sum(weights), weights)))
    print(weights)
       
   
    return performances, weights, crit 

#************** Affichage des donnees ligne par ligne
def print_data(performances):
    for p in performances.keys():
        print(p+" => " +str(performances[p]))

# ***********************************************************
# **************** normalisation des donnees ****************
# ***********************************************************
def normalisation_data(performances):
    print("\n\n")
    print("************************ donnees normalisees ************************")
    for p in performances.keys():
        somme_ligne = sum(performances[p].values())
        for c in performances[p].keys():
            performances[p][c] = performances[p][c]/somme_ligne
        
    print_data(performances)
    print("\n\n\n\n")                      

# ******** calcul de la somme pondere pour un salarie s
# ******** retuorne pour chaque salarie le tuple (cles, somme_pondere)
def somme_pond(s):
    sp = 0.0
    for critere in Criteres:
        poids = Poids[critere]
        gs = Performances[s][critere]
        sp += (gs * poids)
        
    return s, sp

# ******** itere sur les salarie et calcul de la somme pondere de chaqu'un
# ******** retourne une liste trié par ordre decroissant par rapport à la somme pondere
def scores():
    salarie_score = dict()
    for s in Performances.keys():
        key, val = somme_pond(s)
        salarie_score[key] = val
    
    sorted_x = sorted(salarie_score.items(), key=operator.itemgetter(1), reverse=True)
    
    print("******** salariees trie *************")
    for x, y in sorted_x:
        print(x+" "+str(y))
    return sorted_x
        
# ********************************************
# ******************* main *******************
# ********************************************
if len(sys.argv) > 1:
    filename = sys.argv[1]
else: 
    filename= "data.xlsx"

Performances, Poids, Criteres = get_all(filename)

workbook = xlsxwriter.Workbook('resultat.xlsx')
worksheet = workbook.add_worksheet('Resultat')
worksheet.write(0, 0, 'Liste Salaries trie')
    
normalisation_data(Performances)
list_sorted = scores()


counter = 0
for key, value in list_sorted:
    counter += 1
    worksheet.write(counter, 0, key)

workbook.close()
a = raw_input()