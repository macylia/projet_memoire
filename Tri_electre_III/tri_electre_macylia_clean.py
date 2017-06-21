#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import division
import sys
from xlrd import open_workbook
from xlwt import Workbook
import openpyxl
import xlsxwriter
import os

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
    print(Perf)
    nbCriteres = len(Perf[0])
    print(nbCriteres)
    nbActions = len(Perf)
    print(nbActions)
    crit = ["competences_et_resultats",	"appreciation_sociale",	"penibilite_du_travail", "anciennete_dans_lentreprise",	"assiduite"]
            #[ 'g%d'%i for i in range(1, nbCriteres) ]
    print(crit)
    for i in Perf:
        performances[i[0]] = dict(zip(crit, i[1:]))
    print("perforances : ")
    print(performances)
    act = performances.keys()
    print(act)
    print("weights :")
    weights = get_matrix_data(filename, 2)[0][1:]
    print(weights)
    weights = dict(zip(crit, map(lambda p: p/sum(weights), weights)))
    print(weights)
    
    seuil_preference = dict(zip(crit, get_matrix_data(filename, 3)[0][1:]))
    seuil_indiference = dict(zip(crit, get_matrix_data(filename, 3)[0][1:]))
    print(seuil_preference)
    print(seuil_indiference)
    vals = get_matrix_data(filename, 1)

    print("vetos : ")
    vetos = get_matrix_data(filename, 1)[0][1:]
    print(vetos)
    vetos = dict(zip(crit,  vetos))
    print(vetos)

    return performances, vetos, weights, act, crit, seuil_preference, seuil_indiference 

# ***************************************************
# ******************* discordance *******************
# ***************************************************
# retourne vrai si au moins 1 critiere est discordant
def appliquer_disc(s1, s2, conc):
    for critere in Criteres:
            veto = vetos[critere]
            gs1 = Performances[s1][critere]
            gs2 = Performances[s2][critere]
            d = discord_parti(gs1, gs2, veto, critere)
            if d > 0:
                print("        appliquer disc")
                return True
    print("        il n'a pas de discordance")
    return False   

# calcule dicordance partielle
# retourne 1 si le critere s'oppose fortement
# retourne 0 si le critere ne s'oppose pas
# retourne ]0, 1[ sinon
def discord_parti(gs1, gs2, veto, critere):
    diff = gs2-gs1
    if (gs1 < 0) & (gs2 < 0) :
        diff = abs(gs2 - gs1)
        
    if diff >= veto:
        return 1
    elif diff <= seuil_preference[critere]:
        return 0
    else :
        res = 1 - ((veto - diff)/(veto - seuil_preference[critere]))
        return res

# calcule discordance globale
# retourne 1 si il y'a pas discordance
# retourne [0, 1[ si il y'a discordance
def discord(s1, s2, conc):
    app_disc = appliquer_disc(s1, s2, conc)
    disc = 1.0
    if app_disc == True:
        for critere in Criteres:
            veto = vetos[critere]
            gs1 = Performances[s1][critere]
            gs2 = Performances[s2][critere]
            d = discord_parti(gs1, gs2, veto, critere)
            if d > conc :
                tmp = (1-d)/(1-conc)
                disc *=  tmp
                   
        print("        discordance entre : "+s1+" et "+s2+" est : "+str(disc) )
    return disc


# ***************************************************
# ******************* concordance *******************
# ***************************************************
# calcule concordance partielle
def concord_parti(gs1, gs2, critere):
    if gs1 >= gs2 - seuil_indiference[critere]:  
        return 1
    elif gs1 <= gs2 - seuil_preference[critere]:  
        return 0
    else:
        return (seuil_preference[critere] - (gs2 - gs1)) / (seuil_preference[critere] - seuil_indiference[critere])

# calcule concordance globale
# retourne la somme des WjCj
def concord(s1, s2):
    conc = 0.0
    for critere in Criteres:
        poids = Poids[critere]
        
        gs1 = Performances[s1][critere]
        gs2 = Performances[s2][critere]
        c = concord_parti(gs1, gs2, critere)
        conc += c * poids
        
    print("    concordance entre : "+s1+" et "+s2+" est : "+str(conc) )
    return conc    

# *************************************************************
# ******************* indice de credibilite *******************
# *************************************************************
def credibilite(ks1, ks2):
    conc = concord(ks1, ks2)
    disc = discord(ks1, ks2, conc)
    cred = conc * disc
    return cred

# ***************************************************************
# ******************* classement des salaries *******************
# ***************************************************************
# retourne meilleur entre 2 salaries
def compare_2_salarie(ks1, vs1, ks2, vs2):
    print("Calcule credibilite entre "+ks1+" et "+ks2+" et l'inverse : ")
    cred1 = credibilite(ks1, ks2)
    cred2 = credibilite(ks2, ks1)
    print("credibilite "+ks1+" et "+ks2+" : "+str(cred1))
    print("credibilite "+ks2+" et "+ks1+" : "+str(cred2))
    if cred1 >= cred2:
        print("*********** "+ks1+" surclasse "+ks2+" *********** ")
        return ks1, vs1
    else :
        print("*********** "+ks2+" surclasse "+ks1+" *********** ")
        return ks2, vs2

# retourne le meilleur salarie de la liste courrente
def compare_one_to_all(all_salaries):
    key_best = all_salaries.keys()[0]
    value_best = all_salaries[all_salaries.keys()[0]]
    print("\n\n\n\nSalaries restants : "+str(all_salaries.keys()))
    for k, v in all_salaries.items():
        if (k != key_best):
            print("\n\nComparing "+key_best+" to "+k)
            key_best, value_best = compare_2_salarie(key_best, value_best, k, v)
    all_salaries.pop(key_best)
    return key_best, value_best

# retourne liste ordonnee du meilleur au plus mauvais
def get_ordered_list():
    if len(Performances) == 1:
        key_best = Performances.keys()[0]
        value_best = Performances[Performances.keys()[0]]
        sorted_salarie.append((key_best, value_best))
    else:
        key_best, value_best = compare_one_to_all(Performances)
        print("\n >>>>>>>>>>>>>>>>>>>> current best : "+key_best+" <<<<<<<<<<<<<<<<<<<<")
        sorted_salarie.append((key_best, value_best))
        get_ordered_list()

# ********************************************
# ******************* main *******************
# ********************************************
if len(sys.argv) > 1:
    filename = sys.argv[1]
else: 
    filename= "data.xlsx"


Performances, vetos, Poids, Actions, Criteres, seuil_preference, seuil_indiference = get_all(filename)

workbook = xlsxwriter.Workbook('resultat.xlsx')
worksheet = workbook.add_worksheet('Resultat')
worksheet.write(0, 0, 'Liste Salaries trie')
worksheet.write(0, 1, 'Classification Optimiste')
worksheet.write(0, 2, 'Classification Pessimiste')

sorted_salarie = []
get_ordered_list()
print(sorted_salarie)

counter = 0
for row in sorted_salarie:
    counter += 1
    print("meilleur salarie "+row[0])
    worksheet.write(counter, 0, row[0])



workbook.close()
toBlock = raw_input()