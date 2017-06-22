# projet_memoire
ELECTRE III sort + somme ponderee
### Description
Ce repos contient 3 dossiers :
1. memoire_macylia : qui contient le source du projet Latex.
2. Tri_electre_III : l'implementation de l'algorithme de tri multicritères ELECTRE III.
3. Somme_ponderee : l'implementation d'un algorithme de tri en utilisant la somme pondéré.



### Pré-requis:
* **Python** déjà installé

### Installation: 
Pour installer les 4 bibliotheques requises :
1. XlsxWriter 
2. openpyxl
3. xlrd
4. xlwt

`sudo pip install -r requirements.txt`


Le résultat de la commande devrait être comme suit : 

**Successfully installed XlsxWriter openpyxl xlrd xlwt**

### Input : 
Pour chaque algorithme un fichier **data.xlsx** existe déjà pré-rempli, et pourrais être éditer à votre guise.


### How to Run : 
en se plaçant à la racine du repos 
1. Pour Run l'algo Tri Electre II : `cd Tri_electre_III/;python tri_electre_macylia_clean.py`
2. Pour Run l'algo Somme pondérée : `cd Somme_ponderee/;python tri_sp_macylia_clean.py`

### Résultat : 
Le résultat sera affiché sur la console et aussi exporté vers le fichier **resultat.xlsx**.






