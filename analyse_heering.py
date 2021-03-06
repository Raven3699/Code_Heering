import pandas as pd 
import os
import glob
import math
import time 
from matplotlib import pyplot as plt

start_execution = time.process_time() # Temps d'exécution total du Scripte

############################################################################################
################################### Variable changeable ####################################
############################################################################################

temperature_max = 29
temperature_min = 26
CO2_max = 1800

Chemin_telechargement = 'C://Users\p36200223\Documents\Base_heering\Scrapy_T\*.xls'
Fichier_blanc = 'C://Users/p36200223/Documents/Base_heering/S4/file_1.xls'
Nom_du_fichier = 'Base_heering_python.xlsx'

############################################################################################
###################################### Ne pas toucher ######################################
############################################################################################

list_file = glob.glob(Chemin_telechargement)

# Initialisation

petite_base_inte = []
petite_base = []
name = []
name_split = []
truck_id_not = []
truck_id = []
n_trajet_not = []
n_trajet = []
grosse_base = []
n = 0

start_recuperation = time.process_time() # Temps réucpération des fichiers Excel

# Supression des fichiers non exploitable

for i in range(len(list_file)):
    if len(pd.read_excel(list_file[i])) > 50 and pd.read_excel(list_file[i])['GPS'][3] == 1 and pd.read_excel(list_file[i])['Unnamed: 5'][0] == 'Ville' and pd.read_excel(list_file[i])['Température [°C]'][3] > 0:
        petite_base_inte.append(list_file[i]) # récupère le chemin des fichiers gardés
        petite_base.append(pd.read_excel(list_file[i])) # récupère les fichiers gardés
        if n == 50: # Graph de l'avancement de la récupération des fichiers
            a = i/len(list_file)*100
            plt.plot(1,a,'go')
            plt.show()
            n = 0
        n = n + 1

print('Temps Récupération fichiers Excel : ', time.process_time() - start_recuperation)

# Rajouter colonne avec immatriculation tracteur et numéro de trajet

for i in range(len(petite_base_inte)) :
    name.append(os.path.basename(petite_base_inte[i])) # récupère le nom du fichier
    name_split.append(name[i].split('_')) # Séparer la chaine de caractères en fonction des underscore
    n_trajet_not.append(name_split[i][2].split('.')) # Séparer l'une des chaines de la séparation précédente en fonciton du point
    n_trajet.append(n_trajet_not[i][0]) # Récupère le numéro de trajet
    truck_id.append(name_split[i][1]) # Récupère le numéro correspondant au camion
    # Sur Heering Link les semi sont différenciées suivant les numéros correspondant au truck_id, je les modifies pour l'exploitation des données
    if truck_id[i] == '301' :
        truck_name = 'MASTER 4'
    else :
        if truck_id[i] == '98' :
            truck_name = 'S4'
        else : 
            if truck_id[i] == '162' :
                truck_name = 'S5'
            else :
                if truck_id[i] == '230' :
                    truck_name = 'S6'
                else :
                    if truck_id[i] == '295' :
                        truck_name = 'S7'
                    else :
                        truck_name = 'error'  # Au cas où un fichier réucpère un autre numéro
    petite_base[i]['immat'] = truck_name
    petite_base[i]['n°trajet'] = n_trajet[i]
    
# Fonction qui permet de séparer la dataFrame d'un trajet en x trajets (1 trajet par déchargment)
# Retourne donc une liste avec les x nouveaux trajets

def separation (marks_data) :
    
    # Initialisation
    pause = 0
    stop = False
    base_inte = []
    
    # Fonction
    marks_data.dropna(subset=['GPS'], inplace=True) # enleve les lignes où il y a une case vide dans la colonne GPS
    marks_data = marks_data.reset_index(drop = True) 
    marks_data['Date / heure'] = pd.to_datetime(marks_data['Date / heure']) 
    marks_data = marks_data.sort_values('Date / heure').reset_index(drop = True) # Trier le fichier par heure croissant
    marks_data = marks_data[10:].copy() # Suprimer les 10 premières valeurs
    marks_data = marks_data.reset_index(drop = True)
    dataframe_inte = pd.read_excel(Fichier_blanc)
    dataframe_inte['immat'] = marks_data['immat'][2]
    dataframe_inte['n°trajet'] = marks_data['n°trajet'][2]
    
    stop = False # Lorsque les sondes s'active le camion est arrété, mais j'initialise de sorte que le premier 'if' soit vrai
    marks_data_size = len(marks_data) # Nombre de prise de mesure pendant le trajet
    for i in range(marks_data_size) : # 'For' pour parcourir toutes les lignes
        if marks_data['Vitesse [KM/H]'][i] <= 5 and stop == False : # Si le camion s'arrete
            stop = True # Camion arrété
            heure_arret = Heure(marks_data,i) # Heure arret
            # Heure est une fonction qui renvoie l'heure quand on lui donne les données d'un trajet entier avec le numéro de ligne souhaité
        else : # Sinon le camion ne s'arrète pas
            if marks_data['Vitesse [KM/H]'][i] > 5 and stop == True : # Si le camion démarre
                heure_depart = Heure(marks_data,i-1) # Heure avant départ
                if heure_depart - heure_arret > 0.25 : # Si pause supérieur à 15 min
                    stop = False # Camion roule
                    pause = pause + 1 # Une pause de plus 
                    if pause != 1 : # Si pas la première pause 
                        base_inte.append(marks_data[:i-marks_data_size].copy()) # Ajouter un trajet avec les i premières données 
                else: # Sinon pause inférieur ou égale à 15 min
                    stop = False # Camion roule
        if i == marks_data_size-1  : # Si dernière valeur 
              base_inte.append(marks_data) # Ajouter le trajet entier à la liste
    return base_inte # Retourne une liste avec tout les trajets séparés
    
# Fonction qui renvoie l'heure de la ligne i d'un trajet (format X,XXh ex : 7,75h = 7h45)

def Heure (marks_data,i):

    marks_data = marks_data.astype({'Date / heure': str}) # Changement de type -> chaine de caractères
    
    t_date_heure = marks_data['Date / heure'][i] # Récupère la date/Heure de la i ème ligne
    t_date_heure_split = t_date_heure.split() # Sépare en fonciton de l'espace entre la date et l'heure
    t_heure_minute = t_date_heure_split[1] # Récupère la partie Heure
    t_heure_minute_split = t_heure_minute.split(":") # Sépare l'heure des minutes
    
    heure = float(t_heure_minute_split[0]) # Récupère l'heure
    minute = float(t_heure_minute_split[1]) / 60 # Récupère les minutes mais en heure (ex : 30 min -> 0.5h)
    
    heure = heure + minute # Calcule de l'heure exacte en heure
        
    return heure

# Définition de la fonction qui ressort les informations voulu à partir d'une base de donné 
# dataFrame en entré et ressort une liste 

def analyse(marks_data) : 
    
    # 1) Initialisation
    
    grande_temp = 0
    petite_temp = 0
    grand_CO2 = 0
    prelevement_garde = len(marks_data)
    arret = []
    stop = False
    pause = -1
    temps_de_pause_tot = 0
    heure_depart = 60
    dif_pos = 0
    dif_neg = 0
    dif_pos_pause = 0
    dif_neg_pause = 0
    
    # 2) Moyenne température sur tous le trajet
    
    moyenne = marks_data['Température [°C]'].mean()
    moyenne = int(moyenne*100) / 100 # Arrondir à deux chiffres après la virgule
    
    moyenne_avant = marks_data['Unnamed: 9'].mean()
    moyenne_avant = int(moyenne_avant*100) / 100
        
    # 3) Nombre de valeur de température plus fort que temperature_max et inférieur à temperature_min (défini en début de programme)
    # Puis pondéré par le nombre de mesures prises
    
    grande_temp = len(marks_data[marks_data['Température [°C]'] > temperature_max]) #récupère la taille de la dataframe sans les température inferieur à temp_max
    
    petite_temp = len(marks_data[marks_data['Température [°C]'] < temperature_min]) #récupère la taille de la dataframe sans les température supérieur à temp_min
                         
    grande_temp_pond = (grande_temp / prelevement_garde) * 100 # nb de haute température par rapport au nombre de valeur prise (en %)
    grande_temp_pond = int(grande_temp_pond*100) / 100
    
    petite_temp_pond = (petite_temp / prelevement_garde) * 100 # nb de faible température par rapport au nombre de valeur prise (en %)
    petite_temp_pond = int(petite_temp_pond*100) / 100
     
    # 4) Moyenne référence
    
    moyenne_ref = marks_data['Unnamed: 15'].mean()
    moyenne_ref = int(moyenne_ref*100) / 100
    
    # 5) Moyenne C02 et CO2 au dessus de 1800 ppm puis pondéré par nombre de mesures prises
    
    moyenne_CO2 = marks_data['CO2 [PPM]'].mean()

    grand_CO2 = len(marks_data[marks_data['CO2 [PPM]'] > CO2_max])
                
    grand_CO2_pond = (grand_CO2 / prelevement_garde) * 100
    grand_CO2_pond = int(grand_CO2_pond*100) / 100
    
    # 6) Isoler la date 
    
    date_heure = [None] * (prelevement_garde-1)
    date_heure_split = [None] * (prelevement_garde-1)
    heure_minute = [None] * (prelevement_garde-1)
    heure_minute_split = [None] * (prelevement_garde-1)
    
    marks_data = marks_data.astype({'Date / heure': str})
    
    for i in range(prelevement_garde-1) :
        date_heure [i] = marks_data['Date / heure'][i]
        date_heure_split[i] = date_heure[i].split()
        heure_minute[i] = date_heure_split[i][1]
        heure_minute_split[i] = heure_minute[i].split(":")
        
    date = date_heure_split[0][0]
    
    # 7) Calucler durée trajet en heure
    
    heure = float(heure_minute_split[prelevement_garde - 2][0]) - float(heure_minute_split[0][0])
    
    if heure < 0 :
        heure = -heure
    minute = float(heure_minute_split[prelevement_garde - 2][1]) - float(heure_minute_split[0][1])
    
    duree = heure + minute / 60
    
    minute = math.floor((duree - math.floor(duree))*100)/100*60
    
    if int(minute) >= 10:    
        duree = str(int(heure)) + ':' + str(int(minute))
    else:
        duree = str(int(heure)) + ':0' + str(int(minute))
    
    # 8) nombre de pause / temps de pause

    for i in range(prelevement_garde) :
        if i != prelevement_garde-1 :    
            if marks_data['Vitesse [KM/H]'][i] <= 5 and stop == False and marks_data['Vitesse [KM/H]'][i+1] <= 5:
                stop = True
                heure_arret = Heure(marks_data,i)
            else :
                if marks_data['Vitesse [KM/H]'][i] > 5 and stop == True and marks_data['Vitesse [KM/H]'][i+1] > 5:
                    backup = heure_depart
                    heure_depart = Heure(marks_data,i-1)
                    if heure_depart - heure_arret > 0.25 :
                        stop = False
                        pause = pause + 1
                        if pause > 0 :                        
                            temps_de_pause = heure_depart - heure_arret
                            temps_de_pause_tot = temps_de_pause_tot + temps_de_pause
                            heure = math.floor(temps_de_pause)
                            minute = math.floor((temps_de_pause - math.floor(temps_de_pause))*100)/100*60                           
                            if int(minute) >= 10:    
                                duree_2 = str(int(heure)) + ':' + str(int(minute))
                            else:
                                duree_2 = str(int(heure)) + ':0' + str(int(minute))
                            arret.append(duree_2)
                    else :
                        stop = False
                        heure_depart = backup
        else :
            temps_depuis_arret = heure_arret - heure_depart
            heure_x = math.floor(temps_depuis_arret)
            if heure_x < 0 :
                heure_x = 0
            minute_x = math.floor((temps_depuis_arret - math.floor(temps_depuis_arret))*100)/100*60
            if int(minute_x) >= 10:    
                temps_depuis_arret = str(int(heure_x)) + ':' + str(int(minute_x))
            else:
                temps_depuis_arret = str(int(heure_x)) + ':0' + str(int(minute_x))
                
            heure_depart = Heure(marks_data,i)
            stop = False
            pause = pause + 1
            if pause > 0 :                        
                temps_de_pause = heure_depart - heure_arret
                temps_de_pause_tot = temps_de_pause_tot + temps_de_pause
                heure = math.floor(temps_de_pause)
                minute = math.floor((temps_de_pause - math.floor(temps_de_pause))*100)/100*60
                if int(minute) >= 10:    
                    duree_2 = str(int(heure)) + ':' + str(int(minute))
                else:
                    duree_2 = str(int(heure)) + ':0' + str(int(minute))
                arret.append(duree_2)
    
    heure_1 = math.floor(temps_de_pause_tot)
    if heure_1 < 0 :
        heure_1 = 0
    minute_1 = math.floor((temps_de_pause_tot - math.floor(temps_de_pause_tot))*100)/100*60
    if int(minute_1) >= 10:    
        arret_tot = str(int(heure_1)) + ':' + str(int(minute_1))
    else:
        arret_tot = str(int(heure_1)) + ':0' + str(int(minute_1))
    
    # 9) Récupération Nom de la ville de déchargement/Code postal/immatriculation semi/n° de lot

    ville = marks_data['Unnamed: 5'][prelevement_garde-1]
    
    code_postal = marks_data['Unnamed: 7'][prelevement_garde-1]
    
    truck_id =  marks_data['immat'][2]
    
    n_trajet = marks_data['n°trajet'][2]
    
    # 10) Regarder la plus grosse différence de température entre moyenne et moyenne des trois sondes
    
    for i in range(prelevement_garde-1):
        devant = marks_data['Unnamed: 9'][i]
        milieu = marks_data['Unnamed: 10'][i]
        derriere = marks_data['Unnamed: 11'][i]
        if marks_data['immat'][2] == 'MASTER 4' :
            moyenne_fixe = (devant + milieu) / 2
        else :
            moyenne_fixe = (devant + milieu + derriere) / 3 # Moyenne des trois sondes
        moyenne_heering =  marks_data['Température [°C]'][i] # Moyenne calculés par Heering
        dif_variable = moyenne_fixe - moyenne_heering # Différence entre les deux moyennes
        if marks_data['Vitesse [KM/H]'][i] > 5 : # Si le camion est à l'arret 
            if abs(dif_variable) > 0.1 : # Pour éviter d'avoir une difference de moyenne très faible et donc pas utile
                if dif_variable > 0 : # Moyenne sonde est plus grande que moyenne heering
                    if dif_pos < dif_variable: # Permet de modifier la différence positive si 
                        dif_pos = dif_variable # la différence de la ligne en cour est plus élevé
                else : # Moyenne sonde est plus faible que moyenne heering
                    if dif_neg > dif_variable: 
                        dif_neg = dif_variable 
        else : # Si le camion est arrété
            if abs(dif_variable) > 0.1 :
                if dif_variable > 0 : # moyenne sonde est plus grande que moyenne heering
                    if dif_pos_pause < dif_variable:
                        dif_pos_pause = dif_variable
                else :
                    if dif_neg_pause > dif_variable: # Permet de modifier la différence négative si 
                        dif_neg_pause = dif_variable # la différence de la ligne en cour est plus faible
                        
        dif_pos = int(dif_pos*100) / 100
        dif_neg = int(dif_neg*100) / 100    
        dif_pos_pause = int(dif_pos_pause*100) / 100
        dif_neg_pause = int(dif_neg_pause*100) / 100  
        
    # 11) fin fonction : création liste avec tous les éléments voulu
    
    info_transport  =  [n_trajet,date,truck_id,ville,code_postal, duree,temps_depuis_arret,
                        moyenne, moyenne_avant, moyenne_ref, grande_temp, grande_temp_pond,
                        petite_temp, petite_temp_pond, moyenne_CO2, grand_CO2,
                        grand_CO2_pond,dif_pos,dif_neg,dif_pos_pause,
                        dif_neg_pause,pause,arret_tot,arret]
    
    return info_transport 

# Créer la grosse_base qui contient tous les trajets découpé

start_separation = time.process_time() # Temps pour séparer les trajets suivant les pauses

for i in range(len(petite_base)):
    grosse_base = grosse_base + separation(petite_base[i])
    if n == 50: # Graph de l'avancement de la séparation des trajets
        b = i/len(petite_base)*100
        plt.plot(2,b,'bo')
        plt.show()
        n = 0
    n = n + 1
    
print('Temps boucle séparation : ', time.process_time() - start_separation)

# Mettre les valeurs souhaitées en DataFrame

    # Initialisation de la dataFrame avec le premier trajet
    


initialisation_dataframe = analyse(grosse_base[0])

info_transport = pd.DataFrame ({'n°trajet' : {0 : initialisation_dataframe[0]},
                                'Date' : {0 : initialisation_dataframe[1]},
                                'Camion' : {0 : initialisation_dataframe[2]},
                                'ville' : {0 : initialisation_dataframe[3]},
                                'Code postal' : {0 : initialisation_dataframe[4]},
                                'Durée trajet' : {0 : initialisation_dataframe[5]},
                                'temps trajet après pause' : {0 : initialisation_dataframe[6]},
                                'Moyenne' : {0 : initialisation_dataframe[7]},
                                'Moyenne_avant' : {0 : initialisation_dataframe[8]},
                                'Réglage' : {0 : initialisation_dataframe[9]},    
                                'T_h' : {0 : initialisation_dataframe[10]}, 
                                'T_h_p' : {0 : initialisation_dataframe[11]},
                                'T_b' : {0 : initialisation_dataframe[12]}, 
                                'T_b_p' : {0 : initialisation_dataframe[13]},
                                'moyenne_CO2' : {0 : initialisation_dataframe[14]},
                                'CO2_h' : {0 : initialisation_dataframe[15]},
                                'CO2_h_p' : {0 : initialisation_dataframe[16]},
                                'dif_pos' : {0 : initialisation_dataframe[17]},
                                'dif_neg' : {0 : initialisation_dataframe[18]},
                                'dif_pos_pause' : {0 : initialisation_dataframe[19]},
                                'dif_neg_pause' : {0 : initialisation_dataframe[20]},
                                'Pause' : {0 : initialisation_dataframe[21]},
                                't_pause_tot' : {0 : initialisation_dataframe[22]},
                                't_pause' : {0 : initialisation_dataframe[23]}})


    # Remplir la DataFrame avec tous les autres trajet
    
start_analyse = time.process_time() # Temps pour analyser les bases

for j in range (len(grosse_base) - 1) :
    j = j + 1
    info_courante = analyse(grosse_base[j])
    info_transport.loc[j] = [info_courante[0], info_courante[1],
                             info_courante[2], info_courante[3],
                             info_courante[4], info_courante[5],
                             info_courante[6],info_courante[7],
                             info_courante[8],info_courante[9],
                             info_courante[10],info_courante[11],
                             info_courante[12],info_courante[13],
                             info_courante[14],info_courante[15],
                             info_courante[16],info_courante[17],
                             info_courante[18],info_courante[19],
                             info_courante[20],info_courante[21],
                             info_courante[22],info_courante[23]]
    if n == 50: # Graph de l'avancement de l'analyse'
        c = j/len(grosse_base)*100
        plt.plot(3,c,'ro')
        plt.show()
        n = 0
    n = n + 1

# Supression lignes inexploitable (à mettre en commentaire pour débugage)

###############################################################################

info_transport = info_transport[info_transport['Pause'] > 0]
info_transport.dropna(subset=['ville'], inplace=True)

info_transport = info_transport.reset_index(drop = True)

###############################################################################
print('Temps analyse : ', time.process_time() - start_analyse)

# Exporter sur Excel
  
info_transport.to_excel(Nom_du_fichier) 
print('DataFrame is written to Excel File successfully.')

print('Temps d éxécution : ', time.process_time() - start_execution)
#Fin