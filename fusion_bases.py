import pandas as pd 

############################################################################################
################################### Variable changeable ####################################
############################################################################################

Heering_data = pd.read_excel('C://Users/p36200223/Documents/Test/Base_heering_python.xlsx')
Kerhis_data = pd.read_excel('C://Users/p36200223/Documents/2022-03-04_transport_c22_2021.xls')
Visio_data = pd.read_excel('C://Users/p36200223/Documents/all_data_per_breeder_and_flock_pierre_2022-06-10.xlsx')
Nom_du_fichier = 'Fusion_heering_visio.xlsx'

############################################################################################
###################################### Ne pas toucher ######################################
############################################################################################

Kerhis_data['Code postal'] = pd.to_numeric(Kerhis_data['Code postal'].str.replace(' ','',regex = False),errors = 'coerce') #pd.to_numeric(...,errors = 'coerce')
Kerhis_data['Date only'] = Kerhis_data['Date livraison'].dt.date #pd.to_datetime(...)

Heering_data['Code postal'] = pd.to_numeric(Heering_data['Code postal'],errors = 'coerce')
Heering_data['Date'] = pd.to_datetime(Heering_data['Date']).dt.date

dataframe_join = pd.merge(
    left = Heering_data,
    right = Kerhis_data,
    left_on =  ['Date','Code postal','Camion'],
    right_on = ['Date only','Code postal','Véhicule'],
    how = 'inner') 

dataframe_join["N° INUAV"]= dataframe_join["N° INUAV"].str.lower() 

#inner(intersection) / outer(tous) /
#validate = True pour enlever les doublons

dataframe_final = pd.merge(
    left = dataframe_join,
    right = Visio_data,
    left_on =  ['N° INUAV'],
    right_on = ['house_id'],
    how = 'inner', 
    indicator = True)

dataframe_join.to_excel('Fusion_heering_kerhis.xlsx') 
dataframe_final.to_excel(Nom_du_fichier) 

print('Fusion is written to Excel file successfully ')

#temp haute en fonction de mortalité

#introduction de 10 min à pandas (vidéo) https://pandas.pydata.org/pandas-docs/stable/user_guide/10min.html
