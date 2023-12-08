import pandas as pd
import time
import re

while True:
    def extract_question(x):
        if 'Question' in x:
            return re.findall(r'Question:(.*?)]', x)
        else:
            return x
        
    def extract_topic(x):
        if 'Question' in x:
            if 'Topic' in x:
                return re.findall(r'Topic:(.*)Question', x)
            else:
                return re.findall(r'(.*)Question', x)
        else:
            return x
        
    def extract_topic2(x):
        if 'Question' in x:
            if 'img src' in x:
                return re.findall(r'>(.*)Question', x)
            else:
                return re.findall(r'(.*)Question', x)
        else:
            return x
        
    #importo il file excel
    path='data/dataset_indagine_stati_animo.xlsx';
    domande=pd.read_excel(path, sheet_name= 'Struttura dei dati', skiprows=5, nrows=109 );
    #creo due nuove variabili con le parti
    domande['Question'] = domande['Etichetta'].apply(lambda x: extract_question(x));
    domande['Topics'] = domande['Etichetta'].apply(lambda x: extract_topic(x));
    #elimino le variabili che non mi servono
    domande=domande.drop(['Posizione', 'Etichetta'], axis=1);

    #importo il secondo file excel
    statidanimo=pd.read_excel(path, sheet_name= 'Labels');
    #faccio un pivot per poter fare report più facilmente con tableau
    statidanimo=pd.melt(statidanimo, id_vars=['ID_Rispondente'],var_name='id_domande', value_name='repliche');
    #faccio una join tra le due tabelle
    statidanimo=statidanimo.merge(domande, how='left', left_on='id_domande', right_on='Variabile');
    #elimino variabile che non mi serve più per rendere la tabella più snella
    statidanimo=statidanimo.drop('Variabile', axis=1);
    #scarico il csv
    statidanimo.to_csv('datacleaned/statidanimo.csv', index=False);

    #mi creo la seconda tabella che mi serve con gli stessi passaggi 
    path2='data/dataset_profilo_intervistati.xlsx';
    domande2=pd.read_excel(path2, sheet_name= 'Struttura dei dati', skiprows=5, nrows=294 );
    domande2['Question'] = domande2['Etichetta'].apply(lambda x: extract_question(x));
    domande2['Topics'] = domande2['Etichetta'].apply(lambda x: extract_topic2(x));
    domande2=domande2.drop(['Posizione', 'Etichetta'], axis=1);
    profilointervistati=pd.read_excel(path2, sheet_name= 'Labels' );
    profilointervistati=pd.melt(profilointervistati, id_vars=['ID_Rispondente'],var_name='id_domande', value_name='repliche');
    profilointervistati=profilointervistati.merge(domande2, how='left', left_on='id_domande', right_on='Variabile');
    profilointervistati=profilointervistati.drop('Variabile', axis=1);
    profilointervistati.to_csv('datacleaned/profilointervistati.csv', index=False);

    time.sleep(60*60*24*7);