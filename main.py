
import pandas as pd
import warnings
from tkinter import messagebox
from datetime import datetime
import os

warnings.filterwarnings('ignore')

link = r'C:\Users\ChristophRazek\Emea\06_Qualitymanagement - Dokumente\01_QS\04_MPS\Vorabmuster.xlsx'

#Zeitstempel Letzte Änderung!
c_time = os.path.getmtime(link)
dt_c = datetime.fromtimestamp(c_time)

#Kopieren von MPS File zu Laufwerk
bestellungen = pd.read_excel(link)


bestellungen[['Fixposnr','Belegart', 'Belegnr']] = bestellungen[['Fixposnr','Belegart', 'Belegnr']].fillna(0).astype('int64')
bestellungen['Vorabmuster'].replace({'Ja': 1, 'Nein': 0}, inplace=True)
bestellungen['Freigabe'].replace({'Ja': 1, 'Nein': 0}, inplace=True)

bestellungen.to_csv(r'L:\Q\Vorabmuster.csv', sep=';', index=False)

#Log File
with open(r'S:\EMEA\Kontrollabfragen\Vorabmuster.txt', 'w') as f:
    f.write(f'Last MPS copied at: {dt_c}')
    f.close()


messagebox.showinfo('Update Erfolgreich!', f'Das Vorabmuster-Update mit letzter Änderung vom {dt_c} wurde erfolgreich durchgeführt.')