# -*- coding: utf-8 -*-

# def jarvidatan_kaanto(kansiopolku):

"""
Muokkaa kansiosta loytyvat xls ja xlsx- tiedostot vanhasta excel-muodosta
dataframe-muotoon, joka edelleen voidaan muokata LajiGISsiin syötettäväksi.

Args:
    kansio (str) : kansio, jossa muutettavat taulukot sijaitsevat.

Returns:
    Kansiosta lyötyvät taulukot yhdistyvät yhteen dataframeen, joka
    tallentuu valitun kansion alakansioon xlsx-muodossa.
"""

import glob
import pandas as pd
import xlwings as xw
import numpy as np
import time

# Define the folder 
kansio = "C:/Users/Rasmusbo/Documents/VesiKasvitKuoppala/JarviTestiKansio/"
# Define the folder where you want to save the excel files
tal_kansio = "C:/Users/Rasmusbo/Documents/VesiKasvitKuoppala/Muokatut_df_muodossa/"

# Grabs the files with .xls tai .xlsx extension
tiedostot = glob.glob(kansio + "*.xls*")


#%%
  
# Create an empty pandas df to append the data in later  

perus_df = pd.DataFrame({"jarvi_nimi" : [],           
                        "jarvi_nro2" : [],
                        "linja_tunnus" : [],
                        "pvm" : [],
                        "alkuaika" : [],
                        "loppuaika" : [],
                        "kokonaisaika" : [],
                        "tekijat" : [],
                        "nakosyvyys" : [],
                        "huomiot" : [],
                        "maxsyvyys_uposleht" : [],
                        "laji_uposlehtiset" : [],
                        "maxsyvyys_pohjleht" : [],
                        "laji_pohjleht" : [],
                        "linj_pituus" : [],
                        "linja_pinta_ala" : [],
                        
                        "alku_koord_x" : [],           
                        "alku_koord_y" : [],
                        "alkupist_maamerkki" : [],
                        "alku_pist_kuvanro" : [],
                        "loppupist_maamerkki" : [],
                        "loppupist_kuvanro" : [],
                        "linjan_suunta" : [],
                        "muut_kuvat_nro" : [],
                        "luty_lehto" : [],
                        "luty_lehtokangas" : [],
                        "luty_tuorekangas" : [],
                        "luty_kuivakangas" : [],
                        "luty_avokallio" : [],
                        "luty_rantaluhta" : [],
                        "luty_korpi" : [],
                        "luty_rame" : [],
                        "luty_neva" : [],
                        "luty_lahteisyys" : [],
                        "luty_muu_mika" : [],
                        "luty_ihmistoiminta" : [],
                        "rantapen_jyrk_loiva" : [],
                        "rantapen_jyrk_keskikalt" : [],
                        "rantapen_jyrk_pysty" : [],
                        "linjatyyp_yleislinja" : [],
                        "linjatyyp_rehev_herkka" : []
                        })
    
# dfj = xw.Book(tiedostot[tiedosto]) 

vyohykkeet = pd.DataFrame()
lajit_final = pd.DataFrame()

    
#%%
for i in range(0, len(tiedostot)):
    print(tiedostot)
#%%
# Loop through the sheets in excel to append the data into given dataframe.
# Assumption: sheets named Linja* are in standard format AND
# If B4 is empty, then there's no other data either


alkuaika = time.time()

for tiedosto in range(0, len(tiedostot)):  
    dfj = xw.Book(tiedostot[tiedosto]) # Open the workbook with xlwings
    print(dfj) 
    for vlehti in dfj.sheets: # Loop through sheets of the workbook          
            if 'Linja' in vlehti.name and vlehti.range("B4").value != None:  # Besides the real data, there's other and empty sheets in the workbook. Leave them out
                
                print(vlehti.name + " käyty läpi")
                
                assert vlehti.range("A13").value == "Laji_uposlehtiset"
                assert vlehti.range("B41").value == "Laji_id" # Check that everything is in order and cells in right places.
                
                # Append data into the dataframe defined above
                ptiedot = pd.DataFrame({"jarvi_nimi" : [vlehti.range("B2").value],           
                                        "jarvi_nro2" : [vlehti.range("B3").value],
                                        "linja_tunnus" : [vlehti.range("B4").value],
                                        "pvm" : [vlehti.range("B5").value],
                                        "alkuaika" : [vlehti.range("B6").value],
                                        "loppuaika" : [vlehti.range("B7").value],
                                        "kokonaisaika" : [vlehti.range("B8").value],
                                        "tekijat" : [vlehti.range("B9").value],
                                        "nakosyvyys" : [vlehti.range("B10").value],
                                        "huomiot" : [vlehti.range("B11").value],
                                        "maxsyvyys_uposleht" : [vlehti.range("B12").value],
                                        "laji_uposlehtiset" : [vlehti.range("B13").value],
                                        "maxsyvyys_pohjleht" : [vlehti.range("B14").value],
                                        "laji_pohjleht" : [vlehti.range("B15").value],
                                        "linj_pituus" : [vlehti.range("B16").value],
                                        "linja_pinta_ala" : [vlehti.range("B17").value],
                                        
                                        "alku_koord_x" : [vlehti.range("L2").value],           
                                        "alku_koord_y" : [vlehti.range("S2").value],
                                        "alkupist_maamerkki" : [vlehti.range("K3").value],
                                        "alku_pist_kuvanro" : [vlehti.range("X3").value],
                                        "loppupist_maamerkki" : [vlehti.range("K4").value],
                                        "loppupist_kuvanro" : [vlehti.range("X4").value],
                                        "linjan_suunta" : [vlehti.range("K5").value],
                                        "muut_kuvat_nro" : [vlehti.range("W5").value],
                                        "luty_lehto" : [vlehti.range("O8").value],
                                        "luty_lehtokangas" : [vlehti.range("O9").value],
                                        "luty_tuorekangas" : [vlehti.range("O10").value],
                                        "luty_kuivakangas" : [vlehti.range("O11").value],
                                        "luty_avokallio" : [vlehti.range("O12").value],
                                        "luty_rantaluhta" : [vlehti.range("O13").value],
                                        "luty_korpi" : [vlehti.range("O14").value],
                                        "luty_rame" : [vlehti.range("X8").value],
                                        "luty_neva" : [vlehti.range("X9").value],
                                        "luty_lahteisyys" : [vlehti.range("X10").value],
                                        "luty_muu_mika" : [vlehti.range("X11").value],
                                        "luty_ihmistoiminta" : [vlehti.range("X12").value],
                                        "rantapen_jyrk_loiva" : [vlehti.range("O15").value],
                                        "rantapen_jyrk_keskikalt" : [vlehti.range("U15").value],
                                        "rantapen_jyrk_pysty" : [vlehti.range("X15").value],
                                        "linjatyyp_yleislinja" : [vlehti.range("O16").value],
                                        "linjatyyp_rehev_herkka" : [vlehti.range("O17").value]
                                        })
    
                pohjat = vlehti.range('A26:D39').options(pd.DataFrame, transpose = True).value # Transpose the data
                pohjat = pohjat.dropna(how = "all") # Drop empty values
                perus_df = perus_df.append(ptiedot, ignore_index=False) #
                
                # Vyohyketiedot as separate df
                
                vyoh_df = vlehti.range('B18:V25').options(pd.DataFrame, transpose = True).value
                vyoh_filt = vyoh_df[vyoh_df.index.notnull()]
                
                id_for_vyoh = pd.DataFrame({'jarvi_nro2' : np.tile([vlehti.range("B3").value], len(vyoh_filt)),
                                            'linja_tunnus' : np.tile([vlehti.range('B4').value], len(vyoh_filt))
                                            })
    
                vyohykkeet_df = pd.concat([id_for_vyoh, vyoh_filt.reset_index(drop=False)], axis=1)
                vyohykkeet = vyohykkeet.append(vyohykkeet_df, ignore_index = False)
                
                # And finally species data as separate df
                Lajit = vlehti.range('A41:E350').options(pd.DataFrame).value
                Lajit.dropna(subset = ['Y', 'P'], inplace = True) # Remove variables with no species data in Y and P columns
                
                id_for_lajit = pd.DataFrame({'jarvi_nro2' : np.tile([vlehti.range("B3").value], len(Lajit)), 
                                             'linja_tunnus' : np.tile([vlehti.range('B4').value], len(Lajit))
                                             })
    
                lajit_df = pd.concat([id_for_lajit, Lajit.reset_index(drop = False)], axis = 1)
                lajit_final = lajit_final.append(lajit_df, ignore_index = True)

# Close the excel files
xw.apps.active.quit()
    

loppuaika = time.time()
running_time = loppuaika - alkuaika

print("Aikaa kului: " + str(running_time) + " sekuntia")

#%%

# Write these to excel files. These could be combined and/or modified also here.
perus_df.to_excel(tal_kansio + "perustiedot.xlsx")
vyohykkeet.to_excel(tal_kansio + "vyohyketiedot.xlsx")
lajit_final.to_excel(tal_kansio + "lajidata.xlsx")



