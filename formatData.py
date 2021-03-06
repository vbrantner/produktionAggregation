#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
sys.path.append('~/Documents/coding/bbn/produktionAggregation/')

import pandas as pd
from pathlib import Path
from os import listdir
from os.path import isfile, join
from sendMail import sendMail
from createDF import saveDF
import pickle

def main():
    files18 = [f for f in listdir(Path('./data/18/')) if f.endswith(".xlsx") if isfile(join(Path('./data/18'), f))]
    savedData = Path('savedData.txt') 
    # create savedData.txt if not exist
    if savedData.is_file():
      with open("savedData.txt", "rb") as fp:
          existing_files18 = pickle.load(fp)
    else:
      with open('savedData.txt', 'wb') as fp:
        pickle.dump([], fp)
    # check if there are new files to add
    if files18 == existing_files18:
        print('There are no new files to add')
    else:
        print("there are files to add...")
        setNewFiles = set(files18)
        setExistingFiles = set(existing_files18)
        unmatched = setNewFiles.symmetric_difference(setExistingFiles)
        print('File with name {} wil be added to the df'.format(unmatched))
        addFiles(True, list(unmatched)[0])
        saveDF()
        sendMail('v.brantner@brantner.de', 'formatting df', 'formatting df completed')
        with open('savedData.txt', 'wb') as fp:
            pickle.dump(files18, fp)

# rewrite, first check which files already in the df
def addFiles(existingDF, fileToAdd=""):
  if existingDF == True:
    df = pd.read_csv("~/Documents/coding/bbn/produktionAggregation/data/dataframe.csv")
    print(fileToAdd)
    dfToAdd = pd.read_excel(Path('./data/18/' + fileToAdd))
    df = df.append(dfToAdd)
  else:
    first_df = pd.read_excel(Path('./data/16/data_16_10-12.xlsx'))
    files17 = [f for f in listdir(Path('./data/17/')) if f.endswith(".xlsx") if isfile(join(Path('./data/17'), f))]
    files18 = [f for f in listdir(Path('./data/18/')) if f.endswith(".xlsx") if isfile(join(Path('./data/18'), f))]
    df = pd.read_excel(Path('./data/17/'+files17[0]))
    df2 = pd.read_excel(Path('./data/17/'+files17[1]))
    df = first_df.append(df)
    df = df.append(df2)
    for indexFile in range(0, len(files18)):
      df = df.append(pd.read_excel(Path('./data/18/'+files18[indexFile])))
      print('added {}'.format(files18[indexFile]))

  df = df.rename(index=str, columns={
    'Kunde.Nummer':'KundenNr',
    'Kunde.Name1':'KundenName', 
    'Bezeichnung 1':'ArtBezeichnung',
    'Lief.': 'GelieferteMenge',
    'Ret.':'Retoure',
    'VK-Menge': 'VerkaufteMenge',
    'Ges.-Menge': 'GesamteMenge',
    'Diff.': 'Differenz',
    'Kor.' : 'Korrektur',
    'LVKP 1 Brutto': 'LVKPBrutto',
    'Back-Gewicht': 'BackGewicht'})

  # prepare the data
  # df = df.drop(['Diff.-Wert', 'VK-Wert', 'VK[Netto]', 'LVKP 1  [Netto]', 'VK[Bew.Preis]', 'Retoure', 'VerkaufteMenge', 'GesamteMenge', 'Differenz', 'LVKPBrutto'], axis=1)
  df = df.drop(['Diff.-Wert', 'VK[Netto]', 'LVKP 1  [Netto]', 'VK[Bew.Preis]', 'GesamteMenge', 'Differenz', 'LVKPBrutto'], axis=1)
  df = df.fillna(0) # replace nan with 0
  df['Datum'] = df['Datum'].apply(lambda x: '{}-{}-{}'.format(x[9:13],x[6:8],x[3:5])) # format date
  df['Datum'] = pd.to_datetime(df['Datum'])
  df['ArtBezeichnung'] = df['ArtBezeichnung'].apply(lambda x: str(x).replace('"', '_'))
  df['KundenName'] = df['KundenName'].apply(lambda x: str(x).replace('"', '_'))

  # change Artikelgruppe
  df['Artikelgruppe'] = df['Artikelgruppe'].apply(lambda x: str(x).replace('Mehlbrot', 'Brote'))
  df['Artikelgruppe'] = df['Artikelgruppe'].apply(lambda x: str(x).replace('Körnerbrot', 'Brote'))
  df['Artikelgruppe'] = df['Artikelgruppe'].apply(lambda x: str(x).replace('Vollkornbrote', 'Brote'))
  df['Artikelgruppe'] = df['Artikelgruppe'].apply(lambda x: str(x).replace('Vollkornbrot', 'Brote'))

  df.to_csv('./data/dataframe.csv', index=False)

if __name__ == '__main__':
  main()