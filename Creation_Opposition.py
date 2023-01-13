import csv
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, messagebox as msg, ttk
from tkinter.messagebox import showinfo

import pandas as pd
import pyexcel_ods3 as pe
from PIL import Image, ImageTk
from pandas.io.formats import info
from pandas_ods_reader import read_ods
from pandastable import Table, config
from pyexcel_ods import save_data
from pynput.keyboard import Controller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager

keyboard = Controller()


# Fonction pour retrouver le chemin d'accès
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


# Procédure pour la consultation des oppositions
def main():
    # Délai entre opérations automate. Pour des numéros non entiers il faut utiliser le point pas la virgule
    delay = 1

    # ########################################

    # ##Saisie nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier:
    numeroDossier = EnterTable6.get()

    wd_options = Options()
    # wd_options.headless = True

    wd_options.set_preference('detach', True)
    wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)

    wd.get(
        'http://medoc.ia.dgfip:8141/medocweb/presentation/md2oagt/ouverturesessionagent/ecran/ecOuvertureSessionAgent'
        '.jsf')

    ## Saisir utilisateur
    time.sleep(delay)
    # wd.find_element(By.ID, 'identifiant').send_keys(login)
    wd.find_element(By.ID, 'identifiant').send_keys("youssef.atigui")

    ## Saisie mot de pass
    time.sleep(delay)
    # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
    wd.find_element(By.ID, 'secret_tmp').send_keys("1")

    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)

    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))

    ## Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')
    time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    ## Saisir habilitation
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys('1')
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)

    ## Boucle sur le fichier selon le nombre de lignes indiquées

    ## Création d'un Redevable
    ## Arriver à la transaction 3-17

    time.sleep(delay)
    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3')
    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
    time.sleep(delay)
    WebDriverWait(wd, 20).until(
        EC.presence_of_element_located((By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee')))
    wd.find_element(By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee').click()

    ## Saisie numéro de Dossier
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))

    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numeroDossier)
    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)

    # try:
    #     WebDriverWait(wd, 20).until(EC.text_to_be_present_in_element((By.CLASS_NAME, 'ui-messages-error-summary'),
    #                                                                  'DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - '
    #                                                                  'ATTENTE OU ABANDON -                 '))
    #     showinfo("Affichage opposition", "Le dossier de l'opposant " + numeroDossier + " est déjà utilisé et doit être "
    #                                                                                    "purgé ")
    #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
    #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
    #     wd.quit()
    # except FileNotFoundError as e:
    #     print(e)
    #     msg.showerror('Error in opening file', e)
    # finally:
    #     pass
    # if EC.text_to_be_present_in_element((By.CLASS_NAME, 'ui-messages-error-summary'),
    #                                                                 'DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - '
    #                                                                 'ATTENTE OU ABANDON -                 '):
    #     showinfo("Affichage opposition", "Le dossier de l'opposant " + numeroDossier + " est déjà utilisé et doit être "
    #                                                                                    "purgé ")
    # else:
    #     pass

    ## Saisie du choix Lister
    time.sleep(delay)
    time.sleep(delay)
    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))

    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('I')
    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)

    ## Récupération du tableau des oppositions
    time.sleep(delay)
    time.sleep(delay)
    time.sleep(delay)

    currentUrl = wd.current_url
    compteur = 0
    k = 0
    while True:
        if currentUrl == 'http://medoc.ia.dgfip:8141/medocweb/presentation/transactions/redevable/pa33g/ecran' \
                         '/Pa33GTx317.jsf':
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            try:
                globals()[f"webtable_df{k}"] = pd.read_html(
                    wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[
                    1]
                webtable_df1 = \
                    pd.read_html(
                        wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[
                        1]
                time.sleep(delay)
                time.sleep(delay)
                time.sleep(delay)
                # dataTable = pd.concat([globals()[f"webtable_df{k}"]], ignore_index=True)
                time.sleep(delay)
                time.sleep(delay)
                time.sleep(delay)
                # print("dataTable au tour" + str(k) + ": \n", globals()[f"webtable_df{k}"])
                last_index = globals()[f"webtable_df{k}"]['Unnamed: 0']
                index_list = globals()[f"webtable_df{k}"]['Unnamed: 0'].isnull().values.any()
                if index_list:
                    print(index_list)
                    print(last_index)
                    print(globals()[f"webtable_df{k}"].dtypes)
                    print("fin: " + str(compteur))
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    dataTable = pd.DataFrame()
                    for k in range(0, compteur + 1):
                        time.sleep(delay)
                        time.sleep(delay)
                        dataTable = dataTable.append([globals()[f"webtable_df{k}"]], ignore_index=True)
                        # print("dataTable au tour" + str(k + 1) + ": \n", globals()[f"webtable_df{k}"])
                        time.sleep(delay)
                        time.sleep(delay)
                    print("dataTable à la fin: \n", dataTable)
                    indice = pd.to_numeric(dataTable['Unnamed: 0']).fillna(0).astype(int)
                    FRP = pd.to_numeric(dataTable['Unnamed: 1']).fillna(0).astype(int)
                    name = dataTable['Unnamed: 2']
                    credit = pd.to_numeric(dataTable['Unnamed: 3']).fillna(0)
                    montant = dataTable['Unnamed: 10']
                    levee = dataTable['Unnamed: 16']
                    fields = {'id': indice, 'FRP': FRP, 'DENOMINATION': name, ' CREDIT D\'IMPOT': credit,
                              'Montant': montant,
                              'LEVEE': levee}
                    table = pd.DataFrame(fields)
                    filename = EnterTable6.get() + '_liste_créances_' + datetime.now().strftime(
                        '%Y-%m-%d-%H-%M-%S') + '.csv'
                    table.to_csv(filename, columns=fields, index=FALSE)

                    try:
                        time.sleep(delay)
                        time.sleep(delay)
                        time.sleep(delay)
                        liste = csv.reader(open(filename), delimiter=',')
                        tabControl.add(tab3, text='liste des oppositions')
                        table1 = Table(tab3, dataframe=table, read_only=True, index=FALSE)
                        table1.place(y=120)
                        table1.show()

                    except FileNotFoundError as e:
                        print(e)
                        msg.showerror('Error in opening file', e)
                    ## Validation de la sortie du formulaire
                    time.sleep(delay)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                else:
                    # time.sleep(delay)
                    time.sleep(delay)
                    time.sleep(delay)
                    if not globals()[f"webtable_df{k}"].empty:
                        dataTable = pd.concat([globals()[f"webtable_df{k}"]], ignore_index=True)
                        wd.find_element(By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi').send_keys('S')
                        wd.find_element(By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi').send_keys(Keys.ENTER)
                        print("Avant: " + str(compteur))
                        compteur += 1
                        k += 1
                        print("après: " + str(compteur))
                    else:
                        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                        wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                        wd.quit()
                        showinfo("Affichage opposition",
                                 "L'opposant " + numeroDossier + " n'a pas d'opposition en cours ")
            except:
                WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                wd.quit()
                showinfo("Affichage opposition", "L'opposant " + numeroDossier + " n'a pas d'opposition en cours ")
                break

    # finally:
    #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
    #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
    #     wd.quit()
    #     showinfo("Affichage opposition", "L'opposant " + numeroDossier + " n'a pas d'opposition en cours ")


def create_opposant():
    delay = 1

    ##Prend la ligne du fichier depuis laquelle commencer à lire
    while True:
        line = EnterTable2.get()
        if line.isnumeric():  ##vérifie que ça soit un numéro
            line = int(line)  ##ajuste l'indice
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()

    ##Combien de lignes du fichier traiter
    while True:
        line_amount = EnterTable3.get()
        if line_amount.isnumeric():
            line_amount = int(line_amount)
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()

    ## Prend les données depuis le fichier, crée une liste de listes (ou "array"), oú chaque liste est
    ## une ligne du fichier Calc. Il faut faire ça parce que pyxcel_ods prend les données sous forme
    ## de dictionaire.
    donnees_creation_opposition = pe.get_data(File_path)
    donnees_creation_opposition_sortie = pe.get_data(File_path)
    data = [i for i in donnees_creation_opposition['Feuille1']]

    # Condition qui vérifie que chaque cellule de la colonne rib, à part le header, est vide,
    # d'après le besoin case vide = rang 1, si l'item correspondant au rang est vide il prend la valeur "1" utilisable dans
    # la boucle d'automatisation. Cette condition sert à s'assurer que l'on aura une valeur pour le rang, s'il n'y a
    # pas de valeur la liste est vide et ça génère une erreur
    # taille_data donne le nombre d'items+1 dans le dico, puisque python boucle à partir de 0,
    #  dans notre cas c'est le nombre de listes, qui est de 11 ( 10 + liste headers)
    # C'est pour cela que je boucle de 0 à taille_data - 2 pour ne pas inclure la liste des headers
    taille_data = len(data)
    last_item_index0 = len(data[0]) - 1
    last_item_index1 = len(data[1]) - 1
    for i in range(taille_data - 2):
        if last_item_index0 != len(data[i + 1]) - 1:
            data[i + 1].append(str(1))
    #########################################

    # while True:
    #     montant_Creance = EnterTable8.get()
    #     if montant_Creance.isnumeric():  ##vérifie que ça soit un numéro
    #         montant_Creance = int(montant_Creance)
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()

    # while True:
    #     jour_d_effet = EnterTable9.get().split('/')[0]
    #     if jour_d_effet.isnumeric():  ##vérifie que ça soit un numéro
    #         jour_d_effet = jour_d_effet
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()
    #
    # while True:
    #     mois_d_effet = EnterTable9.get().split('/')[1]
    #     if mois_d_effet.isnumeric():  ##vérifie que ça soit un numéro
    #         mois_d_effet = mois_d_effet
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()

    # while True:
    #     annee_d_effet = EnterTable9.get().split('/')[2]
    #     if annee_d_effet.isnumeric():  ##vérifie que ça soit un numéro
    #         annee_d_effet = annee_d_effet
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()
    ## Saisie du nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier:
    # numeroDossier = EnterTable6.get()

    ## Saisie de la référence de jugement:
    # reference_de_jugement = EnterTable10.get()

    wd_options = Options()
    # wd_options.headless = True

    wd_options.set_preference('detach', True)
    wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)
    wd.get(
        'https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb%2Fcas%2Fvalidation')

    ## Saisir utilisateur
    time.sleep(delay)
    # wd.find_element(By.ID, 'identifiant').send_keys(login)
    wd.find_element(By.ID, 'identifiant').send_keys("youssef.atigui")

    ## Saisie mot de pass
    time.sleep(delay)
    # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
    wd.find_element(By.ID, 'secret_tmp').send_keys("1")

    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)

    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))

    ## Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')
    time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    ## Saisir habilitation
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys('1')
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)

    ## Boucle sur le fichier selon le nombre de lignes indiquées
    for i in range(line_amount):
        ## Création d'un Redevable
        ## Arriver à la transactionv 3-17

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3')
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee')))
        wd.find_element(By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee').click()

        ## Saisie numéro de Dossier
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))

        # wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numeroDossier)
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[line][0])
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)
        print(data[line][0])
        ## Saisie du choix Créer
        time.sleep(delay)
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))

        wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('C')
        wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)

        ## Saisie du numéro de dossier créancier
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
        # wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numero_creancier_opposant)
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[line][1])
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.TAB)
        print(data[line][1])

        ## Saisie de la suite
        time.sleep(delay)
        time.sleep(delay)
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33gsuitYa33G002ReponseSuite')))
        wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys('S')
        wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys(Keys.TAB)

        ## SAISIE DES REFERENCES DE L'OPPOSITION
        ## Transport de créance
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys('N')
        wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.TAB)

        ## Saisie ATD

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys('O')
        wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.TAB)

        ## Saisie crédit

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys('N')
        wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.TAB)

        ## Saisie Empêchement

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys('N')
        wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.TAB)

        ## Saisie Montant

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(data[line][2])
        wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.TAB)
        print(data[line][2])

        ## Saisie Date d'Effet

        date_d_effet = data[line][3]
        print(date_d_effet.day)
        # jour_d_effet = date_d_effet[0]
        # mois_d_effet = date_d_effet[1]
        # annee_d_effet = date_d_effet[2]
        # print("jour : " + jour_d_effet + " mois : " + mois_d_effet + " année : " + annee_d_effet)

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(date_d_effet.day)
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.TAB)

        ## Saisie Mois d'Effet

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(date_d_effet.month)
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.TAB)

        ## Saisie Année d'Effet

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(date_d_effet.year)
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.TAB)

        ## Saisie de la référence de jugement

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[line][4])
        wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.TAB)
        print(data[line][4])
        ## Saisie de la date d'exécution de jugement

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.TAB)

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.TAB)

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.TAB)

        ## Saisie de la date de renouvellement

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.TAB)

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.TAB)

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)

        ## Validation de la non saisie des dates

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
        wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
        wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)

        ## Validation de la suite

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
        wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
        wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys(Keys.TAB)

        ## Validation de la saisie de l'opposition

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
        wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys('O')
        wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys(Keys.TAB)

        ## Vérification du message

        # time.sleep(delay)
        # WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-info-summary')))
        # message_de_succes = wd.find_element(By.CLASS_NAME, 'ui-messages-info-summary')
        # assert (message_de_succes == "MESSAGE ACQUITTE")
        # print("tout est bon !")
        # time.sleep(delay)

        ## Saisie de la fin de saisie

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition')))
        wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys('N')
        wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys(Keys.TAB)

        ## Validation de la sortie du formulaire

        time.sleep(delay)
        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        wd.find_element(By.ID, 'barre_outils:touche_f2').click()

        ## Marquage tâche faîte dans le fichier
        donnees_creation_opposition_sortie['Feuille1'][line].append('X')
        line += 1

        filename = 'donnees_creation_opposition_sortie' + datetime.now().strftime('_%Y-%m-%d') + '.ods'

        save_data(filename, donnees_creation_opposition_sortie)
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            sheet = "Feuille1"
            table = read_ods(filename, sheet)
            tabControl.add(tab4, text='liste des oppositions')
            table1 = Table(tab4, dataframe=table, read_only=True, index=FALSE)
            table1.place(y=120)

            options = {'colheadercolor': 'green', 'floatprecision': 0}
            config.apply_options(options, table1)
            table1.autoResizeColumns()
            table1.show()

        except FileNotFoundError as e:
            print(e)
            msg.showerror('Error in opening file', e)
    wd.quit()


# Procédure pour
def open_file():
    global File_path
    global l1
    file = filedialog.askopenfile(mode='r', filetypes=[('Ods Files', '*.ods')])
    if file:
        filepath = os.path.abspath(file.name)
        filepath = filepath.replace(os.sep, "/")
        label_path.configure(text="Le fichier sélectionné est : " + Path(filepath).stem)
        File_path = filepath


# Procédure pour la gestion de l'interface Tkinter
Interface = Tk()
Interface.geometry('1000x600')
Interface.title('SATD DGE')
paramx = 10
paramy = 170

tabControl = ttk.Notebook(Interface)
tab1 = Frame(tabControl, bg='#C7DDC5')
label1 = Label(tab1, text='Afficher la liste des oppositions', font=('Arial', 15), fg='Black', bg='#ffffff',
               relief="sunken")
label1.place(x=400, y=paramx)

creancierButton = Button(tab1, text='Afficher la liste', command=main)
creancierButton.place(x=paramx + 240, y=paramy + 40)

tab2 = Frame(tabControl, bg='#E3EBD0')
label2 = Label(tab2, text='Créer des oppositions', font=('Arial', 15), fg='Black', bg='#ffffff', relief="sunken")
label2.place(x=400, y=paramx)
tabControl.add(tab1, text='Liste des oppositions')
tabControl.add(tab2, text='Création des oppositions')
tabControl.pack(expand=1, fill="both")

tab3 = Frame(tabControl, bg='#E3EBD0')
tab4 = Frame(tabControl, bg='#E3EBD0')

# Etablissement de l'image de fermeture
img = Image.open('C:/Users/meddb-jean-francoi01/Documents/Application de Creation d\'Opposant/close-button.png')
img_resize = img.resize((30, 30), Image.LANCZOS)
closeIcon = ImageTk.PhotoImage(img_resize)
closeButton1 = Button(Interface, image=closeIcon, command=lambda: tabControl.forget(tab3))
closeButton1.pack(side=LEFT)
closeButton2 = Button(Interface, image=closeIcon, command=lambda: tabControl.forget(tab4))
closeButton2.pack(side=LEFT)

EnterTable1 = StringVar()
EnterTable2 = StringVar()
EnterTable3 = StringVar()
EnterTable4 = StringVar()
EnterTable5 = StringVar()
EnterTable6 = StringVar()
EnterTable7 = StringVar()
EnterTable8 = StringVar()
EnterTable9 = StringVar()
EnterTable10 = StringVar()

labelNumeroDossier = Label(tab1, text='Numéro Dossier Opposant:', relief="sunken")
labelNumeroDossier.place(x=250, y=paramy - 30)
entryNumeroDossier = Entry(tab1, textvariable=EnterTable6, justify='center')
entryNumeroDossier.place(width=225, x=paramx + 490, y=paramy - 30)

creerOpposition = Button(tab2, text='Créer les Oppositions', relief="ridge", command=create_opposant)
creerOpposition.place(x=paramx + 240, y=paramy + 300)

# labelNumeroDossierCreancierOpposant = Label(tab2, text="Saisir le numéro d\'un créancier opposant :")
# labelNumeroDossierCreancierOpposant.place(x=paramx + 250, y=paramy + 20)
# NumeroDossierCreancierOpposant = Entry(tab2, textvariable=EnterTable7, justify='center')
# NumeroDossierCreancierOpposant.place(x=paramx + 500, y=paramy + 20)
#
# labelMontantCreance = Label(tab2, text="Saisir le montant de la créance :")
# labelMontantCreance.place(x=paramx + 250, y=paramy + 45)
# montantCreance = Entry(tab2, textvariable=EnterTable8, justify='center')
# montantCreance.place(x=paramx + 500, y=paramy + 45)

# labelDateEffet = Label(tab2, text="Saisir la date d'effet :")
# labelDateEffet.place(x=paramx + 250, y=paramy + 70)

# now = datetime.today()
# date_d_effet = DateEntry(tab2, selectmode='day', textvariable=EnterTable9, locale='fr_FR', year=now.year,
#                          month=now.month, day=now.day)
# date_d_effet.place(x=paramx + 500, y=paramy + 70)

# label_reference_de_jugement = Label(tab2, text="Référence jugement Validité :")
# label_reference_de_jugement.place(x=paramx + 250, y=paramy + 100)
# reference_de_jugement = Entry(tab2, textvariable=EnterTable10, justify='center')
# reference_de_jugement.place(x=paramx + 500, y=paramy + 100)


# Methode pour la mise à jour de la saisie de la date en manuellement
# def my_upd(i):
#     i: int
#     l1 = Label(tab2, bg='yellow')
#     l1.config(text=EnterTable9.get().split('/')[i])
#     l1.place(x=paramx + 650 + i * 20, y=paramy + 70)
#
#
# for i in [0, 1, 2]:
#     EnterTable9.trace('w', my_upd(i))

# label2 = Label(tab1, text='Saisir le délai entre les opérations de l\'automate en secondes :',relief="sunken")
# label2.place(x=paramx + 250, y=paramy + 120)
# entry1 = Entry(tab1, textvariable=EnterTable1, justify='center')
# entry1.place(x=paramx + 600, y=paramy + 120)
# label3 = Label(tab1, text='Saisir la ligne du début: ',relief="sunken")
# label3.place(x=paramx + 250, y=paramy + 155)
# entry2 = Entry(tab1, textvariable=EnterTable2, justify='center')
# entry2.place(x=paramx + 600, y=paramy + 155)
# label4 = Label(tab1, text='Saisir le nombre de lignes à traiter: ',relief="sunken")
# label4.place(x=paramx + 250, y=paramy + 185)
# entry3 = Entry(tab1, textvariable=EnterTable3, justify='center')
# entry3.place(x=paramx + 600, y=paramy + 185)

# label2 = Label(tab2, text='Saisir le délai entre les opérations de l\'automate en secondes :', relief="sunken")
# label2.place(x=paramx + 250, y=paramy + 150)
# entry1 = Entry(tab2, textvariable=EnterTable1, justify='center')
# entry1.place(x=paramx + 600, y=paramy + 150)
label3 = Label(tab2, text='Saisir la ligne du début: ', relief="sunken")
label3.place(x=paramx + 240, y=paramy + 45)
entry2 = Entry(tab2, textvariable=EnterTable2, justify='center')
entry2.place(width=225, x=paramx + 490, y=paramy + 45)
label4 = Label(tab2, text='Saisir le nombre de lignes à traiter: ', relief="sunken")
label4.place(x=paramx + 240, y=paramy + 105)
entry3 = Entry(tab2, textvariable=EnterTable3, justify='center')
entry3.place(width=225, x=paramx + 490, y=paramy + 105)

# login et mot de passe
label5 = Label(tab1, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab1, textvariable=EnterTable4, justify='center')
entry4.place(x=340, y=70)
label6 = Label(tab1, text='Mot de passe: ', relief="sunken")
label6.place(x=500, y=70)
entry5 = Entry(tab1, textvariable=EnterTable5, justify='center')
entry5.place(x=600, y=70)

label5 = Label(tab2, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab2, textvariable=EnterTable4, justify='center')
entry4.place(x=340, y=70)
label6 = Label(tab2, text='Mot de passe: ', relief="sunken")
label6.place(x=500, y=70)
entry5 = Entry(tab2, textvariable=EnterTable5, justify='center')
entry5.place(x=600, y=70)

button2 = Button(tab2, text='Choisir le fichier d\'entrée', command=open_file)
button2.place(x=paramx + 240, y=paramy - 30)
label_path = Label(tab2)
label_path.place(x=paramx + 490, y=paramy - 30)
# button1 = Button(Interface, text='Lancer le programme', command=main)
# button1.place(x=350, y=560)
# QUIT = Button(Interface, text='Quitter', fg='Red', command=Interface.destroy)
# QUIT.place(x=550, y=560)

Interface.mainloop()
