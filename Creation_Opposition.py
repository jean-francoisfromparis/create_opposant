import datetime
import hashlib
import os
import sys
import time
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox
import tkinter as tk

import pandas as pd
from pynput.keyboard import Controller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tkcalendar import Calendar, DateEntry
from webdriver_manager.firefox import GeckoDriverManager

keyboard = Controller()


# Fonction pour retrouver le chemin d'accès
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def main():
    # Délai entre opérations automate. Pour des numéros non entiers il faut utiliser le point pas la virgule
    delay = 1

    # Prend la ligne du fichier depuis laquelle commencer à lire
    # while True:
    #     line = EnterTable2.get()
    #     if line.isnumeric():  ##vérifie que ça soit un numéro
    #         line = int(line)  ##ajuste l'indice
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()

    # Combien de lignes du fichier traiter
    line_amount = 1
    # while True:
    #     line_amount = EnterTable3.get()
    #     if line_amount.isnumeric():
    #         line_amount = int(line_amount)
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()

    # Prend les données depuis le fichier, crée une liste de listes (ou "array"), oú chaque liste est
    # une ligne du fichier Calc. Il faut faire ça parce que pyxcel_ods prend les données sous forme
    # de dictionnaire.
    # donnees_entree = pe.get_data()
    # data = [i for i in donnees_entree['Database']]

    # Condition qui vérifie que chaque cellule de la colonne rib, à part le header, est vide, d'après le besoin case
    # vide = rang 1, si l'item correspondant au rang est vide il prend la valeur "1" utilisable dans la boucle
    # d'automatisation. Cette condition sert à s'assurer que l'on aura une valeur pour le rang, s'il n'y a pas de
    # valeur la liste est vide et ça génère une erreur taille_data donne le nombre d'items+1 dans le dico,
    # puisque python boucle à partir de 0, dans notre cas, c'est le nombre de listes qui est de 11 (10 + liste
    # headers) C'est pour cela que je boucle de 0 à taille_data - 2 pour ne pas inclure la liste des headers
    # taille_data = len(data) last_item_index0 = len(data[0]) - 1 last_item_index1 = len(data[1]) - 1 for i in range(
    # taille_data - 2): if last_item_index0 != len(data[i + 1]) - 1: data[i + 1].append(str(1))
    # ########################################

    # ##Saisie nom utilisateur et mot de passe
    # login = pe.get_data('C:/Users/meddb-el-farouki01/Desktop/Rembursement_DGE/Programme/login.ods')['Database'][0]
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier:
    numeroDossier = EnterTable6.get()
    ## Lancement webdriver Selenium
    s = Service(resource_path("geckodriver"))
    # wd = webdriver.Firefox(service=s)
    wd = webdriver.Firefox(executable_path=GeckoDriverManager().install())
    wd_options = Options()
    wd_options.set_preference('detach', True)
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

        ## Saisie du choix Lister
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))

        wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('I')
        wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)

        ## Récupération d'un des oppositions

        time.sleep(delay)
        webtable_df1 = \
            pd.read_html(
                wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[1]

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi')))

        wd.find_element(By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi').send_keys('S')
        wd.find_element(By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi').send_keys(Keys.ENTER)

        time.sleep(delay)
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]')))
        if EC.presence_of_element_located((By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]')):
            webtable_df2 = \
                pd.read_html(
                    wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[1]
        time.sleep(delay)
        # WebDriverWait(wd, 20).until(
        #     EC.presence_of_element_located((By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]')))
        # if EC.presence_of_element_located((By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]')):
        #     webtable_df3 = \
        #         pd.read_html(
        #             wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[1]

        webtable_df = pd.concat([webtable_df1, webtable_df2])
        webtable_df.to_csv('temp_data.csv')
        # Printing the URL
        print(webtable_df)

        # wd.quit()

        ## Incrémentation des lignes du tableau
        # line += 1

    # wd.quit()


def create_opposant():
    delay = 1

    # while True:
    #     numero_creancier_opposant = EnterTable7.get()
    #     if numero_creancier_opposant.isnumeric():  ##vérifie que ça soit un numéro
    #         numero_creancier_opposant = int(numero_creancier_opposant)
    #         break
    #     else:
    #         messagebox.OK('Saisie incorrecte, réessayez')
    #         exit()

    while True:
        montant_Creance = EnterTable8.get()
        if montant_Creance.isnumeric():  ##vérifie que ça soit un numéro
            montant_Creance = int(montant_Creance)
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()

    while True:
        jour_d_effet = EnterTable9.get().split('/')[0]
        if jour_d_effet.isnumeric():  ##vérifie que ça soit un numéro
            jour_d_effet = jour_d_effet
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()

    while True:
        mois_d_effet = EnterTable9.get().split('/')[1]
        if mois_d_effet.isnumeric():  ##vérifie que ça soit un numéro
            mois_d_effet = mois_d_effet
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()

    while True:
        annee_d_effet = EnterTable9.get().split('/')[2]
        if annee_d_effet.isnumeric():  ##vérifie que ça soit un numéro
            annee_d_effet = annee_d_effet
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()
    ## Saisie du nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier:
    numeroDossier = EnterTable6.get()

    ## Saisie de la référence de jugement:
    reference_de_jugement = EnterTable10.get()

    ## Lancement webdriver Selenium
    s = Service(resource_path("geckodriver"))
    # wd = webdriver.Firefox(service=s)
    wd = webdriver.Firefox(executable_path=GeckoDriverManager().install())
    wd_options = Options()
    wd_options.set_preference('detach', True)
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
    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys('617382')
    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)

    ## Saisie du choix Créer
    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))

    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('C')
    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)

    ## Saisie du numéro de dossier créancier
    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
    # wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numero_creancier_opposant)
    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys('203268')
    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.TAB)

    ## Saisie de la suite
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
    wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(montant_Creance)
    wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.TAB)

    ## Saisie Date d'Effet

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(jour_d_effet)
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.TAB)

    ## Saisie Mois d'Effet

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(mois_d_effet)
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.TAB)

    ## Saisie Année d'Effet

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(annee_d_effet)
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.TAB)

    ## Saisie de la référence de jugement

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(reference_de_jugement)
    wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.TAB)

    ## Saisie de la date d'exécution de jugement

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.TAB)

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.TAB)

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.TAB)

    ## Saisie de la date de renouvellement

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.TAB)

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.TAB)

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)

    ## Validation de la non saisie des dates

    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
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

    # wd.quit()


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


Interface = Tk()
Interface.geometry('1000x600')
Interface.title('Création Opposition')

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

paramx = 10
paramy = 170
label1 = Label(Interface, text='Création Opposition', font=('Arial', 15), fg='Black', bg='#ffffff')
label1.place(x=400, y=1)
labelNumeroDossier = Label(Interface, text='Numéro Dossier Opposant:')
labelNumeroDossier.place(x=250, y=100)
entryNumeroDossier = Entry(Interface, textvariable=EnterTable6, justify='center')
entryNumeroDossier.place(x=paramx + 400, y=100)

creancierButton = Button(Interface, text='Afficher le créancier', command=main)
creancierButton.place(x=paramx + 250, y=paramy - 20)

creerOpposition = Button(Interface, text='Créer une Opposition', command=create_opposant)
creerOpposition.place(x=paramx + 450, y=paramy - 20)

labelNumeroDossierCreancierOpposant = Label(Interface, text="Saisir le numéro d\'un créancier opposant :")
labelNumeroDossierCreancierOpposant.place(x=paramx + 250, y=paramy + 20)
NumeroDossierCreancierOpposant = Entry(Interface, textvariable=EnterTable7, justify='center')
NumeroDossierCreancierOpposant.place(x=paramx + 500, y=paramy + 20)

labelMontantCreance = Label(Interface, text="Saisir le montant de la créance :")
labelMontantCreance.place(x=paramx + 250, y=paramy + 45)
montantCreance = Entry(Interface, textvariable=EnterTable8, justify='center')
montantCreance.place(x=paramx + 500, y=paramy + 45)

labelDateEffet = Label(Interface, text="Saisir la date d'effet :")
labelDateEffet.place(x=paramx + 250, y=paramy + 70)

now = datetime.datetime.today()
date_d_effet = DateEntry(Interface, selectmode='day', textvariable=EnterTable9, locale='fr_FR', year=now.year,
                         month=now.month, day=now.day)
date_d_effet.place(x=paramx + 500, y=paramy + 70)

label_reference_de_jugement = Label(Interface, text="Référence jugement Validité :")
label_reference_de_jugement.place(x=paramx + 250, y=paramy + 100)
reference_de_jugement = Entry(Interface, textvariable=EnterTable10, justify='center')
reference_de_jugement.place(x=paramx + 500, y=paramy + 100)

def my_upd(i):
    i: int
    l1 = Label(Interface, bg='yellow')
    l1.config(text=EnterTable9.get().split('/')[i])
    l1.place(x=paramx + 650 + i * 20, y=paramy + 70)


for i in [0, 1, 2]:
    EnterTable9.trace('w', my_upd(i))

label2 = Label(Interface, text='Saisir le délai entre les opérations de l\'automate en secondes :')
label2.place(x=paramx + 250, y=paramy + 120)
entry1 = Entry(Interface, textvariable=EnterTable1, justify='center')
entry1.place(x=paramx + 600, y=paramy + 120)
label3 = Label(Interface, text='Saisir la ligne du début: ')
label3.place(x=paramx + 250, y=paramy + 155)
entry2 = Entry(Interface, textvariable=EnterTable2, justify='center')
entry2.place(x=paramx + 600, y=paramy + 155)
label4 = Label(Interface, text='Saisir le nombre de lignes à traiter: ')
label4.place(x=paramx + 250, y=paramy + 185)
entry3 = Entry(Interface, textvariable=EnterTable3, justify='center')
entry3.place(x=paramx + 600, y=paramy + 185)

# login et mot de passe
label5 = Label(Interface, text='Login:')
label5.place(x=250, y=50)
entry4 = Entry(Interface, textvariable=EnterTable4, justify='center')
entry4.place(x=300, y=50)
label6 = Label(Interface, text='Mot de passe: ')
label6.place(x=500, y=50)
entry5 = Entry(Interface, textvariable=EnterTable5, justify='center')
entry5.place(x=600, y=50)

label_path = Label(Interface)
label_path.place(x=360, y=200)
button1 = Button(Interface, text='Lancer le programme', command=main)
button1.place(x=350, y=560)
QUIT = Button(Interface, text='Quitter', fg='Red', command=Interface.destroy)
QUIT.place(x=550, y=560)

Interface.mainloop()
