
import csv
import os
from queue import Empty
import time
import sys
import pyexcel_ods3 as pe

from pynput.keyboard import Controller, Key
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from pathlib import Path
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.service import Service


keyboard = Controller()

# Fonction pour retrouver le chemin d'accès
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def main():

    
  

    ##Délai entre opérations automate. Pour des numéros non entiers il faut utiliser le point pas la virgule
    while True:
        try:
            delay = float(EnterTable1.get())
            break
        except ValueError:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()


    ##Prend la ligne du fichier depuis laquelle commencer à lire 
    while True:
        line = EnterTable2.get()
        if line.isnumeric():  ##vérifie que ça soit un numéro
            line = int(line) - 1  ##ajuste l'indice
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
    donnees_entree = pe.get_data(File_path)
    data = [i for i in donnees_entree['Feuille1']]

    
    # Condition qui vérifie que chaque cellule de la colonne rib, à part le header, est vide, 
    # d'après le besoin case vide = rang 1, si l'item correspondant au rang est vide il prend la valeur "1" utilisable dans
    # la boucle d'automatisation. Cette condition sert à s'assurer que l'on aura une valeur pour le rang, s'il n'y a
    # pas de valeur la liste est vide et ça génère une erreur
    # taille_data donne le nombre d'items+1 dans le dico, puisque python boucle à partir de 0,
    #  dans notre cas c'est le nombre de listes, qui est de 11 ( 10 + liste headers)
    #C'est pour cela que je boucle de 0 à taille_data - 2 pour ne pas inclure la liste des headers
    taille_data=len(data)
    last_item_index0=len(data[0])-1
    last_item_index1=len(data[1])-1
    for i in range(taille_data-2):
        if last_item_index0!=len(data[i+1])-1:  
            data[i+1].append(str(1))
      #########################################        
        
    # ##Saisie nom utilisateur et mot de passe
    #login = pe.get_data('C:/Users/meddb-el-farouki01/Desktop/Rembursement_DGE/Programme/login.ods')['Feuille1'][0]
    login =EnterTable4.get()
    mot_de_passe= EnterTable5.get()
    ##Lancement webdriver Selenium
    s=Service(resource_path("geckodriver"))
    wd = webdriver.Firefox(service=s)
    wd_options = Options()
    wd_options.set_preference('detach',True)
    wd.get('https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb%2Fcas%2Fvalidation')

    ##Saisir utilisateur
    time.sleep(delay)
    wd.find_element(By.ID, 'identifiant').send_keys(login)

    ##Saisie mot de pass
    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)


    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)

    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    
    ##Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')
    time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    ##Saisir habilitation
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys('1')
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)
    
    ##Boucle sur le fichier selon le nombre de lignes indiquées 
    
    for i in range(line_amount):

        ##Création d'un Redevable
        ## Arriver à la transactionv 3-2-4

        time.sleep(delay)
        wd.find_element(By.ID, 'bmenuxtableMenus:9:outputBmenuxBrmenx04LibelleLigneProposee').send_keys('324')

        ##Saisir sous-dossier: "DIV"
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrdos02Yrdos021NatureSousDossier')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrdos02Yrdos021NatureSousDossier').send_keys('DIV')

        ##Creation Redevable
        ##Capture et Saisie Dénomination

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrtitnomNomprfNomProfession')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrtitnomNomprfNomProfession').send_keys(data[line][0])
        wd.find_element(By.ID, 'inputBrtitnomNomprfNomProfession').send_keys(Keys.ENTER)

        ##Saisie d'une tabulation d'échappement de la saisie de la profession

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrtitnomPrslibLibelleProfessionRaisonSoc')))
        wd.find_element(By.ID, 'inputBrtitnomPrslibLibelleProfessionRaisonSoc').send_keys(Keys.TAB)

        ##Saisie d'une tabulation d'échappement du code

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB326codeYa326CodeCodeSirOuSpi')))
        wd.find_element(By.ID, 'inputB326codeYa326CodeCodeSirOuSpi').send_keys(Keys.TAB)

        ##Saisie de la conservation du code

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
        wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')

        ##Saisie de l'adresse du redevable
        ##Saisie du Complément d'Adresse

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01CpladrComplementAdressage')))
        wd.find_element(By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01CpladrComplementAdressage').send_keys(data[line][23])

        ##Saisie du Numéro de rue

        if data[line][20]!= 0:
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01NuvoiDonNumeroVoirie')))
            wd.find_element(By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01NuvoiDonNumeroVoirie').send_keys(data[line][20])

        ##Saisie de la Boîte postale ou d'une Course Spéciale

        if data[line][6]!= 0:
            WebDriverWait(wd, 20).until(EC.presence_of_element_located(
                (By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibvoLibelleVoie')))
            wd.find_element(By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibvoLibelleVoie').send_keys(
                data[line][6])

        ##Saisie de la Commune

        WebDriverWait(wd, 20).until(EC.presence_of_element_located(
                (By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibcoLibeleCommune')))
        wd.find_element(By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibcoLibeleCommune').send_keys(
                data[line][27])


        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBndordNatureDegrevement')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordNatureDegrevement').send_keys('REMTF')

        wd.find_element(By.ID, 'inputBndordNatureDegrevement').send_keys(Keys.TAB)
        
        ##Capture et reutilisation de la date journee comptable
        djc_capture = wd.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text

        djc = djc_capture.split('/')

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordDatordDateOrdonnancementJour').send_keys(djc[0])

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordDatordDateOrdonnancementMois').send_keys(djc[1])

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordDatordDateOrdonnancementAnnee').send_keys(djc[2])
        
        ##Saisir montant
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordMntordMontantOrdonnance').send_keys(data[line][1])

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordMntordMontantOrdonnance').send_keys(Keys.ENTER)

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('o')
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'outputBdnuordAnOrdAnneeOrdonnancement')))
        valeur_rib=str(data[line][1])
        ## Création d'une liste temporaire avec numéro ordonnancement et numéro d'opération d'ordonnancement
        ## NB le numéro d'ordonnancement est divisé sur deux cellules dans MEDOC
        ## Cette liste sera finalement collée comme ligne dans le fichiers de donnees de sortie
        temp_data = []
        temp_data.append(str(data[line][0])) ##FRP  ##0 (indice dans temp_data)
        temp_data.append(str(data[line][1])) ##Montant  ##1 (indice dans temp_data)
        temp_data.append(str(data[line][5])) ##Interets moratoires  ##2 (indice dans temp_data)
        
        

        ##Numero ordonnancement
        temp_data.append(wd.find_element(By.ID, 'outputBdnuordAnOrdAnneeOrdonnancement').text +
        wd.find_element(By.ID, 'outputBdnuordNuOrdNuopesPourOrdonnancement').text)  ##4 (indice dans temp_data)
        ##numero operation ordonnancement
        #try:
        temp_data.append(wd.find_element(By.ID, 'outputBdnuordNuopet1ErCarNuopeF2').text +
        wd.find_element(By.ID, 'outputBdnuordNuopes5DerniersCarNuope').text)  ##5 (indice dans temp_data)
        #except:
            #pass
        ##Cree un fichier txt de securité avec les donnees en sortie au cas où le script plante avant 
        ##ajouter la ligne dans le fichier csv
        #with open(resource_path('temp_safety_file.txt'), 'w') as f:
        # f.write(' '.join(temp_data))
        time.sleep(delay)
        wd.find_element(By.ID,'barre_outils:image_f2').click()
        
        #wd.find_element(By.ID,'inputBdnuordYc94401AcqBarreEspace').send_keys(Keys.F2)

        ##Cree un fichier txt de securité avec les donnees en sortie au cas où le script plante avant 
        ##ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
         f.write(' '.join(temp_data))

        ##DEPENSE MONTANT DEGREVEMENT
        ## Arriver à la transactionv 21-2
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('212')

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)

        ##Saisir nature et montant
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
        
        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('REMTF')

        time.sleep(delay)
        keyboard.tap(Key.tab)
        time.sleep(delay)
        keyboard.tap(Key.tab)

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][1])
        
        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        ##Saisir libelle
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))

        time.sleep(delay)
        libelle = 'REMBT DGVT TF ' + str(data[line][2]) + '-' + str(data[line][3]) + '-' + str(data[line][4])
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(libelle)

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)

        ##Saisir numero ordonnancement
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBddeord1NuordodemNumeroOrdonancementOuDemande')))
        delai_qui_debloque = 2
        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'inputBddeord1NuordodemNumeroOrdonancementOuDemande').send_keys(temp_data[3])

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'inputBddeord1NuordodemNumeroOrdonancementOuDemande').send_keys(Keys.ENTER)

        ##Saisie credit et date imputation. Dans cette étape il faut utilisé delai double parce que medoc est particulierment capricieux.
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))

        time.sleep(delai_qui_debloque )
        keyboard.tap(Key.tab)

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('VIRT')

        time.sleep(delai_qui_debloque )
        keyboard.tap(Key.tab)

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys('E')

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][1])

        time.sleep(delai_qui_debloque )
        keyboard.tap(Key.enter)

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])

        time.sleep(delay)
        keyboard.tap(Key.tab)
        time.sleep(delay)
        keyboard.tap(Key.tab)

        ##Saisie numero dossier
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(temp_data[0])


        ##ATTENTE POUR MESSAGE D'ERREUR RÉLATIF À LA PRÉSENCE DE RAR. 
        ##SI LE MESSAGE NE S'AFFICHE PAS DANS 20 SECONDS LE SCRIPT CONTINUE
        ##EFFACER OU "COMMENT-OUT"LES LIGNES SUIVANTES SI CE MESSAGE N'APPARAITRE JAMAIS DANS LES CAS REELS
        try:
            WebDriverWait(wd, 10).until(EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
            time.sleep(delay)
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('o')
        except:
            pass
        ##JUSQ'A ICI


        ## Choix RIB
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBibanremYaribchoixSaisieChoix')))

        time.sleep(delay)

        wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(data[line][7])

        #if data[line][7] =="2":
            #numero_rang_rib=2
            #wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(numero_rang_rib)
        #elif data[line][7] =="3":
            #numero_rang_rib=3
            #wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(numero_rang_rib)
        #else:
            #numero_rang_rib=1
            #wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(numero_rang_rib)
        
        ##Saisie libelle virement emis et validation
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis')))

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(libelle)

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(Keys.ENTER)

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('o')

        ##Capture numero ordre de depense et numero operation
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(' ')

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))

        temp_data.append(wd.find_element(By.ID, 'outputBcvcs04Cr17R27CodeR17OuR27').text + ':' +
        wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)  ##5 (indice dans temp_data)

        temp_data.append(wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text + 
        wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)  ##6 (indice dans temp_data)

        ##Cree un fichier txt de securité avec les donnees en sortie au cas où le script plante avant 
        ##ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
            f.write(' '.join(temp_data))

        time.sleep(delay)
        keyboard.tap(Key.f2)

        ##ORDONNANCEMENT DE LA DEPENSE INTERETS MORATOIRES
        ## Arriver à la transactionv 26-3-1
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('2631')

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)

        ##Saisir numero dossier
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[line][0])

        ##Creation REMTF
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBndordTypeModifChoixModificationCMA')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordTypeModifChoixModificationCMA').send_keys('c')

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBndordNatureDegrevement')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordNatureDegrevement').send_keys('REMIMO')

        time.sleep(delay)
        
        ##Capture et reutilisation de la date journee comptable
        djc_capture = wd.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text

        djc = djc_capture.split('/')

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordDatordDateOrdonnancementJour').send_keys(djc[0])

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordDatordDateOrdonnancementMois').send_keys(djc[1])

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordDatordDateOrdonnancementAnnee').send_keys(djc[2])
        
        ##Saisir montant
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordMntordMontantOrdonnance').send_keys(data[line][5])

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBndordMntordMontantOrdonnance').send_keys(Keys.ENTER)

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('o')

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'outputBdnuordAnOrdAnneeOrdonnancement')))

        ## Ajoute numero ordonnancemnt des interets moratoires et le numero de l'operation correspondante
        ## à liste temporaire créée precedemment
        ## NB le numéro d'ordonnancement est divisé sur deux cellules dans MEDOC
        ## Cette liste sera finalement collée come ligne dans le fichiers de donnees de sortie

        ##Numero ordonnancement     
        temp_data.append(wd.find_element(By.ID, 'outputBdnuordAnOrdAnneeOrdonnancement').text +
        wd.find_element(By.ID, 'outputBdnuordNuOrdNuopesPourOrdonnancement').text)  ##7 (indice dans temp_data)

        ##numero operation ordonnancement
        temp_data.append(wd.find_element(By.ID, 'outputBdnuordNuopet1ErCarNuopeF2').text +
        wd.find_element(By.ID, 'outputBdnuordNuopes5DerniersCarNuope').text)  ##8 (indice dans temp_data)

        ##Cree un fichier txt de securité avec les donnees en sortie au cas où le script plante avant 
        ##ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
            f.write(' '.join(temp_data))

        time.sleep(delay)
        keyboard.tap(Key.f2)

        ##DEPENSE INTERETS MORATOIRES
        ## Arriver à la transaction 21-2
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('212')

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)

        ##Saisir nature et montant
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
        
        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('REMIMO')

        time.sleep(delay)
        keyboard.tap(Key.tab)

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][5])
        
        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        ##Saisir libelle
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))

        time.sleep(delay)
        libelle = 'REMBT IM TEOM DGVT TF ' + str(data[line][2]) + '-' + str(data[line][3]) + '-' + str(data[line][4])
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(libelle)

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)

        ##Saisir numero ordonnancement
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBddeord1NuordodemNumeroOrdonancementOuDemande')))

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'inputBddeord1NuordodemNumeroOrdonancementOuDemande').send_keys(temp_data[7])

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'inputBddeord1NuordodemNumeroOrdonancementOuDemande').send_keys(Keys.ENTER)

        ##Saisie credit et date imputation
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))

        time.sleep(delai_qui_debloque )
        keyboard.tap(Key.tab)

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('VIRT')

        time.sleep(delai_qui_debloque )
        keyboard.tap(Key.tab)

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys('E')

        time.sleep(delai_qui_debloque )
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][5])

        time.sleep(delai_qui_debloque )
        keyboard.tap(Key.enter)

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])

        keyboard.tap(Key.tab)
        time.sleep(delay)
        keyboard.tap(Key.tab)

        ##Saisie numero dossier
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(temp_data[0])


        ## ATTENTE POUR MESSAGE D'ERREUR RÉLATIF À LA PRÉSENCE DE RAR. 
        ## SI LE MESSAGE NE S'AFFICHE PAS DANS 20 SECONDS LE SCRIPT CONTINUE.
        ## SI CE MESSAGE N'APPARAITRE JAMAIS DANS LES CAS REELS ON PEUT EFFACER OU 
        ## "COMMENT-OUT" LES LIGNES SUIVANTES.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
        try:
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))

            time.sleep(delay)
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('o')
        except:
            pass
        ##JUSQ' ICI


        ## Choix RIB
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBibanremYaribchoixSaisieChoix')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(data[line][7])

        ##Saisie libelle virement emis et validation
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis')))

        time.sleep(delay)
        libelle = 'REMBT IMTEOM TF ' + str(data[line][2]) + '-' + str(data[line][3]) + '-' + str(data[line][4])
        wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(libelle)

        time.sleep(delay)
        wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(Keys.ENTER)

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('o')

        ##Capture numero ordre de depense et numero operation
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(' ')

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))

        temp_data.append(wd.find_element(By.ID, 'outputBcvcs04Cr17R27CodeR17OuR27').text + ':' +
        wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)  ##9 (indice dans temp_data)

        temp_data.append(wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text + 
        wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)  ##10 (indice dans temp_data)

        time.sleep(delay)
        keyboard.tap(Key.f2)

        ##Cree un fichier txt de securité avec les donnees en sortie au cas où le script plante avant 
        ##ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
            f.write(' '.join(temp_data))

        ##Les données en sortie son ajouté dans un fichier csv
        with open('donnees_sortie.csv', 'a', newline='\n') as f:
            writer_object = csv.writer(f)
            writer_object.writerow(temp_data)
    
        ## La ligne du fichier BIS des données d'entrée est marquée comme traitée
        ## Pour faire ça on ajoute un element à la fin de la bonne ligne de l'"array" des données
        ## d'entrée et on écrase le fichier. C'est pour ça qu'on le fait sur un copie, comme mésure 
        ## de sécurité extra
        donnees_entree_bis = pe.get_data('donnees_entree_bis.ods')
        donnees_entree_bis['Feuille1'][line].append('X') 
        pe.save_data('donnees_entree_bis.ods', donnees_entree_bis)

        ##On passe à la ligne suivante
        line += 1
    wd.quit()

#if __name__ == '__main__':
    #main()

# Procédure pour 
def open_file():
   global File_path
   file = filedialog.askopenfile(mode='r', filetypes=[('Ods Files', '*.ods')])
   if file:
      filepath = os.path.abspath(file.name)
      filepath =filepath.replace(os.sep,"/")
      label_path.configure(text="Le fichier sélectionné est : " + Path(filepath).stem)   
      File_path=filepath
                                                                                                  

Interface = Tk()
Interface.geometry('1000x500')
Interface.title('Remboursement TEOM')


EnterTable1 = StringVar()
EnterTable2 = StringVar()
EnterTable3 = StringVar() 
EnterTable4 = StringVar() 
EnterTable5 = StringVar() 
#label_file_explorer = Label(Interface,text = "chemin du fichier",
#                            width = 100, height = 4,
#                            fg = "blue")
#label_file_explorer.grid(row=30,column=4)
paramx=10
paramy=170
label1=Label(Interface, text='Remboursement TEOM', font=('Arial',15), fg='Black',bg='#ffffff')
label1.place(x= 400,y=1)
label2=Label(Interface,text='Saisir le delai entre les opérations de l\'automate en secondes :')
label2.place(x=paramx + 250,y=paramy+30)
entry1=Entry(Interface, textvariable=EnterTable1, justify='center')
entry1.place(x= paramx +600,y=paramy+30)
label3=Label(Interface,text='Saisir la ligne du début: ')
label3.place(x= paramx +250,y=paramy+60)
entry2=Entry(Interface, textvariable=EnterTable2, justify='center')
entry2.place(x= paramx +600,y=paramy+60)
label4=Label(Interface,text='Saisir le nombre de lignes à traiter: ')
label4.place(x= paramx +250,y=paramy+90)
entry3=Entry(Interface, textvariable=EnterTable3, justify='center')
entry3.place(x= paramx +600,y=paramy+90)
#login et mot de passe 
label5=Label(Interface,text='Login:')
label5.place(x=  250,y=60)
entry4=Entry(Interface, textvariable=EnterTable4, justify='center')
entry4.place(x= 300,y=60)
label6=Label(Interface,text='Mot de passe: ')
label6.place(x= 500,y=60)
entry5=Entry(Interface, textvariable=EnterTable5, justify='center')
entry5.place(x= 600,y=60)
#entry6=Entry(Interface, textvariable=EnterTable, justify='center').grid(row=6, column=6)
#label1=Label(Interface,text='aaa').grid(row=7,column=6)            
button2=Button(Interface, text='Choisir le fichier d\'entrée', command=open_file)
button2.place(x= 400,y=120)
label_path=Label(Interface)
label_path.place(x= 360,y=150)
button1=Button(Interface, text='Lancer le programme', command=main)
button1.place(x= 350,y=370)  
QUIT=Button(Interface,text='Quitter', fg='Red', command=Interface.destroy);
QUIT.place(x= 550,y=370)
     
 
 
Interface.mainloop()




    



    



