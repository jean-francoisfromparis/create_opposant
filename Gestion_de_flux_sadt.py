import csv
import time
import os
import sys
import pyexcel_ods3 as pe
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from odf import *
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from pprint import pprint


# Fonction pour retrouver le chemin d'accès
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def main():
    
    #Délai entre opérations automate. Pour des numéros non entiers il faut utiliser le point pas la virgule
    while True:
        try:
            delay = float(delai_entre_operations.get())
            break
        except ValueError:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()


    #Prend la ligne du fichier depuis laquelle commencer à lire 
    while True:
        line = ligne_de_debut.get()
        if line.isnumeric():  ##vérifie que ça soit un numéro
            line = int(line) - 1  ##ajuste l'indice
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()

    #Combien de lignes du fichier traiter
    while True:
        line_amount = nb_lignes_a_traiter.get()
        if line_amount.isnumeric():
            line_amount = int(line_amount)
            break
        else:
            messagebox.OK('Saisie incorrecte, réessayez')
            exit()
    
    #Prendre les données depuis le fichier,
    donnees_entree = pe.get_data(File_path)
    data = [i for i in donnees_entree['Feuille1']]

    #Lancement webdriver Selenium
    ser=Service(resource_path("geckodriver"))
    browser = webdriver.Firefox(service=ser)
    wd_options = Options()
    wd_options.set_preference('detach',True)
    browser.get('http://medoc.ia.dgfip:8141/medocweb/presentation/md2oagt/ouverturesessionagent/ecran/ecOuvertureSessionAgent.jsf')

    login = login_interface.get()
    mot_de_passe= mot_de_passe_interface.get()

    numero_de_service= '0070100'
    habilitation = 1

    #Saisir utilisateur
    time.sleep(delay)
    browser.find_element(By.ID, 'identifiant').send_keys(login)

    #Saisir mot de pass et valider
    time.sleep(delay)
    browser.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)

    time.sleep(delay)
    browser.find_element(By.ID, 'secret_tmp').send_keys(Keys.ENTER)

    #Saisir service
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))

    browser.find_element(By.ID, 'nomServiceChoisi').send_keys(numero_de_service)
    time.sleep(delay)
    browser.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    #Saisir habilitation
    time.sleep(delay)
    browser.find_element(By.ID, 'habilitation').send_keys(habilitation)
    time.sleep(delay)
    browser.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)
    
    #Boucle sur le fichier selon le nombre de lignes indiquées 
    for i in range(line_amount):

        #Saisir la transaction 21-2
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')

        #Création affaire service au code R17 "7055"
        #Saisir la nature "AFF" pour debit 473-0
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)

        #Saisir ENTREE pour type de montant
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)

        #Saisir le montant X
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][3])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        #Saisir une identification
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(data[line][4])
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)

        #Saisir le numéro d'affaire
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[line][1])

        #Confirmer le libelle de l'affaire
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
        
        #Messege informatif
        try :
            if (browser.find_element(By.CSS_SELECTOR,'.Ui-messages-error-summary').is_displayed):
                time.sleep(delay)
                browser.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
        except :
           pass
        
        #Saisir le code R27 "7370"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)

        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')

        #Saisir le numéro du compte 477-0
        time.sleep(delay)
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('477-0')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys(Keys.ENTER)
        
        #Saisir la nature "AFF" pour crédit 477-0
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)

        #Saisir ENTREE pour type de montant
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)

        #Saisir le montant X
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][3])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        #Saisir la date
        #Capture et réutilisation de la date journee comptable
        djc_capture = browser.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text

        djc = djc_capture.split('/')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
        
        #Saisir le numéro d'affaire
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(Keys.ENTER)
        
        #Saisir le numéro de dossier
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Bcaff032RedevServOuRlce').send_keys('REDEV')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Bcaff036Car2A7NuordNumDos').send_keys(data[line][0])
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Bcaff038Cplnum').send_keys('0')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Bcaff038Cplnum').send_keys(Keys.ENTER)
        
        #Saisir le libellé
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(data[line][4])
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
        
        #Saisir le code R27 "7055"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Cr17R27CodeR17OuR27').send_keys('7055')
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(data[line][3])
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')

        #Validation de la transaction
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')

        #Création d'une liste temporaire avec numéro d'ordre de dépenses, le numéro de l'affaire créée et le numéro de l'opération
        #Le numéro de l'opération est divisé sur deux cellules dans MEDOC
        #Cette liste sera finalement collée comme ligne dans le fichier des donnees de sortie
        liste_temporaire_data = []
        liste_temporaire_data.append(str(data[line][0])) #FRP indice #0 dans liste_temporaire_data
        liste_temporaire_data.append(str(data[line][3])) #Montant indice #1 dans liste_temporaire_data  

        #Numero de l'ordre de depense
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs022NoDepense')))
        #Numero de l'ordre de depense indice #2 dans liste_temporaire_data
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
                
        #Pour afficher la suite
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)

        #Numero de l'affaire créée
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs021NumAffaireCreee')))
        #Numero de l'affaire créée indice #3 dans liste_temporaire_data
        numero_affaire_creee = browser.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
        liste_temporaire_data.append(numero_affaire_creee)

        try :
            if (liste_temporaire_data[3] != numero_affaire_creee or liste_temporaire_data[3] =='' ):
                time.sleep(delay)
                numero_affaire_creee_v = browser.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
                liste_temporaire_data[3]=numero_affaire_creee_v
                pprint(liste_temporaire_data)
        except :
           pass

        #Pour afficher la suite
        delai_qui_debloque = 2
        time.sleep(delai_qui_debloque)
        browser.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)

        #Numero de l'opération
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
        #Numero de l'opération indice #4 dans liste_temporaire_data
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text + browser.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)

        #Pour afficher la suite
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')

        #Fin de la transaction 21-2 et retour à la page d'accueil
        time.sleep(delay)
        browser.find_element(By.ID,'barre_outils:image_f2').click()
        
        

        #Créer un fichier txt de securité avec les donnees de sortie en cas de plantage
        #Ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
         f.write(' '.join(liste_temporaire_data))


        #Saisir la transaction 21-2
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')

        #Création affaire service au code R27 "8755"
        #Saisir la nature "AFF" pour debit 473-0
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)

        #Saisir ENTREE pour type de montant
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)

        #Saisir le montant X
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][3])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        #Saisir une identification
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(data[line][4])
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)

        #Saisir le numéro d'affaire créée  précédemment
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(liste_temporaire_data[3])

        #Confirmer le libelle de l'affaire
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
        
        #Saisir le code R27 "8755"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
        
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)

        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')

        #Répondre à la question "Soldez-vous l'affaire?"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
        
        #Valider CREDIT
        time.sleep(delai_qui_debloque)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
        
        #Saisir la nature "OVIRT" 
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('OVIRT')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)

        #Saisir ENTREE pour type de montant
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)

        #Saisir le montant X
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[line][3])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        #Saisir le codique du service bénéficiaire
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBn4f3001Bn4F300101ZoneCodiqueService').send_keys(data[line][5])
        
        #Appuyer sur Entrer pour continuer  
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre').send_keys(Keys.ENTER)
        
        #Validation de la transaction
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
        
        #Numero de l'ordre de depense
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
        #Numero de l'ordre de depense indice #5 dans liste_temporaire_data
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
                
        #Pour afficher la suite
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)

        #Numero de l'opération
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
        #Numero de l'opération indice #6 dans liste_temporaire_data      
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text + browser.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)

        #Pour afficher la suite
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')

        #Créer un fichier txt de securité avec les donnees de sortie en cas de plantage
        #Ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
         f.write(' '.join(liste_temporaire_data))


        #Fin de la transaction 21-2 et retour à la page d'accueil
        time.sleep(delay)
        browser.find_element(By.ID,'barre_outils:image_f2').click()
        
        #Saisir la transaction 3-8-2
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('3')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('8')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')

        #Saisir le numéro d'affaire à partir des données d'entrées
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBrsdo03Nuaff1NumeroAffaire').send_keys(data[line][1])

        #Saisir le type de l'affaire "64"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBrsdo03NasdoNatureSousDossier').send_keys('64')
        
        #Récuperer le nouveau solde de l'affaire au code 1760 et enregistrer le sous indice #7 dans liste_temporaire_data  
        time.sleep(delay)
        browser.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text     
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text)

        #Récuperer le nom de l'entreprise à rembourser et enregistrer le sous indice #8 dans liste_temporaire_data  
        time.sleep(delay)
        browser.find_element(By.ID, 'outputBrtit04NomprfNomProfession').text     
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBrtit04NomprfNomProfession').text + "/SOLDE RCTVA" )
        
        #Pour afficher la suite 
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'labelBrval18BarreEspace0')))
        browser.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)

        #Pour afficher la suite encore une fois en cas de besoin
        time.sleep(delay)
        browser.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)

        #Créer un fichier txt de securité avec les donnees de sortie en cas de plantage
        #Ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
         f.write(' '.join(liste_temporaire_data))

        #Fin de la transaction 3-8-2 et retour à la page d'accueil
        time.sleep(delay)
        browser.find_element(By.ID,'barre_outils:image_f2').click()

        #Saisir la transaction 21-2 
        #Remboursement du solde à la société débitrice
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')

        #Saisir la nature "AFF" pour debit 473-0
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)

        #Saisir ENTREE pour type de montant
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)

        #Saisir le montant X
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(liste_temporaire_data[7])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)

        #Saisir une identification
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(liste_temporaire_data[8])
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)

        #Saisir le numéro d'affaire créée  précédemment
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[line][1])

        #Confirmer le libelle de l'affaire
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
        
        #Saisir le code R27 "7370"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
        
        
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)

        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')

        #Répondre à la question "Soldez-vous l'affaire?"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
        
        #Valider CREDIT
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
        
        #Saisir le numéro du compte 512-96

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
        
        #Saisir la nature "OVIRT" 
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('VIRT')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)

        #Saisir ENTREE pour type de montant
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)

        #Saisir le montant X
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(liste_temporaire_data[7])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
        
        #Saisir la date
        #Capture et réutilisation de la date journee comptable
        djc_capture = browser.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text

        djc = djc_capture.split('/')

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])

        #Saisir le numéro de dossier
        time.sleep(delay)
        browser.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[line][0])
        
        #Continuer en cas d'existence de RAR
        try :
            if (browser.find_element(By.ID,'outputBrep9081Txt9081TexteDemandeConfirmation').is_displayed):
                time.sleep(delai_qui_debloque)
                browser.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
        except Exception :
           pass
        

        #Saisir le numéro de l'IBAN
        time.sleep(delai_qui_debloque)
        WebDriverWait(browser,20).until(EC.presence_of_all_elements_located((By.ID, 'inputBibanremYaribmess1LibelleMessage')))
        browser.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(data[line][6])
        
        #Libelle du virement emis
        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(data[line][7])

        time.sleep(delay)
        browser.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(Keys.ENTER)

        #Répondre à la question "Voulez-vous valider?"
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
        
        #Numero de l'ordre de depense
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
        #Numero de l'ordre de depense indice #9 dans liste_temporaire_data
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
                
        #Pour afficher la suite
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)

        #Numero de l'opération
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
        #Numero de l'opération indice #10 dans liste_temporaire_data      
        liste_temporaire_data.append(browser.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text + browser.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)

        #Pour afficher la suite
        time.sleep(delay)
        browser.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')

        #Créer un fichier txt de securité avec les donnees de sortie en cas de plantage
        #Ajouter la ligne dans le fichier csv
        with open('temp_safety_file.txt', 'w') as f:
         f.write(' '.join(liste_temporaire_data))

        #Fin de la transaction 21-2 et retour à la page d'accueil
        time.sleep(delay)
        browser.find_element(By.ID,'barre_outils:image_f2').click()

        #Les données de sortie sont ajoutées dans un fichier csv
        with open('donnees_sortie.csv', 'a', newline='\n') as f:
            writer_object = csv.writer(f)
            writer_object.writerow(liste_temporaire_data)
    
        #Pour marquer une ligne traité dans le fichier "donnees_entree_bis", "X" est ajoutée à la fin de chaque ligne 
        donnees_entree_bis = pe.get_data('donnees_entree_bis.ods')
        donnees_entree_bis['Feuille1'][line].append('X') 
        pe.save_data('donnees_entree_bis.ods', donnees_entree_bis)

        #Passer à la ligne suivante
        line += 1

    #Fermer le webdriver
    browser.quit()



# Procédure pour choisir un fichier d'entrée  
def open_file():
   global File_path
   file = filedialog.askopenfile(mode='r', filetypes=[('Ods Files', '*.ods')])
   if file:
      filepath = os.path.abspath(file.name)
      filepath =filepath.replace(os.sep,"/")
      label_path.configure(text="Le fichier sélectionné est : " + Path(filepath).stem)   
      File_path=filepath

#Paramétrage de l'interface
Interface = Tk()
Interface.geometry('1000x500')
Interface.title('Gestion du flux de SADT')

login_interface = StringVar()
mot_de_passe_interface = StringVar()
delai_entre_operations = StringVar() 
ligne_de_debut = StringVar()
nb_lignes_a_traiter = StringVar() 

#Saisir le nombre de lignes à traiter et le délai entre opérations
paramx=0
paramy=150

label_1=Label(Interface, text='Gestion du flux de SADT', font=('Arial',15), fg='Black')
label_1.place(x= 400,y=10)

label_2=Label(Interface,text='Saisir le delai entre les opérations de l\'automate en secondes :')
label_2.place(x=paramx + 250,y=paramy+50)

entry_1=Entry(Interface, textvariable=delai_entre_operations, justify='center')
entry_1.place(x= paramx +600,y=paramy+50)

label_3=Label(Interface,text='Saisir la ligne du début: ')
label_3.place(x= paramx +250,y=paramy+80)

entry_2=Entry(Interface, textvariable=ligne_de_debut, justify='center')
entry_2.place(x= paramx +600,y=paramy+80)

label_4=Label(Interface,text='Saisir le nombre de lignes à traiter: ')
label_4.place(x= paramx +250,y=paramy+110)

entry3=Entry(Interface, textvariable=nb_lignes_a_traiter , justify='center')
entry3.place(x= paramx +600,y=paramy+110)

#login et mot de passe 
label_5=Label(Interface,text='Login:')
label_5.place(x= 250,y=80)
entry_4=Entry(Interface, textvariable=login_interface, justify='center')
entry_4.place(x= 300,y=80)

label_6=Label(Interface,text='Mot de passe: ')
label_6.place(x= 500,y=80)
entry_5=Entry(Interface, textvariable=mot_de_passe_interface, justify='center')
entry_5.place(x= 600,y=80)

#Boutton pour choisir un fichier d'entrée       
button_2=Button(Interface, text='Choisir le fichier d\'entrée', command=open_file)
button_2.place(x= 400,y=120)

label_path=Label(Interface)
label_path.place(x= 360,y=160)

#Boutton pour commencer les opérations et boutton pour quitter
button_1=Button(Interface, text='Lancer le programme', command=main)
button_1.place(x= 350,y=350)  
QUIT=Button(Interface,text='Quitter', fg='Red', command=Interface.destroy);
QUIT.place(x= 550,y=350)
     
Interface.mainloop()




