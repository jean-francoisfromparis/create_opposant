
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
    donnees_entree = pe.get_data(File_path)
    data = [i for i in donnees_entree['Database']]

    
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
    #login = pe.get_data('C:/Users/meddb-el-farouki01/Desktop/Rembursement_DGE/Programme/login.ods')['Database'][0]
    login =EnterTable4.get()
    mot_de_passe= EnterTable5.get()
    ##Lancement webdriver Selenium
    s=Service(resource_path("geckodriver"))
    # wd = webdriver.Firefox(service=s)
    wd = webdriver.Firefox(executable_path=GeckoDriverManager().install())
    wd_options = Options()
    wd_options.set_preference('detach',True)
    wd.get('https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8121%2Fmedocweb%2Fcas%2Fvalidation')

    ##Saisir utilisateur
    time.sleep(delay)
    # wd.find_element(By.ID, 'identifiant').send_keys(login)
    wd.find_element(By.ID, 'identifiant').send_keys("youssef.atigui")

    ##Saisie mot de pass
    time.sleep(delay)
    # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
    wd.find_element(By.ID, 'secret_tmp').send_keys("1")


    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)

    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    
    ##Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('6200100')
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
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('324')
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)

        ##Saisir sous-dossier: "DIV"
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrdos02Yrdos021NatureSousDossier')))

        wd.find_element(By.ID, 'inputBrdos02Yrdos021NatureSousDossier').send_keys('DIV')
        wd.find_element(By.ID, 'inputBrdos02Yrdos021NatureSousDossier').send_keys(Keys.ENTER)


        ##Saisie du Champ du Genre

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrtitnomTitredTitreCftitdes')))
        wd.find_element(By.ID, "inputBrtitnomTitredTitreCftitdes").click()
        wd.find_element(By.ID, "inputBrtitnomTitredTitreCftitdes").send_keys("M")
        wd.find_element(By.ID, "inputBrtitnomTitredTitreCftitdes").send_keys(Keys.ENTER)

        ##Saisie du Champ du Destinataire

        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrtitnomNomprfNomProfession')))
        wd.find_element(By.ID, "inputBrtitnomNomprfNomProfession").click()
        wd.find_element(By.ID, "inputBrtitnomNomprfNomProfession").send_keys(data[line][0])
        wd.find_element(By.ID, "inputBrtitnomNomprfNomProfession").send_keys(Keys.ENTER)

        ##Saisie du Champ de la raison sociale
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located(
            (By.ID, 'inputBrtitnomPrslibLibelleProfessionRaisonSoc')))
        wd.find_element(By.ID,
                        "inputBrtitnomPrslibLibelleProfessionRaisonSoc").click()
        wd.find_element(By.ID,
                        "inputBrtitnomPrslibLibelleProfessionRaisonSoc").send_keys(
            Keys.ENTER)

        ##Saisie du Champ du code
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located(
            (By.ID, 'inputB326codeYa326CodeCodeSirOuSpi')))
        wd.find_element(By.ID,
                        "inputB326codeYa326CodeCodeSirOuSpi").click()
        wd.find_element(By.ID,
                        "inputB326codeYa326CodeCodeSirOuSpi").send_keys(
            Keys.ENTER)

        ##Saisie du Champ du
        WebDriverWait(wd, 20).until(EC.presence_of_element_located(
            (By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
        time.sleep(delay)
        wd.find_element(By.ID,
                        "inputBrep9081Rep9082ReponseUtilisateurON").click()
        wd.find_element(By.ID,
                        "inputBrep9081Rep9082ReponseUtilisateurON").send_keys(
            "O")

        ##Saisie du Champ du Code DSF
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'repeatBradr01SaisieUneAdresse:0:inputBradr01NudsfCodeDsfAdresseCorrespondance')))
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01NudsfCodeDsfAdresseCorrespondance").send_keys(Keys.TAB)

        ##Saisie du Champ du Code Commune
        wd.find_element(By.ID,
                        "repeatBradr01SaisieUneAdresse:0:inputBradr01CcomCodeCommune").send_keys(
            Keys.TAB)

        ##Saisie du Champ du numéro de voie
        time.sleep(delay)
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01TopnuvoiNumeroVoie").click()
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01TopnuvoiNumeroVoie").send_keys(
            data[line][20])

        ##Saisie du Champ du libellé de la voie
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibvoLibelleVoie").click()
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibvoLibelleVoie").send_keys(
            data[line][23])

        ##Saisie du Champ du libellé de la commune
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibcoLibeleCommune").click()
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01ToplibcoLibeleCommune").send_keys(
            data[line][32])

        ##Saisie du Champ du libellé code postal
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01TopcodpoCodePostal").click()
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01TopcodpoCodePostal").send_keys(
            data[line][13])

        ##Saisie du Champ du libellé du bureau distributeur
        time.sleep(delay)
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01TopburBureauDistributeur").click()


        time.sleep(delay)
        wd.find_element(By.ID,
                                 "repeatBradr01SaisieUneAdresse:0:inputBrval01ReponseOOuN").send_keys("N")

        time.sleep(delay)
        wd.find_element(By.ID,
                                 "repeatBradr01SaisieUneAdresse:0:inputBradr01TitdesTitreDestinataire").send_keys("M")
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01TitdesTitreDestinataire").send_keys(
            Keys.TAB)

        time.sleep(delay)
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01NomdesNomDestinataire").send_keys(
            data[line][0])
        wd.find_element(By.ID, "repeatBradr01SaisieUneAdresse:0:inputBradr01NomdesNomDestinataire").send_keys(
            Keys.TAB)

        ##Saisie du Champ DU Formulaire d'adresse
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrval01Yrval011ReponseOOuN')))
        wd.find_element(By.ID, 'inputBrval01Yrval011ReponseOOuN').send_keys('N')

        ##Saisie du Champ de Validation du Formulaire
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBr351val')))
        wd.find_element(By.ID, 'inputBr351val').send_keys('O')

        ##Capture du numéro de dossier
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'brmes01Row6Label2')))
        numero_dossier_capture = wd.find_element(By.ID, 'brmes01Row6Label2').text


        ##Capture du numéro de la Clé du dossier

        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'brmes01Row6Label4')))
        cle_capture = wd.find_element(By.ID, 'brmes01Row6Label4').text

        ##Enregistrement de données captures dans un tableau temporaire
        ## Cette liste sera finalement collée comme ligne dans le fichier de données de sortie
        temp_data = []
        temp_data.append(str(data[line][0]))  ##FRP  ##0 (indice dans temp_data)
        temp_data.append(str(numero_dossier_capture))  ##Numero de dossier  ##1
        temp_data.append(str(cle_capture))  ##Cle du dossier ##2

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
        donnees_entree_bis['Database'][line].append('X')
        pe.save_data('donnees_entree_bis.ods', donnees_entree_bis)

        ##On passe à la ligne suivante
        # line += 1

        ## Retour au menu principal
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        wd.find_element(By.ID, 'barre_outils:touche_f2').click()

        ##Création d'un Redevable
        ## Arriver à la transactionv 3-15-2

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3152')
        wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)

        ## Saisie du N° de Dossier
        numero_dossier = numero_dossier_capture[1:][:-3]
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib002YaribnumSaisieNuord').click()
        wd.find_element(By.ID, 'inputBrib002YaribnumSaisieNuord').send_keys(numero_dossier)

        ## Saisie de la ligne NOUVEAU COMPTE
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib002RibtypTypeCompteCfTexte').click()
        wd.find_element(By.ID, 'inputBrib002RibtypTypeCompteCfTexte').send_keys('1')

        ## Saisie de la Création
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib003YaribchxChoixCompteATraiter').click()
        wd.find_element(By.ID, 'inputBrib003YaribchxChoixCompteATraiter').send_keys('1')

        ## Saisie de l'IBAN
            ## Saisie du code pays
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban1Bqriban1OuBqeiban1').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban1Bqriban1OuBqeiban1').send_keys(data[line][36])

           ## Saisie du code Banque
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban2Bqriban2OuBqeiban2').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban2Bqriban2OuBqeiban2').send_keys(data[line][38])

           ## Saisie du code Guichet
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban3Bqriban3OuBqeiban3').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban3Bqriban3OuBqeiban3').send_keys(data[line][40])

           ## Saisie du code Agence
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban4Bqriban4OuBqeiban4').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban4Bqriban4OuBqeiban4').send_keys(data[line][42])

           ## Saisie du Compte tranche 1
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban5Bqriban5OuBqeiban5').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban5Bqriban5OuBqeiban5').send_keys(data[line][44])

           ## Saisie du Compte tranche 2
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban6Bqriban6OuBqeiban6').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban6Bqriban6OuBqeiban6').send_keys(data[line][46])

           ## Saisie du Compte tranche 3
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban7Bqriban7OuBqeiban7').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban7Bqriban7OuBqeiban7').send_keys(data[line][48])

        ## Tabulation dernier champs
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban8Bqriban8OuBqeiban8').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban8Bqriban8OuBqeiban8').send_keys(Keys.TAB)

        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004Yaiban9Bqriban9OuBqeiban9').click()
        wd.find_element(By.ID, 'inputBrib004Yaiban9Bqriban9OuBqeiban9').send_keys(Keys.TAB)

           # Saisie du BIC
        time.sleep(delay)
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBrib004YabicCodeBic')))
        wd.find_element(By.ID, 'inputBrib004YabicCodeBic').click()
        wd.find_element(By.ID, 'inputBrib004YabicCodeBic').send_keys(data[line][34])

          ## Valider la saisie du formulaire
        time.sleep(delay)
        wd.find_element(By.ID, 'inputBrib004YaribvalValidationEcran').send_keys('O')

         ## Validation retour au menu


    # wd.quit()




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
Interface.title('Création Redevable')


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
label1=Label(Interface, text='Création Redevable', font=('Arial',15), fg='Black',bg='#ffffff')
label1.place(x= 400,y=1)
label2=Label(Interface,text='Saisir le délai entre les opérations de l\'automate en secondes :')
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




    



    



