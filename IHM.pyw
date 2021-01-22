from tkinter import *
from tkinter import filedialog, messagebox, END
import tkinter.font as tkFont
from functools import partial
import os
import json
from tools import *
from yaml import Loader, load
###############################################################
swd = os.path.dirname(os.path.realpath(sys.argv[0]))
# Chargement du fichier YAML
f = open(swd + "/theme.yaml", 'r',  encoding='utf8')
thor = load(f, Loader=Loader)
f.close()

###############################################################
# fenetre explorateur de fichier pour la selection d'un nom de fichier
# INPUT : extension : extension autorisées pour la selection de fichier


def choose_a_file(extension):
        global lastdir
        options = {}
        options['defaultextension'] = extension
        options['filetypes'] = [(extension.upper(), extension.lower())]
        options['initialdir'] = lastdir
        options['initialfile'] = ''
        options['title'] = 'Selectionnez un fichier'

        chosenFile = filedialog.askopenfilename(**options)
        if chosenFile:
            lastdir = os.path.dirname(chosenFile)
            return chosenFile
        else:
            return ''

#####################################################################################################################
# fenetre explorateur de fichier pour la selection d'un nom de fichier
# INPUT : extension : extension autorisées pour la selection de fichier


def choose_a_fileName(extension):
        global lastdir
        options = {}
        options['defaultextension'] = extension
        options['filetypes'] = [(extension.upper(), extension.lower())]
        options['initialdir'] = lastdir
        options['initialfile'] = ''
        options['title'] = 'Choisissez un nom de fichier'

        chosenFile = filedialog.asksaveasfilename(**options)
        if chosenFile:
            lastdir = os.path.dirname(chosenFile)
            return chosenFile
        else:
            return ''

#####################################################################################################################
# fonction permettant de selectionner un nom de fichier puis de mettre à jour le champ correspondant dans l'interface graphique
# INPUT : extension : extension attendue du fichier
# INPUT : field : champ de l'interface à nmettre à jour


def update_file(extension, field, saveas=False):
    if saveas:
        fileName = choose_a_fileName(extension)
    else:
        fileName = choose_a_file(extension)
    if fileName:
        field.set(fileName)

#####################################################################################################################
# fonction permettant de mettre à jour l'iHM en crééant de nouveau champs pour les scenarios strategiques
# INPUT : field : pointeur vers la velaur du champ à modifier
# INPUT : atel : contenaire des champs
# INPUT : root : identifiant de la fenetre
# INPUT : numrow : positionnement vertical dans le container
# INPUT : numpart : numpart positionnement vertical de l'atel dans la fenetre


def update_ihm_strat(field, atel, inputs, root, numrow, numpart, mod):
    c, r = atel.grid_size()  # taille de l'atelier en colonne et ligne
    modele = mod["Scénario_stratégique"]  # modele des scenario strategiques
    modeleImage = mod["Image_scénario_stratégique"]  # modele des imges associées au scenarios stratégiques
    if int(field.get()) > len(context["scenariosStrategiques"]):  # Si le nombre de scenario a crée est supérieur à celui existant
        numrow = numrow + 2*len(context["scenariosStrategiques"])  # On se place apres les scenarios existants
        # On decale vers le bas les widget présent après les scenarios strategiques
        for x in range(numrow, r):  # pour chaque ligne jusqu'à la fin de l'atelier
            for widget in atel.grid_slaves(row=x):  # pour chaque widget de la ligne
                widget.grid(row=x+2*(int(field.get())-len(context["scenariosOperationnels"])))  # on décale les widgets existants
        for x in range(len(context["scenariosStrategiques"]), int(field.get())):  # Pour chaque scenario a créér
            if not "str"+str(x+1) in inputs.keys():  # si la cle n'existe pas
                inputs["str"+str(x+1)] = StringVar(root)
            if not "Image str"+str(x+1) in inputs.keys():  # si la cle n'existe pas
                inputs["Image str"+str(x+1)] = StringVar(root)
            newLabel(atel, modele['label'].replace("{{ sc }}", str(x+1)), bold).grid(column=0, row=numrow)  # on créér le label du cficher excel
            newEntry(atel, normal, inputs["str"+str(x+1)]).grid(column=1, row=numrow)  # on créer le champ du fichier excel
            Button(atel, text=' search', font=bold, image=excelIcon, compound=LEFT, command=partial(update_file, "."+modele["extension"], inputs["str"+str(x+1)])).grid(column=2, row=numrow, padx=5)  # onc créer le bouron associé au fichier excel
            numrow = numrow+1  # ligne suivante
            newLabel(atel, modeleImage['label'].replace("{{ sc }}", str(x+1)), bold).grid(column=0, row=numrow)  # on créer le label de l'image
            newEntry(atel, normal, inputs["Image str"+str(x+1)]).grid(column=1, row=numrow)  # on créer le champdu fichier image
            Button(atel, text=' search', font=bold, image=jpgIcon, compound=LEFT, command=partial(update_file, "."+modeleImage["extension"], inputs["Image str"+str(x+1)])).grid(column=2, row=numrow, padx=5)  # on créér le bouton associé au fichier image
            numrow = numrow+1  # ligne suivante
            context["scenariosStrategiques"].append(x+1)  # on ajoute le scenario à la liste des scenarios
        atel.grid(row=numpart, column=0)
    elif int(field.get()) < len(context["scenariosStrategiques"]):  # Si le nombre de scenario à créér est inférieur au nombre existant
        for x in range(int(field.get()), 2*len(context["scenariosStrategiques"])):  # pour chaque scnenario en trop
            for widget in atel.grid_slaves(row=x+2+numrow):  # pour chaque widget de la ligne correspondante
                widget.destroy()  # on detruit le widget
            if 'str'+str(x+1) in inputs.keys():  # si la cle existe
                del inputs['str'+str(x+1)]  # destruction de la cle du fichier excel
                del inputs['Image str'+str(x+1)]  # destruction de la cle de l'image
                context["scenariosStrategiques"].remove(x+1)  # suppression du scenario dans la liste des scenarios
    return numrow  # on retourne le numero de ligne pour continuer l'affichage de l'IHM

#####################################################################################################################
# fonction permettant de mettre à jour l'iHM en crééant de nouveau champs pour les scenarios opérationnels
# INPUT : field : pointeur vers la velaur du champ à modifier
# INPUT : atel : contenaire des champs
# INPUT : root : identifiant de la fenetre
# INPUT : numrow : positionnement vertical dans le container
# INPUT : numpart : numpart positionnement vertical de l'atel dans la fenetre


def update_ihm_oper(field, atel, inputs, root, numrow, numpart, mod):
    c, r = atel.grid_size()  # on recupere le nmobre de ligne et de ccolonne de l'atelier dans l'interface graphique
    modeleImage = mod["Image_scénario_opérationnel"]  # on recupere le modele pour les image
    if int(field.get()) > len(context["scenariosOperationnels"]):  # Si le nombre de scenario a crée est supérieur à celui existant
        numrow = numrow + len(context["scenariosOperationnels"])  # on  se place à la fin des scénario existants
        # on decale vers le bas les widget présent après les scenarios strategiques existant
        for x in range(numrow, r):  # pour chaque ligne jusqu'à la fin de l'atelier
            for widget in atel.grid_slaves(row=x):  # pour chaque widget de la ligne
                widget.grid(row=x+(int(field.get())-len(context["scenariosOperationnels"])))  # on le décale du nombre de scenario à inserer (lignes par scenario)
        for x in range(len(context["scenariosOperationnels"]), int(field.get())):  # pour chaque nouveau scenario
            if not "Image op"+str(x+1) in inputs.keys():  # si la cle n'existe pas dans les inputs on la créée
                inputs["Image op"+str(x+1)] = StringVar(root)
            newLabel(atel, modeleImage['label'].replace("{{ sc }}", str(x+1)), bold).grid(column=0, row=numrow)  # On place le label du fichier excel
            newEntry(atel, normal, inputs["Image op"+str(x+1)]).grid(column=1, row=numrow)  # on place le champ de saisi du fichier excel
            Button(atel, text=' search', font=bold, image=jpgIcon, compound=LEFT, command=partial(update_file, "."+modeleImage["extension"], inputs["Image op"+str(x+1)])).grid(column=2, row=numrow, padx=5)  # on place le bouton lié au champ de saisi du fichier excel
            numrow = numrow+1  # ligne suivante
            context["scenariosOperationnels"].append(x+1)  # on ajoute le scenario dans la liste des scenarios
        atel.grid(row=numpart, column=0)
    elif int(field.get()) < len(context["scenariosOperationnels"]):  # si le nombre de scenarios a crée est inférieur au nombre dejà existant
        for x in range(int(field.get()), len(context["scenariosOperationnels"])):  # pour chaque scenario en trop
            for widget in atel.grid_slaves(row=x+numrow):  # chaque widhet de la ligne correspondante
                widget.destroy()  # on detruit le widget
            if 'Image op'+str(x+1) in inputs.keys():  # Si la cle existe
                del inputs['Image op'+str(x+1)]  # suppression de la cle
                context["scenariosOperationnels"].remove(x+1)  # suppression de la liste des scenarios opérationnels
    return numrow
#####################################################################################################################
# fonction pour regénérer les champs de la fenetre


def redraw():
    list = scrollable_frame.grid_slaves()
    for l in list:
        l.destroy()
    context["scenariosStrategiques"] = []  # effacement
    context["scenariosOperationnels"] = []  # effacement
    initWin()
#####################################################################################################################
# fonction aide pour la création de label pour une Frame dans l'interface graphique. elle permet que tous les labels aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : title : texte du label
# INPUT : font : mise en forme du label (en gras ou normal)


def newLabelFrame(parent, title, font):
        return LabelFrame(parent, bd=2, relief='solid', text=title, padx=0, pady=10, font=font, background="#FFFFFF")

#####################################################################################################################
# fonction aide pour la création de label dans l'interface graphique. elle permet que tous les labels aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : title : texte du label
# INPUT : font : mise en forme du label (en gras ou normal)


def newLabel(parent, text, font):
        return Label(parent, background="#FFFFFF", text=text, width=45, font=font, anchor="e")

#####################################################################################################################
# fonction aide pour la création de label de Titre dans l'interface graphique. elle permet que tous les labels aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : title : texte du label
# INPUT : font : mise en forme du label (en gras ou normal)


def newLabelTitle(parent, text, font):
        return Label(parent, text=text, foreground="#892222", background="#CCCCCC", width=100, font=font)

#####################################################################################################################
# fonction aide pour la création d'un champ texte dans l'interface graphique. elle permet que tous les champs aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : font : mise en forme du label (en gras ou normal)
# INPUT : textvariable : variable contenant la donnée utile


def newEntry(parent, font, textvariable):
        return Entry(parent, width=95, font=font, textvariable=textvariable)

#####################################################################################################################
# Handler permettant de prendre en compte la molette de la souris pour scroller l'ecran


def MouseWheelHandler(event):
    def delta(event):
        if event.num == 5 or event.delta < 0:
            return 1
        return -1

    canvas.yview_scroll(delta(event), UNITS)

#####################################################################################################################
# fonction permettant charger dans les champs texte les valeurs sauvegardées dans un fichier json
# INPUT : inputs : tableau contenant toutes les variables des champs textes


def load_config(inputs, fileName=None):
    try:
        if not fileName:
            fileName = choose_a_file('.json')  # ouverture de l'exploration de fichier pour selectionner le fichier de config
        file = open(fileName)  # ouverture du fichier de config
        data = json.load(file)  # lectude du fichier json
        for key in data.keys():  # pour chaque clé du fichier
            if not key in inputs.keys():
                inputs[key] = StringVar()
            inputs[key].set(data[key])  # on charge la clé
        file.close()  # fermeture du fichier de configuration
        redraw()
    except:
        if thor["debug"]:  # si mode debug activ"
            sys.exc_info()[0]  # On affiche l'erreur
            raise  # levée de l'erreur
        else:
            log.insert(END, "\nWARNING : Le fichier config.json est innexistant ou invalide")

#####################################################################################################################
# fonction permettant de sauvegarder les données des champs texte dans un fichier json
# INPUT : config : tableau contenant toutes les variables des champs textes


def save_config(config):
    #récupération de la valeur des inputs
    config = dict()
    for key in inputs.keys():
        config[key] = inputs[key].get()  # conversion StringVar en String
    #sauvegarde de la configuration
    try:
        file = open(config["Config_file"], "w")  # ouverture du fchier de config
        file.write(json.dumps(config))  # ecriture du fichier de config
        file.close()  # fermeture du fichier de config
    except:
        if thor["debug"]:  # si mode debug activ"
            sys.exc_info()[0]  # On affiche l'erreur
            raise  # levée de l'erreur
        else:
            log.insert(END, "\nERROR : la sauvegarde de la configuration a échouée")
    return config


##############################################################################################
# Insertion des scenarios stratégiques dans le tableau Thor pour la generation du word
# Pour la génération du word, il faut ajouter les scenarios au bon endroit dans l'arborescence de thor
def insert_thor_strat(thor):
    global context
    for atelierKey, atelier in thor["tableaux"].items():  # niveau 2 # Pour chaque atelier
        for titleKey, title in atelier.items():  # niveau 3 # Pour chaque sous-titre
            for tabKey, tab in title.items():  # niveau 4 # Pour chaque tableau
                if tabKey == "nbScenariosStrategiques":  # si la cle est nbScenariosStrategiques
                    modele = thor["modeles"]["Scénario_stratégique"]  # on recupere le modele des scenarios strategique
                    for x in context["scenariosStrategiques"]:  # pour chaque scenario stratégique
                        title["str"+str(x)] = {}  # on créer un nouvel objet dans thor
                        for key, value in modele.items():  # Pour chaque couple cle/valeur du modele
                            title["str"+str(x)][key] = value  # on recopie la clé et la valeur
                        title["str"+str(x)]["style"] = {}  # de meme pour le style
                        for key, value in modele["style"].items():
                            title["str"+str(x)]["style"][key] = value
                        title["str"+str(x)]["keyWord"] = title["str"+str(x)]["keyWord"].replace("{{ sc }}", str(x))  # personnalise le mot cle pour le scenario en remplacant {{ sc }} par l'indice du scenario
                        title["str"+str(x)]["type"] = "file"  # on personnalise le type du scenario pour le passer de "generique" à "file"

                        # Meme principe pour l'image associee au scenario
                        modele2 = thor["modeles"]["Image_scénario_stratégique"]
                        title["Image str"+str(x)] = {}
                        for key, value in modele2.items():
                            title["Image str"+str(x)][key] = value
                        title["Image str"+str(x)]["style"] = {}
                        for key, value in modele2["style"].items():
                            title["Image str"+str(x)]["style"][key] = value
                        title["Image str"+str(x)]["keyWord"] = title["Image str"+str(x)]["keyWord"].replace("{{ sc }}", str(x))
                        title["Image str"+str(x)]["type"] = "image"
                    # une fois que l'on a inséré les scenarios on peut quitter le fonction
                    return


##############################################################################################
# Insertion des scenarios opérationnels dans le tableau Thor
# Pour la génération du word, il faut ajouter les scenarios au bon endroit dans l'arborescence de thor
# le principe est le meme que pour la fonction insert_thor_strat
def insert_thor_oper(thor):
    global context
    for atelierKey, atelier in thor["tableaux"].items():  # niveau 2
        for titleKey, title in atelier.items():  # niveau 3
            for tabKey, tab in title.items():  # niveau 4
                if tabKey == "nbScenariosOperationnels":
                    modele = thor["modeles"]["Image_scénario_opérationnel"]
                    for x in context["scenariosOperationnels"]:
                        title["Image op"+str(x)] = {}
                        for key, value in modele.items():
                            title["Image op"+str(x)][key] = value
                        title["Image op"+str(x)]["style"] = {}
                        for key, value in modele["style"].items():
                            title["Image op"+str(x)]["style"][key] = value
                        title["Image op"+str(x)]["keyWord"] = title["Image op"+str(x)]["keyWord"].replace("{{ sc }}", str(x))
                        title["Image op"+str(x)]["type"] = "image"
                    # une fois que l'on a inséré les scenarios on peut quitter le fonction
                    return

#####################################################################################################################
# fonction permettant de de lancer la generation du rapport et d'informer de resultat de la génération
# INPUT : config : tableau contenant toutes les variables des champs textes


def launch_rapport(config, log, thor):
    global context
    if not check_config(config, log):
            return False
    config = save_config(confi
    g);  # sauvegarde de la configuration
    # recherche de la configuraiton des scenarios stratefique
    insert_thor_strat(thor)  # insertion des scenario strategiques
    insert_thor_oper(thor)  # insertion des scenarios opérationnels
    nbError = generate_rapport(config, context, log, thor)  # generation du rapport
    # Message de fin
    if nbError == 0:
        messagebox.showinfo(title="Final", message="la génération du rapport est terminée avec succès")
    else:
        messagebox.showerror(title="alert", message="la génération du rapport terminée avec "+str(nbError)+" erreurs")


def check_config(config, log):
    if config["Config_file"].get() == "":
        messagebox.showerror(title="alert", message="le fichier de configuration n'est pas renseigné")
        return False
    if config["Rapport_input"].get() == "":
        messagebox.showerror(title="alert", message="le fichier Word d'entrée n'est pas renseigné")
        return False
    if config["Rapport_output"].get() == "":
        messagebox.showerror(title="alert", message="le fichier Word de sortie n'est pas renseigné")
        return False
    return True
#####################################################################################################################
## MAIN ##
#####################################################################################################################
# Récupération du répertoire d'execution du script


lastdir = os.getcwd()  # par défaut, la fenetre de recherche de fichier s'ouvre dans le repertoire courant

# création de la fenetre
root = Tk()
root.configure(background="#FFFFFF")
root.title('Generation du rapport Word')  # Ajout d'un titre
root.resizable(True, True)  # autoriser le redimensionnement vertical.

# Gestion de la molette de la souris
root.bind("<MouseWheel>", MouseWheelHandler)
root.bind("<Button-4>", MouseWheelHandler)
root.bind("<Button-5>", MouseWheelHandler)

# variables de police
big = tkFont.Font(family='Arial', size=14, weight='bold')
bold = tkFont.Font(family='Arial', size=12, weight='bold')
normal = tkFont.Font(family='Arial', size=12)


# déclaration des images utilisées
excelIcon = PhotoImage(file=r""+swd+"\images\excel.jpg")
addIcon = PhotoImage(file=r""+swd+"\images\\add.png")
wordIcon = PhotoImage(file=r""+swd+"\images\word.jpg")
jsonIcon = PhotoImage(file=r""+swd+"\images\json.jpg")
logIcon = PhotoImage(file=r""+swd+"\images\log.png")
jpgIcon = PhotoImage(file=r""+swd+"\images\jpg.png")

#Configuration de la barre de defilement
# container avec la scrollbar
container = Frame(root, background="#FFFFFF")
# Canvas avec le contenu de la page
canvas = Canvas(container, width=1500, height=700, background="#FFFFFF")
#Barre de défilement verticale
scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = Frame(canvas, background="#FFFFFF")
scrollable_frame.bind(
    "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

# Window contenant le programme (dans le canvas)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

#####################################################################################################################
#contexte pour  la generation du template Word
context = {}
context["scenariosStrategiques"] = []  # utilisatble pour le template word, contient la liste des scenarios strategique
context["scenariosOperationnels"] = []  # utilisatble pour le template word, contient la liste des scenarios opérationnels

#tableau de configuration pour la génération du rapport
inputs = dict()
inputs["Rapport_input"] = StringVar(root)  # création d'une varibale de configuration
inputs["Rapport_output"] = StringVar(root)  # création d'une varibale de configuration
inputs["Config_file"] = StringVar(root)  # création d'une varibale de configuration
inputs['LOG'] = StringVar(root)  # création d'une varibale de configuration
for partie in ["echelles", "tableaux"]:  # niveau 1
    for atelierKey, atelier in thor[partie].items():  # niveau 2
        for titleKey, title in atelier.items():  # niveau 3
            for tabKey, tab in title.items():  # niveau 4
                inputs[tabKey] = StringVar(root)  # création d'une varibale de configuration

#####################################################################################################################
# fonction de chargement de l'interface graphique


def initWin():
    numPart = 0  # Identification d'une ligne sur la grille principale de l'interface graphique
    numRow = 0  # identification d'une ligne au sein d'un atelier
    newLabelTitle(scrollable_frame, 'THOR v'+str(thor["version"])+' – Traitement Hybride pour l’Optimisation du Rapport', bold).grid(column=0, row=numPart, pady=10)
    numPart = numPart+1
    #Menu
    Button(scrollable_frame, text='Load config', image=jsonIcon, compound=LEFT, font=bold, command=partial(load_config, inputs)).grid(column=0, row=numPart)
    numPart = numPart+1

    #####################################################################################################################
    #Premiere partie, les options concernant le script
    rapport = newLabelFrame(scrollable_frame, "Rapport", bold)

    newLabel(rapport, 'Document Word en entrée: ', bold).grid(column=0, row=numRow)
    newEntry(rapport, normal, inputs["Rapport_input"]).grid(column=1, row=numRow)
    Button(rapport, text=' search', font=bold, image=wordIcon, compound=LEFT, command=partial(update_file, ".docx", inputs["Rapport_input"])).grid(column=2, row=numRow, padx=10)
    numRow = numRow+1

    newLabel(rapport, 'Document Word en sortie: ', bold).grid(column=0, row=numRow)
    newEntry(rapport, normal, inputs["Rapport_output"]).grid(column=1, row=numRow)
    Button(rapport, text=' search', font=bold, image=wordIcon, compound=LEFT, command=partial(update_file, ".docx", inputs["Rapport_output"], True)).grid(column=2, row=numRow, padx=10)
    numRow = numRow+1

    newLabel(rapport, 'Fichier de configuration: ', bold).grid(column=0, row=numRow)
    newEntry(rapport, normal, inputs["Config_file"]).grid(column=1, row=numRow)
    Button(rapport, text=' search', font=bold, image=jsonIcon, compound=LEFT, command=partial(update_file, ".json", inputs["Config_file"], True)).grid(column=2, row=numRow, padx=10)
    numRow = numRow+1

    rapport.grid(row=numPart, column=0, padx=20, pady=10)
    numPart = numPart+1

    #####################################################################################################################
    numRow = 0

    for partie in ["echelles", "tableaux"]:  # niveau 1
        for atelierKey, atelier in thor[partie].items():  # niveau 2
            atel = newLabelFrame(scrollable_frame, atelierKey, bold)
            for titleKey, title in atelier.items():  # niveau 3
                newLabelTitle(atel, titleKey, big).grid(column=0, row=numRow, columnspan=3)
                numRow = numRow+1
                for tabKey, tab in title.items():  # niveau 4
                    if tab["type"] == "file":  # pour les fichiers
                        newLabel(atel, tab["label"], bold).grid(column=0, row=numRow)  # Label
                        newEntry(atel, normal, inputs[tabKey]).grid(column=1, row=numRow)  # Champ de saisi
                        Button(atel, text=' search', font=bold, image=excelIcon, compound=LEFT, command=partial(update_file, "."+tab["extension"], inputs[tabKey])).grid(column=2, row=numRow, padx=5)  # Bouton
                    elif tab["type"] == "scénariosStrategiques":  # pour les scenarios stratégiques
                        newLabel(atel, tab["label"], bold).grid(column=0, row=numRow)  # Label
                        newEntry(atel, normal, inputs[tabKey]).grid(column=1, row=numRow)  # Champ de saisi
                        Button(atel, text=' update', font=bold, image=addIcon, compound=LEFT, command=partial(update_ihm_strat, inputs[tabKey], atel, inputs, root, numRow+1, numPart, thor["modeles"])).grid(column=2, row=numRow, padx=5)  # Bouton
                        if inputs[tabKey].get():  # Si le nombre de scenario est déjà saisi (load_config)
                            numRow = update_ihm_strat(inputs[tabKey], atel, inputs, root, numRow+1, numPart, thor["modeles"])  # on insere les scenarios
                        numRow = numRow+1
                    elif tab["type"] == "scénariosOperationnels":  # pour les scenarios opérationnels
                        newLabel(atel, tab["label"], bold).grid(column=0, row=numRow)  # Label
                        newEntry(atel, normal, inputs[tabKey]).grid(column=1, row=numRow)  # Champ de saisi
                        Button(atel, text=' update', font=bold, image=addIcon, compound=LEFT, command=partial(update_ihm_oper, inputs[tabKey], atel, inputs, root, numRow+1, numPart, thor["modeles"])).grid(column=2, row=numRow, padx=5)  # Bouton
                        if inputs[tabKey].get():  # Si le nombre de scenario est déjà saisi (load_config)
                            numRow = update_ihm_oper(inputs[tabKey], atel, inputs, root, numRow+1, numPart, thor["modeles"])  # on insere les scenarios
                        numRow = numRow+1
                    elif tab["type"] == "image":  # pour les images
                        newLabel(atel, tab["label"], bold).grid(column=0, row=numRow)  # Label
                        newEntry(atel, normal, inputs[tabKey]).grid(column=1, row=numRow)  # Champ de saisi
                        Button(atel, text=' search', font=bold, image=jpgIcon, compound=LEFT, command=partial(update_file, "."+tab["extension"], inputs[tabKey])).grid(column=2, row=numRow, padx=5)  # Bouton
                    numRow = numRow+1  # ligne suivante de l'atelier
            atel.grid(row=numPart, column=0)  # affichage de l'atelier
            numPart = numPart+1  # ligne suivante sur la grille principale

    #####################################################################################################################
    numrow = 0  # positionnement au debut de l'atelier
    atel = newLabelFrame(scrollable_frame, "Journalisation", bold)  # creation d'un atelier journalisation
    log = Text(atel, width=100, height=20)  # Création d'un champ texte qui contiendra les journaux
    Button(scrollable_frame, text=' Generate', image=wordIcon, compound=LEFT, font=bold, command=partial(launch_rapport, inputs, log, thor)).grid(column=0, row=numPart, pady=10)  # Bouton de generation du rapport
    numPart = numPart+1  # ligne suivante sur la grille principale
    Label(atel, image=logIcon).grid(column=0, row=numrow, pady=10, padx=20)  # Label
    log.grid(column=1, row=numrow, pady=10)  # affichage des journaux
    numrow = numrow+1  # ligne suivante
    atel.grid(row=numPart, column=0)  # affichage de l'atelier
    numPart = numPart+1  # ligne suivante sur la grille principale
    #####################################################################################################################
    container.pack()  # affichage du container
    canvas.pack(side="left", fill="both", expand=True)  # affichage de l'interface graphique
    scrollbar.pack(side="right", fill="y")  # affichage de la barre de défilement


#####################################################################################################################
initWin()  # initialisation de la fenetre
root.mainloop()  # main loop
