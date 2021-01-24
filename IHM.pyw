import json
import tkinter.font as tkfont
from functools import partial
from tkinter import *
from tkinter import filedialog, messagebox
from yaml import Loader, load
from tools import *

###############################################################
swd = os.path.dirname(os.path.realpath(sys.argv[0]))
# Chargement du fichier YAML
f = open(swd + "/theme.yaml", 'r', encoding='utf8')
thor = load(f, Loader=Loader)
f.close()


###############################################################
# fenetre explorateur de fichier pour la selection d'un nom de fichier
# INPUT : extension : extension autorisées pour la selection de fichier


def choose_a_file(extension):
    global lastdir
    options = {'defaultextension': extension,
               'filetypes': [(extension.upper(), extension.lower())],
               'initialdir': lastdir,
               'initialfile': '',
               'title': 'Selectionnez un fichier'}

    chosenfile = filedialog.askopenfilename(**options)
    if chosenfile:
        lastdir = os.path.dirname(chosenfile)
        return chosenfile
    else:
        return ''


###############################################################
# fenetre explorateur de fichier pour la selection d'un nom de fichier
# INPUT : extension : extension autorisées pour la selection de fichier


def choose_a_filename(extension):
    global lastdir
    options = {'defaultextension': extension,
               'filetypes': [(extension.upper(), extension.lower())],
               'initialdir': lastdir,
               'initialfile': '',
               'title': 'Choisissez un nom de fichier'}

    chosenfile = filedialog.asksaveasfilename(**options)
    if chosenfile:
        lastdir = os.path.dirname(chosenfile)
        return chosenfile
    else:
        return ''


###############################################################
# fonction permettant de selectionner un nom de fichier
# puis de mettre à jour le champ correspondant dans l'interface graphique
# INPUT : extension : extension attendue du fichier
# INPUT : field : champ de l'interface à nmettre à jour


def update_file(extension, field, saveas=False):
    if saveas:
        filename = choose_a_filename(extension)
    else:
        filename = choose_a_file(extension)
    if filename:
        field.set(filename)


###############################################################
# fonction permettant de mettre à jour l'iHM en crééant
# de nouveau champs pour les scenarios strategiques
# INPUT : field : pointeur vers la velaur du champ à modifier
# INPUT : atel : contenaire des champs
# INPUT : root : identifiant de la fenetre
# INPUT : numrow : positionnement vertical dans le container
# INPUT : numpart : numpart positionnement vertical de l'atel dans la fenetre


def update_ihm_strat(field, atel, inputs, root, numrow, numpart, mod):
    # taille de l'atelier en colonne et ligne
    c, r = atel.grid_size()
    # modele des scenario strategiques
    modele = mod["Scénario_stratégique"]
    # modele des imges associées au scenarios stratégiques
    modeleimage = mod["Image_scénario_stratégique"]
    # Si le nombre de scenario a crée est supérieur à celui existant
    if int(field.get()) > len(context["scenariosStrategiques"]):
        # On se place apres les scenarios existants
        numrow = numrow + 2 * len(context["scenariosStrategiques"])
        # On decale vers le bas les widget présent après les scenarios
        # pour chaque ligne jusqu'à la fin de l'atelier
        for x in range(numrow, r):
            # pour chaque widget de la ligne
            for widget in atel.grid_slaves(row=x):
                # on décale les widgets existants
                widget.grid(row=x + 2 * (int(field.get()) -
                                         len(context["scenariosOperationnels"])))
        # Pour chaque scenario a créér
        for x in range(len(context["scenariosStrategiques"]),
                       int(field.get())):
            # si la cle n'existe pas
            if not "str" + str(x + 1) in inputs.keys():
                inputs["str" + str(x + 1)] = StringVar(root)
            # si la cle n'existe pas
            if not "Image str" + str(x + 1) in inputs.keys():
                inputs["Image str" + str(x + 1)] = StringVar(root)
            newlabel(atel, modele['label'].replace("{{ sc }}",
                                                   str(x + 1)), bold).grid(
                column=0, row=numrow)  # positionnement
            newentry(atel, normal, inputs["str" + str(x + 1)]).grid(
                column=1, row=numrow)  # positionnement
            Button(atel, text=' search', font=bold,
                   image=excelicon, compound=LEFT,
                   command=partial(update_file, "." + modele["extension"],
                                   inputs["str" + str(x + 1)])).grid(
                column=2, row=numrow, padx=5)
            numrow = numrow + 1  # ligne suivante
            newlabel(atel, modeleimage['label'].replace("{{ sc }}",
                                                        str(x + 1)), bold).grid(
                column=0, row=numrow)  # on créer le label de l'image
            newentry(atel, normal, inputs["Image str" + str(x + 1)]).grid(
                column=1, row=numrow)  # positionnement
            Button(atel, text=' search',
                   font=bold,
                   image=jpgicon,
                   compound=LEFT,
                   command=partial(update_file, "." + modeleimage["extension"],
                                   inputs["Image str" + str(x + 1)])).grid(
                column=2, row=numrow, padx=5)

            numrow = numrow + 1  # ligne suivante
            # on ajoute le scenario à la liste des scenarios
            context["scenariosStrategiques"].append(x + 1)
        atel.grid(row=numpart, column=0)
    # Si le nombre de scenario à créér est inférieur au nombre existant
    elif int(field.get()) < len(context["scenariosStrategiques"]):
        # pour chaque scnenario en trop
        for x in range(int(field.get()), 2 * len(
                context["scenariosStrategiques"])):
            # pour chaque widget de la ligne correspondante
            for widget in atel.grid_slaves(row=x + 2 + numrow):
                widget.destroy()  # on detruit le widget
            if 'str' + str(x + 1) in inputs.keys():  # si la cle existe
                # destruction de la cle du fichier excel
                del inputs['str' + str(x + 1)]
                # destruction de la cle de l'image
                del inputs['Image str' + str(x + 1)]
                # suppression du scenario dans la liste des scenarios
                context["scenariosStrategiques"].remove(x + 1)
    # on retourne le numero de ligne pour continuer l'affichage de l'IHM
    return numrow


###############################################################
# fonction permettant de mettre à jour l'iHM en crééant de nouveaux
# champs pour les scenarios opérationnels
# INPUT : field : pointeur vers la velaur du champ à modifier
# INPUT : atel : contenaire des champs
# INPUT : root : identifiant de la fenetre
# INPUT : numrow : positionnement vertical dans le container
# INPUT : numpart : numpart positionnement vertical de l'atel dans la fenetre


def update_ihm_oper(field, atel, inputs, root, numrow, numpart, mod):
    # on recupere le nmobre de ligne et de ccolonne de
    # l'atelier dans l'interface graphique
    c, r = atel.grid_size()
    # on recupere le modele pour les image
    modeleimage = mod["Image_scénario_opérationnel"]
    # Si le nombre de scenario a crée est supérieur à celui existant
    if int(field.get()) > len(context["scenariosOperationnels"]):
        # on  se place à la fin des scénario existants
        numrow = numrow + len(context["scenariosOperationnels"])
        # on decale vers le bas les widget présent
        # après les scenarios strategiques existant.
        # Pour chaque ligne jusqu'à la fin de l'atelier
        for x in range(numrow, r):
            # pour chaque widget de la ligne
            for widget in atel.grid_slaves(row=x):
                # on le décale du nombre de scenario à inserer
                widget.grid(row=x + (int(field.get()) - len(
                    context["scenariosOperationnels"])))
        # pour chaque nouveau scenario
        for x in range(len(context["scenariosOperationnels"]),
                       int(field.get())):
            # si la cle n'existe pas dans les inputs on la créée
            if not "Image op" + str(x + 1) in inputs.keys():
                inputs["Image op" + str(x + 1)] = StringVar(root)
            # On place le label du fichier excel
            newlabel(atel, modeleimage['label'].replace("{{ sc }}", str(x + 1)),
                     bold).grid(column=0, row=numrow)
            # on place le champ de saisi du fichier excel
            newentry(atel, normal,
                     inputs["Image op" + str(x + 1)]).grid(column=1,
                                                           row=numrow)
            Button(atel, text=' search',
                   font=bold,
                   image=jpgicon,
                   compound=LEFT,
                   command=partial(update_file, "." + modeleimage["extension"],
                                   inputs["Image op" + str(x + 1)])).grid(
                column=2, row=numrow, padx=5)  # positionnement
            numrow = numrow + 1  # ligne suivante
            # on ajoute le scenario dans la liste des scenarios
            context["scenariosOperationnels"].append(x + 1)
        atel.grid(row=numpart, column=0)
    # si le nombre de scenarios a crée est inférieur au nombre dejà existant
    elif int(field.get()) < len(context["scenariosOperationnels"]):
        # pour chaque scenario en trop
        for x in range(int(field.get()), len(
                context["scenariosOperationnels"])):
            # chaque widhet de la ligne correspondante
            for widget in atel.grid_slaves(row=x + numrow):
                widget.destroy()  # on detruit le widget
            if 'Image op' + str(x + 1) in inputs.keys():  # Si la cle existe
                del inputs['Image op' + str(x + 1)]  # suppression de la cle
                # suppression de la liste des scenarios opérationnels
                context["scenariosOperationnels"].remove(x + 1)
    return numrow


###############################################################
# fonction pour regénérer les champs de la fenetre


def redraw(log):
    list = scrollable_frame.grid_slaves()
    for l in list:
        l.destroy()
    context["scenariosStrategiques"] = []  # effacement
    context["scenariosOperationnels"] = []  # effacement
    initwin()


###############################################################
# fonction aide pour la création de label pour une
# Frame dans l'interface graphique.
# elle permet que tous les labels aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : title : texte du label
# INPUT : font : mise en forme du label (en gras ou normal)


def newlabelframe(parent, title, font):
    return LabelFrame(parent,
                      bd=2,
                      relief='solid',
                      text=title,
                      padx=0,
                      pady=10,
                      font=font,
                      background="#FFFFFF")


###############################################################
# fonction aide pour la création de label dans l'interface graphique.
# elle permet que tous les labels aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : title : texte du label
# INPUT : font : mise en forme du label (en gras ou normal)


def newlabel(parent, text, font):
    return Label(parent,
                 background="#FFFFFF",
                 text=text, width=45,
                 font=font,
                 anchor="e")


###############################################################
# fonction aide pour la création de label de Titre dans l'interface graphique.
# elle permet que tous les labels aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : title : texte du label
# INPUT : font : mise en forme du label (en gras ou normal)


def newlabeltitle(parent, text, font):
    return Label(parent,
                 text=text,
                 foreground="#892222",
                 background="#CCCCCC",
                 width=100,
                 font=font)


###############################################################
# fonction aide pour la création d'un champ texte dans l'interface graphique.
# elle permet que tous les champs aient la même configuration
# INPUT : parent : parent contenant le nouveau label
# INPUT : font : mise en forme du label (en gras ou normal)
# INPUT : textvariable : variable contenant la donnée utile


def newentry(parent, font, textvariable):
    return Entry(parent, width=95, font=font, textvariable=textvariable)


###############################################################
# Handler permettant de prendre en compte
# la molette de la souris pour scroller l'ecran


def mousewheelhandler(event):
    def delta(event):
        if event.num == 5 or event.delta < 0:
            return 1
        return -1

    canvas.yview_scroll(delta(event), UNITS)


###############################################################
# fonction permettant charger dans les champs texte les
# valeurs sauvegardées dans un fichier json
# INPUT : inputs : tableau contenant toutes les variables des champs textes


def load_config(log, inputs, filename=None):
    try:
        if not filename:
            # ouverture de l'exploration de fichier pour
            # selectionner le fichier de config
            filename = choose_a_file('.json')
        file = open(filename)  # ouverture du fichier de config
        data = json.load(file)  # lectude du fichier json
        for key in data.keys():  # pour chaque clé du fichier
            if key not in inputs.keys():
                inputs[key] = StringVar()
            inputs[key].set(data[key])  # on charge la clé
        file.close()  # fermeture du fichier de configuration
        redraw(log)
    except:
        if thor["debug"]:  # si mode debug active"
            sys.exc_info()[0]  # On affiche l'erreur
            raise  # levée de l'erreur
        else:
            log.insert(END, "\nWARNING : \
            Le fichier config.json est innexistant ou invalide")


###############################################################
# fonction permettant de sauvegarder les données des
# champs texte dans un fichier json
# INPUT : config : tableau contenant toutes les variables des champs textes


def save_config(log, config):
    # récupération de la valeur des inputs
    config = dict()
    for key in inputs.keys():
        config[key] = inputs[key].get()  # conversion StringVar en String
    # sauvegarde de la configuration
    try:
        # ouverture du fchier de config
        file = open(config["Config_file"], "w")
        file.write(json.dumps(config))  # ecriture du fichier de config
        file.close()  # fermeture du fichier de config
    except:
        if thor["debug"]:  # si mode debug activ"
            sys.exc_info()[0]  # On affiche l'erreur
            raise  # levée de l'erreur
        else:
            log.insert(END, "\nERROR : la sauvegarde \
                de la configuration a échouée")
    return config


###############################################################
# Insertion des scenarios stratégiques dans le tableau
# Thor pour la generation du word
# Pour la génération du word, il faut ajouter les scenarios
# au bon endroit dans l'arborescence de thor
def insert_thor_strat(thor):
    global context
    # niveau 2 # Pour chaque atelier
    for _atelierkey, atelier in thor["tableaux"].items():
        # niveau 3 # Pour chaque sous-titre
        for _titlekey, title in atelier.items():
            # niveau 4 # Pour chaque tableau
            for tabkey, _tab in title.items():
                # si la cle est nbScenariosStrategiques
                if tabkey == "nbScenariosStrategiques":
                    # on recupere le modele des scenarios strategique
                    modele = thor["modeles"]["Scénario_stratégique"]
                    # pour chaque scenario stratégique
                    for x in context["scenariosStrategiques"]:
                        # on créer un nouvel objet dans thor
                        title["str" + str(x)] = {}
                        # Pour chaque couple cle/valeur du modele
                        for key, value in modele.items():
                            # on recopie la clé et la valeur
                            title["str" + str(x)][key] = value
                        # de meme pour le style
                        title["str" + str(x)]["style"] = {}
                        for key, value in modele["style"].items():
                            title["str" + str(x)]["style"][key] = value
                        # personnalise le mot cle pour le scenario en
                        # remplacant {{ sc }} par l'indice du scenario
                        title["str" + str(x)]["keyWord"] = \
                            title["str" + str(x)]["keyWord"].replace(
                                "{{ sc }}", str(x))
                        # on personnalise le type du scenario pour le
                        # passer de "generique" à "file"
                        title["str" + str(x)]["type"] = "file"

                        # Meme principe pour l'image associee au scenario
                        modele2 = thor["modeles"]["Image_scénario_stratégique"]
                        title["Image str" + str(x)] = {}
                        for key, value in modele2.items():
                            title["Image str" + str(x)][key] = value
                        title["Image str" + str(x)]["style"] = {}
                        for key, value in modele2["style"].items():
                            title["Image str" + str(x)]["style"][key] = value
                        title["Image str" + str(x)]["keyWord"] = \
                            title["Image str" + str(x)]["keyWord"].replace(
                                "{{ sc }}", str(x))
                        title["Image str" + str(x)]["type"] = "image"
                    # une fois que l'on a inséré les scenarios
                    # on peut quitter le fonction
                    return


###############################################################
# Insertion des scenarios opérationnels dans le tableau Thor
# Pour la génération du word, il faut ajouter les scenarios
# au bon endroit dans l'arborescence de thor
# le principe est le meme que pour la fonction insert_thor_strat
def insert_thor_oper(thor):
    global context
    for _atelierkey, atelier in thor["tableaux"].items():  # niveau 2
        for _titlekey, title in atelier.items():  # niveau 3
            for tabkey, _tab in title.items():  # niveau 4
                if tabkey == "nbScenariosOperationnels":
                    modele = thor["modeles"]["Image_scénario_opérationnel"]
                    for x in context["scenariosOperationnels"]:
                        title["Image op" + str(x)] = {}
                        for key, value in modele.items():
                            title["Image op" + str(x)][key] = value
                        title["Image op" + str(x)]["style"] = {}
                        for key, value in modele["style"].items():
                            title["Image op" + str(x)]["style"][key] = value
                        title["Image op" + str(x)]["keyWord"] = \
                            title["Image op" + str(x)]["keyWord"].replace(
                                "{{ sc }}", str(x))
                        title["Image op" + str(x)]["type"] = "image"
                    # une fois que l'on a inséré les scenarios
                    # on peut quitter le fonction
                    return


###############################################################
# fonction permettant de de lancer la generation du rapport
# et d'informer de resultat de la génération
# INPUT : config : tableau contenant toutes les variables des champs textes


def launch_rapport(config, log, thor):
    global context
    if not check_config(config):
        return False
    config = save_config(log, config)
    # sauvegarde de la configuration
    # recherche de la configuraiton des scenarios stratefique
    insert_thor_strat(thor)  # insertion des scenario strategiques
    insert_thor_oper(thor)  # insertion des scenarios opérationnels
    # generation du rapport
    nberror = generate_rapport(config, context, log, thor)
    # Message de fin
    if nberror == 0:
        messagebox.showinfo(title="Final", message=
        "la génération du rapport est terminée avec succès")
    else:
        messagebox.showerror(title="alert", message=
        "la génération du rapport terminée avec " + str(nberror) + " erreurs")
    return True


def check_config(config):
    if config["Config_file"].get() == "":
        messagebox.showerror(title="alert", message=
        "le fichier de configuration n'est pas renseigné")
        return False
    if config["Rapport_input"].get() == "":
        messagebox.showerror(title="alert", message=
        "le fichier Word d'entrée n'est pas renseigné")
        return False
    if config["Rapport_output"].get() == "":
        messagebox.showerror(title="alert", message=
        "le fichier Word de sortie n'est pas renseigné")
        return False
    return True


###############################################################
## MAIN ##
###############################################################
# Récupération du répertoire d'execution du script
# par défaut, la fenetre de recherche de fichier


# s'ouvre dans le repertoire courant
lastdir = os.getcwd()

# création de la fenetre
root = Tk()

root.configure(background="#FFFFFF")
root.title('Generation du rapport Word')  # Ajout d'un titre
root.resizable(True, True)  # autoriser le redimensionnement vertical.

# Gestion de la molette de la souris
root.bind("<MouseWheel>", mousewheelhandler)
root.bind("<Button-4>", mousewheelhandler)
root.bind("<Button-5>", mousewheelhandler)

# variables de police
big = tkfont.Font(family='Arial', size=14, weight='bold')
bold = tkfont.Font(family='Arial', size=12, weight='bold')
normal = tkfont.Font(family='Arial', size=12)

# déclaration des images utilisées
excelicon = PhotoImage(file=r"" + swd + "\images\excel.jpg")
addicon = PhotoImage(file=r"" + swd + "\images\\add.png")
wordicon = PhotoImage(file=r"" + swd + "\images\word.jpg")
jsonicon = PhotoImage(file=r"" + swd + "\images\json.jpg")
logicon = PhotoImage(file=r"" + swd + "\images\log.png")
jpgicon = PhotoImage(file=r"" + swd + "\images\jpg.png")

# Configuration de la barre de defilement
# container avec la scrollbar
container = Frame(root, background="#FFFFFF")
# Canvas avec le contenu de la page
canvas = Canvas(container, width=1500, height=700, background="#FFFFFF")
# Barre de défilement verticale
scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = Frame(canvas, background="#FFFFFF")
scrollable_frame.bind(
    "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

# Window contenant le programme (dans le canvas)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

###############################################################
# contexte pour  la generation du template Word
context = {"scenariosStrategiques": [],
           "scenariosOperationnels": []}
# utilisatble pour le template word,
# contient la liste des scenarios strategique
# contient la liste des scenarios opérationnels

# tableau de configuration pour la génération du rapport
inputs = dict()
# création d'une varibale de configuration
inputs["Rapport_input"] = StringVar(root)
# création d'une varibale de configuration
inputs["Rapport_output"] = StringVar(root)
# création d'une varibale de configuration
inputs["Config_file"] = StringVar(root)
# création d'une varibale de configuration
inputs['LOG'] = StringVar(root)
for partie in ["echelles", "tableaux"]:  # niveau 1
    for _atelierkey, atelier in thor[partie].items():  # niveau 2
        for _titlekey, title in atelier.items():  # niveau 3
            for tabkey, _tab in title.items():  # niveau 4
                # création d'une varibale de configuration
                inputs[tabkey] = StringVar(root)


###############################################################
# fonction de chargement de l'interface graphique


def initwin():
    ###############################################################
    # Identification d'une ligne sur la grille
    # principale de l'interface graphique
    numpart = 0
    numrow = 0  # identification d'une ligne au sein d'un atelier
    ###############################################################
    numrow = 0  # positionnement au debut de l'atelier
    # creation d'un atelier journalisation
    journaux = newlabelframe(scrollable_frame, "Journalisation", bold)
    # Création d'un champ texte qui contiendra les journaux
    log = Text(journaux, width=100, height=20)

    Label(journaux, image=logicon).grid(column=0, row=numrow, pady=10, padx=20)
    log.grid(column=1, row=numrow, pady=10)  # affichage des journaux
    numrow = numrow + 1  # ligne suivante
    # les journaux seront positionnés en bas de la fenetre, donc le positionnement
    # de l'atelier se fera en dernier
    ###############################################################
    newlabeltitle(scrollable_frame,
                  'THOR v' + str(thor["version"]) +
                  ' – Traitement Hybride pour l’Optimisation du Rapport', bold).grid(
        column=0, row=numpart, pady=10)
    numpart = numpart + 1
    # Menu
    Button(scrollable_frame, text='Load config',
           image=jsonicon, compound=LEFT,
           font=bold,
           command=partial(load_config, log, inputs)).grid(
        column=0, row=numpart)
    numpart = numpart + 1

    ###############################################################
    # Premiere partie, les options concernant le script
    rapport = newlabelframe(scrollable_frame, "Rapport", bold)

    newlabel(rapport, 'Document Word en entrée: ', bold).grid(
        column=0, row=numrow)
    newentry(rapport, normal, inputs["Rapport_input"]).grid(
        column=1, row=numrow)
    Button(rapport, text=' search',
           font=bold,
           image=wordicon,
           compound=LEFT,
           command=partial(update_file, ".docx",
                           inputs["Rapport_input"])).grid(
        column=2, row=numrow, padx=10)
    numrow = numrow + 1

    newlabel(rapport, 'Document Word en sortie: ', bold).grid(
        column=0, row=numrow)
    newentry(rapport, normal, inputs["Rapport_output"]).grid(
        column=1, row=numrow)
    Button(rapport, text=' search',
           font=bold,
           image=wordicon,
           compound=LEFT,
           command=partial(update_file, ".docx",
                           inputs["Rapport_output"], True)).grid(
        column=2, row=numrow, padx=10)
    numrow = numrow + 1

    newlabel(rapport, 'Fichier de configuration: ', bold).grid(
        column=0, row=numrow)
    newentry(rapport,
             normal,
             inputs["Config_file"]).grid(
        column=1, row=numrow)
    Button(rapport,
           text=' search',
           font=bold,
           image=jsonicon,
           compound=LEFT,
           command=partial(update_file, ".json",
                           inputs["Config_file"], True)).grid(
        column=2, row=numrow, padx=10)
    numrow = numrow + 1

    rapport.grid(row=numpart, column=0, padx=20, pady=10)
    numpart = numpart + 1

    ###############################################################
    numrow = 0

    for partie in ["echelles", "tableaux"]:  # niveau 1
        for atelierkey, atelier in thor[partie].items():  # niveau 2
            atel = newlabelframe(scrollable_frame, atelierkey, bold)
            for titlekey, title in atelier.items():  # niveau 3
                newlabeltitle(atel, titlekey, big).grid(
                    column=0, row=numrow, columnspan=3)
                numrow = numrow + 1
                for tabkey, tab in title.items():  # niveau 4
                    if tab["type"] == "file":  # pour les fichiers
                        newlabel(atel, tab["label"], bold).grid(
                            column=0, row=numrow)  # Label
                        newentry(atel, normal, inputs[tabkey]).grid(
                            column=1, row=numrow)  # Champ de saisi
                        Button(atel,
                               text=' search',
                               font=bold,
                               image=excelicon,
                               compound=LEFT,
                               command=partial(update_file, "." + tab["extension"],
                                               inputs[tabkey])).grid(
                            column=2, row=numrow, padx=5)  # Bouton
                    # pour les scenarios stratégiques
                    elif tab["type"] == "scénariosStrategiques":
                        newlabel(atel, tab["label"], bold).grid(
                            column=0, row=numrow)  # Label
                        newentry(atel, normal, inputs[tabkey]).grid(
                            column=1, row=numrow)  # Champ de saisi
                        Button(atel,
                               text=' update',
                               font=bold,
                               image=addicon,
                               compound=LEFT,
                               command=partial(update_ihm_strat,
                                               inputs[tabkey],
                                               atel,
                                               inputs,
                                               root,
                                               numrow + 1,
                                               numpart,
                                               thor["modeles"])).grid(
                            column=2, row=numrow, padx=5)  # Bouton
                        # Si le nombre de scenario est déjà saisi (load_config)
                        if inputs[tabkey].get():
                            numrow = update_ihm_strat(inputs[tabkey],
                                                      atel,
                                                      inputs,
                                                      root,
                                                      numrow + 1,
                                                      numpart,
                                                      thor["modeles"])  # on insere les scenarios
                        numrow = numrow + 1
                    # pour les scenarios opérationnels
                    elif tab["type"] == "scénariosOperationnels":
                        newlabel(atel, tab["label"], bold).grid(
                            column=0, row=numrow)  # Label
                        newentry(atel, normal, inputs[tabkey]).grid(
                            column=1, row=numrow)  # Champ de saisi
                        Button(atel,
                               text=' update',
                               font=bold,
                               image=addicon,
                               compound=LEFT,
                               command=partial(update_ihm_oper, inputs[tabkey],
                                               atel,
                                               inputs,
                                               root,
                                               numrow + 1,
                                               numpart,
                                               thor["modeles"])).grid(
                            column=2, row=numrow, padx=5)
                        # Si le nombre de scenario est déjà saisi (load_config)
                        if inputs[tabkey].get():
                            numrow = update_ihm_oper(inputs[tabkey],
                                                     atel,
                                                     inputs,
                                                     root,
                                                     numrow + 1,
                                                     numpart,
                                                     thor["modeles"])  # on insere les scenarios
                        numrow = numrow + 1
                    elif tab["type"] == "image":  # pour les images
                        newlabel(atel, tab["label"], bold).grid(
                            column=0, row=numrow)  # Label
                        newentry(atel, normal, inputs[tabkey]).grid(
                            column=1, row=numrow)  # Champ de saisi
                        Button(atel,
                               text=' search',
                               font=bold,
                               image=jpgicon,
                               compound=LEFT,
                               command=partial(update_file, "." + tab["extension"],
                                               inputs[tabkey])).grid(
                            column=2, row=numrow, padx=5)
                    numrow = numrow + 1  # ligne suivante de l'atelier
            atel.grid(row=numpart, column=0)  # affichage de l'atelier
            numpart = numpart + 1  # ligne suivante sur la grille principale

    ###############################################################
    # positionnement du bouton pour générer le rapport
    Button(scrollable_frame,
           text=' Generate',
           image=wordicon,
           compound=LEFT,
           font=bold,
           command=partial(launch_rapport, inputs, log, thor)).grid(
        column=0, row=numpart, pady=10)
    numpart = numpart + 1  # ligne suivante sur la grille principale
    # positionnement des journaux
    journaux.grid(row=numpart, column=0)  # affichage de l'atelier
    numpart = numpart + 1  # ligne suivante sur la grille principale
    ###############################################################
    container.pack()  # affichage du container
    # affichage de l'interface graphique
    canvas.pack(side="left", fill="both", expand=True)
    # affichage de la barre de défilement
    scrollbar.pack(side="right", fill="y")


###############################################################
initwin()  # initialisation de la fenetre
root.mainloop()  # main loop
