# THOR
Python : Script d'import de tableau Excel dans un document Word
## Licence
<table><tr>
<td><strong>Creative Commons<strong> </td><td>Pas d'utilisation commerciale <br>
Partage dans les mêmes conditions</td><td> <a href="https://creativecommons.org/licenses/by-nc-nd/4.0/"><img src="images/licence.png?raw=true"/></a> </td>
</tr></table>
  
##	Présentation
###	Constat
Lors de la réalisation d’une étude de risque avec la méthode EBIOS RM, l’utilisation d’une application dédiée facilite la production des différents livrables au format  Excel et images mais ne permet pas de réaliser un modèle de rapport personnalisé.
### Besoin
Notre retour d’expérience montre qu’il est difficile d’échanger avec un client en utilisant des fichiers Excel, il est plus aisé d’utiliser un fichier Word contenant des éléments de contexte associés à chaque tableau et permettant d’utiliser le mode révision pour le suivi des modifications ainsi que les commentaires. Il a donc fallu trouver une solution permettant d’associer les fichiers Excel issus de l’application et un modèle de rapport Word.
THOR est un script, réalisé en python 3, prenant en entrée un document Word ainsi qu’un fichier de configuration (en JSON) permettant d’associer chaque tableau du document Word à un fichier Excel.
THOR va alors copier chaque tableau Excel dans le document Word en conservant la mise en forme du tableau Excel. Les entêtes sont ignorés afin de pouvoir dissocier les entêtes du tableau Word de celles du tableau Excel issues de l’application et qui ne sont pas personnalisables. De ce fait, le nombre et l’ordre des colonnes entre les tableaux Word et Excel doivent correspondre.
## 	Mise en place des sources python
THOR étant un script python il est cross-plateforme et fonctionne donc aussi bien sous Windows que sous Linux. Il nécessite en revanche l’installation de certains modules python. En fonction des l’installation de base dont vous disposez, certaines dépendances peuvent ne pas être listées infra.
### 	Windows
*	PYTHON 3.8 ou supérieur avec le module TCL et pylauncher activés lors de l’installation 
###	Linux
*	PYTHON 3.8 ou supérieur
*	PYTHON3-TK
*	PYTHON-PIP
*	Interface graphique
### Modules python a installé avec PIP
*	PYTHON-DOCX => Gestion des fichiers DOCX
*	XLRD => gestion des fichiers XLS
*	COLORAMA = gestion des couleurs
*	PYYAML
*	DOCXTPL => gestion de templates word
## 	Mise en place de la version compilée
Sous Windows, il est possible d'utiliser la version compilée. Pour ce faire, telecherger le zip présent dans le repertoire EXE. Après avoir extrait les fichiers du zip, executer le fichier THOR.exe. Cette version ne necessite aucune installation.
