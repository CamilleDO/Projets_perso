Projet permettant :
- Création de documents à chaque modifications
- Archivage des documents 
- Paramètrage des infos de l'entreprise
- Numérotation automatique des documents
- Visiblité de l'état de paiement et modification rapide de ceux-ci
- Fonctionnalité d'impression rapide des documents (PDF) 


1) Demande de prix par client
2) Devis par fournisseur
3) Bon de commande par client
4) Marchandise + bon de livraison par fournisseur	
5) Facture par fournisseur (pièce qui sera comptabilisé)
6) Paiement (chèque, CB, espèce...)


1ère fonctionnalité de l'application:
L'application doit avoir une numérotation des documents de manière automatique, chronologique, continue et sans rupture pour chaque document.
"F-2023-01-0001"
D-Devis / C-Bon de commande / L-Bon de livraison / F-facture
Année
Mois
Numéro : devra repartir a 0 (0001) au début de l'année 2024

2ème fonctionnalité :
Archivage des documents dans des dossiers spécifiques
Il faut d'abord créer un fichier documents puis dedans créer un dossier pour commande / devis / facture / livraison


AFFICHER CLIENT sous forme de liste
Fonction INDIRECT (Dans Données/validation des données) autoriser: liste / source=
Source:   =INDIRECT("clients[Nom]") --> page para: tableau Clients / col nom

par la suite faire correspondre adresse et n° de tel des clients : Fonction RECHERCHEV (verticale) 
=RECHERCHEV(cellule afficher client;Clients;2; FAUX -->Adress:  ..V(cellule où on veut que ça corresponde J13; tableau Client page para; colonne 2 pour adresse; faux car correspondance exacte)
=RECHERCHEV(cellule afficher client;Clients;3; FAUX -->N° tel : ..V(cellule où on veut que ça corresponde J13; tableau Client page para; colonne 3 pour n° tel; faux car correspondance exacte)

Pareil pour tableau de remplissage devis : mettre formule indirect etc puis réf / prix mettre formule recherche V
ATTENTION))))) En revanche on voit des messages d'erreur pour quand le champs est vide :
Solution ---> mettre SIERREUR(RECHERCHEV....);"") mettre "" = rien)

Pour le texte en bas :
Formule =CONCATENER : bien fusionner les celulles -- problème pour saut de ligne "....";CAR(10);"Sinon appelez-nous au "; SELECTION DU TEL DS PARA)

Pour les boutons : copier coller Redimenssionner couleur ect
Aligner à gauche 
Aligner vertical

Insertion macro : 
Mettre onglet développeur (fichier/option/personnaliser ruban/cliquer développeur)
Onglet développeur/visual basic/CLIQUE DROIT sur VBAProject (Application)/Insertion/Module 
MACRO Annuler :
Sub annuler_devis()
    Sheets("Devis").Range("F20:F32").Value = ""
    Sheets("Devis").Range("H20:H32").Value = ""
    Sheets("Devis").Range("J20:J32").Value = ""
    Sheets("Devis").Range("I13").Value = ""
End Sub

Croix rouge / clique droit bouton annuler/Associer une macro/selectionner nom de la macro : annuler_devis/test

Voir les autres fichiers text : annuler / nous / impression
Attention pour chaque nouvelle feuille il faut adapter les macro les refaire dans un nouveau module 

Faire des listes : sélectionner une cell / données / validation des données / selectionner autoriser=liste / dans source mettre tous les mots que l'on veut voir apparaitre séparé d'un ';' 
Possibilité dans les autres onglets de mettre des mots si on passe la souris dessus ou de mettre un msg d'erreur personnalisé si erreur!
Mise en forme conditionnelle pour état de paiement = en gros le client et les info de la facture seront pré-enregistré et de manière automatique l'état de paiement de la facture sera en impayée et en rouge ce qui faudra modifier manuellement pour mettre payé et précisé le mode de paiement etc...


!!!!!! SI PROBLEME LIGNE IMPRESSION !!!!!!!!!
Vérif si code ok
Sinon aller dans onglet mise en page / largeur et hauteur mettre 1 page au lieu de automatique 