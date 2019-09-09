
var colonnesAjoutees = ['Mail','Nom','Prenom','Total'];
var colonnesAjouteesLargeur = [50,100,100,50];
var IDDossierRacine="1Lwnqw9Cx_SpA15kH48LcY6rOP-jJGOTw";//Dossier Corrections Eleves

//var NOM_RANGE_INFO_CLASSES="sioParMail";//col1:mail, col3:classe, col4:nom, prénom
var reqColNom="Col2", reqColClasse="Col1", reqColMail="Col4",reqColPrenom="Col3";
var NOM_RANGE_INFO_CLASSES="infosClasse"; //(colonnes MAILS,NOM, PRENOM)

var CODE_EVAL_FONCTION = "feval";

var horodateurTexte="Horodateur";
var LISTE_CORRECTEURS=["friedelleprof@gmail.com","sfriedelmeyer@ac-toulouse.fr","correction","correcteur"]


var NOM_RANGE_NOTES="notes";

var NOM_COLONNE_CHECK_ENVOI_MAIL="Envoi Mail";
var NOM_COLONNE_CHECK_MAIL_DOC="Check Mail Doc";
var NOM_COLONNE_LIENS_CORRIGES="Corrigé";
var NOM_COLONNE_CHECK_COPIE_DOC="Copie dossier Doc";

var NbDeQuestions;

var itemsEvaluesFonction = ["Le nom de la fonction correspond à la demande",
                            "Le rôle de la fonction est bien écrit en commentaire de façon précise",
                            "Les variables d'entrée (paramètres) correspondent à la demande, par le(s) NOM(s) et pour le(s) TYPE(s)",
                            "Le type de la FONCTION (sortie) correspond à la demande",
                            "L'algorithme est correct"
                           ];

var erreursB = ["Il y a une petite erreur",
                "Le commentaire n'est pas suffisamment précis",
                "Des erreurs ou imprécisions sur le NOM ou le TYPE des paramètres",
                "Des imprécisions sur le TYPE de la fonction",
                "L'algorithme est globalement correct mais pourrait être amélioré"
               ];

var rule = SpreadsheetApp.newDataValidation()
.requireNumberBetween(1, 100)
.requireValueInList(['A','B','C'], true)
.setAllowInvalid(true)
.build();