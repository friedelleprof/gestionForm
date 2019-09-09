ss = SpreadsheetApp.getActiveSpreadsheet();
sheetGestionFormulaires=ss.getSheetByName("gestion");

//file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW) 

TEST=true; //Passer à true en phase de test pour sauvegarder
ANNEE='2019';

/********************
 * 
 * 
 *  NOMS DE FEUILLES ET DE RANGE
 * 
 */

NOM_FEUILLE_INFO        ="infos";
NOM_FEUILLE_DONNEES_ELEVES="Données Elèves";
NOM_FEUILLE_LISTE_FORMULAIRES="FormulairesNotés";
NOM_FEUILLE_REPONSEA    ="Réponses élèves"  // nom feuille des réponses générée par le formulaire
NOM_FEUILLE_REPONSE     ="Réponses au formulaire 1" // AUtre nom possible ### Il faudrait le retrouver autrement//Pb pays
NOM_FEUILLE_LOGS        ="Logs"

FEUILLE_LOG             =ss.getSheetByName(NOM_FEUILLE_LOGS);
NOM_FEUILLE_CORRECTION  ="Corrections";
NOM_FEUILLE_QR          ="QuestionsReponses";

/***********
 * 
 * COnstantes TEXTES
 * 
 */

TEXT_CONTIENT_ERROR     ="texte contient ERROR";
MAIL_ENVOYE             ="mail envoyé";
ERROR                   = "ERREUR:";
PB_QUOTA                ="Pb quota";
OLD_REPONSES            ="old reponses";

ID_FEUILLE_SUIVI        ="Id feuille suivi";


LISTE_DOSSIERS_ETUDIANTS="listeDossiers";
LISTE_MAILS_ETUDIANTS   ="listeMails";

nomRangeNbDocsCopies    ="nbCopiesDossier";
nomRangeDocsCrees       ="nbDocsCrees";
nomRangeInfosDoc        ="InfosDoc";
nomRangeURLForm         ="URLFormulaire";
nomRangeNomsDossiers    ="NomsDossiers";
nomRangeURLSheet        ="URLFeuilleReponse";
nomRangeNom             ="NomFormulaire";
nomRangeScript          ="scriptFormulaire";
nomRangeInfoScript      ="infoScriptFormulaire";
nomRangeNomFeuille      ="feuilleFormulaire";
nomRangeURLDossier      ="URLDossier";
nomRangeCheckOPEN       ="checkOpen";
nomRangeNbReponse	      ="nbReponses";
//nomRangeCheckCorrecteur	="checkCorrecteur";
nomRangeCheckTrigger    ="checkTrigger";
nomRangeNbMail          ="nbMailEnvoyes";
nomRangeCheckCollecteMail="checkCollecteMail";
nomRangeCheckEnvoiReponses="checkEnvoiReponses";
nomRangeCheckModification="checkModification";
nomRangeNbDocsEnvoyes   ="nbDocsEnvoyes";
nomRangeDateDebut       ="dateDebut";
nomRangeDateFin         ="dateFin";

rangeNbDocsCopies       =ss.getRangeByName(nomRangeNbDocsCopies);
rangeInfosDoc           =ss.getRangeByName(nomRangeInfosDoc);
rangeNbDocsCrees        =ss.getRangeByName(nomRangeDocsCrees);
rangeNbMail             =ss.getRangeByName(nomRangeNbMail);
rangeCheckTrigger       =ss.getRangeByName(nomRangeCheckTrigger);
rangeNomsDossiers       =ss.getRangeByName(nomRangeNomsDossiers);
rangeURLDossier         =ss.getRangeByName(nomRangeURLDossier);
rangeURLForm            =ss.getRangeByName(nomRangeURLForm);
rangeURLSheet           =ss.getRangeByName(nomRangeURLSheet);
rangeNom                =ss.getRangeByName(nomRangeNom);
rangeScript             =ss.getRangeByName(nomRangeScript);
rangeInfoScript         =ss.getRangeByName(nomRangeInfoScript);
rangeCheckOPEN          =ss.getRangeByName(nomRangeCheckOPEN);
rangeNbReponse	        =ss.getRangeByName(nomRangeNbReponse);
//rangeCheckCorrecteur	=ss.getRangeByName(nomRangeCheckCorrecteur);
rangeNomFeuille         =ss.getRangeByName(nomRangeNomFeuille);
rangeCheckCollecteMail  =ss.getRangeByName(nomRangeCheckCollecteMail);
rangeCheckEnvoiReponses =ss.getRangeByName(nomRangeCheckEnvoiReponses);
rangeCheckModification  =ss.getRangeByName(nomRangeCheckModification);
rangeNbDocsEnvoyes      =ss.getRangeByName(nomRangeNbDocsEnvoyes);
rangeDateDebut          =ss.getRangeByName(nomRangeDateDebut);
rangeDateFin            =ss.getRangeByName(nomRangeDateFin);


/*********
 * 
 *  CHERGEMENT EVENTUEL DES DONNEES
 */

dataChargees=false;

function openDatas() {
  
  if(dataChargees==false) {
  dataChargees=true;
  //dataNbDocCopies=rangeNbDocCopies.getValues();
    //dataNbDocsCrees=rangeNbDocsCrees.getValues();
    dataInfoDoc=rangeInfosDoc.getValues();
    //datanbMail=rangeNbMail.getValues();
    //dataNomsDossiers=rangeNomsDossiers.getValues();
    //dataURLDossier=rangeURLDossier.getValues();
    dataURLForm=rangeURLForm.getValues();
    dataURLSheet=rangeURLSheet.getValues();
    dataNomFormulaire=rangeNom.getValues();
    dataScript =rangeScript.getValues();
    dataInfoScript =rangeInfoScript.getValues();
    dataCheckOPEN=rangeCheckOPEN.getValues();
    dataNbReponse	=rangeNbReponse.getValues();
    //dataCheckCorrecteur	=rangeCheckCorrecteur.getValues();
    //dataNomFeuille=rangeNomFeuille.getValues();
    //dataCheckTrigger=rangeCheckTrigger.getValues();
    //dataCheckCollecteMail=rangeCheckCollecteMail.getValues();
    //dataCheckEnvoiReponses=rangeCheckEnvoiReponses.getValues();
    //dataCheckModification=rangeCheckModification.getValues();
    //dataNbDocsEnvoyes=rangeNbDocsEnvoyes.getValues();
    //dataDateDebut=rangeDateDebut.getValues();
    //dataDateFin=rangeDateFin.getValues();
  }
}
/******************$
 * 
 * SUIVI DES ELEVES
 * 
 */

fichierSS           =DriveApp.getFileById(ss.getId()); /* POUR LE PARTAGE DES DONNEES AVEC LES ELEVES */





/**************
 * 
 * DIVERS
 * 
 */

dateAujourdhui=maintenant();
retourChariot=String.fromCharCode(10);
RC=retourChariot;
