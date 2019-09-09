/********
* 
* GESTION // CREATION DE LA FEUILLE DONNEES ELEVES
* 
*/

//PLAGES NOMMEES
/*
infosClasse	    Données Elèves	A2:D77
infosEleves	    Données Elèves	A2:F77
listePrenomEleves	Données Elèves	C2:C77
listeDossiers	    Données Elèves	F2:F77
listeScripts	    Données Elèves	G2:G77
listeURLSuivi	    Données Elèves	E2:E77
listeCheck	        Données Elèves	I2:I77
listeMails	        Données Elèves	D2:D77
listeInfoScript	Données Elèves	H2:H77
listeClasseEleves	Données Elèves	A2:A77
listeNomEleve	    Données Elèves	B2:B77
*/

NOMBRE_MAX_ELEVES=100;
DOSSIER_ELEVES_ID     = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("URL_DossierSuivis").getValue();//PREVOIR UNE RECHERCHE
DOSSIER_ELEVES        = DriveApp.getFolderById(DOSSIER_ELEVES_ID);

//Création de la page si inexistante

nomRangeInfosClasse = "infosClasse";
nomRangeInfosEleves = "infosEleves";
nomRangePrenomEleve = "listePrenomEleves";
nomRangeDossiers    = "listeDossiers";
nomRangeCheckMails  = "listeCheck";//vérifie que tout est OK, à décocher si élève absent
nomRangeMail        = "listeMails";//colonne Mail
nomRangeinfoScripts = "listeInfoScript";
nomRangeClasse      = "listeClasseEleves";
nomRangeNomEleve    = "listeNomEleve";


//Fonction appelée lors de la création de la page

var creePageDonneesEleves = function (sheet) {
  //Fonction apellée pour créer la page si elle n'existe pas
  //renvoie une feuille
  var enTetes=['CLASSE','NOMS','PRENOMS','MAIL','DOSSIER','CHECK','INFOS'];
  var tailles=[2,4,2,6,2,3,1];
  var rangeEntete=sheet.getRange(1,1,1,enTetes.length);
  rangeEntete.setValues([enTetes]);
  for(var i=0;i<tailles.length;i++) {
    sheet.setColumnWidth(i+1,tailles[i]*30);
  }
  rangeEntete.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setFontSize(12).setBackground(orange1);
  //Supression des lignes en trop
  if(sheet.getMaxRows()>NOMBRE_MAX_ELEVES) {  
    sheet.deleteRows(NOMBRE_MAX_ELEVES+1,sheet.getMaxRows()-NOMBRE_MAX_ELEVES);
  }
  if(sheet.getMaxColumns()>enTetes.length+1) {
    sheet.deleteColumns(enTetes.length+1,sheet.getMaxColumns()-enTetes.length);
  }
  //Format auto
  var newRange=sheet.getRange(1,1,100,enTetes.length)
  newRange.activate().createFilter();
  newRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = newRange.getBandings()[0];
  banding.setHeaderRowColor(orange1)
  .setFirstRowColor('#ffffff')
  .setSecondRowColor(orange2)
  .setFooterRowColor(null);
  
  //FREEZE LIGNE+COLONNES
  sheet.setFrozenColumns(2)
  sheet.setFrozenRows(1);
  
  return sheet;
}

var FEUILLE_ELEVES,
    rangeInfosClasse,rangeInfosEleves,rangePrenomEleve,rangeDossiers,
    rangeCheck,rangeMails,rangeInfoScript,rangeClasseEleve,rangeNomEleve,
    dataPrenomEleve     ,dataNomEleve,dataClasseEleve,dataMails,dataCheck,
    dataInfoScript,dataDossiers;
var dataElevesChargees=false;

function chargeRanges() {
  FEUILLE_ELEVES      = openSheet(NOM_FEUILLE_DONNEES_ELEVES,ss,true,creePageDonneesEleves);
  rangeInfosClasse    = openRange(ss,nomRangeInfosClasse,FEUILLE_ELEVES,"A2:D"+NOMBRE_MAX_ELEVES,true);
  rangeInfosEleves    = openRange(ss,nomRangeInfosEleves,FEUILLE_ELEVES,"A2:F"+NOMBRE_MAX_ELEVES,true);
  rangePrenomEleve    = openRange(ss,nomRangePrenomEleve,FEUILLE_ELEVES,"C2:C"+NOMBRE_MAX_ELEVES,true);
  rangeDossiers       = openRange(ss,nomRangeDossiers,FEUILLE_ELEVES,"E2:E"+NOMBRE_MAX_ELEVES,true);
  rangeCheck          = openRange(ss,nomRangeCheckMails,FEUILLE_ELEVES,"F2:F"+NOMBRE_MAX_ELEVES,true);
  rangeMails          = openRange(ss,nomRangeMail,FEUILLE_ELEVES,"D2:D"+NOMBRE_MAX_ELEVES,true);
  rangeInfoScript     = openRange(ss,nomRangeinfoScripts,FEUILLE_ELEVES,"G2:G"+NOMBRE_MAX_ELEVES,true);
  rangeClasseEleve    = openRange(ss,nomRangeClasse,FEUILLE_ELEVES,"A2:A"+NOMBRE_MAX_ELEVES,true);
  rangeNomEleve       = openRange(ss,nomRangeNomEleve,FEUILLE_ELEVES,"B2:B"+NOMBRE_MAX_ELEVES,true);
  //openDataEleves(true);
}

function openDataEleves(recharger) {
  chargeRanges();
  if(!dataElevesChargees || recharger) {
    dataPrenomEleve     =rangePrenomEleve.getValues();
    dataNomEleve        =rangeNomEleve.getValues();
    dataClasseEleve     =rangeClasseEleve.getValues();
    dataMails           =rangeMails.getValues();
    dataCheck           =rangeCheck.getValues();
    dataInfoScript      =rangeInfoScript.getValues();
    dataDossiers        =rangeDossiers.getValues();
    dataElevesChargees=true;
  }
}

/********************
* 
* FONCTIONS DE GESTION
* 
*/

function genererDossiers() {
  chargeRanges();
  //try {
  openDataEleves(true);
  for (var i = 0; i < dataNomEleve.length; i++) {
    var nomEleve = dataNomEleve[i][0];
    var prenomEleve = dataPrenomEleve[i][0];
    var mailEleve = dataMails[i][0];
    var classeEleve = dataClasseEleve[i][0];
    
    if (nomEleve != "" && prenomEleve != "" && classeEleve != "" && mailEleve != "") {
      var cellCheck = rangeCheck.getCell(i + 1, 1);
      var cellURLDossier = rangeDossiers.getCell(i + 1, 1);//Pour écriture
      var cellInfo = rangeInfoScript.getCell(i + 1, 1);//Pour écriture
      var urlDossier = cellURLDossier.getFormula(); //dataDossiers[i][0];
      var check = dataCheck[i][0];
      var retour = null;
      
      //Génère le dossier de l'élève
      
      var nouveauDossier, nouvelleFeuille;
      var nomDossier = "SUIVI " + classeEleve + " - " + ANNEE + " :" + prenomEleve + " " + nomEleve;
      
      if (urlDossier == "" || check==false) { //Si vide
        //On crée le dossier
        nouveauDossier = DOSSIER_ELEVES.createFolder(nomDossier);
        //Droits 
        nouveauDossier.addViewer(adresseMail);
        nouveauDossier.addViewer("fotozenne@gmail.com");
        urlDossier = nouveauDossier.getUrl();
        cellURLDossier.setFormula('=HYPERLINK("' + urlDossier + '";"' + nomDossier + '")')
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false).setFontSize(8);
        cellCheck.insertCheckboxes().setValue(true);
      } else {
        //Dossier existant
        //url de la forme =HYPERLINK("https://drive.google.com/drive/folders/1_asa8OABC5VrMXmmF3DBCGkjcS94NYkD";"SUIVI NSI-1 - 2020 :Guilhem FAURE")
        urlDossier=urlDossier.replace('=HYPERLINK("https://drive.google.com/drive/folders/',"");
        urlDossier=urlDossier.substring(0,urlDossier.indexOf("\";"));
        try {
          nouveauDossier = DriveApp.getFolderById(urlDossier);
          nouveauDossier.addViewer(mailEleve);
          nouveauDossier.addViewer("fotozenne@gmail.com");
          cellCheck.insertCheckboxes().setValue(true);
        } catch (e) {
          //cellCheck.insertCheckboxes().setValue(false);
          Logger.log(e);
          cellInfo.setValue(e);
          //cellCheck.insertCheckboxes().setValue(true);
        }
      }
    }
    
  }

  
}

