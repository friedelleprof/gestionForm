function removeTriggers(IDclasseur, handler) {
  Logger.log(RC + "SUPRESSION DE " + handler + " SUR " + IDclasseur + RC);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    // Logger.log(infoTrigger(triggers[i]));
    if (triggers[i].getHandlerFunction() == handler && triggers[i].getTriggerSourceId() == IDclasseur) {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log("SUPRESSION");
    }
  }
}

function removeTriggerMiseAJour(IDclasseur) {
  var cellCheck = rangeCheckTrigger.getCell(numLigneScript, 1);

  Logger.log("Appel removeTriggerMiseAJour " + IDclasseur + "/" + cellCheck.getA1Notation());

  try {
    //Ajout trigger
    var handler = "miseAJourClasseur";
    //On les tue
    removeTriggers(IDclasseur, handler);
    cellCheck.setValue(false).setBackground(vert1).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
    return new Info(numLigneScript, null, null, "Supression trigger miseAJourClasseur", null, null, null);
  }
  catch (e) {
    return callErreur(e);//erreur type Info
  }
}

function ajoutTriggerMiseAJour(IDclasseur) {
  var cellCheck = rangeCheckTrigger.getCell(numLigneScript, 1);
  Logger.log("Appel ajoutTriggerMiseAJour " + IDclasseur + "/" + cellCheck.getA1Notation());

  try {
    //Ajout trigger
    var handler = "miseAJourClasseur";
    var classeur = SpreadsheetApp.openById(IDclasseur)
    //On les tue
    removeTriggerMiseAJour(IDclasseur, cellCheck);
    //On le rajoute
    ScriptApp.newTrigger("miseAJourClasseur")
      .forSpreadsheet(classeur)
      .onFormSubmit()
      .create();

    metAJourFeuilleInfo(classeur, "AJout trigger miseAJourClasseur");
    cellCheck.setValue(true).setBackground(vert1).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false);
    return new Info(numLigneScript, null, null, "Ajout trigger miseAJourClasseur", null, null, null);
  }
  catch (e) {
    return callErreur(e);//erreur type Info
  }
}


function miseAJourClasseur(e) {

  //La différence de n° de rang entre Réponses et Corrections est constante, égale à +5 pour correction
  //Ligne 2--> ligne 7 etc...

  var classeur = e.source;
  var numLigneAjoutee = e.range.getRow();//On doit recopier à +5//10
  var numLigneCopiee = numLigneAjoutee + 5;
  var sheetCorrection = classeur.getSheetByName(NOM_FEUILLE_CORRECTION);

  //Tant que le Nb de lignes est plus petit que numLigneAjoutee+5, on en ajoute:
  while (sheetCorrection.getMaxRows() < numLigneCopiee) {
    sheetCorrection.appendRow([""]);
  }
  //Mise en forme ligne ajoutée

  e.range.setBackground(bleu5);
  classeur.getSheetByName(NOM_FEUILLE_REPONSE).setRowHeight(numLigneAjoutee, 20);

  var rangeACopier = sheetCorrection.getRange(7, 1, 1, sheetCorrection.getMaxColumns());
  var endroit = sheetCorrection.getRange(numLigneCopiee, 1, 1, sheetCorrection.getMaxColumns());
  rangeACopier.copyTo(endroit);
  sheetCorrection.setRowHeight(numLigneCopiee, 20);
  try {
    metAJourFeuilleInfo(classeur, "Ligne ajoutée n°" + numLigneAjoutee + "\nLigne copiée n°" + numLigneCopiee + "\n" + e.range.getCell(1, 2).getValue());//Mail
  } catch (e) { }
  //Le mail est mis à false
  var range = rangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL, 7);

  var column = range.getColumn();
  sheetCorrection.getRange(numLigneCopiee, column).setValue(false);

  //Ainsi que doc et checkEnvoiDoc
  var numColonne = numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_MAIL_DOC);

  if (numColonne) {
    sheetCorrection.getRange(numLigneCopiee, numColonne).setValue(false);
  }

  numColonne = numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_COPIE_DOC);
  if (numColonne) {
    sheetCorrection.getRange(numLigneCopiee, numColonne).setValue(false);
  }

  numColonne = numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_COPIE_DOC);
  if (numColonne) {
    sheetCorrection.getRange(numLigneCopiee, numColonne).setValue(false);
  }

  numColonne = numRangeColonneNommee(sheetCorrection, NOM_COLONNE_LIENS_CORRIGES)
  if (numColonne) {
    sheetCorrection.getRange(numLigneCopiee, numColonne).setValue("");
  }

  //Mise à jour des mails et des infos sur la page de gestion
  var retour = checkEtEnvoiMail(classeur.getUrl(), false);

  //Mise à jour des infos sur classeur gestion

  metAJourInfos(classeur, retour);
  metAJourFeuilleInfo(classeur, retour.retour)

}



function metAJourInfos(classeur, retour) {

  Logger.log("Appel metAJourInfo " + classeur.getId() + "\n"); retour.log();
  var feuilleInfos, IDclasseurGestion, classeurGestion;
  //La feuille gestion est elle reliée ?
  try {
    feuilleInfos = classeur.getSheetByName(NOM_FEUILLE_INFO);
    IDclasseurGestion = classeur.getRangeByName(camelize(ID_FEUILLE_SUIVI)).getValue();
    classeurGestion = SpreadsheetApp.openById(IDclasseurGestion);
  } catch (e) {
    Logger.log("problème ouverture classeur Gestion\n");
    Logger.log(e.name + RC + "ligne:" + e.lineNumber + RC + "->" + e.stack);
    metAJourFeuilleInfo(classeur, e.name + "\n" + "ligne:" + e.lineNumber + "\n" + "->" + e.message);
    return null;
  }

  try {

    //on a retour =  new Info(-1,rangeNbMail.getColumn(),nbDocMailes+nbDocDejaMailes,null,compteRendu,null,null);

    //On doit chercher l'URL dans URLFormulaire
    var dataURLFormulaires = transpose(classeurGestion.getRangeByName(nomRangeURLSheet).getValues())[0];//Commence à ligne 2
    var num = dataURLFormulaires.indexOf(classeur.getUrl());
    Logger.log("Formulaire trouvé ligne :" + num);
    if (num >= 0) {//On a trouvé la ligne, on met à jour le classeur de gestion

      var rangeNbReponses = classeurGestion.getRangeByName(nomRangeNbReponse);
      var rangeNB = rangeNbReponses.getCell(num + 1, 1);
      rangeNB.setValue(rangeNB.getValue() + 1).setBackground(rose2);

      var rangeNbMailsEnvoyes = classeurGestion.getRangeByName(nomRangeNbMail);
      rangeNbMailsEnvoyes.getCell(num + 1, 1).setValue(retour.retour);//nbDocMailes+nbDocDejaMailes

      var rangeInfoDoc = classeurGestion.getRangeByName(nomRangeInfosDoc);//Commence à ligne 2
      rangeInfoDoc.getCell(num + 1, 1).setValue(retour.infoDoc);
    }
  } catch (e) {
    var info = e.name + "\n" + "ligne:" + e.lineNumber + "\n" + "->" + e.message
    var cellInfoScript = classeurGestion.getRangeByName(nomRangeInfoScript).getCell(num, 1);
    cellInfoScript.setValue(maintenant() + info + RC + cellInfoScript.getValue()).setBackground(rouge1);
    metAJourFeuilleInfo(classeur, info);
  }
}

function metAJourFeuilleInfo(classeur, message) {
  var feuilleInfos = classeur.getSheetByName(NOM_FEUILLE_INFO);
  if (feuilleInfos) {
    var l = feuilleInfos.getLastRow() + 1;
    feuilleInfos.getRange(l, 3).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yy à HH:mm:ss")).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(8);
    feuilleInfos.getRange(l, 4).setValue(message).setHorizontalAlignment('left').setVerticalAlignment('middle').setFontSize(8);
  }
}

function testTrig() {
  listTriggersFunctions(ss);
}

function listTriggersFunctions(classeur) {
  Logger.log("Appel listTriggersFunction");
  //renvoie la liste des triggers du classeur donné
  var IDclasseur = "";
  if (classeur) {
    IDclasseur = classeur.getId().toString();
  } else {
    IDclasseur = SpreadsheetApp.getActiveSpreadsheet().getId();
  }
  Logger.log("ID:" + IDclasseur);

  var test = false;
  var triggers = ScriptApp.getProjectTriggers();
  var liste = new Array();
  var nbTriggers = triggers.length;
  for (var i = 0; i < nbTriggers; i++) {

    var event = triggers[i].getEventType();
    var handler = triggers[i].getHandlerFunction();
    var source = triggers[i].getTriggerSource();
    var sourceID = triggers[i].getTriggerSourceId().toString();
    var ID = triggers[i].getUniqueId();
    var sheet = SpreadsheetApp.openById(sourceID).getName();
    if (sourceID == IDclasseur) {
      test = true;
      liste[handler] = ID;
      Logger.log("----->" + sheet + RC + event + RC + handler + RC + source + RC + sourceID + RC + ID + RC);
    } else {
    }
  }
  return test;
}
function testTriggers(classeur) {
  //renvoie test vrai si le classeur a déjà des triggers
  var ID = "";
  if (classeur) {
    ID = classeur.getId();
  } else {
    ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  }
  Logger.log("CLASSEUR ID:" + ID + RC + RC);
  var test = false;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    var sourceID = triggers[i].getTriggerSourceId();
    if (sourceID == ID) {
      test = true;
    }
    Logger.log(infoTrigger(triggers[i]));
  }
  return test;
}



function infoTrigger(trigger) {
  return "Evénément:" + trigger.getEventType() + RC + "Handler:" + trigger.getHandlerFunction() + RC + "Source:" + trigger.getTriggerSource() + RC + "Source ID:" + trigger.getTriggerSourceId() + RC + "Unique ID:" + trigger.getUniqueId() + RC;

}


