//////////////////////////////////////////////////////////////
// INFOS Sur le nouveau classeur entré par URL
//////////////////////////////////////////////////////////////

function getInfos(URLForm, URLSheet) {
  try {
    //Vérification cohérence URL
    var cellURLForm = rangeURLForm.getCell(numLigneScript, 1);
    var cellURLSheet = rangeURLSheet.getCell(numLigneScript, 1);
    checkURLS(URLForm, URLSheet, cellURLForm, cellURLSheet).affiche();//Renvoie une info
    URLForm = rangeURLForm.getCell(numLigneScript, 1).getValue();
    URLSheet = rangeURLSheet.getCell(numLigneScript, 1).getValue();

    var classeur = SpreadsheetApp.openByUrl(URLSheet);
    var formulaire = FormApp.openByUrl(URLForm);

    //Mise à jour classeur simple
    miseAJourINFOSClasseur(formulaire, classeur);

    //Nom Formulaire
    rangeNom.getCell(numLigneScript, 1)
      .setValue(getNom(formulaire)).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false).setFontSize(14).setBackground(blanc);

    //Nom feuille
    rangeNomFeuille.getCell(numLigneScript, 1)
      .setValue(getNomFeuille(classeur)).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false).setFontSize(7).setBackground(rose1);

    //Dossier
    var fileFormulaire = DriveApp.getFileById(formulaire.getId());
    rangeURLDossier.getCell(numLigneScript, 1)
      .setValue(getURLDossier(fileFormulaire)).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false).setFontSize(10).setBackground(blanc);

    rangeNomsDossiers.getCell(numLigneScript, 1)
      .setValue(getNomDossier(fileFormulaire)).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false).setFontSize(14).setBackground(blanc);

    //Open / Close ?
    rangeCheckOPEN.getCell(numLigneScript, 1)
      .setValue(getOpen(formulaire)).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(vert1);


    rangeCheckCollecteMail.getCell(numLigneScript, 1)
      .setValue(getCollecteMail(formulaire)).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(vert1);

    rangeCheckEnvoiReponses.getCell(numLigneScript, 1)
      .setValue(getEnvoiReponses(formulaire)).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(vert1);

    rangeCheckModification.getCell(numLigneScript, 1)
      .setValue(getModificationReponses(formulaire)).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(vert1);

    //Existence correcteur 
    //rangeCheckCorrecteur.getCell(numLigneScript,1)
    //.setValue(checkCorrecteur(classeur)).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
    //.setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(vert1);


    //Triggers en marche ?
    rangeCheckTrigger.getCell(numLigneScript, 1)
      .setValue(getTriggers(classeur)).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(vert1);

    //Nb réponses
    rangeNbReponse.getCell(numLigneScript, 1)
      .setValue(getNbReponse(formulaire))
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(gris1);

    //Nb MAILS envoyés
    rangeNbMail.getCell(numLigneScript, 1)
      .setValue(getMailEnvoyes(classeur))
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(gris1);

    //Nb CORRECTIONS créées
    rangeNbDocsCrees.getCell(numLigneScript, 1)
      .setValue(getDocsCrees(classeur))
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(gris1);


    //Nb CORRECTIONS envoyées
    rangeNbDocsEnvoyes.getCell(numLigneScript, 1)
      .setValue(getDocsEnvoyes(classeur))
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(gris1);

    //Nb CORRECTIONS copiées dans le dossier
    rangeNbDocsCopies.getCell(numLigneScript, 1)
      .setValue(getDocsCopies(classeur))
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(gris1);

    // Virgules ou autre spb dans les énoncés
    checkPresenceVirguleDansReponsesCB(URLForm);
    verifieNomItemsDifferents(URLForm);
    return new Info(numLigneScript, null, null, "INFOS mises à jour", "", null, null);
  }
  catch (e) {
    return callErreur(e);//erreur type Info
  }
}

function checkURLS(URLForm, URLSheet, cellURLForm, cellURLSheet) {
  //Vérifie la présence d'au moins une URL et complète
  if ((URLForm == "" && URLSheet == "") || (URLForm != "" && URLSheet != "")) {
    return new Info(numLigneScript, null, null, "pas de chgt d'URLs", "", null, null);
  }
  else if (URLForm == "" && URLSheet != "") {
    try {
      //Mise à jour URL Formulaire
      var sheetReponse = SpreadsheetApp.openByUrl(URLSheet);
      URLForm = sheetReponse.getFormUrl();
      cellURLForm.setValue(URLForm).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false);
      return new Info(numLigneScript, null, null, "URLForm à jour", "", null, null);
    }
    catch (e) {
      return callErreur(e);//erreur type Info//erreur type Info
    }
  }
  else if (URLForm != "" && URLSheet == "") {
    try {
      //Mise à jour URL Formulaire
      var formulaire;
      try {
        formulaire = FormApp.openByUrl(URLForm);
      } catch (e) {
        var ID = URLForm.replace("https://drive.google.com/open?id=", "");
        formulaire = FormApp.openById(ID);
        URLForm = formulaire.getEditUrl();
        cellURLForm.setValue(URLForm);
      }
      try {
        var IDsheet = formulaire.getDestinationId();
      } catch (e) {
        //Création de la feuille réponse si elle n'existe pas
        IDsheet = creeClasseurAssocie(formulaire).getId();
      }
      URLSheet = SpreadsheetApp.openById(IDsheet).getUrl();
      cellURLSheet.setValue(URLSheet).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false);
      return new Info(numLigneScript, null, null, "URLSheet à jour", "", null, null);
    }
    catch (e) {
      return callErreur(e);//erreur type Info//erreur type Info
    }
  }
}

function creeClasseurAssocie(formulaire) {

  //Sauvegarde de l'original dans un dossier spécial

  //duplique baseEvaluation et on recopie telles-quelles les autres feuilles
  var URLclasseurModele = "https://docs.google.com/spreadsheets/d/1louJso_6sDfbixDX6vRfMIxNz5IZEF28jErb8rKQ9XU/edit#gid=0";
  //On ouvre le modèle
  var classeurModele = SpreadsheetApp.openByUrl(URLclasseurModele);
  var nomFormulaire = formulaire.getTitle();
  var fileFormulaire = DriveApp.getFileById(formulaire.getId());

  var IDdossierFormulaire = fileFormulaire.getParents();

  // On prend le premier dossier.
  var dossierFormulaire = IDdossierFormulaire.next();
  var idDossierFormulaire = dossierFormulaire.getId();
  //Copie du modèle
  var classeurCopie = classeurModele.copy(nomFormulaire + "(réponses)");
  //Ou a été faite la copie ?
  var dossierBase = DriveApp.getFileById(classeurCopie.getId()).getParents().next();
  //Dans quel fichier ?
  var fileCopie = DriveApp.getFileById(classeurCopie.getId());
  //Droits pour visibilité
  fileCopie.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  //On le copie dans le bon dossier et en recupère son ID
  var IDnouveauClasseur = fileCopie.makeCopy(nomFormulaire, dossierFormulaire).getId();
  //On l'ouvre et on supprime l'autre version
  var nouveauClasseur = SpreadsheetApp.openById(IDnouveauClasseur);
  dossierBase.removeFile(fileCopie);

  //L'associe au formulaire
  formulaire.setDestination(FormApp.DestinationType.SPREADSHEET, IDnouveauClasseur);

  var feuilleInfo = nouveauClasseur.insertSheet(NOM_FEUILLE_INFO);
  //Vérification script déjà installé

  ajoutInformationDansFeuilleInfo(nouveauClasseur, 2, "SCRIPT:" + maintenant(), "Mise en place SCRIPT");
  //Modifie la feuille associée au formulaire

  return nouveauClasseur;
}

function getNom(formulaire) {
  try {
    return formulaire.getTitle();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}
function getNomFeuille(classeur) {
  try {
    var sheetReponseAuFormulaire = classeur.getSheets();
    var texteNoms = "";
    for (i in sheetReponseAuFormulaire) {
      texteNoms += sheetReponseAuFormulaire[i].getName() + RC;
    }
    return texteNoms;
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}
function getURLDossier(fileFormulaire) {
  try {
    var tousLesDossiers = fileFormulaire.getParents();
    var dossier = tousLesDossiers.next();
    return dossier.getUrl();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}

function getNomDossier(fileFormulaire) {
  try {
    var tousLesDossiers = fileFormulaire.getParents();
    var dossier = tousLesDossiers.next();
    return dossier.getName();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}

function getOpen(formulaire) {
  //renvoie true ou false si le formulaire est ouvert aux réponses
  try {
    return formulaire.isAcceptingResponses();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}

function getCollecteMail(formulaire) {
  //renvoie true ou false si le formulaire collecte les mails
  try {
    return formulaire.collectsEmail();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}

function getModificationReponses(formulaire) {
  //renvoie true ou false si le formulaire est modifiable après soumission
  try {
    return formulaire.canEditResponse();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}

function getEnvoiReponses(formulaire) {


  //renvoie true ou false si le formulaire enoi un mail
  try {
    return formulaire.isPublishingSummary();
  }
  catch (e) {
    callErreur(e).affiche();//erreur type Info
  }
}

function getTriggers(feuille) {
  //renvoie true si des triggers sont en place sur la feuille
  var rep = false;
  try {
    var ID = feuille.getId();
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length && rep == false; i++) {
      if (triggers[i].getTriggerSourceId() == ID) {
        rep = true;
      }
    }
    return rep;
  } catch (e) {
    return callErreur(e);//erreur type Info//erreur type Info
  }
}

function getNbReponse(formulaire) {
  //renvoie le Nb de réponses
  try {
    return formulaire.getResponses().length;
  }
  catch (e) {
    return callErreur(e);//erreur type Info//erreur type Info
  }
}

function getMailEnvoyes(classeur) {
  //Compte le Nb de mails réponse envoyés
  try {
    var sheetCorrection = classeur.getSheetByName(NOM_FEUILLE_CORRECTION);
    if (sheetCorrection) {
      var rangeCheckEnvoiMail = rangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL, 7);
      if (rangeCheckEnvoiMail) {
        var dataCheckMail = rangeCheckEnvoiMail.getValues();
        Logger.log(dataCheckMail);
        var compteur = 0;
        for (var i = 0; i < dataCheckMail.length; i++) {
          if (dataCheckMail[i][0] == true) compteur++;
        }
        return compteur;
      }
    }
    return 0;
  }
  catch (e) {
    return callErreur(e);//erreur type Info//erreur type Info
  }
}

function getDocsEnvoyes(classeur) {
  //Compte le Nb de docs correction envoyés

  try {
    if (classeur.getSheetByName(NOM_FEUILLE_CORRECTION)) {
      var rangeCheckMailDoc = rangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_MAIL_DOC, 7);
      if (rangeCheckMailDoc) {
        var dataCheckMail = rangeCheckMailDoc.getValues();
        var compteur = 0;
        for (var i = 0; i < dataCheckMail.length; i++) {
          if (dataCheckMail[i][0] == true) compteur++;
        }
        return compteur;
      }
    }
    return 0;
  }

  catch (e) {
    return callErreur(e);//erreur type Info
  }
}

function getDocsCrees(classeur) {
  try {
    if (classeur.getSheetByName(NOM_FEUILLE_CORRECTION)) {
      var rangeCheckDocCrees = rangeColonneNommee(sheetCorrection, NOM_COLONNE_LIENS_CORRIGES, 7);

      if (rangeCheckDocCrees) {
        var dataCheckDoc = rangeCheckDocCrees.getValues();
        var compteur = 0;
        for (var i = 0; i < dataCheckDoc.length; i++) {
          if (dataCheckDoc[i][0].toString().indexOf("https://docs.google.com/document/") >= 0) compteur++;
        }
        return compteur;
      }
    }
    return 0;

  }

  catch (e) {
    return callErreur(e);//erreur type Info
  }
}

function getDocsCopies(classeur) {
  try {
    sheetCorrection = classeur.getSheetByName(NOM_FEUILLE_CORRECTION);
    if (sheetCorrection) {
      var rangeNbDocCopies = rangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_COPIE_DOC, 7);
      if (rangeNbDocCopies) {
        var dataCheckDoc = rangeNbDocCopies.getValues();
        var compteur = 0;
        for (var i = 0; i < dataCheckDoc.length; i++) {
          if (dataCheckDoc[i][0] == true) compteur++;
        }
        return compteur;
      }
    }
    return 0;
  }

  catch (e) {
    return callErreur(e);//erreur type Info
  }
}

/*function checkCorrecteur(classeur) {
  //pour si le formulaire contient un des correcteurs
  //On compte, on met "correcteur" systématiquement

  var retour=false;
  try {

    var sheetReponseAuFormulaire=classeur.getSheets()[0];//A priori 1
    var MailRange=rangeColonneDuTexte(EMAIL_TEXTE,1,sheetReponseAuFormulaire);
    var dataMailRange=MailRange.getValues();
    for(var i in LISTE_CORRECTEURS) {
      if(TrouveNumLigneDatas(LISTE_CORRECTEURS[i],dataMailRange)) {
        retour=true;
      }
    }
    return retour
  }
  catch(e) {
    return callErreur(e);//erreur type Info//erreur type Info
    return "erreur "+e;
  }
}*/



