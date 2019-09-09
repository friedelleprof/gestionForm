//function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {

function miseAJourScriptClasseur(formulaire,classeur) {
  //Renvoie l'URL du classeur (nouveau ou pas changé) dans retour
  return MAJSimple(formulaire,classeur);
}
function miseAJourINFOSClasseur(formulaire,classeur) {
  AjoutInfos(classeur);
  ajoutMoyennesCorrectionEtFormatConditionnels(classeur);
  return "ok";
}

function MAJSimple(formulaire,ancienClasseur) {
  //Installation du script par copie de classeur modèle
  try {
    var URLSheetID=formulaire.getDestinationId();
    var feuilleInfo=ancienClasseur.getSheetByName(NOM_FEUILLE_INFO);
    //Vérification script déjà installé
    
    if(ancienClasseur.getRangeByName(camelize("Mise en place SCRIPT"))!=null) {
      //Le script est en place
      return new Info(numLigneScript, rangeURLSheet.getColumn(),ancienClasseur.getUrl(),"Script déjà installé",null,null,null);
    }
    
    //Sauvegarde de l'original dans un dossier spécial    
    //duplique baseEvaluation et on recopie telles-quelles les autres feuilles
    var URLclasseurModele="https://docs.google.com/spreadsheets/d/1louJso_6sDfbixDX6vRfMIxNz5IZEF28jErb8rKQ9XU/edit#gid=0";
    var IDDossierSauvegarde="1gSjfwBIMZplYUtrVFJ5uyLKAaBWAiyYh";
    //On ouvre le modèle
    var classeurModele=SpreadsheetApp.openByUrl(URLclasseurModele);
    var nomFormulaire=formulaire.getTitle();
    var fileFormulaire=DriveApp.getFileById(formulaire.getId());
    
    var IDdossierFormulaire=fileFormulaire.getParents();
    var dossierSauvegardes=DriveApp.getFolderById(IDDossierSauvegarde);
    
    // On prend le premier dossier.
    var dossierFormulaire = IDdossierFormulaire.next();
    var idDossierFormulaire=dossierFormulaire.getId();
    //Copie du modèle
    var classeurCopie=classeurModele.copy(nomFormulaire+ "(réponses)");
    //Ou a été faite la copie ?
    var dossierBase=DriveApp.getFileById(classeurCopie.getId()).getParents().next();
    //Dans quel fichier ?
    var fileCopie=DriveApp.getFileById(classeurCopie.getId());
    //Droits pour visibilité
    fileCopie.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    //On le copie dans le bon dossier et en recupère son ID
    var IDnouveauClasseur=fileCopie.makeCopy(nomFormulaire, dossierFormulaire).getId();
    //On l'ouvre et on supprime l'autre version
    var nouveauClasseur=SpreadsheetApp.openById(IDnouveauClasseur);
    dossierBase.removeFile(fileCopie);
    
    //L'associe au formulaire
    formulaire.setDestination(FormApp.DestinationType.SPREADSHEET, IDnouveauClasseur);
    
    //Supression de 'test' dans le nouveau classeur
    var feuilleTests=nouveauClasseur.getSheetByName("tests");
    nouveauClasseur.deleteSheet(feuilleTests);
    //Récupération des anciennes feuilles
    
    var toutesFeuilles=ancienClasseur.getSheets();
    //La zéro on la copie dans "old" sauf si old existe déjà
    if(ancienClasseur.getSheetByName(OLD_REPONSES)==null) {
      toutesFeuilles[0].copyTo(nouveauClasseur).setName(OLD_REPONSES);
    }
    //Copie de QuestionReponses
    if(ancienClasseur.getSheetByName(NOM_FEUILLE_QR)) {
      ancienClasseur.getSheetByName(NOM_FEUILLE_QR).copyTo(nouveauClasseur).setName(NOM_FEUILLE_QR);
    } else if(ancienClasseur.getSheetByName("Question - Reponses")) {
      ancienClasseur.getSheetByName("Question - Reponses").copyTo(nouveauClasseur).setName(NOM_FEUILLE_QR);
    }
    //Copie de Correction
    if(ancienClasseur.getSheetByName(NOM_FEUILLE_CORRECTION)) {
      ancienClasseur.getSheetByName(NOM_FEUILLE_CORRECTION).copyTo(nouveauClasseur).setName(NOM_FEUILLE_CORRECTION);
    }
    
    //On déplace l'ancien classeur
    var fileClasseurReponse=DriveApp.getFileById(ancienClasseur.getId());
    dossierSauvegardes.addFile(fileClasseurReponse.setName(nomFormulaire+" (sauvegarde)"));
    dossierFormulaire.removeFile(fileClasseurReponse);
    
    var feuilleInfo=nouveauClasseur.insertSheet(NOM_FEUILLE_INFO);
    //Vérification script déjà installé
    
    ajoutInformationDansFeuilleInfo(nouveauClasseur,2,"SCRIPT:"+maintenant(),"Mise en place SCRIPT");
    //Modifie la feuille associée au formulaire
    return new Info(numLigneScript, rangeURLSheet.getColumn(),nouveauClasseur.getUrl(),"Script installé",null,null,null);    
  } 
  catch(e) {
    callErreur(e).affiche();
    return new Info(numLigneScript, rangeURLSheet.getColumn(),ancienClasseur.getUrl(),"ERREUR lor de la mise en place du script",null,null,null);        
  }
}


function ajoutInformationDansFeuilleInfo(classeur,ligne,information,clef) {
  //Ajout de la feuille info si nécessaire
  if(classeur.getSheetByName(NOM_FEUILLE_INFO)==null) {
    classeur.insertSheet(NOM_FEUILLE_INFO).hideSheet();
  }
  var feuilleInfo=classeur.getSheetByName(NOM_FEUILLE_INFO);
  feuilleInfo.getRange(ligne,1).setValue(clef);
  feuilleInfo.getRange(ligne,2).setValue(information);
  classeur.setNamedRange(camelize(clef),feuilleInfo.getRange(ligne,2));
  return feuilleInfo;
}

function lisInfoDansFeuilleInfo(classeur,clef) {
  try {
    return classeur.getRangeByName(camelize(clef)).getValue();
  } catch(e) {
    Logger.log(e);
  }
}


function AjoutInfos(classeur) {
  var id=ss.getId();
  var feuilleInfo=ajoutInformationDansFeuilleInfo(classeur,1,id,ID_FEUILLE_SUIVI);
  feuilleInfo.getRange(1,3).setValue("DATE").setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground(orange1).setFontSize(12);
  feuilleInfo.getRange(1,4).setValue("HISTORIQUE").setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground(orange1).setFontSize(12);
  feuilleInfo.setColumnWidth(1, 180);
  feuilleInfo.setColumnWidth(2, 250);
  feuilleInfo.setColumnWidth(3, 180);
  feuilleInfo.setColumnWidth(4, 300);
  return id;
}

function ajoutMoyennesCorrectionEtFormatConditionnels(classeur) {
  var sheetCorrection=classeur.getSheetByName(NOM_FEUILLE_CORRECTION);
  if(sheetCorrection) {
    var nbRows=sheetCorrection.getMaxRows()-7;
    var rangeB=classeur.getRangeByName("reponseBCorrection");
    var rangeC=classeur.getRangeByName("reponseCCorrection");
    var rangeP=classeur.getRangeByName("pointsCorrection");
    var rangeT=classeur.getRangeByName("titresCorrection");
    var dataT= rangeT.getValues()[0];
    Logger.log(dataT.length);
    for(var i=0;i<dataT.length;i++) {
      var titre=dataT[i];
      if(titre.toString().indexOf("POINTS Q")>=0) {
        Logger.log("MAJ:"+titre);
        nomRangeCol=sheetCorrection.getRange(7,i+1,nbRows,1).getA1Notation();
        nomRangeCol2=sheetCorrection.getRange(7,3,nbRows,1).getA1Notation();
        nomCel1=rangeB.getCell(1,i+2).getA1Notation();
        nomCel2=rangeP.getCell(1,i+1).getA1Notation();
        
        rangeB.getCell(1,i+1).setFormula('=COUNTIF('+nomRangeCol+';">0")').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(12).setNumberFormat("00");
        rangeC.getCell(1,i+1).setFormula('=COUNTIF('+nomRangeCol+';"=0")').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(12).setNumberFormat("00");
        rangeB.getCell(1,i+2).setFormula('=AVERAGE('+nomRangeCol+')').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(12).setNumberFormat("0.0");
        rangeC.getCell(1,i+2).setFormula('='+nomCel1+'/max(split('+nomCel2+';";"))').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontSize(12).setNumberFormat("00%");
        
        setRulesCorrections(sheetCorrection,sheetCorrection.getRange(7,i+1,nbRows,1));
      } else if(titre.toString().indexOf("POINTS F")>=0) {//Cas des fonctions//critères
        setRulesCorrections(sheetCorrection,sheetCorrection.getRange(7,i+1,nbRows,1));
      }
    }
  }
}


function setRulesCorrections(feuille,ranges) {
  var rules = feuille.getConditionalFormatRules();
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberEqualTo(0)  
  .setBackground(rouge2)
  .setRanges([ranges])
  .build();
  
  rules.push(rule);
  feuille.setConditionalFormatRules(rules);
};
