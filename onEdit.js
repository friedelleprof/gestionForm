function createSpreadsheetOpenTrigger() {
  var ss = SpreadsheetApp.getActive();
  try {
    //On le tue
    removeTriggers(ss.getId(),"triggerOnEdit");
  } catch(e) {
  }
  
  ScriptApp.newTrigger('triggerOnEdit')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
}

function triggerOnEdit(e) {
  if(e.oldValue!=e.value) {
    
    Logger.log("Appel onedit sur "+e.range.getA1Notation());
    openDatas();
    try {
      var range = e.range;
      var sheet=range.getSheet();
      var nomFeuille=sheet.getName();
      var ligneDebut=rangeURLForm.getRow();
      //Détermination  de l'action
      var colonne=range.getColumn();
      numLigneScript=range.getRow()-ligneDebut+1;
      var URLForm=getForm(numLigneScript-1);
      var retour;
      
      
      if(nomFeuille.substr(0,5)=="SUIVI") {
        Logger.log("Mise à jour suivi");
        retour=metAJourColonne(sheet,colonne);
        Logger.log("set value false");
        range.setValue(false);
      } else if(URLForm) {
        if(colonne==rangeCheckOPEN.getColumn() ){
          Logger.log("Appel setOpenClose sur "+e.range.getA1Notation()+RC+URLForm);
          retour=setOpenClose(URLForm);
        } else if(colonne==rangeCheckCollecteMail.getColumn()) {
          Logger.log("Appel setCollecteMail sur "+e.range.getA1Notation()+RC+URLForm);
          retour=setCollecteMail(URLForm);
        } else if(colonne== rangeCheckEnvoiReponses.getColumn()) {
          Logger.log("Appel setEnvoiReponses sur "+e.range.getA1Notation()+RC+URLForm);
          retour=setEnvoiReponses(URLForm);
        } else if(colonne== rangeCheckModification.getColumn()) {
          Logger.log("Appel setModificationReponses sur "+e.range.getA1Notation()+RC+URLForm);
          retour=setModificationReponses(URLForm);
        } else if(colonne==rangeCheckTrigger.getColumn()) {
          Logger.log("Appel setTriggerMiseAJour sur "+e.range.getA1Notation()+RC+URLForm);
          var IDclasseur=FormApp.openByUrl(URLForm).getDestinationId();
          retour=setTriggerMiseAJour(IDclasseur);
        } else {
          Logger.log("rien sur "+e.range.getA1Notation()+RC+URLForm);
          retour =null;
        }
      } 
      else {
        Logger.log("rien sur "+e.range.getA1Notation()+RC+URLForm);
        retour =null;
      }
      if(retour) retour.affiche();
      
    } catch(e) {
      return callErreur(e);//erreur type Info
    }
    
  }
  
}

function getForm(ligne) {
  Logger.log("Ligne:"+ligne);
  try {
    var formURL=dataURLForm[ligne][0];
    Logger.log("formURL:"+formURL);
    return formURL;
  } catch(e) {
    callErreur(e).affiche();
    return null;
  }
}

