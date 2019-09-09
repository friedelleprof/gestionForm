//function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {

function setOpenClose(URLForm) {
  try {
    var formulaire=FormApp.openByUrl(URLForm);
    var cellCheck=rangeCheckOPEN.getCell(numLigneScript,1);
    var result=formulaire.setAcceptingResponses(cellCheck.getValue()).isAcceptingResponses();
    if(result) {
      return new Info(numLigneScript,null,null,"Le formulaire accepte les réponses",null,null,null) ;
    } else {
      return new Info(numLigneScript,null,null,"Le formulaire n'accepte plus les réponses",null,null,null) ;
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function setModificationReponses(URLForm) {
  try {
    var formulaire=FormApp.openByUrl(URLForm);
    var cellCheck=rangeCheckModification.getCell(numLigneScript,1);
    var result=formulaire.setAllowResponseEdits(cellCheck.getValue()).canEditResponse();
    if(result) {
      return new Info(numLigneScript,null,null,"Modification réponse autorisée",null,null,null) ;
    } else {
      return new Info(numLigneScript,null,null,"Modification réponse non autorisée",null,null,null) ;
    }  
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}
function setCollecteMail(URLForm) {
  //renvoie true ou false si le formulaire collecte les mails
  try {
    var formulaire=FormApp.openByUrl(URLForm);
        var cellCheck=rangeCheckCollecteMail.getCell(numLigneScript,1);

    var result=formulaire.setCollectEmail(cellCheck.getValue()).collectsEmail();
if(result) {
      return new Info(numLigneScript,null,null,"Collecte mails",null,null,null) ;
    } else {
      return new Info(numLigneScript,null,null,"Ne collecte pas les mails",null,null,null) ;
    } 
    } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function setEnvoiReponses(URLForm) {
  //renvoie true ou false si le formulaire collecte les mails
  try {
    var formulaire=FormApp.openByUrl(URLForm);
            var cellCheck=rangeCheckEnvoiReponses.getCell(numLigneScript,1);

    var result=formulaire.setPublishingSummary(cellCheck.getValue()).isPublishingSummary();
if(result) {
      return new Info(numLigneScript,null,null,"Envoi des réponses par mail",null,null,null) ;
    } else {
      return new Info(numLigneScript,null,null,"Pas d'envoi des réponses",null,null,null) ;
    }
    } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function setTriggerMiseAJour(IDclasseur) {
  Logger.log("Appel setTrigger "+IDclasseur+"/"+numLigneScript);
    var cellCheck=rangeCheckTrigger.getCell(numLigneScript,1);

  try {
    if(cellCheck.getValue()==true) {
      //Installation trigger
      return ajoutTriggerMiseAJour(IDclasseur);
    } else {
      //On l'enlève
      return removeTriggerMiseAJour(IDclasseur);
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}


function cleanFeuilles(URLSheet) {
  //ne garde QUE QuestionReponses, Correction et feuille 1
  var tableau=[NOM_FEUILLE_CORRECTION,NOM_FEUILLE_QR,NOM_FEUILLE_INFO,NOM_FEUILLE_REPONSE];
  var texte="feuilles suprimées:";
  try {
    var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);
    var feuilles=classeurReponse.getSheets(),tfeuille="";;
    var result=false;//Pas de supression
    for(i in feuilles) {
      var feuille=feuilles[i];
      tfeuille+=feuille.getName()+RC;
      if(tableau.indexOf(feuille.getName())==-1) {
        texte+=feuille.getName()+"/";
        classeurReponse.deleteSheet(feuille);
        result=true;
      }
    }
    if(result) {            
      return new Info(numLigneScript,rangeNomFeuille.getColumn(),tfeuille,texte,null,null,null) ;
    } else {
      return new Info(numLigneScript,rangeNomFeuille,tfeuille,"aucune feuille supprimée",null,null,null) ;
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}
