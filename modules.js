function verifieNomItemsDifferents(URLForm) {
  //Vérifie que tous les items ont des noms différents, sinon rajoute des " "
  try {
    var formulaire=FormApp.openByUrl(URLForm);
    var listeItems=formulaire.getItems();
    var result=false;
    var nb=0,texte="",itemI,itemJ;
    for(var i in listeItems) {
      itemI=listeItems[i].getTitle();
      for( var j in listeItems) {
        itemJ=listeItems[j].getTitle();
        if(j>i && itemI==itemJ) {
          //result=true;
          nb++;
          texte+=itemJ+RC;
          listeItems[j].setTitle(itemI+" ");
        }
      }
    }
    if(result==true) {
      return new Info(numLigneScript,null,null,nb+" items corrigés"+RC+texte,"",null,null);
    }
    else {
      return new Info(numLigneScript,null,null,"Items différents","",null,null);
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function checkPresenceVirguleDansReponsesCB(URLForm) {
  //Vérifie qu'aucun check box ne contient de réponses avec des ',' pour éviter bug
  try {
    var formulaire=FormApp.openByUrl(URLForm);
    var listeItems=formulaire.getItems(FormApp.ItemType.CHECKBOX);
    var nb=0,texte="";
    for(var i in listeItems) {
      var result=false;
      var item=listeItems[i];
      var choices=item.asCheckboxItem().getChoices();
      for (var j in choices) {
        var choice=choices[j];
        if(choice.getValue().toString().indexOf(",")>0) {
          if(!result) texte+=item.getTitle()+":"+RC;
          result=true;
          texte+=choice.getValue()+"/";
        }
      }
      if(result==true) {
        texte+=RC;
        nb++;
      }
    }
    if(nb>0) {
      return new Info(numLigneScript,null,null,nb+" items à corriger"+RC+texte,"",null,null);
    }
    else {
      return new Info(numLigneScript,null,null,nb+" aucun item à corriger"+RC+texte,"",null,null);
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function supprimerToutesReponses(URLForm,memeCorrecteur) {
  //Supprime toutes les réponses enregistrées 
  try {
    var formulaire=FormApp.openByUrl(URLForm);
    if(memeCorrecteur) {
      formulaire.deleteAllResponses();
    }
    else {
      var tabReponses=formulaire.getResponses();
      for(var i in tabReponses) {
        var reponse=tabReponses[i];
        var id=reponse.getId();
        var mail=reponse.getRespondentEmail();
        if(LISTE_CORRECTEURS.indexOf(mail)==-1) {
          formulaire.deleteResponse(id);
        }
      }
    }
    //Suppression dans la feuille
    var ID=formulaire.getDestinationId()
    var classeur=SpreadsheetApp.openById(ID);      
    var sheetReponseAuFormulaire=classeur.getSheetByName(NOM_FEUILLE_REPONSE);
    if(sheetReponseAuFormulaire==null) {
      sheetReponseAuFormulaire=classeur.getSheetByName(NOM_FEUILLE_REPONSEA);
      sheetReponseAuFormulaire.setName(NOM_FEUILLE_REPONSE);
    }
    var mailRange=rangeColonneDuTexte(EMAIL_TEXTE,1,sheetReponseAuFormulaire);
    if(mailRange==null) mailRange=rangeColonneDuTexte(EMAIL_TEXTE2,1,sheetReponseAuFormulaire);
    if(mailRange!=null) {
      //On peut supprimer toutes les lignes saufs celles de correction
      //Il faut le faire par en bas
      var dataMailRange=mailRange.getValues();
      for(var i=dataMailRange.length-1;i>=0;i--) {
        var mail=dataMailRange[i][0];
        if(mail !="" && LISTE_CORRECTEURS.indexOf(mail)<0) {
          //var estCorrecteur=false;
          /*for(var cor in LISTE_CORRECTEURS) {
          //Logger.log(mail+"/"+LISTE_CORRECTEURS[cor]+(mail==LISTE_CORRECTEURS[cor]));
          if(LISTE_CORRECTEURS[cor]==mail) estCorrecteur=true;
        }
        if(estCorrecteur==false) {*/
          //On peut supprimer la ligne
          try {
            sheetReponseAuFormulaire.deleteRow(1+i);
            //Logger.log("ligne "+(i+1));
          } catch(e) {
            //rien si on n'arrive pas à supprimer la ligne
          }
        }
      }
    }
    //Suppression des lignes dans CORRECTION
    var sheetCorrection=classeur.getSheetByName(NOM_FEUILLE_CORRECTION);
    
    if(sheetCorrection) {
      var dataMailReponses=sheetCorrection.getRange(7,1,sheetCorrection.getMaxRows()).getValues();
      //On peut supprimer toutes les lignes saufs celles de correction
      //Il faut le faire par en bas
      for(var i=dataMailReponses.length-1;i>=0;i--) {
        var mail=dataMailReponses[i][0];
        if(mail!="" && LISTE_CORRECTEURS.indexOf(mail)<0) {
/*        var estCorrecteur=false;
        for(var cor in LISTE_CORRECTEURS) {
          //Logger.log(mail+"/"+LISTE_CORRECTEURS[cor]+(mail==LISTE_CORRECTEURS[cor]));
          if(LISTE_CORRECTEURS[cor]==mail) estCorrecteur=true;
        }
        if(estCorrecteur==false) {*/
          //On peut supprimer la ligne
          try {
            sheetCorrection.deleteRow(7+i);
            Logger.log("ligne cor"+(i+7)+"/"+mail);
          } catch(e) {
            //rien si on n'arrive pas à supprimer la ligne
          }
        }
      }
    }
    
    
    var rep= formulaire.getResponses();
    var nbRep=rep.length;
    return new Info(numLigneScript,rangeNbReponse.getColumn(),nbRep,"get NbRep à jour:"+nbRep,"",null,null);
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function setNom(URLForm,URLSheet,nom) {
  try {
    var sheetReponse=SpreadsheetApp.openByUrl(URLSheet);  
    
    var formulaire=FormApp.openByUrl(URLForm);
    formulaire.setTitle(nom);
    var fichierFeuilleReponse=DriveApp.getFileById(sheetReponse.getId());
    fichierFeuilleReponse.setName(nom+ "(reponses)");
    var fichierformulaire=DriveApp.getFileById(formulaire.getId());
    fichierformulaire.setName(nom+ "(form)");
    return new Info(numLigneScript,null,null,"Nouveau nom"+nom,"",null,null);
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}