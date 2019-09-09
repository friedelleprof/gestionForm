function creeDocReponse(URLSheet,URLForm,forcer) {
  
  try {
    setStyles();
    //Vérifie que le range n'est pas déjà créé, sinon le crée ainsi que le dossier
    var nouveauDossier=creeRangeURLDocetDossier(URLSheet,URLForm);
    if(nouveauDossier) {
      var retour=creeDoc2Reponse(URLSheet,URLForm,nouveauDossier,forcer);
      metAJourFeuilleInfo(SpreadsheetApp.openByUrl(URLSheet),retour.retour)
      return retour;
    } else {
      return new Info(numLigneScript,null,null,"Impossible de crééer le dossier",null,rouge1,null);
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function envoiDocReponse(URLSheet,URLForm,forcer) {
  //forcer = true pour renvoyer
  try {
    setStyles();
    //Vérifie que le range n'est pas déjà créé, sinon le crée ainsi que le dossier
    var nouveauDossier=creeRangeURLDocetDossier(URLSheet,URLForm);
    if(nouveauDossier) {
      var retour=sendMailDocReponse(URLSheet,URLForm,forcer);
      metAJourFeuilleInfo(SpreadsheetApp.openByUrl(URLSheet),retour.retour);
      return retour;
    } else {
      return new Info(numLigneScript,null,null,"Le dossier n'existe pas",null,rouge1,null);
    }
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function copieDocToDossierSuivi(URLSheet,URLForm) {
  try {
    //Vérifie que le range n'est pas déjà créé, sinon le crée ainsi que le dossier
    var retour=copieDocToDossierSuivi2(URLSheet,URLForm);
    metAJourFeuilleInfo(SpreadsheetApp.openByUrl(URLSheet),retour.retour)
    return retour;
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}


// Création du doc REPONSE
function creeRangeURLDocetDossier(URLSheet,URLForm) {
  
  var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);
  var sheetCorrection=classeurReponse.getSheetByName(NOM_FEUILLE_CORRECTION);
  var sheetReponse=classeurReponse.getSheetByName(NOM_FEUILLE_REPONSE);
  
  //Correctif sur début du range texteMail
  
  var r=classeurReponse.getRangeByName(camelize("URL DOSSIER CORRECTION"));
  
  //Si le range n'existe pas, on le crée et on crée le dossier associé
  if(r==null) {
    var formulaire=FormApp.openByUrl(URLForm);
    var nomFormulaire=formulaire.getTitle();
    var nomNouveauDossier="Corrigés de "+nomFormulaire+"/"+maintenant();
    var dossierRacine=DriveApp.getFolderById(IDDossierRacine);
    dossierRacine.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
    var numColonneMail=numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL);
    sheetCorrection.insertColumnAfter(numColonneMail);
    //Création du dossier
    var dossierNouveau=dossierRacine.createFolder(nomNouveauDossier);//.setName(nomNouveauDossier);
    sheetCorrection.getRange(1, numColonneMail+1).setValue(dossierNouveau.getUrl());
    sheetCorrection.getRange(6, numColonneMail+1).setValue(dossierNouveau.getId());
    ajoutInformationDansFeuilleInfo(classeurReponse,3,dossierNouveau.getUrl(),"URL DOSSIER CORRECTION");
    ajoutInformationDansFeuilleInfo(classeurReponse,4,dossierNouveau.getId(),"ID DOSSIER CORRECTION");
    
    sheetCorrection.getRange(2,numColonneMail+1,1,1).setValue(NOM_COLONNE_LIENS_CORRIGES).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
    
    return dossierNouveau;
  } else { //Le dossier existe
    try {
      var IDDossier=classeurReponse.getRangeByName(camelize("ID DOSSIER CORRECTION")).getValue();
      var dossierNouveau=DriveApp.getFolderById(IDDossier);
      return dossierNouveau;
    } catch(e) {
      Looger.log("Erreur dans création dossier:"+e);
      return null;
    }
  } 
}

//Création de tous les docs dans le dossier

function creeDoc2Reponse(URLSheet,URLForm,nouveauDossier,forcer) {
  try {
    var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);
    var sheetCorrection=classeurReponse.getSheetByName(NOM_FEUILLE_CORRECTION);
    var sheetReponse =classeurReponse.getSheetByName(NOM_FEUILLE_REPONSE);
    
    var formulaire=FormApp.openByUrl(URLForm);
    var nomFormulaire=formulaire.getTitle();
    
    //Récupération nom du TP = nom du formulaire
    //Normalisation
    var rangeInfosEleve=sheetCorrection.getRange("A7:D"+sheetCorrection.getLastRow())
    var numColonneMail=numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL);

    //var rangeReponseEtCorrection=sheetCorrection.getRange(7,5,sheetReponse.getLastRow()+5,numColonneMail-2); ///A MODIFIER
    //var dataReponseEtCorrection=rangeReponseEtCorrection.getValues();
    var titresUnique=sheetCorrection.getRange(1,5,1,numColonneMail-1).getValues()[0];
    var intitulesItems=sheetCorrection.getRange(2,5,1,numColonneMail-1).getValues()[0];
    var pointsQ=sheetCorrection.getRange(6,5,1,numColonneMail-1).getValues()[0];
    var dataInfoEleve=rangeInfosEleve.getValues();
    var rangeURLDoc=rangeColonneNommee(sheetCorrection,NOM_COLONNE_LIENS_CORRIGES,7);    
    rangeURLDoc.setVerticalAlignment('middle').setFontSize(8).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    var dataURLDoc=rangeURLDoc.getValues();
    var total=classeurReponse.getRangeByName("totalPoints").getValue();
    var surRange=classeurReponse.getRangeByName("sur"),sur;
    if(surRange) {
      sur=surRange.getValue();
    } else {
      sur=20;
    }
    var nbDocCrees=0,nbDocRecrees=0,nbDocErreurs=0,nbDocExistants=0;
  } catch (e) {
    Logger.log("Erreur dans les variables de crée doc reponse");
  }
  
  //POUR CHAQUE REPONSE
  
  for(var i=0; i<dataInfoEleve.length;i++) { // dataInfoEleve.length;i++) {
    var mailEleve=dataInfoEleve[i][0];
    var nomEleve=dataInfoEleve[i][1];
    var prenomEleve=dataInfoEleve[i][2];
    var noteEleve=dataInfoEleve[i][3];
    var URLDoc=dataURLDoc[i][0];
    var docEleve,docEleveURL;
    var cellURLDoc=rangeURLDoc.getCell(i+1,1)//.setVerticalAlignment('middle').setFontSize(8).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    var rangeReponseEtCorrection=sheetCorrection.getRange(7+i,5,1,numColonneMail-1); ///A MODIFIER
    var dataReponseEtCorrection=rangeReponseEtCorrection.getValues();
    
    //S'il n'y a rien
    Logger.log("Valeur dans URLDoc: "+cellURLDoc.getValue());
    Logger.log("Nom:"+nomEleve+", prenom:"+prenomEleve);
    
    //Si le doc n'existe pas, on le crée
    if(nomEleve!="" && prenomEleve!="" && URLDoc=="") {
      var titreDoc=nomEleve+"_"+prenomEleve+": Corrigé "+nomFormulaire;
      docEleveURL=creeNouveauDocDansDossier(nouveauDossier,titreDoc);
      try {
        docEleve=DocumentApp.openByUrl(docEleveURL);
        cellURLDoc.setValue(docEleveURL).setBackground(vert2).setNote("créé le "+maintenant());
        Logger.log("Création doc:"+cellURLDoc.getA1Notation()+docEleveURL);
        //Logger.log(dataReponseEtCorrection[0]);
        rempliDoc(docEleve,titresUnique,intitulesItems,pointsQ,dataReponseEtCorrection[0],dataInfoEleve[i],total,sur,classeurReponse);
        nbDocCrees++;
      } catch(e) {
        Logger.log("Problème sur "+titreDoc+RC+e);
        cellURLDoc.setNote(ERROR+" LECTURE"+RC+e).setBackground(rouge1);
        docEleve=null;
        callErreur(e);
      }
    } 
    
    //Sinon, il y a quelque chose, mais FORCER est VRAI
    
    else if( forcer==true && nomEleve!="" && prenomEleve!="" ) {
      //On récupère le doc existant si on doit le recréer
      var URLDoc=cellURLDoc.getValue();
      Logger.log("Doc existant:"+URLDoc);
      try {
        docEleve=DocumentApp.openByUrl(URLDoc);
        Logger.log(docEleve.getName());
        rempliDoc(docEleve,titresUnique,intitulesItems,pointsQ,dataReponseEtCorrection[0],dataInfoEleve[i],total,sur,classeurReponse);
        cellURLDoc.setBackground(vert3).setNote("recréé le "+maintenant());
        nbDocRecrees++;
      } catch(e) {
        if(URLDoc.toString().indexOf("https://docs.google.com/open?id=")>-1) {
          try { 
            var ID=URLDoc.replace("https://docs.google.com/open?id=","");
            docEleve=DocumentApp.openById(ID);
            Logger.log(docEleve.getName());
            rempliDoc(docEleve,titresUnique,intitulesItems,pointsQ,dataReponseEtCorrection[0],dataInfoEleve[i],total,sur,classeurReponse);
            cellURLDoc.setBackground(vert3).setNote("recréé le "+maintenant());
            nbDocRecrees++;
          } catch(e) {
            Logger.log("Problème sur "+titreDoc+RC+e);
            cellURLDoc.setNote(ERROR+" LECTURE"+RC+e).setBackground(rouge1);
            docEleve=null;
          }
        } else {
          Logger.log("Problème sur "+titreDoc+RC+e);
          cellURLDoc.setNote(ERROR+" LECTURE"+RC+e).setBackground(rouge1);
          docEleve=null;
          nbDocErreurs++;
        }
      } 
    }
    else if ((nomEleve=="" || prenomEleve=="") && mailEleve!="") {
      //problème sur adresse mail
      cellURLDoc.setNote(ERROR +" PRENOM").setBackground(rouge1);
      var numRow=cellURLDoc.getRow();
      sheetCorrection.getRange(numRow, 1, 1,4).setBackground(orange1);
      nbDocErreurs++;
    }
    else {
      Logger.log("La cellule "+cellURLDoc.getA1Notation()+" n'est pas vide:"+URLDoc+RC+", on ne fait rien");
      nbDocExistants++
        //Il y a déjà un doc, on ne fait rien.
    }
  }
  var texte=nbDocCrees+" docs créés"+RC+nbDocRecrees+" docs refaits"+RC+nbDocErreurs+" erreurs"+RC+nbDocExistants+" déjà existants";
  return new Info(numLigneScript,rangeNbDocsCrees.getColumn(),nbDocExistants+nbDocCrees,"Création de "+nbDocCrees+" docs",texte,null,null) 
  
}

function creeNouveauDocDansDossier(nouveauDossier,titreDoc) {
  try {
    //Crée le document et le place dans le dossier;
    Logger.log(nouveauDossier.getUrl());
    Logger.log(titreDoc);
    
    var document=DocumentApp.create(titreDoc);
    Logger.log(document.getUrl());

    var fileDocument=DriveApp.getFileById(document.getId());
    var doc=fileDocument.makeCopy(nouveauDossier).setName(titreDoc);
    DriveApp.removeFile(fileDocument);
    return doc.getUrl();
  } catch(e) {
    callErreur(e);//erreur type Info
    return "";
  }
}

function rempliDoc(docEleve,titresUnique,intitulesItems,pointsQ,dataReponseEtCorrection,dataInfoEleve,total,sur,classeurReponse) {
  var body = docEleve.getBody();
  var margeDroite=body.getMarginRight(),margeGauche=body.getMarginLeft();
  var width=body.getPageWidth()-margeDroite-margeGauche;
  var mailEleve=dataInfoEleve[0];
  var nomEleve=dataInfoEleve[1];
  var prenomEleve=dataInfoEleve[2];
  var noteEleve=dataInfoEleve[3];
  body.clear();
  body.appendParagraph("CORRIGE").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var tabTitre=body.appendTable().setBorderWidth(0);
  var rangT=tabTitre.appendTableRow();
  var cel1rT=rangT.appendTableCell().clear().setWidth(width*0.7);
  var cel2rT=rangT.appendTableCell().clear();
  cel1rT.appendParagraph(mailEleve).setHeading(DocumentApp.ParagraphHeading.NORMAL);//MAIL
  cel1rT.appendParagraph(nomEleve+" "+prenomEleve).setHeading(DocumentApp.ParagraphHeading.NORMAL);//Nom prenom
  noteEleve=Math.round(noteEleve*2)/2;
  cel2rT.appendParagraph(noteEleve+" points /"+total).setAttributes(stylePoints);
  var noteSur20=Math.round(noteEleve*sur/total*2)/2;
  cel2rT.appendParagraph("soit:"+noteSur20+"/"+sur).setAttributes(stylePoints);
  
  //Pour chaque question
  //Si pas ":" et commence par "F": début d'une fonction/programme
  //Tout dans un tableau
  
  var tableauPrincipal=body.appendTable();
  var tableauCriteres=body.appendTable();
  
  for(var i=0;i<titresUnique.length;i+=2) {
    var code=titresUnique[i];
    var question =intitulesItems[i];
    var points =pointsQ[i];
    if(typeof points === 'string') {
      points=points.split(";")[0];
    }
    if(code.toString().substring(0, 6)=="POINTS" && points>0) {
      //On ne prend pas les questions à 0
      //Logger.log(question+RC+dataReponseEtCorrection[i]+RC+dataReponseEtCorrection[i+1]);
      
      var rang1=tableauPrincipal.appendTableRow();
      var rang2=tableauPrincipal.appendTableRow();
      var cel1r1=rang1.appendTableCell().clear();
      cel1r1.insertParagraph(0, question.trim()).setAttributes(styleTitreItem);
      if(dataReponseEtCorrection[i]>points/2) {
        cel1r1.setBackgroundColor(vert2);
      } else if (dataReponseEtCorrection[i]>points/5) {
        cel1r1.setBackgroundColor(orange1);
      } else
      {
        cel1r1.setBackgroundColor(rouge2);
      }
      var cel2r1=rang1.appendTableCell().clear().setWidth(width*0.1);
      cel2r1.insertParagraph(0,Math.round(dataReponseEtCorrection[i]*4)/4+"/"+points).setAttributes(stylePoints);
      var cel1r2=rang2.appendTableCell().clear();
      cel1r2.setWidth(width*0.9).insertParagraph(0,dataReponseEtCorrection[i+1]).setAttributes(styleCorrection);
    }
    else if(code.toString().indexOf("Total F")!=-1) { //FONCTION
      //Début de traitement de critères
      //En tête
      //Numero
      var nomRangeCorrection="Corrige"+code.toString().replace("Total ","");
      var nomRangeIntitule="intituleQ"+code.toString().replace("Total ","");
      var copieReponse=dataReponseEtCorrection[i+1]; //La réponse
      var totalPointsExo=dataReponseEtCorrection[i];//Les points
      var rang1=tableauCriteres.appendTableRow();
      //Logger.log("Critères "+question+RC+nomRangeCorrection+RC+nomRangeIntitule);
      var cel1r1=rang1.appendTableCell().clear().setWidth(width*0.6);
      var intitule=classeurReponse.getRangeByName(nomRangeIntitule).getValue();
      cel1r1.insertParagraph(0, intitule).setAttributes(styleTitreItem);
      var cel1r2=rang1.appendTableCell().clear().setWidth(width*0.4);
      cel1r2.insertParagraph(0,"Critères: "+Math.round(totalPointsExo*4)/4+"/"+points).setAttributes(stylePoints);
      
      
      if(totalPointsExo>points/2) {
        cel1r1.setBackgroundColor(vert2);
        cel1r2.setBackgroundColor(vert2);
      } else if (totalPointsExo>points/5) {
        cel1r1.setBackgroundColor(orange1);
        cel1r2.setBackgroundColor(orange1);
      } else
      {
        cel1r1.setBackgroundColor(rouge2);
        cel1r2.setBackgroundColor(rouge2);
      }
      
      var rang2=tableauCriteres.appendTableRow();
      var cel1r2=rang2.appendTableCell();
      cel1r2.insertParagraph(0,classeurReponse.getRangeByName(nomRangeCorrection).getValue()).setAttributes(styleCode);//Corrigé
      cel1r2.insertParagraph(0,NOM_COLONNE_LIENS_CORRIGES).setAttributes(styleTitreItem);
      
      cel1r2.insertParagraph(0,copieReponse).setAttributes(styleCode);//Ajout de la réponse de l'élève
      
      
      var cel2r2=rang2.appendTableCell();
      var tableauInsere=cel2r2.appendTable();
      var quest = new Array();
      var pts = new Array();
      var reps = new Array();
      
      //Combien de critères ?
      
      var indice=i+2;var ncrt=0,cr="";
      do {
        //On cherche un intitulé "correction"
        question =intitulesItems[indice];
        points =pointsQ[indice];
        var ptRep=dataReponseEtCorrection[indice+1];
        ncrt++;//Nb critères
        if(points>0) {
        //On évite les questions à 0 points
        quest.push(question);
        pts.push(ptRep+"/"+points);
        }
        indice+=2;
        //Logger.log("Critères "+ncrt+" "+question+RC+code);
        code=titresUnique[indice];
      } while (code.toString().indexOf("CORRECTION F")==-1);
      
      for(var p=0;p<quest.length;p++) {
        //Logger.log(intitulesItems[p*2+1+i]);    
        reps.push(RC+dataReponseEtCorrection[indice+p]);
      }
      //Ecriture dans le doc
      for(var p=0;p<quest.length;p++) {
        var rang=tableauInsere.appendTableRow();
        var cel1r2=rang.appendTableCell().clear();
        cel1r2.insertParagraph(0, reps[p]).setAttributes(styleCorrection);        
        cel1r2.insertParagraph(0, pts[p]).setAttributes(stylePoints);
        
        cel1r2.insertParagraph(0, quest[p]).setAttributes(styleTitreItem);
      }
      
      i+=ncrt;//ncrt colonnes par question
    } else {
      //Logger.log("NONPOINT"+question+RC+dataReponseEtCorrection[i]+RC+dataReponseEtCorrection[i+1]);
    }
  }
  
}


function sendMailDocReponse(URLSheet,URLForm,forcer) {
  
  var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);
  var sheetCorrection=classeurReponse.getSheetByName(NOM_FEUILLE_CORRECTION);
  var sheetReponse=classeurReponse.getSheetByName(NOM_FEUILLE_REPONSE);
  
  var formulaire=FormApp.openByUrl(URLForm);
  var nbRep= formulaire.getResponses().length;
  var nomFormulaire=formulaire.getTitle();
    
  var numColonneMail=numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL);

  //Si le range n'existe pas, on le crée 
  if(numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_MAIL_DOC)==null) {
    sheetCorrection.insertColumnAfter(numColonneMail);
    sheetCorrection.getRange(2,numColonneMail+1,1,1).setValue(NOM_COLONNE_CHECK_MAIL_DOC).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  
  var rangeCheckMailDoc=rangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_MAIL_DOC,7);
  
  var rangeInfosEleve=sheetCorrection.getRange("A7:D"+sheetCorrection.getLastRow());
  
  var rangeURLDoc=rangeColonneNommee(sheetCorrection,NOM_COLONNE_LIENS_CORRIGES,7);
  
  //On veut le même nombre d'enregistrements
  
  var dataInfoEleve=rangeInfosEleve.getValues();
  var dataURLDoc=rangeURLDoc.getValues();
  var dataCheckMailDoc=rangeCheckMailDoc.getValues();
  
  var nbDocMailes=0,nbDocRemailes=0,nbDocErreurs=0,nbDocDejaMailes=0;
  
  for(var i=0; i<dataURLDoc.length;i++) { // dataInfoEleve.length;i++) {
    var mailEleve=dataInfoEleve[i][0];
    var nomEleve=dataInfoEleve[i][1];
    var prenomEleve=dataInfoEleve[i][2];
    var noteEleve=dataInfoEleve[i][3];
    var URLDoc=dataURLDoc[i][0];
    var checkMailDoc=dataCheckMailDoc[i][0];
    var cellcheckMailDoc=rangeCheckMailDoc.getCell(i+1,1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
    .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(null).setFontColor(vert1);
    
    var docEleve,docEleveID;
    var cellURLDoc=rangeURLDoc.getCell(i+1,1);
    //S'il n'y a rien
    Logger.log(cellURLDoc.getValue());
    var rep;
    if(nomEleve!="" && prenomEleve!="" && (checkMailDoc==false || forcer==true)) {
      //On récupère le doc existant
      var URLDoc=cellURLDoc.getValue();
      try {
        docEleve=DocumentApp.openByUrl(URLDoc);
        Logger.log("GO:"+docEleve.getName());
        rep=envoi(docEleve,prenomEleve,mailEleve);
        if(rep==MAIL_ENVOYE) {
          cellcheckMailDoc.setValue(true).setFontColor(vert1).setNote("envoyé le "+maintenant());
          if(forcer) { nbDocRemailes++; } else { nbDocMailes++ };
        } else if(rep=="Pb quota") {
          cellcheckMailDoc.setValue(false).setFontColor(rouge1).setNote("Pb quota le "+maintenant());
          nbDocErreurs++;
        } else {
          cellcheckMailDoc.setValue(false).setFontColor(rouge2).setNote(rep);
          nbDocErreurs++;
        }
      }
      catch(e) {
        if(URLDoc.toString().indexOf("https://docs.google.com/open?id=")>-1) {
          try { 
            var ID=URLDoc.replace("https://docs.google.com/open?id=","");
            docEleve=DocumentApp.openById(ID);
            Logger.log("GO:"+docEleve.getName());
            rep=envoi(docEleve,prenomEleve,mailEleve);
            if(rep==MAIL_ENVOYE) {
              cellcheckMailDoc.setValue(true).setFontColor(vert1).setNote("envoyé le "+maintenant());
              if(forcer) { nbDocRemailes++; } else { nbDocMailes++ };
            } else if(rep=="Pb quota") {
              cellcheckMailDoc.setValue(false).setFontColor(rouge1).setNote("Pb quota le "+maintenant());
              nbDocErreurs++;
            } else {
              cellcheckMailDoc.setValue(false).setFontColor(rouge2).setNote(rep);
              nbDocErreurs++;
            }
          }
          catch(e) {
            Logger.log("Problème sur "+mailEleve+RC+e);
            cellcheckMailDoc.setValue(false).setFontColor(rouge1).setNote(ERROR+" ENVOI"+RC+e);
            docEleve=null;
            nbDocErreurs++;
          }
        } else {
          Logger.log("Problème sur "+mailEleve+RC+e);
          cellcheckMailDoc.setValue(false).setFontColor(rouge1).setNote(ERROR+" LECTURE"+RC+e);
          docEleve=null;
          nbDocErreurs++;
        }
      } 
    }
    else if(nomEleve=="" || prenomEleve=="") {
      cellURLDoc.setNote(ERROR+" NOM\n"+maintenant()).setBackground(rouge1);
      nbDocErreurs++;
    } 
    else if(checkMailDoc==true) {
      nbDocDejaMailes++;
    }
  }
  var compteRendu="";
  if(nbDocMailes>0) compteRendu+=nbDocMailes+" docs envoyés\n";
  if(nbDocRemailes>0) compteRendu+=nbDocRemailes+" docs remailés\n";
  if(nbDocErreurs>0) compteRendu+=nbDocErreurs+" erreurs\n";
  if(nbDocDejaMailes>0) compteRendu+=nbDocDejaMailes+" déjà faits\n";
  
  return new Info(numLigneScript,rangeNbDocsEnvoyes.getColumn(),nbDocMailes+nbDocDejaMailes,"Envoi de "+nbDocMailes+" docs",compteRendu,null,null);
  
}


function envoi(docEleve,prenomEleve,mailEleve) {
  try {
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    Logger.log("Remaining email quota: " + emailQuotaRemaining);
    if(emailQuotaRemaining>0) {
      var mailBody = "<H1><B>Bonjour "+prenomEleve+"</B></H1>";
      mailBody+="<p>Tu trouveras une nouvelle correction dans ce lien:</p><br>";
      mailBody+="<a href='"+docEleve.getUrl()+"'>"+docEleve.getName()+"</a>";
      mailBody+="<H3>S. Friedelmeyer</H3>";
      
      //mailEleve="friedelleprof@gmail.com";
      
      docEleve.addViewer(mailEleve);
      MailApp.sendEmail({
        to: mailEleve,
        subject: "Envoi corrigé"+docEleve.getName(),
        htmlBody: mailBody,
        noReply:true
      });
      return MAIL_ENVOYE;
    }
    
    else {
      return PB_QUOTA;
    }
  }
  catch(e) {
    callErreur(e).affiche();
    return ERROR +e;
  }
}

function copieDocToDossierSuivi2(URLSheet,URLForm) {
  
  var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);
  //  var formulaire=FormApp.openByUrl(URLForm);
  var rangeURLFolderEtudiants=ss.getRangeByName(LISTE_DOSSIERS_ETUDIANTS);//récupérations des dossiers étudiants
  var rangeListeMails=ss.getRangeByName(LISTE_MAILS_ETUDIANTS);//Et des mails
  
  var dataListeMails=transpose(rangeListeMails.getValues())[0];
  
  var sheetCorrection=classeurReponse.getSheetByName(NOM_FEUILLE_CORRECTION);
  var sheetReponse=classeurReponse.getSheetByName(NOM_FEUILLE_CORRECTION);
  
  var rangeURLDoc=rangeColonneNommee(sheetCorrection,NOM_COLONNE_LIENS_CORRIGES,7);
  
  //Si le range n'existe pas, on le crée et on crée le dossier associé
  if(numRangeColonneNommee(sheetCorrection,NOM_COLONNE_CHECK_COPIE_DOC)==null) {
    var colonneURL=rangeURLDoc.getColumn();
    sheetCorrection.insertColumnAfter(colonneURL);
    //Création range
    sheetCorrection.getRange(2,colonneURL+1,1,1).setValue(NOM_COLONNE_CHECK_COPIE_DOC).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
  } 
  
  rangecheckDossierDoc=rangeColonneNommee(sheetCorrection,NOM_COLONNE_CHECK_COPIE_DOC,7);
  
  var rangeInfosEleve=sheetCorrection.getRange("A7:D"+sheetCorrection.getLastRow());
 
  
  var dataURLDoc=rangeURLDoc.getValues();
  
  var dataInfoEleve=rangeInfosEleve.getValues();
  
  var datacheckDossierDoc=rangecheckDossierDoc.getValues();
  
  var nbDoc=0,nbDocCopies=0,nbDocErreurs=0,nbDocDejaCopies=0;
  
  
  for(var i=0; i<dataInfoEleve.length;i++) {
    var mailEleve=dataInfoEleve[i][0].trim();
    
    var URLDoc=dataURLDoc[i][0];
    var checkDossierDoc=datacheckDossierDoc[i][0];
    var docEleve,docEleveID,rangeURLFolder,num;
    var cellURLDoc=rangeURLDoc.getCell(i+1,1);
    var cellDossierDoc=rangecheckDossierDoc.getCell(i+1,1).setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
    .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(null).setFontColor(vert1);
    
    //S'il n'y a rien
    var URLDoc=cellURLDoc.getValue();
    
    Logger.log(URLDoc);
    if(checkDossierDoc==false) {
      //On récupère le doc existant
      try {
        var IDfichier=DocumentApp.openByUrl(URLDoc).getId();
        Logger.log("copie:"+IDfichier);
        
        //Récupération dossier élève
        num=dataListeMails.indexOf(mailEleve);
        if(num>-1) {
          rangeURLFolder=rangeURLFolderEtudiants.getCell(num+1,1);
          copieDocDansDossier(IDfichier,rangeURLFolder)
          cellDossierDoc.setValue(true).setFontColor(vert1).setNote("copié le "+maintenant());
          nbDocCopies++;
        }
        else {
          //Pas de dossier associé à cette adresse mail
          Logger.log(mailEleve);
          cellDossierDoc.setValue(false).setFontColor(orange1).setNote("pas de dossier associé "+maintenant());
          nbDocErreurs++;
        }
      } catch(e) {
        //Problème ID du fichier 
        if(URLDoc.toString().indexOf("https://docs.google.com/open?id=")>-1) {
          try { 
            var IDfichier=URLDoc.replace("https://docs.google.com/open?id=","");
            //Récupération dossier élève
            var num=dataListeMails.indexOf(mailEleve);
            if(num>-1) {
              rangeURLFolder=rangeURLFolderEtudiants.getCell(num+1,1);
              copieDocDansDossier(IDfichier,rangeURLFolder)
              cellDossierDoc.setValue(true).setFontColor(vert1).setNote("copié le "+maintenant());
              nbDocCopies++;
            }
            
          } catch(e) {
            Logger.log("Problème sur "+URLDoc+RC+e);
            cellDossierDoc.setValue(false).setFontColor(orange1).setNote(ERROR+" DEP"+RC+e);
          }
        } else {
          Logger.log("Problème sur "+URLDoc+RC+e);
          cellDossierDoc.setValue(false).setFontColor(orange1).setNote(ERROR+" LECTURE"+RC+e);
        }
      } 
    }
    else {
      //Sinon rien
      cellDossierDoc.setValue(true).setFontColor(vert1);
      nbDocDejaCopies++
    }
  }
  
  var compteRendu="";
  if(nbDocCopies>0) compteRendu+=nbDocCopies+" docs copiés \n";
  if(nbDocErreurs>0) compteRendu+=nbDocErreurs+" erreurs \n";
  if(nbDocDejaCopies>0) compteRendu+=nbDocDejaCopies+" déjà faits \n";
  
  return new Info(numLigneScript,rangeNbDocsCopies.getColumn(),nbDocCopies+nbDocDejaCopies,"Déplacement de "+nbDocCopies+" docs",compteRendu,null,null);
  
}


function copieDocDansDossier(IDfichier,rangeURLFolderEtudiants) {
  try {
    var fichier=DriveApp.getFileById(IDfichier);
    var ID=lienURL(rangeURLFolderEtudiants).replace("https://drive.google.com/drive/folders/","");
    var dossier=DriveApp.getFolderById(ID);
    dossier.addFile(fichier);
    return "OK";
  }
  catch(e) {
    callErreur(e).affiche();
    return ERROR +" lors du déplacement "+e;
  }
}

function lienURL(range) {
  var formulas = range.getFormulas();
  var output = [];
  for (var i = 0; i < formulas.length; i++) {
    var row = [];
    for (var j = 0; j < formulas[0].length; j++) {
      var url = formulas[i][j].match(/=hyperlink\("([^"]+)"/i);
      row.push(url ? url[1] : '');
    }
    output.push(row);
  }
  return output[0][0];
}