

function mailReponse(adresseMail,prenom,texteMail,titreMail,cellCheck) {
  try {
    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    Logger.log("Remaining email quota: " + emailQuotaRemaining);
    if(emailQuotaRemaining>0) {
      if(texteMail.indexOf("#ERROR!")!=-1) {
        return TEXT_CONTIENT_ERROR;
      } 
      else
      {
        var mailBody = "<H1><B>Bonjour "+prenom+"</B></H1>"+texteMail;
        //adresseMail="sfriedelmeyer@ac-toulouse.fr";
        
        MailApp.sendEmail({
          to: adresseMail,
          subject: titreMail,
          htmlBody: mailBody,
          noReply:true
        });
        return MAIL_ENVOYE;
      }
    }
    else {
      return PB_QUOTA;
    }
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function checkEtEnvoiMail(URLSheet,forcer) {
  
  //Envoi de mail soit par commande directe soit via trigger
  try {
    var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);

    var URLForm=classeurReponse.getFormUrl();
    var formulaire=FormApp.openByUrl(URLForm);
    var titreForm=formulaire.getTitle();
    
    var sheetCorrection=classeurReponse.getSheetByName(NOM_FEUILLE_CORRECTION);

    //Si forcer=true, on renvoie même s'il a été déjà envoyé
    //Vérifier que les ranges sont bien une colonne ENTIERE jusquau Nb d'enregistrements
    
    //En cas de 'disparition' du range:
    classeurReponse.setNamedRange(NOM_RANGE_NOTES, sheetCorrection.getRange("D7:E"+sheetCorrection.getLastRow()));
    var rangeInfoEleve=csheetCorrection.getRange("A7:D"+sheetCorrection.getLastRow());

 
    var rangeCheckEnvoiMail=rangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL, 7);

    var dataCheckMail=rangeCheckEnvoiMail.getValues();
    var dataInfosEleve=rangeInfoEleve.getValues();
    var nbDocMailes=0,nbDocRemailes=0,nbDocErreurs=0,nbDocDejaMailes=0;
    
    //Generation du mail
    
    var numColonneMail=numRangeColonneNommee(sheetCorrection, NOM_COLONNE_CHECK_ENVOI_MAIL);

    var titresUnique=sheetCorrection.getRange(1,5,1,numColonneMail-1).getValues()[0];//Première ligne : les titres
    var intitulesItems=sheetCorrection.getRange(2,5,1,numColonneMail-1).getValues()[0];//2nd : les intitulés
    var pointsQ=sheetCorrection.getRange(6,5,1,numColonneMail-1).getValues()[0];//Les points sur ligne 6
    var sur=classeurReponse.getRangeByName("totalPoints").getValue();
    
    
    for(var i=0;i<dataInfosEleve.length;i++) {
      var checkMail=dataCheckMail[i][0];
      
      if(checkMail==false || forcer || checkMail=="") {
        //var texte=dataTexte[i][0];
        
        var note=dataInfosEleve[i][3];
        
        var rangeReponseEtCorrection=sheetCorrection.getRange(7+i,5,1,numColonneMail-1); ///A MODIFIER
        var dataReponseEtCorrection=rangeReponseEtCorrection.getValues()[0];
        var texte=creeTexteMail(titreForm,titresUnique,intitulesItems,pointsQ,dataReponseEtCorrection,note,sur,classeurReponse)
        
        var cellCheck=rangeCheckEnvoiMail.getCell((i+1),1).setFontSize(14);
        var adresseMail=dataInfosEleve[i][0];
        var prenom=dataInfosEleve[i][2];
        var nom=dataInfosEleve[i][1];
        cellCheck.setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build())
        .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(null).setFontColor(vert1);
        //Logger.log("MAIL:"+adresseMail+"/"+texte+"/"+nom+"/"+prenom);
        if(adresseMail!="" && texte!="" && texte.indexOf("#ERROR!")==-1  && nom!="#N/A" && prenom!="") {
          var textSend=mailReponse(adresseMail,prenom,texte,"Evaluation "+classeurReponse.getName(),checkMail);
          if(textSend==MAIL_ENVOYE) {
            cellCheck.setValue(true).setNote(dateAujourdhui);
            fixeData(rangeReponseEtCorrection,dataReponseEtCorrection);
            if(forcer) {
              nbDocRemailes++;
            } else {
              nbDocMailes++;
            }
          } else if(textSend==PB_QUOTA) {
            cellCheck.setValue(false).setNote(PB_QUOTA).setFontColor(rouge1);
            nbDocErreurs++;
          } else {
            cellCheck.setValue(false).setNote(textSend).setFontColor(bleu2);
            nbDocErreurs++;
          }
        }  else if(texte.indexOf("#ERROR!")!=-1) { //Pb texte
          cellCheck.setValue(false).setFontColor(rouge1).setNote("Pb correction (#ERROR)");
          nbDocErreurs++;
        }
        else if(adresseMail.toString().indexOf("@")>0) { //Pb mail
          cellCheck.setValue(false).setFontColor(bleu1).setNote("Pb adresse mail");
          nbDocErreurs++;
        }
      } else if(checkMail==true) {
        nbDocDejaMailes++;
      } else {
        nbDocErreurs++;
        cellCheck.setValue(false).setFontColor(bleu1).setNote("erreur non connue");

      }
    }
    var compteRendu="";
    if(nbDocMailes>0) compteRendu+=nbDocMailes+" mails\n";
    if(nbDocRemailes>0) compteRendu+=nbDocRemailes+" remails\n";
    if(nbDocErreurs>0) compteRendu+=nbDocErreurs+" erreurs\n";
    if(nbDocDejaMailes>0) compteRendu+=nbDocDejaMailes+" déjà faits\n";
    
    //Dans retour : la case à remplir pour avoir le nb de mails
    //function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {
    if(numLigneScript) {//On est en gestion
      
      return new Info(numLigneScript,null,nbDocMailes+nbDocDejaMailes,null,compteRendu,null,null);
    } else {//cas trigger
    
      return new Info(-1,null,nbDocMailes+nbDocDejaMailes,null,compteRendu,null,null);
    }
  }
  catch(e) {
    var r=callErreur(e); if(r) r.affiche();//erreur type Info
    if(numLigneScript) {//On est en gestion
    return new Info(numLigneScript,null,null,"Erreur generation Mails",null,null,null);
    } else {//cas trigger
    return new Info(-1,null,null,"Erreur generation Mails",null,null,null);
    }
    
  }
}

function fixeData(rangeReponseEtCorrection,dataReponseEtCorrection) {
  for(var i=0;i<dataReponseEtCorrection.length;i++) {
    var cell=rangeReponseEtCorrection.getCell(1,i+1);
    cell.setValue(dataReponseEtCorrection[i]);
  }
}

function creeTexteMail(titreForm,titresUnique,intitulesItems,pointsQ,dataReponseEtCorrection,total,sur,classeurReponse) {
  
  var texteMail="<H1>Evaluation du questionnaire "+titreForm+"</H1>"+RC;
  texteMail+="<H2>Voici tes réponses et tes points par question:</H2>"+RC;
  var ilYADesPointsNonEvalues=false;
  
  //Pour chaque question
  //Si pas ":" et commence par "F": début d'une fonction/programme
  //Tout dans un tableau
  
  for(var i=0;i<titresUnique.length;i+=2) {
    var code=titresUnique[i];
    var question =intitulesItems[i];
    var points =pointsQ[i];
    if(typeof points === 'string') {
      points=points.split(";")[0];
    }
    //Logger.log(dataReponseEtCorrection[i+1]+":"+dataReponseEtCorrection[i]);
    if(code.toString().substring(0, 6)=="POINTS" && points>0) {
      texteMail+="<H2>"+question+"</H2>"+RC;
      texteMail+="<P>Ta réponse te rapporte: "+Math.round(dataReponseEtCorrection[i]*4)/4+"/"+points+"</P>"+RC;
    }
    else if(code.toString().indexOf("Total F")!=-1) { //FONCTION
      //Début de traitement de critères
      //En tête
      //Numero
      ilYADesPointsNonEvalues=true;
      var nomRangeIntitule="intituleQ"+code.toString().replace("Total ","");
      var copieReponse=dataReponseEtCorrection[i+1]; //La réponse
      var intitule=classeurReponse.getRangeByName(nomRangeIntitule).getValue();
      texteMail+="<H2>"+intitule+"</H2>"+RC;
      texteMail+="<H4>Ta réponse:</H4>"+RC;
      texteMail+="<P>"+copieReponse+"</P>"+RC;
      texteMail+="<H3>sera évaluée manuellement par le professeur sur les critères suivants :</H3>"+RC;
      
      var indice=i+2;var ncrt=0;
      do {
        //On cherche un intitulé "correction"
        ncrt++;//Nb critères
        texteMail+="<P>"+intitulesItems[indice]+" ("+pointsQ[indice]+" pts)</p>"+RC;
        indice+=2;
        //Logger.log("Critères "+ncrt+" "+question+RC+code);
        code=titresUnique[indice];
      } while (code.toString().indexOf("CORRECTION F")==-1);
      i+=ncrt;//ncrt colonnes par question
    } 
  }
  texteMail+="<H1>Total des points : "+total+"/"+sur+"</H1>";
  if(ilYADesPointsNonEvalues) {
    texteMail+="<H2>(Avant la correction manuelle)</H2>";
  }
  //Logger.log(texteMail);
  return texteMail;
}