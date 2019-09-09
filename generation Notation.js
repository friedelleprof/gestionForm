function genererFeuilleQuestionReponse(formulaire,classeur) {
  //false:ne pas refaire la correction
  return genererFeuilleQuestionReponseTRUEFALSE(formulaire,classeur,false);
}
//function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {

function genererFeuilleQuestionReponseTRUEFALSE(formulaire,classeur,nePasRefaireCorrection) {//Avec modif QUESTIONREPONSES
  try {
    metAJourFeuilleInfo(classeur,"Lancement génération de la correction");
    var newURL=miseAJourScriptClasseur(formulaire,classeur).retour;
    if(classeur.getUrl()!=newURL) {
      //Le classeur a été mis à jour pour le script
      metAJourFeuilleInfo(classeur,"Mise à jour URL:"+newURL);
      //cellURLSheet.setValue(newURL);
      classeur=SpreadsheetApp.openByUrl(newURL);
    }
    var retour=creeFeuilleQuestionsReponses(formulaire,classeur,nePasRefaireCorrection);
    if(retour==null) {//Si pas d'erreurs
      return new Info(numLigneScript,null,newURL,"generation OK/"+nePasRefaireCorrection,"",null,null);
    }
    else {
      retour.retour=newURL;
      return retour;
    }
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function creeFeuilleQuestionsReponses(formulaire,classeur,nePasRefaireCorrection) {
  try {
    //  nePasRefaireCorrection=true:pour ne pas modifier la feuille correction mais utiliser ses informations
    openDatas();
    AjoutInfos(classeur);
    
    //Droits pour visibilité
    var fichierSS=DriveApp.getFileById(classeur.getId());
    fichierSS.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    classeur.setSpreadsheetLocale('fr'); //Les virgules sont des POINTS
    
    //La feuille réponse s'appelle "réponse au formulaire 1"
    var sheetReponseAuFormulaire=classeur.getSheetByName(NOM_FEUILLE_REPONSE);
    if(sheetReponseAuFormulaire==null) {
      sheetReponseAuFormulaire=classeur.getSheetByName(NOM_FEUILLE_REPONSEA);
      sheetReponseAuFormulaire.setName(NOM_FEUILLE_REPONSE);
    }
    //Suppression éventuelle de 'feuille 1'
    var sheetFeuille1=classeur.getSheetByName("Feuille 1");
    if(sheetFeuille1) {
      classeur.deleteSheet(sheetFeuille1);
    }
    var sheetQuestionReponses, sheetCorrection;
    var dataRang1=sheetReponseAuFormulaire.getRange(1,1,1,sheetReponseAuFormulaire.getMaxColumns()).getValues();
    
    var formulaTotal="=0";
    //On cherche la ligne question
    var questionsRange=sheetReponseAuFormulaire.getRange("1:1").setBackground(orange1).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(12).setWrap(true);
    var mailRange=rangeColonneDuTexte(EMAIL_TEXTE,1,sheetReponseAuFormulaire);
    if(mailRange==null) mailRange=rangeColonneDuTexte(EMAIL_TEXTE2,1,sheetReponseAuFormulaire);
    if(mailRange==null) return ERROR+"pas de champs MAIL";
    
    var dataMailRange=mailRange.getValues();
    var ligneReponses;
    //Séparateur des CHECKBOX : ";"
    var separateurCB=";"; 
    for(var i in LISTE_CORRECTEURS) {
      var l=TrouveNumLigneDatas(LISTE_CORRECTEURS[i],dataMailRange);
      if(l) ligneReponses=l;
    }
    
    //Récupération des items par type 
    var typeItems=infoFormulaire(formulaire);var typeItem;
    var helpText=getHelpText(formulaire);
    
    //Création des feuilles, sauvegarde éventuelle
    if(classeur.getSheetByName(NOM_FEUILLE_QR)==null || !nePasRefaireCorrection) {
      sheetQuestionReponses=creeNouvelleFeuille(classeur,NOM_FEUILLE_QR,TEST);
      nePasRefaireCorrection=false;
    } else {
      sheetQuestionReponses=classeur.getSheetByName(NOM_FEUILLE_QR);
    }
    sheetCorrection=creeNouvelleFeuille(classeur,NOM_FEUILLE_CORRECTION,TEST);
    //Limiter le Nb de rangées
    
    var numFonction=1,ligneQR=2;//Num de la ligne rajoutée, colonne question lue, numero fonctiont traitée
    var numeroQuestion=1,colCorrection=1;
    
    var celOrigine,celprecedente,celMEMOREP; //Pour réponses précédentes et mémorisées #REP#, #PRECED#, #MEMOREP#
    
    
    //POUR CHAQUE ITEM :
    
    for(var titreItem in  typeItems) {
      
      if(titreItem!="") {
        var help=helpText[titreItem];      
        typeItem=typeItems[titreItem];
        var reponse="";
        //Y a-t'il une correction ?
        var colonneITEM= TrouveNumColonneDatas(titreItem, dataRang1) ;
        Logger.log(ligneReponses+" "+colonneITEM);
        if(ligneReponses  && colonneITEM) {
          reponse=sheetReponseAuFormulaire.getRange(ligneReponses,colonneITEM).getValue().toString();
        }
        Logger.log(titreItem+" "+reponse+" type:"+typeItem);
        
        
        if(colonneITEM) { //SI La colonne correction existe !
          
          celprecedente=celOrigine;//POUR FORMULES SUR PLUSIEURS NbS
          
          /****************************
          /
          / Cas des PARAGRAPHES ou FONCTIONS
          /
          /****************************/
          var nbCritere=1;
          if(reponse.indexOf("critères:")>=0) {
            nbCritere=Number(reponse.replace("critères:",""));
          }
          else if(reponse.indexOf("critère:")>=0) {
            nbCritere=Number(reponse.replace("critère:",""));
          }
          else if(reponse.indexOf("critere:")>=0) {
            nbCritere=Number(reponse.replace("critere:",""));
          }
          if(reponse==CODE_EVAL_FONCTION || typeItem==FormApp.ItemType.PARAGRAPH_TEXT) {
            if(nbCritere==NaN || nbCritere==0 || nbCritere==null ) nbCritere=1;
            
            //On masque la colonne
            sheetReponseAuFormulaire.hideColumns(colonneITEM);
            var colCorrectionDebut=colCorrection;
            //Création des items
            // Fonction
            var sum,nomB1,sumCriteres,sumNotes;
            if(reponse==CODE_EVAL_FONCTION) {
              sum=itemsEvaluesFonction.length;
            }
            else {
              sum=nbCritere;
            }
            //La fonction copiée
            sheetCorrection.getRange(1,colCorrection,6).setValues([
              ["Total F"+numFonction],
              ["="+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,1)) ],
              ["A"],
              ["B"],
              ["C"],
              [""]])
            .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);//On calculera le total dans une formule;
            
            var sumCriteres=sheetCorrection.getRange(6,colCorrection);
            var sumNotes=sheetCorrection.getRange(7,colCorrection);
            sheetCorrection.getRange(7,colCorrection,sheetCorrection.getMaxRows(),1).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10).setBackground(bleu5).setBackground(bleu5);
            
            //Ajout du sous-total
            formulaTotal += "+"+sheetCorrection.getRange(7,colCorrection+colonnesAjoutees.length).getA1Notation();
            //Formules de notation ABC
            //On calculera le total dans une formule;
            var nomBl="ABCF"+numFonction;
            
            classeur.setNamedRange(nomBl,sheetCorrection.getRange(1,colCorrection,6,sum*3+2));
            colCorrection++;
            if(!nePasRefaireCorrection) {
              sheetQuestionReponses.getRange(ligneQR,1,1,5).setValues([
                ["F"+numFonction,
                 titreItem+help,
                 "=SUM(R[1]C[0]:R["+sum+"]C[0])",
                 "",
                 ""]])
            }
            classeur.setNamedRange("IntituleQF"+numFonction,sheetQuestionReponses.getRange(ligneQR,2,1,1));
            
            //Copie fonction
            sheetCorrection.getRange(1,colCorrection).setValue("Copie F"+numFonction);
            celOrigine="'"+NOM_FEUILLE_REPONSE+"'!"+sheetReponseAuFormulaire.getRange(2,colonneITEM).getA1Notation();//setFormula("="+celOrigine)
            sheetCorrection.getRange(7,colCorrection,sheetCorrection.getMaxRows()).setValue(null)
            .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true).setFontSize(8).setBackground(gris3);
            
            sheetCorrection.setColumnWidth(colCorrection, 200);
            sheetCorrection.showColumns(colCorrection);
            
            var rangeCorrection = sheetCorrection.getRange(2,colCorrection,5,1);
            
            colCorrection++;
            
            
            ligneQR++;
            var j=2*sum+3; var ajust=2*sum;
            
            for(var i=0;i<sum;i++) {
              
              //CODE unique
              var cel;
              if(!nePasRefaireCorrection) {
                if(nbCritere==0 || reponse==CODE_EVAL_FONCTION) {
                  sheetQuestionReponses.getRange(ligneQR,1,1,8).setValues([["F"+numFonction+":It"+i,itemsEvaluesFonction[i],1,null,null,"OUI",erreursB[i],"NON"]]);
                }
                else {
                  sheetQuestionReponses.getRange(ligneQR,1,1,8).setValues([["F"+numFonction+":It"+i,"Critere "+(i+1),1,null,null,"OUI","","NON"]]);
                }
              }
              var nomPT="pointF"+numFonction+"I"+i;
              classeur.setNamedRange(nomPT,sheetQuestionReponses.getRange(ligneQR,3));
              
              var nomCT="critereF"+numFonction+"I"+i;
              classeur.setNamedRange(nomCT,sheetQuestionReponses.getRange(ligneQR,2));
              
              //Notation
              var indFeuille="='"+NOM_FEUILLE_QR+"'!";
              sheetCorrection.getRange(1,colCorrection).setFormula('='+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,1)))
              sheetCorrection.getRange(2,colCorrection).setFormula('='+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,2)))
              sheetCorrection.getRange(3,colCorrection).setFormula('='+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,6)))
              sheetCorrection.getRange(4,colCorrection).setFormula('='+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,7)))
              sheetCorrection.getRange(5,colCorrection).setFormula('='+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,8)))
              sheetCorrection.getRange(6,colCorrection).setFormula('='+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,3)))
              sheetCorrection.getRange(7,colCorrection).setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('');//Par défaut .setDataValidation(rule)
              sheetCorrection.getRange(1,colCorrection,sheetCorrection.getMaxRows(),1).setBackground(gris1);
              sheetCorrection.setColumnWidth(colCorrection, 50);
              sheetCorrection.showColumns(colCorrection);
              
              //Points
              sheetCorrection.getRange(1,colCorrection+1).setFormula('="POINTS "'&nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,1)));
              sheetCorrection.getRange(2,colCorrection,1,2).merge();
              sheetCorrection.getRange(3,colCorrection,1,2).merge();
              sheetCorrection.getRange(4,colCorrection,1,2).merge();
              sheetCorrection.getRange(5,colCorrection,1,2).merge();
              sheetCorrection.getRange(6,colCorrection,1,2).merge();
              
              cel=sheetCorrection.getRange(7,colCorrection).getA1Notation();
              sheetCorrection.getRange(7,colCorrection+1).setFormula("=IF("+cel+"=\"A\";"+nomPT+";IF("+cel+"=\"B\";"+nomPT+"/2;0))")
              .setHorizontalAlignment('center').setVerticalAlignment('middle');
              sheetCorrection.getRange(1,colCorrection+1,sheetCorrection.getMaxRows(),1).setBackground(bleu4);
              
              //Correction
              //sheetCorrection.setColumnWidth(colCorrection+ajust, 50);
              sheetCorrection.setColumnWidth(colCorrection+1, 20);
              //Masquer la colonne :
              //sheetCorrection.hideColumns(colCorrection+1);
              
              sheetCorrection.getRange(1,colCorrection+ajust).setFormula('="CORRECTION " &'+nomA1(NOM_FEUILLE_QR,sheetQuestionReponses.getRange(ligneQR,1))).setWrap(false);;
              sheetCorrection.getRange(2,colCorrection+ajust).setFormula('='+fullA1Name(sheetQuestionReponses.getRange(ligneQR,2))).setWrap(false);
              sheetCorrection.getRange(3,colCorrection+ajust).setFormula('='+fullA1Name(sheetQuestionReponses.getRange(ligneQR,6))).setWrap(false);
              sheetCorrection.getRange(4,colCorrection+ajust).setFormula('='+fullA1Name(sheetQuestionReponses.getRange(ligneQR,7))).setWrap(false);
              sheetCorrection.getRange(5,colCorrection+ajust).setFormula('='+fullA1Name(sheetQuestionReponses.getRange(ligneQR,8))).setWrap(false);
              sheetCorrection.getRange(6,colCorrection+ajust).setFormula('='+fullA1Name(sheetQuestionReponses.getRange(ligneQR,3))).setWrap(false);
              
              sheetCorrection.getRange(7,colCorrection+ajust).setFormula("IFERROR(VLOOKUP("+cel+";"+nomBl+";"+j+";false);\"non évalué\")").setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
              sheetCorrection.getRange(1,colCorrection+ajust,sheetCorrection.getMaxRows(),1).setBackground(gris3);
              sheetCorrection.hideColumns(colCorrection+ajust);
              //Logger.log("masquer:"+(colCorrection+ajust));
              j++;ajust--;
              
              ligneQR++;colCorrection+=2;
            } //Fin boucle sur critères
            
            if(!nePasRefaireCorrection) {
              
              if(numFonction%2==0) {
                sheetQuestionReponses.getRange(ligneQR-sum,1,sum+1,8).setBackground(vert2);
              }
              else {
                sheetQuestionReponses.getRange(ligneQR-sum,1,sum+1,8).setBackground(vert3);
              }
              sheetQuestionReponses.getRange(ligneQR-1-sum,1,1,8).setBackground(vert1);
              sheetQuestionReponses.getRange(ligneQR-sum,4,sum,2).merge();
              sheetQuestionReponses.getRange(ligneQR-sum-1,4,1,2).merge().setValue(NOM_COLONNE_LIENS_CORRIGES).setFontSize(12).setHorizontalAlignment('center');
            }
            
            classeur.setNamedRange("CorrigeF"+numFonction,sheetQuestionReponses.getRange(ligneQR-sum,4));
            rangeCorrection.merge().setFormula("=CorrigeF"+numFonction).setFontSize(7).setHorizontalAlignment('left');//afficheReponse("+colonneITEM+";CorrigeF"+numFonction+")");
            sheetCorrection.getRange(1,colCorrectionDebut,sheetCorrection.getMaxRows(),colCorrection-colCorrectionDebut).setBorder(false, true, false, true, false, false, rouge1, SpreadsheetApp.BorderStyle.SOLID_THICK);
            numFonction++;colCorrection+=sum;
            sheetCorrection.getRange(1,colCorrectionDebut,sheetCorrection.getMaxRows(),colCorrection-colCorrectionDebut+1).setBorder(false, true, false, true, false, false, noir, SpreadsheetApp.BorderStyle.SOLID_THICK);
            //,
            sumCriteres.setFormula("=SUM(R[0]C["+2+"]:R[0]C["+(sum*2+1)+"])");
            sumNotes.setFormula("=sumEtReponse(SUM(R[0]C[2]:R[0]C["+(sum*2+1)+"]);"+celOrigine+")");
            
            //Logger.log("SOMME CRITERES:"+sumCriteres.getA1Notation()+RC+"Col max"+sheetCorrection.getMaxColumns()+RC+sumCriteres.getFormula());
            
          } 
          
          /****************************
          /
          / AUTRES CAS
          /
          /****************************/
          
          
          //Pas ITEMS
          else {
            
            if(!nePasRefaireCorrection) {
              
              sheetQuestionReponses.getRange(ligneQR,1).setValue("Q"+numeroQuestion+": "+titreItem+help);
              if(typeItem==FormApp.ItemType.CHECKBOX || reponse.indexOf("%cb")>=0) {
                sheetQuestionReponses.getRange(ligneQR,2).setValue(reponse.replace("%cb", "").trim().split(",").join(separateurCB).trim());
                sheetQuestionReponses.getRange(ligneQR,4).setValue("");
                sheetQuestionReponses.getRange(ligneQR,5).setValue(reponse.replace("%cb", "").trim().split(",").join(RC).trim());
                sheetQuestionReponses.getRange(ligneQR,4,1,2).merge();
                sheetQuestionReponses.getRange(ligneQR,3).setValue("2;1");
              }
              else if(reponse.indexOf("%or")>=0) {
                //Pour les OU, écrire A %or B %or C...
                var temp="#"+reponse.trim().split("%or").join("#;#").trim()+"#";
                temp.replace(/# /g,"#").replace(/ #/g,"#");
                sheetQuestionReponses.getRange(ligneQR,2).setValue(temp);
                sheetQuestionReponses.getRange(ligneQR,4).setValue("C'est une bonne réponse");//Ligne réponse
                sheetQuestionReponses.getRange(ligneQR,5).setValue("On attendait :"+reponse.trim().split("%or").join(" ou "));//Ligne réponse fausse
                sheetQuestionReponses.getRange(ligneQR,3).setValue(1);
              } 
              else {
                sheetQuestionReponses.getRange(ligneQR,2).setValue(reponse.trim().split(",").join(separateurCB).trim());
                sheetQuestionReponses.getRange(ligneQR,4).setValue("C'est une bonne réponse");//Ligne réponse
                sheetQuestionReponses.getRange(ligneQR,5).setValue("On attendait :"+reponse.trim());//Ligne réponse fausse
                sheetQuestionReponses.getRange(ligneQR,3).setValue(1);
              } 
              
              if(ligneQR%2==0) {
                sheetQuestionReponses.getRange(ligneQR,1,1,8).setBackground(jaune3);
              }
              else {
                sheetQuestionReponses.getRange(ligneQR,1,1,8).setBackground(gris3);
              }
            }
            
            
            classeur.setNamedRange("nomQ"+numeroQuestion,sheetQuestionReponses.getRange(ligneQR,1));
            classeur.setNamedRange("corre"+numeroQuestion,sheetQuestionReponses.getRange(ligneQR,2));
            classeur.setNamedRange("point"+numeroQuestion,sheetQuestionReponses.getRange(ligneQR,3));
            classeur.setNamedRange("brepo"+numeroQuestion,sheetQuestionReponses.getRange(ligneQR,4));
            classeur.setNamedRange("frepo"+numeroQuestion,sheetQuestionReponses.getRange(ligneQR,5));
            
            //Points question
            //Correction : points
            sheetCorrection.getRange(1,colCorrection).setValue("POINTS Q"+numeroQuestion);
            sheetCorrection.getRange(2,colCorrection).setFormula("=nomQ"+numeroQuestion);//Enoncé
            sheetCorrection.getRange(3,colCorrection).setFormula("=corre"+numeroQuestion);//Réponse
            //En 4 et 5 ajout des moyennes
            sheetCorrection.getRange(6,colCorrection).setFormula('=max(split(point'+numeroQuestion+';";"))');//Points
            
            sheetCorrection.getRange(1,colCorrection,sheetCorrection.getMaxRows()).setBackground(bleu4);
            sheetCorrection.setColumnWidth(colCorrection, 40);
            
            celOrigine=nomA1(NOM_FEUILLE_REPONSE,sheetReponseAuFormulaire.getRange(2,colonneITEM));
            
            var bonneReponse;
            var fausseReponse;
            var celPts;
            
            //Correction
            //IF(FIND('Réponses élèves'!T3; corre19; 1)
            var size=0;
            if(typeItem==FormApp.ItemType.CHECKBOX || reponse.indexOf("%cb")>=0) { 
              formul="=evalueReponse("+ celOrigine+";corre"+numeroQuestion+";point"+numeroQuestion+";brepo"+numeroQuestion+";frepo"+numeroQuestion+")";
              celPts="'"+NOM_FEUILLE_CORRECTION+"'!"+sheetCorrection.getRange(7,colCorrection).setFormula(formul).setNumberFormat("0.00")
              .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).getA1Notation();
              bonneReponse="";//'texteCheckBox2('+celOrigine+';corre'+numeroQuestion+")";
              fausseReponse="";//'texteCheckBox2('+celOrigine+';corre'+numeroQuestion+")";
              size=220;
            } 
            
            else if(reponse.indexOf("TEST:")!=-1) {  
              //Cas des formules personnalisées
              //on remplace flb_response et REP par celOrigine
              
              //Possibilité formule sur REPONSE PRECEDENTE : PREC#
              if(reponse.indexOf("#MEMOREP#")!=-1) {
                //Logger.log("#MEMOREP:"+celOrigine);
                //Mémorisation d'un résultat // Ecrase la précédente mémo
                celMEMOREP=celOrigine;
                celPts="'"+NOM_FEUILLE_CORRECTION+"'!"+sheetCorrection.getRange(7,colCorrection).setFormula("=point"+numeroQuestion).setNumberFormat("0.00")
                .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).getA1Notation();
                bonneReponse='"Tu as répondu :" &'+celOrigine;
                fausseReponse='"Tu as répondu :" &'+celOrigine;
                if(!nePasRefaireCorrection) {
                  sheetQuestionReponses.getRange(ligneQR,5).setValue("");
                }
              }
              else {
                formul="=evalFormule("+ celOrigine+";"+ celprecedente+";"+celMEMOREP+";corre"+numeroQuestion+";point"+numeroQuestion+";brepo"+numeroQuestion+";frepo"+numeroQuestion+")";
                celPts=nomA1(NOM_FEUILLE_CORRECTION,sheetCorrection.getRange(7,colCorrection).setFormula(formul).setNumberFormat("0.00")
                             .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true));
                bonneReponse="";
                fausseReponse="";
                if(!nePasRefaireCorrection) {
                  sheetQuestionReponses.getRange(ligneQR,5).setValue("C'est incorrect");
                }
              }
            }
            else {//Eval normale
              formul="=evalueReponse("+ celOrigine+";corre"+numeroQuestion+";point"+numeroQuestion+";brepo"+numeroQuestion+";frepo"+numeroQuestion+")";
              celPts="'"+NOM_FEUILLE_CORRECTION+"'!"+sheetCorrection.getRange(7,colCorrection).setFormula(formul).setNumberFormat("0.00")
              .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).getA1Notation();
              bonneReponse="";
              fausseReponse="";
            }
            formulaTotal += '+'+sheetCorrection.getRange(7,colCorrection+colonnesAjoutees.length).getA1Notation();//Forume de calcul note
            
            colCorrection++;
            
            //correction question
            sheetCorrection.getRange(1,colCorrection).setValue("CORRECTION Q"+numeroQuestion);
            sheetCorrection.getRange(2,colCorrection-1,1,2).merge();
            sheetCorrection.getRange(3,colCorrection-1,1,2).merge();
            sheetCorrection.getRange(6,colCorrection-1,1,2).merge();
            sheetCorrection.getRange(1,colCorrection,sheetCorrection.getMaxRows()).setBackground(gris3);
            
            
            if(celPts=="" || bonneReponse=="") {
              sheetCorrection.getRange(7,colCorrection).setValue(null)
              .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true).setFontSize(8);
            } else if(bonneReponse==fausseReponse) {
              sheetCorrection.getRange(7,colCorrection).setFormula("="+bonneReponse)
              .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true).setFontSize(8);
            } else {
              sheetCorrection.getRange(7,colCorrection).setFormula("=IF("+celPts+"=point"+numeroQuestion+";"+bonneReponse+";"+fausseReponse+")")
              .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true).setFontSize(8);
            }
            sheetCorrection.setColumnWidth(colCorrection, 80+size);
            
            ligneQR++;numeroQuestion++;colCorrection++;
          }
          
        }
      }
    }
    //FIN BOUCLE
    
    //Couleur reponses
    if(ligneReponses && colonneITEM) {
      sheetReponseAuFormulaire.getRange(ligneReponses,1,1,colonneITEM).setBackground(orange2).setWrap(true).setVerticalAlignment('middle');
    }
    var NbDeQuestions=ligneQR-1;
    //Logger.log(NbDeQuestions);
    var listeQuestions = sheetQuestionReponses.getRange("A2:A"+NbDeQuestions);
    classeur.setNamedRange("listeQuestions", listeQuestions);
    var listeReponses = sheetQuestionReponses.getRange("B2:B"+NbDeQuestions);
    classeur.setNamedRange("listeReponses", listeReponses);
    var pointsQuestions = sheetQuestionReponses.getRange("C2:C"+NbDeQuestions);
    classeur.setNamedRange("listePoints", pointsQuestions);
    
    //Mise en forme
    listeQuestions.setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheetQuestionReponses.setColumnWidth(1,300);//Questions
    listeReponses.setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheetQuestionReponses.setColumnWidth(2,300).getRange("B2:B"+NbDeQuestions).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);//Réponses
    sheetQuestionReponses.setColumnWidth(3,140).getRange("C2:C"+NbDeQuestions).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);//Points
    sheetQuestionReponses.setColumnWidth(4,200).getRange("D2:D"+NbDeQuestions).setVerticalAlignment('middle').setWrap(true);
    sheetQuestionReponses.setColumnWidth(5,200).getRange("E2:E"+NbDeQuestions).setVerticalAlignment('middle').setWrap(true);
    sheetQuestionReponses.setColumnWidth(6,80).getRange("F2:F"+NbDeQuestions).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheetQuestionReponses.setColumnWidth(7,200).getRange("G2:G"+NbDeQuestions).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    sheetQuestionReponses.setColumnWidth(8,80).getRange("H2:H"+NbDeQuestions).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sheetQuestionReponses.getRange("1:1").setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(18).setBackground(orange1);
    
    
    sheetCorrection.getRange("1:1").setBackground(jaune1).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
    sheetCorrection.getRange("3:3").setBackground(vert2).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9).setWrap(true);
    sheetCorrection.getRange("4:4").setBackground(orange1).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(8);
    sheetCorrection.getRange("5:5").setBackground(rouge1).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(8);
    sheetCorrection.getRange("2:2").setBackground(gris1).setVerticalAlignment('middle').setFontSize(6).setWrap(true);
    sheetCorrection.getRange("6:6").setBackground(gris2).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);
    
    // Ajout : check MAIL, fichier correction généré, fichier MAIL généré
    
    sheetCorrection.getRange(2,colCorrection).setValue(NOM_COLONNE_CHECK_ENVOI_MAIL).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheetCorrection.getRange(2,colCorrection+1).setValue("texte Mail").setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
    
    //Insertion nom, mail, total etc
    //On utilise une variable (tableau)
    sheetCorrection.insertColumns(1, colonnesAjoutees.length);
    for(var i=0 ;i<colonnesAjoutees.length;i++) {
      sheetCorrection.getRange(1,i+1).setValue(colonnesAjoutees[i]);
      sheetCorrection.setColumnWidth(i+1,colonnesAjouteesLargeur[i]);
    }
    colCorrection+=colonnesAjoutees.length;
    
    
    //Première ligne de correction
    
    sheetCorrection.getRange(7,1).setFormula("="+nomA1(NOM_FEUILLE_REPONSE,sheetReponseAuFormulaire.getRange(2,mailRange.getColumn())));
    sheetCorrection.getRange(7,2).setFormula('=QUERY(IMPORTRANGE('+camelize(ID_FEUILLE_SUIVI)+';"'+NOM_RANGE_INFO_CLASSES+'");"SELECT '+reqColNom+","+reqColPrenom+' WHERE '+reqColMail+'=\'"&A7&"\' limit 1";-1)');
    
    // A changer si TOTAL change de place
    sheetCorrection.getRange(7,4).setFormula(formulaTotal).setNumberFormat("0.00");
    sheetCorrection.getRange(7,4).copyTo(sheetCorrection.getRange(6,4));
    classeur.setNamedRange("totalPoints", sheetCorrection.getRange("D6"));
    sheetCorrection.getRange("B2").setValue("Total pt").setFontSize(12);
    sheetCorrection.getRange("B3").setFormula("=totalPoints").setFontSize(18);
    sheetCorrection.getRange("C2").setValue("Noté sur").setFontSize(12);
    classeur.setNamedRange("sur", sheetCorrection.getRange("C3"));
    sheetCorrection.getRange("C3").setValue(20).setFontSize(14);
    var celprov=sheetCorrection.getRange(7,colCorrection+1);
    
    classeur.setNamedRange("titresCorrection", sheetCorrection.getRange("1:1"));
    classeur.setNamedRange("questionCorrection", sheetCorrection.getRange("2:2"));
    classeur.setNamedRange("reponseCorrection", sheetCorrection.getRange("3:3"));
    classeur.setNamedRange("reponseBCorrection", sheetCorrection.getRange("4:4"));
    classeur.setNamedRange("reponseCCorrection", sheetCorrection.getRange("5:5"));
    classeur.setNamedRange("pointsCorrection", sheetCorrection.getRange("6:6"));  
    //Masquages et 
    sheetCorrection.hideRows(1);
    sheetCorrection.hideRows(6);
    sheetCorrection.setFrozenColumns(4);
    sheetCorrection.setFrozenRows(6);
    sheetQuestionReponses.setFrozenRows(1);
    sheetQuestionReponses.setFrozenColumns(1);
    
    //Copie du premier rang de correction si déjà des réponses
    
    if(sheetReponseAuFormulaire.getLastRow()>2) {
      sheetCorrection.getRange(7,1,1,sheetCorrection.getMaxColumns()).copyTo(sheetCorrection.getRange(8,1,sheetReponseAuFormulaire.getLastRow()-2,sheetCorrection.getMaxColumns()));
    }
    
    
    //Moyenne, max, min
    var rangeNotes=sheetCorrection.getRange("D7:E"+sheetCorrection.getLastRow());
    classeur.setNamedRange(NOM_RANGE_NOTES, rangeNotes);
    sheetCorrection.getRange("D2").setValue("MOY MAX MIN").setFontSize(12);
    sheetCorrection.getRange("D3").setFormula('=AVERAGEIF(notes;">0")').setFontSize(14).setNumberFormat("0.00");
    sheetCorrection.getRange("D4").setFormula('=MAXIFS(notes;notes;">0")').setFontSize(14).setNumberFormat("0.00");
    sheetCorrection.getRange("D5").setFormula('=MINIFS(notes;notes;">0")').setFontSize(14).setNumberFormat("0.00");
    
    //Suppression des colonnes en trop :
    var nbCol=sheetCorrection.getMaxColumns()-colCorrection-3;
    if (nbCol>0) sheetCorrection.deleteColumns(colCorrection+3, nbCol);
    
    //Titres
    if(!nePasRefaireCorrection) {
      sheetQuestionReponses.getRange(1,1,1,8).setValues([["Questions","Réponses","","Correction si juste","Correction si faux","com A","com B","com C"]]);
    }
    sheetQuestionReponses.getRange(1,3).setValue('="Points:"&totalPoints');
    
    ajoutMoyennesCorrectionEtFormatConditionnels(classeur);
    if(nePasRefaireCorrection) {
      return new Info(numLigneScript,null,null,"generation OK sans refaire correction","",null,null);
    } else {
      return new Info(numLigneScript,null,null,"generation OK + correction refaite","",null,null);
    }
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}












