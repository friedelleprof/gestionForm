function getFeuilleReponses(classeur) {
  
  //La feuille réponse s'appelle "réponse au formulaire 1"
  //A adapter suivant la langue
  
  var sheetReponseAuFormulaire=classeur.getSheetByName(NOM_FEUILLE_REPONSE);
  if(sheetReponseAuFormulaire==null) {
    sheetReponseAuFormulaire=classeur.getSheetByName(NOM_FEUILLE_REPONSEA);
    if(sheetReponseAuFormulaire) {
      sheetReponseAuFormulaire.setName(NOM_FEUILLE_REPONSE);
    } else {
      sheetReponseAuFormulaire=classeur.getSheets()[0];
      sheetReponseAuFormulaire.setName(NOM_FEUILLE_REPONSE);
    }
  }
  return sheetReponseAuFormulaire;
}

function removeFeuille(nomFeuille, classeur) {
  var sheetFeuille1=classeur.getSheetByName(nomFeuille);
  if(sheetFeuille1) {
    classeur.deleteSheet(sheetFeuille1);
  }
}

var EMAIL_TEXTE="Adresse e-mail";EMAIL_TEXTE2="Email Address";



//Séparateur des CHECKBOX : ";"
var SEPARATEUR_CB=";"; 

function reponsesCorrige(feuille) {
  
  //On veut une liste des réponses
  //Dans un tableau indicé par l'ITEM
  //Ou vide
  
  var dataReponses;
  var rangeMail=rangeColonneDuTexte(EMAIL_TEXTE,1,feuille);
  if(rangeMail==null) rangeMail=rangeColonneDuTexte(EMAIL_TEXTE2,1,feuille);
  if(rangeMail==null) { 
    //Pas de correction trouvée
    dataReponses=new Array(feuille.getMaxColumns()); //Même taille que ligneQuestions
  } else {
    //Recherche d'une correction
    var ligneReponses=-1;
    for(var i in LISTE_CORRECTEURS) {
      var l=TrouveNumLigneDatas(LISTE_CORRECTEURS[i],rangeMail.getValues());
      if(l) ligneReponses=l;
    }
    
    if(ligneReponses!=-1) {
      dataReponses=feuille.getRange(ligneReponses,1,1,feuille.getMaxColumns()).getValues();
      feuille.getRange(ligneReponses).setBackground(orange2).setWrap(true).setVerticalAlignment('middle');
      
    } else {
      //Pas de correction trouvée
      dataReponses=new Array(feuille.getMaxColumns()); //Même taille que ligneQuestions
    }
  }
  return dataReponses;
}

function creeListeItems(formulaire,dataReponses,dataQuestions) {
  //Renvoie un tableau d'objets ITEMS
  //function ITEM(numColonne_,type_,titre_,reponse_,helpText) 
  
  var listeItems=formulaire.getItems();
  var result=new Array();
  
  var obItem,item,titre,type,help,reponse,numColonne;
  for(var i in listeItems) {
    item=listeItems[i];
    titre=item.getTitle();
    type=item.getType();
    help=item.getHelpText();
    if(help==undefined || help==null) help=""; else help=" ("+help+")";
    
    if(type== FormApp.ItemType.GRID) { //cas particulier, la réponse est décomposée par ligne
      var itemGrid=item.asGridItem();
      var colonnes=itemGrid.getRows();
      for(var j in colonnes) {
        var nomCol=titre+ " ["+colonnes[j]+"]";
        numColonne=dataQuestions.indexOf(titre);
        if(numColonne) {
          reponse=dataReponses[numColonne];
        } else {
          //Pas de colonne avec ce titre
          reponse="";
          numColonne=1;
        }
        obItem=new ITEM(numColonne+1,"colonne de GRID",nomCol,reponse,help)
      }
    } //Sinon cas normal
    else {
      numColonne=dataQuestions.indexOf(titre);
      if(numColonne) {
        reponse=dataReponses[numColonne];
      } else {
        //Pas de colonne avec ce titre
        reponse="";
        numColonne=-1;
      }
      obItem=new ITEM(numColonne+1,type,titre,reponse,help)
    }
    result.push(obItem);
    Logger.log(obItem.log());
  }
  return result;
}

var LISTE_CODE_CRITERES=["critères","criteres","critere","critère"];

function detNbCritere(reponse) {
  var nbCritere=-1;
  for(var code in LISTE_CODE_CRITERES) {
    if(reponse.indexOf(code)>=0) {
      nbCritere=Number(reponse.replace(code,""));
    }
  }
if(nbCritere==NaN || nbCritere==null ) nbCritere=-1;

  return nbCritere;
}


  
  