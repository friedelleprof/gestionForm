function testOI() {
removeDeadReferences("https://docs.google.com/spreadsheets/d/1FbMnECOK-djGM1PNossqU8sEqHJ4pwFGU1g8JE08JLM/edit");

  //var formulaire=FormApp.openByUrl("https://docs.google.com/forms/d/11zKESc-sOjzNd_232y9WPld0YwEAyQ4lBE_m76i88ZA/edit");
  //var classeur=SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1XkEcX3mHdNG5pRZxwTUS62qdlP0b-n0FVRNZHUo4EjQ/edit");
  //creeFeuilleQuestionReponse(formulaire,classeur);
}


function creeFeuilleQuestionReponse(formulaire,classeur) {
  
  //ICI on considère que le SCRIPT des fonctions perso est installé
  //Crée UNIQUEMENT la feuille QUESTION REPONSE
  
  //Droits pour visibilité
  //var fichierSS=DriveApp.getFileById(classeur.getId());
  //  fichierSS.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  classeur.setSpreadsheetLocale('fr'); //Les virgules sont des POINTS
  
  var sheetReponseAuFormulaire=getFeuilleReponses(classeur)
  var sheetQuestionReponses=creeNouvelleFeuille(classeur,NOM_FEUILLE_QR);
  
  //Récupération des items par type
  
  //Ligne 1 des titres d'items, ordre de la feuille
  var dataQuestions=sheetReponseAuFormulaire.getRange(1,1,1,sheetReponseAuFormulaire.getMaxColumns())
  .setBackground(orange1).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(12).setWrap(true)
  .getValues()[0];
  
  //var dataReponses= reponsesCorrige(sheetReponseAuFormulaire);
  //Liste des réponses dans l'ordre de la feuille
  
  var listeItems=creeListeItems(formulaire,dataQuestions);//Tableau avec des objets ITEMS
  var rowQuestion,numQuestion=1;  
  for(var i=0;i<listeItems.length;i++){
    if(listeItems[i].type==FormApp.ItemType.PARAGRAPH_TEXT) {   
      //On masque la colonne
      //sheetReponseAuFormulaire.hideColumns(listeItems[i].numColonne);
      rowQuestion=new ligneCorrection(listeItems[i],"F"+(i+1),sheetQuestionReponses);
      rowQuestion.draw(i+2);
    } else {
      rowQuestion=new ligneCorrection(listeItems[i],"Q"+(i+1),sheetQuestionReponses);
      rowQuestion.draw(i+2);
    }
  }
  var columnWidth=[8,8,1,80,80,14,1,30,30,30];
  var columnAlignement=['center','center','center','left','left','center','center','left','left','left'];
  var columnSize=[12,14,10,12,9,14,14,12,12,12];
  for(var i=0;i<columnWidth.length;i++) {
    sheetQuestionReponses.setColumnWidth(i+1,columnWidth[i]*5);
    sheetQuestionReponses.getRange(1,i+1,sheetQuestionReponses.getMaxRows(),1).setHorizontalAlignment(columnAlignement[i]).setVerticalAlignment('middle').setWrap(true).setFontSize(columnSize[i]);
  }
  
  sheetQuestionReponses.getRange(1,1,1,10).setValues([["Code","Col","type","Questions","Réponses","=SUM(G2:G"+(listeItems.length+1)+")&\" points\"","max points","Correction si juste","Correction si incorrect","Correction si faux"]])
  .setBackground(orange1).setFontSize(15).setHorizontalAlignment('center');
  
  
}

function creesheetCorrection(classeur) {
  //Uniquement à partir de la feuille questionReponse
  


}





