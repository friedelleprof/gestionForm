//function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {
var numLigneScript;

//A MODIFIER EN ONEDIT

function scriptGeneratif() {
  //Début traitement de la liste
  openDatas();
  
  for(var i=0;i<dataNomFormulaire.length;i++) {
    try {
      var nom=dataNomFormulaire[i][0];
      var URLForm=dataURLForm[i][0];
      var URLSheet=dataURLSheet[i][0];
      var nomScript=dataScript[i][0];
      numLigneScript=1+i;
      var retourInfo=null;
      if(URLForm=="" && URLSheet=="" && nomScript!="") {
        retourInfo=new Info(numLigneScript,null,null,"PAS DE FEUILLE NI DE FORMULAIRE",null,rouge2);
      } else if(nomScript!=""){
        switch (nomScript) {
        
          case "generer": 
            var cellURLSheet=rangeURLSheet.getCell(numLigneScript,1);
            retourInfo=genererFeuilleQuestionReponse(FormApp.openByUrl(URLForm),SpreadsheetApp.openByUrl(URLSheet)) ;
            //retour contient URLSheet
            if(retourInfo) retourInfo.numColonneRetour=rangeURLSheet.getColumn();
            break;
            
          case "INFOS": 
            retourInfo=getInfos(URLForm,URLSheet);
            break;
            
          case "set:nomFormulaire": 
            var nom=rangeNom.getCell(numLigneScript,1).getValue();
            retourInfo=setNom(URLForm,URLSheet,nom) ;
            break;          
            
          case "miseAJourSCRIPT":
            retourInfo=miseAJourScriptClasseur(FormApp.openByUrl(URLForm),SpreadsheetApp.openByUrl(URLSheet));//le script renvoie la nouvelle URL
            break;
            
          case "Envois Mails":
            retourInfo=checkEtEnvoiMail(URLSheet,false);
            retourInfo.numColonneRetour=rangeNbMail.getColumn()
            break;
            
          case "Clean reponses":
            retourInfo=supprimerToutesReponses(URLForm,false);
            break;
            
          case "creeDocReponse":
            retourInfo=creeDocReponse(URLSheet,URLForm,false);
            break;
            
          case "refaireDocReponse":
            retourInfo=creeDocReponse(URLSheet,URLForm,true);
            break;
            
          case "envoiDocReponse":
            retourInfo=envoiDocReponse(URLSheet,URLForm,false);            
            break;
          case "RENVOYER DocReponse":
            retourInfo=envoiDocReponse(URLSheet,URLForm,true);
            break;
            
          case "copieDocToDossierSuivi":
            retourInfo=copieDocToDossierSuivi(URLSheet,URLForm);
            break;
            
          case "cleanFeuilles":
            retourInfo=cleanFeuilles(URLSheet);
            break;
            
          default:
            break;
        }
      }
      if(retourInfo) retourInfo.affiche();
    } //Fin try
    catch (e) {
      callErreur(e).affiche();//erreur type Info
    }
  } //fin for
  
} //fin function





