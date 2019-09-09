//Objet info // retour des getteurs/setteurs pour affichage
function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {
  if (numLigneScript_) this.numLigneScript=numLigneScript_; else this.numLigneScript=-1;
  if(numColonneRetour_) this.numColonneRetour=numColonneRetour_; else this.numColonneRetour=-1;
  if(retour_) this.retour=retour_; else this.retour="";
  if(infoScript_) this.infoScript=infoScript_; else this.infoScript="";
  if(infoDoc_) this.infoDoc=infoDoc_; else this.infoDoc="";
  if(couleurInfoScript_) this.couleurInfoScript=couleurInfoScript_; else this.couleurInfoScript=vert2;
  if(couleurInfoDoc_) this.couleurInfoDoc=couleurInfoDoc_; else this.couleurInfoDoc=vert2;
  
  this.log = function() {
    var text="\nRETOUR:\n";
    for( var i in this) {
      if( this[i].toString().indexOf('function')<0) {
        text+='['+i+'] =>'+this[i]+'\n';
      }
    }
    Logger.log(text);
  };
  this.affiche = function() {
    //Mise à jour des infos sur feuille Formulaires Notés
    if(this.numLigneScript>0) {
      var cellInfoScript=rangeInfoScript.getCell(this.numLigneScript,1);
      cellInfoScript.setValue(maintenant()+this.infoScript+"\n"+cellInfoScript.getValue()).setBackground(this.couleurInfoScript);
      var cellInfoDoc=rangeInfosDoc.getCell(this.numLigneScript,1);
      cellInfoDoc.setValue(this.infoDoc+"\n"+cellInfoDoc.getValue()).setBackground(this.couleurInfoDoc);
      if(this.numColonneRetour>=0) {
        ss.getSheetByName(NOM_FEUILLE_LISTE_FORMULAIRES).getRange(this.numLigneScript+rangeScript.getRow()-1,this.numColonneRetour).setValue(this.retour);
        Logger.log(this.numLigneScript+rangeScript.getRow()-1+"/"+this.numColonneRetour+"/"+this.retour);
      }
      rangeScript.getCell(this.numLigneScript,1).setValue("").setBackground(orange1);
    }
    var vals=[[maintenant(),this.retour,this.infoScript,this.infoDoc]];
    Logger.log(vals);
    FEUILLE_LOG.getRange(FEUILLE_LOG.getLastRow()+1,1,1,4).setValues(vals).setBackground(this.couleurInfoScript);
  }
}

function getURLFeuille(formulaire) {
  //renvoie l
  try {
    var ID= formulaire.getDestinationId();
    if(ID) {
      return DriveApp.getFileById(ID).getUrl();
    } else {
      return "pas de feuille";
    }
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}


function getIDDossier(fileFormulaire) {
  try {
    var tousLesDossiers=fileFormulaire.getParents();
    var dossier = tousLesDossiers.next();
    return dossier.getId();
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function getIDFeuille(formulaire) {
  //renvoie l
  try {
    return formulaire.getDestinationId();
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function getIDFormulaire(formulaire) {
  //renvoie l'ID
  try {
    return formulaire.getId();
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function getDossier(URLForm,cellURLDossier,cellNomsDossiers) {
  try {
    var formulaire=FormApp.openByUrl(URLForm);
    var fileFormulaire=DriveApp.getFileById(formulaire.getId());
    var tousLesDossiers=fileFormulaire.getParents();
    // On prend le premier dossier.
    var noms="",dossier;
    while(tousLesDossiers.hasNext()) {
      dossier = tousLesDossiers.next();
      noms+=dossier.getName()+"("+dossier.getOwner()+")"+RC;
    }
    var tousLesDossiers=fileFormulaire.getParents();
    dossier = tousLesDossiers.next();
    var URLDossier=dossier.getUrl();
    cellURLDossier.setValue(URLDossier).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false);
    cellNomsDossiers.setValue(noms).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
  }
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function cleanNamedRanges(URLSheet) {
  //ne garde QUE QuestionReponses, Correction et feuille 1
  try {
    var classeurReponse=SpreadsheetApp.openByUrl(URLSheet);
    var ranges=classeurReponse.getNamedRanges();
    for(i in ranges) {
      var r=ranges[i].getRange().getA1Notation();
      var n=ranges[i].getName();
      Logger.log(r+":"+n);
    }
    return "ok";
  } 
  catch(e) {
    return callErreur(e);//erreur type Info
  }
}

function removeDeadReferences(URLSheet)
{
  var activeSS=SpreadsheetApp.openByUrl(URLSheet);
  
  var sheets = activeSS.getSheets();
  
  var sheetNamedRanges, loopRangeA1Notation;
  
  var x, i;
  // minimum sheet count is 1, no need to check for empty array, but why not
  if (sheets.length)
  {
    for (x in sheets)
    {
      sheetNamedRanges = sheets[x].getNamedRanges();
      var nomFeuille=sheets[x].getName();
      // check for empty array
      if (sheetNamedRanges.length)
      {
        for (i = 0; i < sheetNamedRanges.length; i++)
        { // get A1 notation of referenced cells for testing purposes
          loopRangeA1Notation = sheetNamedRanges[i].getRange().getA1Notation();
          // check for length to prevent throwing errors during tests
          if (loopRangeA1Notation.length)
          { // check for bad reference
            // note: not sure why the trailing "!" mark is currently omitted
            // ....: so there are added tests to ensure future compatibility
            var r=loopRangeA1Notation;
            var n=sheetNamedRanges[i].getName();
            Logger.log(n+']'+nomFeuille+ ']'+r);
            if (
              loopRangeA1Notation.slice(0,1) === "#"
              || loopRangeA1Notation.slice(-1) === "!"
            || loopRangeA1Notation.indexOf("REF") > -1
            )
            {
              //sheetNamedRanges[i].remove();
            }
          }
        }
      }
    }
  }
}


