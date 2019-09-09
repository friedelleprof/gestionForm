function evalueReponse(reponsesEleve,correction,points,brepo,frepo) {
  Utilities.sleep(getRandomInt(2000));

  if(points.toString().indexOf(";")>0) {
    //Cas CB
    return texteCheckBox4(reponsesEleve,correction,points);
  } else if(correction.toString().indexOf(";")>0) {
    //Cas OU
    return evalOU(reponsesEleve,correction,points,brepo,frepo); 
  } else {
    return evalNormale(reponsesEleve,correction,points,brepo,frepo);
  }
}

function texteCheckBox4(reponsesEleve,correction,points) {
  
  if(reponsesEleve==null || reponsesEleve=="")  {
    return [[0,"aucune réponse"]];
  }
  //Utilities.sleep(1001);
  
  var tableauRep,tableauCorrection;
  if(reponsesEleve.toString().indexOf(",")>=0) {
    tableauRep=reponsesEleve.split(",").map(function(x) {return x.toUpperCase().trim()});
  } else {
    tableauRep=new Array();
    tableauRep.push(reponsesEleve.toString().toUpperCase().trim());
  }  
  if(correction.toString().indexOf(";")>=0) {
    tableauCorrection=correction.split(";").map(function(x) {return x.toUpperCase().trim()});
  } else {
    tableauCorrection=new Array();
    tableauCorrection.push(correction.toUpperCase().trim());
  }
  
  //Pour chaque réponse
  function estDansTableauCorrection(x) {
    return tableauCorrection.indexOf(x)>-1;
  }
  function nestPasDansTableauCorrection(x) {
    return tableauCorrection.indexOf(x)==-1;
  }
  function nestPasDansTableauRep(x) {
    return tableauRep.indexOf(x)==-1;
  }  
  
  var t1=tableauRep.filter(estDansTableauCorrection),
      t2=tableauRep.filter(nestPasDansTableauCorrection),
      t3=tableauCorrection.filter(nestPasDansTableauRep);
  
  var cptVrai=t1.length,cptFaux=t2.length;
  
  var texte= (cptVrai>0 ? "Bonnes réponses :"+cptVrai+"\n"+t1.join('\n') : "Aucune bonne réponse\n")
  + (cptFaux>0 ? "\nRéponses incorrectes :"+cptFaux+"\n"+t2.join('\n') : "\nPas de réponses incorrectes\n")
  + (t3.length>0 ? "\nRéponses attendues :\n"+t3.join('\n') : "\nToutes les bonnes réponses ont été cochées");
  
  var nbRep=tableauCorrection.length;
  
  var tabPoints=points.replace(",",".").split(";");
  var pts=Math.min(Math.max(tabPoints[0]*cptVrai/nbRep-tabPoints[1]*cptFaux/nbRep,0),tabPoints[0])
  return [[pts,texte]];
}

function sumEtReponse(somme, reponse) {
  //Juste pour afficher une réponse d'une FEVAL
  return [[somme,reponse]];
}


function evalFormule(reponseEleve,prec,memo,formule,points,brepo,frepo) {
  //Traduit la formule en rajoutant des IF, ELSE... et l'évalue
  if(reponseEleve=="") return [[0,"Aucune réponse"]];
  var formul=formule;
  formul="("+formul.replace("TEST:","")+") ? "+points+" : 0";
  formule.replace(/#UTILREP#/g,memo);
  formul=formul.replace(/#PREC#/g,prec);
  formul=formul.replace(/#REP#/g,reponseEleve);
  
  formul=formul.replace(/#UTILREP/g,memo);
  formul=formul.replace(/#PREC/g,prec);
  formul=formul.replace(/#REP/g,reponseEleve);
  
  try {
    var totP=eval(formul);
    if(totP>0) return [[totP,"Tu as répondu "+reponseEleve+RC+brepo]];
    else return [[totP,"Tu as répondu "+reponseEleve+RC+frepo]];
    return totP;
  } catch(e) {
    return [[0,e.name]];
  }
}

function evalNormale(reponsesEleve,correction,points,brepo,frepo) {
  //Evaluation basique sur une réponse
  if(reponsesEleve.toString()=="")  {
    return [[0,"aucune réponse"]];
  } else if(correction.toString().toUpperCase().trim()==reponsesEleve.toString().toUpperCase().trim() || Number(correction)==Number(reponsesEleve)) {
    return [[points,"Tu as répondu "+reponsesEleve+"\n"+brepo]];
  } else {
    return [[0,"Tu as répondu "+reponsesEleve+"\n"+frepo]];
  }
}

function evalOU(reponsesEleve,correction,points,brepo,frepo) {
  
  var tableauCorrection,tableauCorrectionNum;
  tableauCorrection=correction.split(";").map(function(x) {return x.toUpperCase().trim()});
  tableauCorrectionNum=correction.split(";").map(function(x) {return Number(x)});
  
  //OU
  
  if(reponsesEleve==null || reponsesEleve=="")  {
    return [[0,"aucune réponse"]];
  } else if(tableauCorrection.indexOf(reponsesEleve)>=0 || tableauCorrectionNum.indexOf(Number(reponsesEleve.toString().replace(/,/g,".")))>0) {
    return [[points,"Tu as répondu "+reponsesEleve+"\n"+brepo]];
  } else {
    return [[0,"Tu as répondu "+reponsesEleve+"\n"+frepo]];
  }
} 

function getRandomInt(max) {
  return Math.floor(Math.random() * Math.floor(max));
}
