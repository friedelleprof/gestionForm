function Item(item_,numColonne_,type_,titre_,helpText_,reponsesPossibles_) {
  this.item=item_;
  this.titre=titre_;
  this.type=type_;
  this.helpText=helpText_;
  this.numColonne=numColonne_;
  this.reponsesPossibles=reponsesPossibles_;
  
  this.log = function() {
    var text="\nITEM:\n";
    for( var i in this) {
      if( this[i].toString().indexOf('function')<0) {
        text+='['+i+'] =>'+this[i]+'\n';
      }
    }
    Logger.log(text);
  };
}

function creeListeItems(formulaire,dataReponses) {
  //Renvoie un tableau d'objets ITEMS
  
  var listeItems=formulaire.getItems();
  var formResponses = formulaire.getResponses();
  var aDesReponses=formulaire.isQuiz();
  var result=new Array();
  //OBJET Item(numColonne_,type_,titre_,helpText_,reponsesPossibles_,points_,reponseExacte_,commentaireSiJuste_,commentaireSiFaux_,lien_) {
  
  var obItem,item,titre,type,helpText,numColonne,reponsesPossibles,lignes,itemGrid,nomCol;
  for(var i in listeItems) {
    item=listeItems[i];
    titre=item.getTitle();
    type=item.getType();
    helpText="",reponsesPossibles=[""];
    itemGrid=itemAsHisType(item);
    
    if(type== FormApp.ItemType.GRID || type==FormApp.ItemType.CHECKBOX_GRID) { //cas particulier, la réponse est décomposée par ligne
      type= (type== FormApp.ItemType.GRID) ?  FormApp.ItemType.MULTIPLE_CHOICE : FormApp.ItemType.CHECKBOX;
      lignes=itemGrid.getRows();
      reponsesPossibles=itemGrid.getColumns().map(function(x) {if(typeof x ==="object") return x.getValue(); else return x;});
      helpText=item.getHelpText();
      if(!helpText) helpText=""; else helpText=" ("+helpText+")";
      //On regarde les lignes et on décompose
      for(var j in lignes) {
        nomCol=titre+ " ["+lignes[j]+"]";
        numColonne=dataReponses.indexOf(nomCol);
        if(numColonne>=0) {
          obItem=new Item(itemGrid,numColonne+1,type,nomCol,helpText,reponsesPossibles);
          result.push(obItem);
        }
      }
    } else if( type==FormApp.ItemType.MULTIPLE_CHOICE || type==FormApp.ItemType.CHECKBOX || type==FormApp.ItemType.LIST) {
      reponsesPossibles=itemGrid.getChoices().map(function(x) {if(typeof x ==="object") return x.getValue(); else return x;});
      numColonne=dataReponses.indexOf(titre);
      helpText=item.getHelpText();
      if(!helpText) helpText=""; else helpText=" ("+helpText+")";
      if(numColonne>=0) {
        obItem=new Item(itemGrid,numColonne+1,type,titre,helpText,reponsesPossibles);
        result.push(obItem);
      }
    } 
    else {//Sinon cas normal
      helpText=item.getHelpText();
      if(!helpText) helpText=""; else helpText=" ("+helpText+")";
      numColonne=dataReponses.indexOf(titre);
      if(numColonne>=0) {
        obItem=new Item(itemGrid,numColonne+1,type,titre,helpText,reponsesPossibles);
        result.push(obItem);
      }
    }
  }
  return result;
}




REPONSE_SI_JUSTE="EXACT";
REPONSE_SI_FAUX="FAUX";
//OBJET Item(numColonne_,type_,titre_,helpText_,reponsesPossibles_) {

function ligneCorrection(item_,code_,feuille_) {
  this.feuille=feuille_;
  this.code=code_;
  this.item=item_;
  var points;
  this.reponseSiJuste=REPONSE_SI_JUSTE;
  this.reponseSiFaux=REPONSE_SI_FAUX;
  this.correctionsB="";
  if(item_.type==FormApp.ItemType.PARAGRAPH_TEXT) {
    this.item.reponsesPossibles=["Entrez ici les critères en passant à la ligne"];
    this.reponseSiJuste="Entrez ici la correction si note A en passant à la ligne";
    this.correctionsB="Entrez ici la correction si note B en passant à la ligne";
    this.reponseSiFaux="Entrez ici la correction si note C en passant à la ligne";
  }
  if(item_.type==FormApp.ItemType.CHECKBOX) {
    this.points="2.0;1.0";
  } else {
    this.points=1.0;
  }
  this.draw = function(ligne) {
    var values;
    if(ligne>0) {
      var values=[[this.code,this.item.numColonne,this.item.type,this.item.titre+this.item.helpText,this.item.reponsesPossibles.join(RC),this.points,"=max(split(R[0]C[-1];\";\"))",this.reponseSiJuste,this.correctionsB,this.reponseSiFaux]]
      var range=this.feuille.getRange(ligne,1,values.length,values[0].length);
      this.feuille.getParent().setNamedRange("ROW_QR_"+ligne,range);
      if(ligne%2==0) {
        range.setBackground(vert2);
      }
      else {
        range.setBackground(vert3);
      }
      range.setValues(values);        
      var p=range.prototype;
      //setStyle(range,this.style);
    }
  };
  
  this.read = function(ligne,feuille) {
    //Lis les données sur une ligne
    var values=feuille.getRange(ligne,1,9,1)[0];
    this.feuille=feuille;
    this.code=values[0];
    this.points=values[5];
    this.reponseSiJuste=values[6];
    this.correctionsB=values[7];
    this.reponseSiFaux=values[8];
    this.item=new Item(values[1],values[2],values[3],"",values[4]); //helpText à vide
  };
  
}

function proprietes(objet) {
  var text="\n\n";
  for( var i in objet) {
    text+='['+i+'] =>'+objet[i];
  }
  return text;
}

