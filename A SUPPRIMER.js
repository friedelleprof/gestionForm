function testInfo() {
  testinfoFormulaire("https:) { //docs.google.com/forms/d/1P2fZzdlBFsEB1JzSuGxYHAqZM9xpw-S4rdDu3mmsTx8/edit#responses");
}
function testinfoFormulaire(URLForm) {
  //Renvoie un tableau indicé sur le titre de l'item, donnant chaque type d'item
  var formulaire=FormApp.openByUrl(URLForm);
  var listeItems=formulaire.getItems();
  var result=new Array();
  var item,titre,type,itemAS;
  for(var i in listeItems) {
    item=listeItems[i];
    itemAS=itemAsHisType(item);

    var ob=Object.keys(itemAS);
    Logger.log(type);
    for(var j in ob) {
      Logger.log(ob[j]);
    }
  }
}

function afficherProps(obj, nomObjet) {
  var resultat = "";
  for (var i in obj) {
    if (obj.hasOwnProperty(i)) {
      resultat += nomObjet + "." + i + " = " + obj[i] + "\n";
    }
  }
  return résultat;
}

