function nomA1(nomFeuille, range) {
  return "'" + nomFeuille + "'!" + range.getA1Notation();
}

function fullA1Name(range) {
  return nomA1(range.getSheet().getName(), range);
}

var suffixeSauvegarde = "SAV_";

function creeNouvelleFeuille(classeur, nomFeuille, sauvegarde) {
  //Création feuille QuestionsReponses
  //if(sauvegarde==undefined) sauvegarde=true;
  var sheet = classeur.getSheetByName(nomFeuille);
  if (sheet && sauvegarde) {
    //On duplique
    //Si déjà sauvegarde, on supprime
    var sheetS = classeur.getSheetByName(suffixeSauvegarde + nomFeuille);
    if (sheetS) classeur.deleteSheet(sheetS);
    sheet.copyTo(classeur).setName(suffixeSauvegarde + nomFeuille);
    classeur.deleteSheet(sheet);
  } else if (sheet && !sauvegarde) {
    classeur.deleteSheet(sheet);
  }
  //On crée une nouvelle
  sheet = classeur.insertSheet().setName(nomFeuille);
  return sheet;
}

function infoFormulaire(formulaire) {
  //Renvoie un tableau indicé sur le titre de l'item, donnant chaque type d'item
  var listeItems = formulaire.getItems();
  var result = new Array();
  var item, titre, type;
  for (var i in listeItems) {
    item = listeItems[i];
    titre = item.getTitle();
    type = item.getType();
    if (type == FormApp.ItemType.GRID) {
      var itemGrid = item.asGridItem();
      var colonnes = itemGrid.getRows();
      for (var j in colonnes) {
        var nomCol = titre + " [" + colonnes[j] + "]";
        result[nomCol] = FormApp.ItemType.MULTIPLE_CHOICE;
      }
    }
    else if (type == FormApp.ItemType.CHECKBOX_GRID) {
      var itemGrid = item.asCheckboxGridItem();
      var colonnes = itemGrid.getRows();
      for (var j in colonnes) {
        var nomCol = titre + " [" + colonnes[j] + "]";
        result[nomCol] = FormApp.ItemType.CHECKBOX;
      }
    }
    else {
      result[titre] = type;
    }
  }
  return result;
}
function getHelpText(formulaire) {
  //Renvoie un tableau indicé sur le titre de l'item, donnant le texte d'aide
  var listeItems = formulaire.getItems();
  var result = new Array();
  var item, titre, help, type;
  for (var i in listeItems) {
    item = listeItems[i];
    titre = item.getTitle();
    type = item.getType();
    if (type == FormApp.ItemType.GRID) {
      var itemGrid = item.asGridItem();
      var colonnes = itemGrid.getRows();
      for (var j in colonnes) {
        var nomCol = titre + " [" + colonnes[j] + "]";
        result[nomCol] = "";
      }
    }
    else if (type == FormApp.ItemType.CHECKBOX_GRID) {
      var itemGrid = item.asCheckboxGridItem()
      var colonnes = itemGrid.getRows();
      for (var j in colonnes) {
        var nomCol = titre + " [" + colonnes[j] + "]";
        result[nomCol] = "";
      }
    }
    else {
      help = item.getHelpText();
      if (help == undefined || help == null || help == "") help = ""; else help = " (" + help + ")";
      result[titre] = help;
    }
  }
  return result;
}

function rangeColonneDuTexte(texte, ligne, feuille) {
  //Anciennement Colonne

  //Renvoie un RANGE colonne ou est situé le TEXTE sur la LIGNE

  try {
    //Logger.log(texte+" "+ligne+" "+feuille);
    if (feuille == undefined) {
      feuille = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    }
    if (ligne == undefined) {
      ligne = 1;
    }

    //Renvoie le range contenant la valeur en ligne i
    var rang = TrouveNumColonneRange(texte, feuille.getRange(ligne + ":" + ligne));
    if (rang) {
      var range = feuille.getRange(1, rang, feuille.getLastRow());
      return range;
    }
    else {
      return null;
    }
  }
  catch (e) {
    return null;
  }
}

function TrouveNumLigneRange(value, range) {
  //Anciennement TrouveLigne
  //Renvoie le numéro de la première ligne du range qui contient la valeur
  var data = range.getValues();
  return TrouveLigneDatas(value, data);//Décalage du range
}

function TrouveNumLigneDatas(value, data) {
  //Renvoie le numéro de la première ligne des datas qui contient la valeur
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == value) {
        return i + 1;
      }
    }
  }
  return null;
}

function rangeColonneNommee(feuille, nom, debut) {
  //Renvoie le RANGE sous forme de colonne, débutant à debut jusqu'à lastRow
  var data = feuille.getRange(2, 1, 1, feuille.getNumColumns()).getValues();//Sur la 2ème ligne 
  num = TrouveNumColonneDatas(nom, data);

  if (num) {
    return feuille.getRange(debut, num,  feuille.getLastRow() - debut + 1,1);
  } else {
    return null;
  }
}
function numRangeColonneNommee(feuille, nom) {
  //Renvoie le RANGE sous forme de colonne, débutant à debut jusqu'à lastRow
  var data = feuille.getRange(1, 1, 1, feuille.getNumColumns()).getValues();
  num = TrouveNumColonneDatas(nom, data);
  if (num) {
    return num;
  } else {
    return null;
  }
}

function TrouveNumColonneRange(value, range) {
  //Anciennement TrouveNumColonne 
  //Renvoie le numero de la colonne contenant la valeur dans le range
  var data = range.getValues();
  return TrouveNumColonneDatas(value, data);//Décalage du range
}

function TrouveNumColonneDatas(value, data) {
  //Renvoie le numero de la colonne contenant la valeur dans les datas
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == value) {
        return j + 1;
      }
    }
  }
  return null;
}

function creePlage(nomPlage, textePlage, feuille) {
  //cherche le texte donné sur ligne 1 de la feuille
  //nomme et renvoi le range trouvé

  var rang = TrouveColonne(textePlage, feuille.getRange("1:1"));
  if (rang != null) {
    var range = feuille.getRange(1, rang, feuille.getMaxRows());
    ss.setNamedRange(nomPlage, range);
    return rang;
  }
  return null;
}

function camelize(str) {
  return str.replace(/(?:^\w|[A-Z]|\b\w)/g, function (letter, index) {
    return index == 0 ? letter.toLowerCase() : letter.toUpperCase();
  }).replace(/\s+|\W/g, '');
}
function transpose(matrix) {
  const rows = matrix.length, cols = matrix[0].length;
  const grid = [];
  for (var j = 0; j < cols; j++) {
    grid[j] = Array(rows);
  }
  for (var i = 0; i < rows; i++) {
    for (var j = 0; j < cols; j++) {
      grid[j][i] = matrix[i][j];
    }
  }
  return grid;
}

function maintenant() {
  return Utilities.formatDate(new Date(), "GMT+1", "dd/MM à HH:mm:ss | ");
}

function metAJourRangeColonne(classeur, nomRange, numLigne) {
  //renvoie le range après l'avoir mis à jour
  var range = classeur.getRangeByName(nomRange);

  if (range != null && (range.getRow() != 7 || range.getNumRows() != numLigne)) {
    var newRange = range.getSheet().getRange(7, range.getColumn(), numLigne, range.getNumColumns());
    classeur.setNamedRange(nomRange, newRange);
    return classeur.getRangeByName(nomRange);
  } else {
    return range;
  }
}

function openRange(classeur, nomRange, feuille, A1Notation, metAJour) {
  var metAJour = (typeof metAJour !== 'undefined') ? metAJour : false;

  //Renvoie le range
  //S'il n'existe pas, le crée
  var range = classeur.getRangeByName(nomRange);
  if (range == null || (metAJour && range.getA1Notation() != A1Notation)) {
    range = feuille.getRange(A1Notation);
    classeur.setNamedRange(nomRange, range);
  }
  return range;
}

function openSheet(nomFeuille, classeur, creer, creation) {
  var creer = (typeof creer !== 'undefined') ? creer : false;
  var creation = (typeof creation !== 'undefined') ? creation : function (sheet) { return sheet; };

  //Renvoie la feuille si elle existe
  //Sinon la crée (vierge)
  var sheet = classeur.getSheetByName(nomFeuille);
  if (sheet == null) {
    sheet = classeur.insertSheet(nomFeuille, classeur.getNumSheets());
    if (creer) {
      sheet = creation(sheet);
    }
  }
  return sheet;
}


function checkQuota() {
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
}


function callErreur(e) {
  var inf = ERROR + e.name + RC + "ligne:" + e.lineNumber + RC + "->" + e.stack;
  var retourInfo;
  if (numLigneScript && rouge1) {
    retourInfo = new Info(numLigneScript, -1, null, inf, null, rouge1, null);
    retourInfo.affiche();
  }
  Logger.log(ERROR + inf);
  return retourInfo;
}

function itemAsHisType(item) {
  var itemAS;
  type = item.getType();
  if (type == FormApp.ItemType.CHECKBOX) { //Enum	A question item that allows the respondent to select one or more checkboxes, as well as an optional "other" field.
    itemAS = item.asCheckboxItem();
  }
  else if (type == FormApp.ItemType.CHECKBOX_GRID) {//	Enum	A question item, presented as a grid of columns and rows, that allows the respondent to select multiple choices per row from a sequence of checkboxes.
    itemAS = item.asCheckboxGridItem();
  }
  else if (type == FormApp.ItemType.DATE) {//	Enum	A question item that allows the respondent to indicate a date.
    itemAS = item.asDateItem();
  }
  else if (type == FormApp.ItemType.DATETIME) { //	Enum	A question item that allows the respondent to indicate a date and time.
    itemAS = item.asDateTimeItem();
  }
  else if (type == FormApp.ItemType.DURATION) { //	Enum	A question item that allows the respondent to indicate a length of time.
    itemAS = item.asDurationItem();
  }
  else if (type == FormApp.ItemType.GRID) { //	Enum	A question item, presented as a grid of columns and rows, that allows the respondent to select one choice per row from a sequence of radio buttons.
    itemAS = item.asGridItem();
  }
  else if (type == FormApp.ItemType.IMAGE) { //	Enum	A layout item that displays an image.
    itemAS = item.asImageItem();
  }
  else if (type == FormApp.ItemType.LIST) { //	Enum	A question item that allows the respondent to select one choice from a drop-down list.
    itemAS = item.asListItem();
  }
  else if (type == FormApp.ItemType.MULTIPLE_CHOICE) { //	Enum	A question item that allows the respondent to select one choice from a list of radio buttons or an optional "other" field.
    itemAS = item.asMultipleChoiceItem();
  }
  else if (type == FormApp.ItemType.PAGE_BREAK) { //	Enum	A layout item that marks the start of a page.
    itemAS = item.asPageBreakItem();
  }
  else if (type == FormApp.ItemType.PARAGRAPH_TEXT) { //	Enum	A question item that allows the respondent to enter a block of text.
    itemAS = item.asParagraphTextItem();
  }
  else if (type == FormApp.ItemType.SCALE) { //	Enum	A question item that allows the respondent to choose one option from a numbered sequence of radio buttons.
    itemAS = item.asScaleItem();
  }
  else if (type == FormApp.ItemType.SECTION_HEADER) { //	Enum	A layout item that visually indicates the start of a section.
    itemAS = item.asSectionHeaderItem();
  }
  else if (type == FormApp.ItemType.TEXT) { //	Enum	A question item that allows the respondent to enter a single line of text.
    itemAS = item.asTextItem();
  }
  else if (type == FormApp.ItemType.TIME) { //	Enum	A question item that allows the respondent to indicate a time of day.
    itemAS = item.asTimeItem();
  }
  else if (type == FormApp.ItemType.VIDEO) { //	Enum	A layout item that displays a YouTube video.
    itemAS = item.asVideoItem();
  }
  return itemAS;
}

