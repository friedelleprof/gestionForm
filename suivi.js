//Creation DES suivis

function test() {
  genereSuivi("SIO1");
  //genereSuivi("SIO2");
  //metAJourColonne(ss.getSheetByName("SUIVI SIO1"),10)
}

function genereSuivi(classe) {
  openDatas();
  var nomFeuille = "SUIVI " + classe;
  //Crée la feuille de suivi ou l'actualise si elle existe
  if (ss.getSheetByName(nomFeuille) == null) {
    ss.insertSheet(nomFeuille, 1);
  }
  var feuilleSuivi = ss.getSheetByName(nomFeuille).clear();
  feuilleSuivi.clearConditionalFormatRules();
  //Supression des lignes en trop
  if (feuilleSuivi.getMaxRows() > 100) {
    feuilleSuivi.deleteRows(100, feuilleSuivi.getMaxRows() - 100);
  }

  var numLastCol = feuilleSuivi.getMaxColumns();
  if (numLastCol < 5 + dataURLSheet.length) {
    //Le nbre de colonnes est 5+nombre formulaires
    feuilleSuivi.insertColumns(numLastCol, 5 + dataURLSheet.length - numLastCol)
  }
  feuilleSuivi.setColumnWidths(1, feuilleSuivi.getMaxColumns(), 80);
  feuilleSuivi.setRowHeights(1, feuilleSuivi.getMaxRows(), 15);

  //Mise en place des éléments
  feuilleSuivi.getRange("A2").setValue("Classe").setBackground(orange1);
  feuilleSuivi.getRange("B2").setValue(classe).setBackground(orange1);
  ss.setNamedRange("nomClasse" + classe, feuilleSuivi.getRange("B2"));

  feuilleSuivi.getRange(1, 1, 1, feuilleSuivi.getMaxColumns()).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(false).setFontSize(7);
  feuilleSuivi.getRange(2, 1, 1, feuilleSuivi.getMaxColumns()).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setFontSize(14).setBackground(orange1);
  feuilleSuivi.getRange(3, 1, 1, feuilleSuivi.getMaxColumns()).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false).setFontSize(12).setNumberFormat("00").setBackground(orange1);;
  feuilleSuivi.getRange(4, 1, 1, feuilleSuivi.getMaxColumns()).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false).setFontSize(12).setNumberFormat("00").setBackground(orange2);;
  feuilleSuivi.getRange(5, 1, 1, feuilleSuivi.getMaxColumns()).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false).setFontSize(12).setNumberFormat("00").setBackground(orange3);
  feuilleSuivi.getRange(6, 1, 1, feuilleSuivi.getMaxColumns()).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false).setBackground(blanc)
    .setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireCheckbox().build()).setFontSize(14);

  feuilleSuivi.getRange("A4:A6").setValue("NOM").setBackground(orange1).merge();
  feuilleSuivi.getRange("B4:B6").setValue("PRENOM").setBackground(orange1).merge();
  feuilleSuivi.getRange("C6").setValue("MAIL").setBackground(orange1);
  feuilleSuivi.setFrozenColumns(3);
  feuilleSuivi.hideColumns(3, 2);
  feuilleSuivi.setColumnWidths(1, 2, 100);
  setAlternance(feuilleSuivi.getRange(6, 1, 100, feuilleSuivi.getMaxColumns()), blanc, gris3, bleu5)

  /**Probleme:
   * Les noms sont dynamiques
   * Les notes statiques
   */

  openDataEleves(true);
  for (var i = 0; i < dataNomEleve.length; i++) {
    var nomEleve = dataNomEleve[i][0];
    var prenomEleve = dataPrenomEleve[i][0];
    var mailEleve = dataMails[i][0];
    var classeEleve = dataClasseEleve[i][0];
    if (classeEleve == classe) {
      feuilleSuivi.getRange(i + 7, 1, 1, 3).setValues([[nomEleve, prenomEleve, mailEleve]]);
    }
  }
  var nomRangeListeMail = "listeMailSuivi" + classe;
  ss.setNamedRange(nomRangeListeMail, feuilleSuivi.getRange("C7:C100"));

  //Mise en place des notes
  // dataURLSheet contient la liste des URLs

  for (var i = 0; i < dataURLSheet.length; i++) {
    var URLSheet = dataURLSheet[i][0];
    var titre = dataNomFormulaire[i][0];

    if (titre) {
      var resultats = suivi(URLSheet, dataMails);//Récupération des notes
      var retour = resultats[1];

      if (retour == "ok") {
        var notes = resultats[0];
        var totalFinal = resultats[2];
        var noteSur = resultats[3];
        var moyenne = resultats[4];
        //Notes
        var range = feuilleSuivi.getRange(7, 5 + i, notes.length, 1);
        range.setValues(notes);
        range.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false).setFontSize(11).setNumberFormat("00");
        setRulesNotation(feuilleSuivi, range);
        //Autres info
        feuilleSuivi.getRange(1, 5 + i, 6, 1).setValues([[URLSheet], [titre], [noteSur], [totalFinal], [moyenne], [null]]);

      } else {
        //Pb sur les résultats
      }
    } else {
      //Vide
      feuilleSuivi.setColumnWidth(5 + i, 5);
      feuilleSuivi.getRange(1, 5 + i, 100, 1).setBackground(noir);
    }
  }
  setRulesNotation(feuilleSuivi, feuilleSuivi.getRange(5, 1, 1, feuilleSuivi.getLastColumn()));
}

function suivi(URLSheet, dataMail) {
  //Récupère les notes dans le classeur
  //Génère un tableau result ordonné sur les mails de dataMail
  Logger.log("Appel suivi");
  Logger.log(URLSheet);
  var result = new Array();
  var dataNotes, retour, moyenne = 0, eff = 0;
  try {
    var feuille = SpreadsheetApp.openByUrl(URLSheet);
    var rangeNotes = feuille.getRange("A7:D"+sheetCorrection.getLastRow());
    var totalFinal = feuille.getRangeByName("totalPoints").getValue();
    var noteSur = feuille.getRangeByName("sur").getValue();
    dataNotes = transpose(rangeNotes.getValues());
    retour = "ok";
  } catch (e) {
    dataNotes = [[null, null, null, e]];
    retour = "vide";
  }

  for (var i = 0; i < dataMail.length; i++) {
    var mail = dataMail[i][0];
    if (mail != "" && mail != null) {
      var index = dataNotes[0].indexOf(mail);
      if (index >= 0) {
        var note = dataNotes[3][index];
        result.push([note / totalFinal * noteSur]);
        moyenne += note;
        eff++;
      } else {
        result.push(["x"]);
      }
    } else {
      result.push([null]);
    }
  }
  if (eff > 0) {
    moyenne = (moyenne / totalFinal * noteSur) / eff;
  }
  else {
    moyenne = 0;
  }
  return [result, retour, totalFinal, noteSur, moyenne];
}

function metAJourColonne(feuilleSuivi, numColonne) {
  Logger.log("Appel metAJourColonne " + numColonne);
  openDatas();
  var classe = feuilleSuivi.getName().replace("SUIVI ", "");
  var nomRangeListeMail = "listeMailSuivi" + classe;
  var dataMails = ss.getRangeByName(nomRangeListeMail).getValues();
  //Mise en place des notes
  // dataURLSheet contient la liste des URLs

  var i = numColonne - 5;
  var URLSheet = dataURLSheet[i][0];
  var titre = dataNomFormulaire[i][0];
  var retour;

  if (titre) {
    Logger.log("titre:" + titre);

    var resultats = suivi(URLSheet, dataMails);
    var retour = resultats[1];
    Logger.log("retour:" + retour);

    if (retour == "ok") {

      Logger.log(resultats);
      var notes = resultats[0];
      var totalFinal = resultats[2];
      var noteSur = resultats[3];
      var moyenne = resultats[4];
      //Notes
      var range = feuilleSuivi.getRange(7, numColonne, notes.length, 1);
      range.setValues(notes);
      //range.setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(false).setFontSize(11).setNumberFormat("00");
      //setRulesNotation(feuilleSuivi,range);
      //Autres info
      feuilleSuivi.getRange(1, numColonne, 6, 1).setValues([[URLSheet], [titre], [noteSur], [totalFinal], [moyenne], [false]]);
      //function Info(numLigneScript_,numColonneRetour_,retour_,infoScript_,infoDoc_,couleurInfoScript_,couleurInfoDoc_) {
      //    return new Info(numLigneScript,-1,null,"Erreur generation Mails",null,null,null);
      Logger.log("OK Mise à jour");
      retour = new Info(null, -1, "mise à jour " + feuilleSuivi.getName(), titre + " OK", "Colonne:" + numColonne, vert2, vert2);
    } else {
      //Pb sur les résultats
      Logger.log("Problème:résultats");
      retour = new Info(null, -1, "mise à jour :" + feuilleSuivi.getName(), "problème rencontré", "Colonne:" + numColonne, rouge2, rouge2);
    }
  } else {
    //Vide
    Logger.log("Vide");
    feuilleSuivi.setColumnWidth(numColonne, 5);
    feuilleSuivi.getRange(1, numColonne, 100, 1).setBackground(noir);
    retour = new Info(null, -1, "mise à jour :" + feuilleSuivi.getName(), "colonne vide", "Colonne:" + numColonne, orange1, orange1);
  }
  return retour;
}


function setRulesNotation(feuille, ranges) {
  var rules = feuille.getConditionalFormatRules();
  Logger.log("Nb de règles:" + rules.length);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([ranges])
    .whenNumberLessThan(6)
    .setBackground(rouge1)
    //.setGradientMinpoint('#FF0000')
    //.setGradientMidpointWithValue('#FF9900', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    //.setGradientMaxpoint('#00FF00')
    .build();
  rules.push(rule);
  feuille.setConditionalFormatRules(rules);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([ranges])
    .whenNumberGreaterThan(17)
    .setBackground(vert1)
    .build();
  rules.push(rule);
  feuille.setConditionalFormatRules(rules);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([ranges])
    .whenTextContains("x")
    .setBackground(rouge2)
    .build();
  rules.push(rule);
  feuille.setConditionalFormatRules(rules);
}

function setAlternance(range, couleur1, couleur2, couleur3) {
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  var banding = range.getBandings()[0];
  banding.setHeaderRowColor(couleur1)
    .setFirstRowColor(couleur2)
    .setSecondRowColor(couleur3)
    .setFooterRowColor(couleur1);
}

function testRules() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    Logger.log(sheet.getName());
    Logger.log(sheet.getConditionalFormatRules().length);
  }
}

