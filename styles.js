
/*

setHorizontalAlignment('center')
setVerticalAlignment('middle')
setHorizontalAlignment('right')
setHorizontalAlignment('left')
setFontSize(6)
*/

//COULEURS

bleu1=   "#00FFFF";
bleu2	="#0000FF";
bleu3	="#00008B";
bleu4	="#ADD8E6";
bleu5	="#ADD8F6";
brun1	="#A52A2A";
brun2	="#A0522D";
gris1	="#DCDCDC";
gris2	="#778899";
gris3	="#F8F8FF";
jaune1	="#FFFF00";
jaune2	="#FFD700";
jaune3	="#F0E68C";
noir	="#000000";
orange1	="#D2691E";
orange2	="#FF7F50";
orange3	="#FF8C00";
rose1	="#FF1493";
rose2	="#FF69B4";
rouge1	="#DC143C";
rouge2	="#DC143C";
vert1	="#008000";
vert2	="#ADFF2F";
vert3	="#7FFF00";
violet	="#8A2BE2";
violet2	="#4B0082";
blanc="#FFFFFF";

styleCorrection = {};
stylePoints = {};
styleCode = {};
styleTitreItem = {};
function setStyles () {
  styleCorrection[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =  DocumentApp.HorizontalAlignment.LEFT;
  styleCorrection[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  styleCorrection[DocumentApp.Attribute.FONT_FAMILY] = 'Consolas';
  styleCorrection[DocumentApp.Attribute.FONT_SIZE] = 14;
  styleCorrection[DocumentApp.Attribute.BOLD] = true;
  styleCorrection[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0000FF';
  
  stylePoints[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =  DocumentApp.HorizontalAlignment.LEFT;
  stylePoints[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  stylePoints[DocumentApp.Attribute.FONT_FAMILY] = 'Droid Sans';
  stylePoints[DocumentApp.Attribute.FONT_SIZE] = 16;
  stylePoints[DocumentApp.Attribute.BOLD] = true;
  stylePoints[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FF0000';
  
  styleCode[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =  DocumentApp.HorizontalAlignment.LEFT;
  styleCode[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  styleCode[DocumentApp.Attribute.FONT_SIZE] = 10;
  styleCode[DocumentApp.Attribute.BOLD] = false;
  styleCode[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  
  styleTitreItem[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =  DocumentApp.HorizontalAlignment.LEFT;
  styleTitreItem[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  styleTitreItem[DocumentApp.Attribute.FONT_FAMILY] = 'Confortaa';
  styleTitreItem[DocumentApp.Attribute.FONT_SIZE] = 14;
  styleTitreItem[DocumentApp.Attribute.BOLD] = true;
  styleTitreItem[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
}

function defStyleGSI(style,gras,souligne,italique) {
  style[DocumentApp.Attribute.BOLD]=gras;
  style[DocumentApp.Attribute.ITALIC]=italique;
  style[DocumentApp.Attribute.UNDERLINE]=souligne;
}
function defStylePOSITION(style,vertical,horizontal) {
  var VA,HA;
  if(vertical=="haut") VA=DocumentApp.VerticalAlignment.TOP;
  else if(vertical=="bas") VA=DocumentApp.VerticalAlignment.BOTTOM;
  else VA=DocumentApp.VerticalAlignment.CENTER;
  if(horizontal=="justifié") HA=DocumentApp.VerticalAlignment.JUSTIFY;
  else if(horizontal=="droite") HA=DocumentApp.VerticalAlignment.RIGHT;
  else if(horizontal=="centre") HA=DocumentApp.VerticalAlignment.CENTER;
  else VA=DocumentApp.VerticalAlignment.LEFT;
  
  style[DocumentApp.Attribute.VERTICAL_ALIGNMENT]=VA;
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]=HA;
}
function defStylePOLICE(style,couleur,taille,famille) {
  var VA,HA;
  style[DocumentApp.Attribute.FOREGROUND_COLOR]=couleur;
  style[DocumentApp.Attribute.FONT_SIZE]=taille;
  style[DocumentApp.Attribute.FONT_FAMILY]=famille;
}

/*
BACKGROUND_COLOR	Enum	The background color of an element (Paragraph, Table, etc) or document.
BOLD	Enum	The font weight setting, for rich text.
FONT_FAMILY	Enum	The font family setting, for rich text.
FONT_SIZE	Enum	The font size setting in points, for rich text.
HORIZONTAL_ALIGNMENT	Enum	The horizontal alignment, for paragraph elements (for example, DocumentApp.HorizontalAlignment.CENTER).
ITALIC	Enum	The font style setting, for rich text.
UNDERLINE	Enum	The underline setting, for rich text.
VERTICAL_ALIGNMENT	Enum	The vertical alignment setting, for table cell elements.

*/