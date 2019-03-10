// Automates Google sheets management of my Magic: the Gathering pauper cube
// Link to my pauper cube Google Sheet: https://docs.google.com/spreadsheets/d/1n7Y204NWFy0I1_D011xKDq5J_cnxDGuPLyCVLAWOQ8A

// This script will pull cards from the 'Change Log' sheet and write everything on the 'Card List' sheet based in the 'Column List' sheet columns and text

// see GitHub for issues: https://github.com/d-ellsworth/PauperCube

// loads menu on open
// menu items to update/re-write the card list and sort by sort column
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet Scripts').addItem('Update Card List','writeCardList').addItem('Sort Card List','sortMe').addToUi()
}

// function to retrive list of cards currently in the cube from the 'Change Log' sheet
// card count values must be in col 6 (H) and names in col 3 (D)
// returns array of unique, exact card names
function getCardList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sChange = ss.getSheetByName('Change Log');
  var changeData = sChange.getDataRange().getValues();
  // initialize card list and list index
  var j = 0;
  var cardList = new Array(changeData[0][3]); 
  // loop through all cards in change log
  for (var i = 1; i < changeData.length; i++) {
    if (changeData[i][6] == 1) {
      // if there is exactly 1 card in the cube add it to the list
      cardList[j] = changeData[i][3];
      j++;
    }  else if (changeData[i][6] == 0) {
      // do nothing if the card is not in the cube
    } else {
      // if the number of cards is not 0 or 1 throw an error
      throw new Error("Card count value not 0 or 1!");
    }
  }
  return cardList;
}

// Log input for debugging log output
// see log in 'view' -> 'logs'
function logMe(input) {
  Logger.log(input);
}

// sorts the card list
function sortMe() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var cList = ss.getSheetByName('Card List');
  var range = cList.getRange(2,1,cList.getMaxRows()-1,cList.getMaxColumns());
  // sort by sort and name
  range.sort([findColumn('Sort')+1,findColumn('Name')+1]);
}

// writes the cards and info, using ScryFall, to the 'Card List' sheet
function writeCardList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var cList = ss.getSheetByName('Card List');
  // this sheet lists all of the columns and what's in them
  var colList = ss.getSheetByName('Column List');
  var colData = colList.getDataRange().getValues();
  // get cards in cube from the Change Log sheet
  var cardList = getCardList();
  // The 'Column List' sheet defines what columns there are
  // The columns are hard coded, if the columns change need to update the sheet and the code here
  // write column headers
  cList.getRange(1, 1, 1, colData[0].length).setValues([colData[0]]);
  // get column order from Column List
  // these are hard coded and must match (case sensitive) what is in the sheet
  var sortCol = findColumn('Sort');
  var colorCol = findColumn('Color');
  var sectionCol = findColumn('Section');
  var nameCol = findColumn('Name');
  var classCol = findColumn('Class');
  var typeCol = findColumn('Type');
  var subTypeCol = findColumn('Subtype');
  var speedCol = findColumn('Speed');
  var manaCostCol = findColumn('Mana Cost');
  var powerCol = findColumn('P');
  var toughnessCol = findColumn('T');
  var cardTextCol = findColumn('Card Text');
  var evasionCol = findColumn('Evasion');
  var drawCol = findColumn('Draw/Filter');
  var tempoCol = findColumn('Tempo');
  var tricksCol = findColumn('Tricks');
  var removalCol = findColumn('Removal');
  var rampCol = findColumn('Ramp/Fixing');
  // sort criteria by Section, Color, Type, CMC
  var sectionSort = ['White','Blue','Black','Red','Green','Multi','Colorless','Land'];
  var colorSort = ['WU','WB','WR','WG','UB','UR','UG','BR','BG','RG'];
  var typeSort = ['Creature','Artifact Creature','Enchantment','Instant','Sorcery','Artifact','Land'];
  // write the cards to the Card List sheet row by row
  for (var i = 0; i < cardList.length; i++) {
    // get the card info JSON from ScryFall
    cardInfo = getCardInfo(cardList[i]);
    // initialize what we're going to write to the sheet
    var rowInfo = [];
    // card name
    rowInfo[nameCol] = cardInfo.name;
    logMe(cardInfo.name);
    // set card color properties (color, section)
    var colorInfo = getColor(cardInfo.color_identity);
    rowInfo[colorCol] = colorInfo[0];
    rowInfo[sectionCol] = colorInfo[1];
    // set card type properties (class, type, subtype, speed)
    // set type and subtype
    var cardType = [];
    // if the card has a subtype it will have a dash
    // don't have a good way to tell if there is a supertype
    if (cardInfo.type_line.indexOf('—') > -1) {
      // type (info before dash)
      cardType = cardInfo.type_line.substr(0,cardInfo.type_line.indexOf('—')-1);
      // subtype (info after dash)
      rowInfo[subTypeCol] = cardInfo.type_line.substr(cardInfo.type_line.indexOf('—')+2, cardInfo.type_line.length);
    } else {
      // type
      cardType = cardInfo.type_line;
      // no subtype
      rowInfo[subTypeCol] = '';
    }
    rowInfo[typeCol] = cardType;
    // set speed. If instant or has flash, instant speed, else sorcery speed
    if (cardType.indexOf('Instant') > -1) {
      rowInfo[speedCol] = 'Instant';
    } else if (/Flash(?!back)/i.test(cardInfo.oracle_text)) {
      // if 'Flash' is in the oracle text, instant. Exclude Flashback.
      rowInfo[speedCol] = 'Instant';
    } else {
      rowInfo[speedCol] = 'Sorcery';
    }
    // set class
    if (cardType.indexOf('Land') > -1) {
      // if it's a land also set the color and section to 'Land'
      rowInfo[colorCol] = 'Land';
      rowInfo[sectionCol] = 'Land';
      rowInfo[classCol] = 'Land';
      // Lands don't have power & toughness
      rowInfo[powerCol] = '';
      rowInfo[toughnessCol] = '';
    } else if (cardType.indexOf('Creature') > -1) {
      rowInfo[classCol] = 'Creature';
      // set power & toughness
      rowInfo[powerCol] = cardInfo.power;
      rowInfo[toughnessCol] = cardInfo.toughness;
    } else {
      rowInfo[classCol] = 'Spell';
      // Spells don't have power & toughness
      rowInfo[powerCol] = '';
      rowInfo[toughnessCol] = '';
    }
    // mana cost
    // don't have a good way to get 'real' cmc (delve, phyrexian mana, etc)
    rowInfo[manaCostCol] = cardInfo.cmc;
    // card text
    rowInfo[cardTextCol] = cardInfo.oracle_text;
    // Set card special properties (evasion, draw, tempo, tricks, removal, ramp)
    var propCol = [evasionCol,drawCol,tempoCol,tricksCol,removalCol,rampCol];
    for (var j = 0; j < propCol.length; j++) {
      rowInfo[propCol[j]] = findString(colData, rowInfo[cardTextCol], propCol[j]);
    }
    // set sort order by section, color, type, mana cost
    var sortNum = 0;
    sortNum += 10000 * ( sectionSort.indexOf(rowInfo[sectionCol]) + 1);
    sortNum += 100   * ( colorSort.indexOf(rowInfo[colorCol]) + 1);
    sortNum += 10    * ( typeSort.indexOf(rowInfo[typeCol]) + 1);
    sortNum += rowInfo[manaCostCol];
    rowInfo[sortCol] = sortNum;
    // write the row
    cList.getRange(i+2, 1, 1, rowInfo.length).setValues([rowInfo]);
  }
  // sort sheet by 'Sort' column
  sortMe();
}

// queries the ScryFall API for cardName and parses it
// cardName must be an exact card name
function getCardInfo(cardName) {
  // qurery ScryFall API, must use exact card name
  var scryfallAPI = "https://api.scryfall.com/cards/named?exact=";
  var cardJSON = UrlFetchApp.fetch(scryfallAPI.concat(cardName));
  // parse the JSON file
  var cardInfo = JSON.parse(cardJSON);
  return cardInfo;
}

// finds what column in 'Column List' specified string is in
function findColumn(string) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var colList = ss.getSheetByName('Column List');
  var colData = colList.getDataRange().getValues(); 
  for (var i = 0; i < colData.length; i++) {
    if (colData[0][i] == string) {
      return i; // return column number
    }
  }
  // couldn't find column. muse be exact match (case sensitive)
  throw new Error("Couldn't find column " + string + " !");
  return (-1);
}

// get the length of a specified column in sheet
function columnLength(sheet, column) {
  for (var length = sheet.length-1; length >=0; length--) {
    if (sheet[length][column] != null && sheet[length][column] != '') {
      return length+1;
    }
  }
}

// find the color of a card
// input an array of colors (order doesn't matter). eg: ['W','B']
// return array with color letter(s) and text. eg: ['WB', 'Multi']
function getColor(colors) {
  var colorCode = []; // color letters, e.g. WU
  var colorText = []; // color text, e.g. White
  var colorListCode = ['W','U','B','R','G'];
  var colorListText = ['White','Blue','Black','Red','Green','Multi'];
  if (colors.length == 0) {
    // colors.length is empty for colorless cards
    colorCode = 'CL';
    colorText = 'Colorless';
  } else {
    for (var i = 0; i < colorListCode.length; i++) {
      if (colors.indexOf(colorListCode[i]) > -1) {
        // will always write color codes in WUBRG order (defined by colorListCode)
        colorCode += colorListCode[i];
        colorText = colorListText[i];
      }
    }
  }
  if (colors.length > 1) {
    // multicolor cards have more than 1 color
    // write 'Multi' instead of each color
    colorText = colorListText[5];
  }
  return [colorCode, colorText];
}

// find a string in the card text
// colData is the data in 'Column List' sheet
// cardText is string of oracle text
// propCol is what column to look in 'Column List', ie which property?
function findString(colData, cardText, propCol) {
  // does this card have this property? default blank/no
  var check = '';
  for (var i = 1; i < columnLength(colData,propCol); i++) {
    // check for strings in 'Column List', case insensitive
    var regex = new RegExp(colData[i][propCol],'i');
    if (regex.test(cardText)) {
      // has the property, mark it with an 'x' and stop loop
      // multiple instances of a property are not counted
      check = 'x';
      break;
    }
  }
  return check;
}
