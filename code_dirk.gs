
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'sort')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

const order = [
"",
"kokend water",


"Dirk",
"mini krieltjes",
"Kipworstje",
"kipreepjes",
"eieren",
"drinken Derk",
"raak framboos",
"raak zero framboos",
"halfvolle melk",
"Patat",
"spinazie",
"drinken Cas",
"snoeptomaatjes",
"geraspte kaas",
"geraspte 30+ kaas",
"appelmoes",



"Jumbo",
"pitabroodjes",
"hamreepjes",
"honing",
"bloem",
"drinken Marieke",
"crystal clear",
"crystalclear",
"limonade",
"aanmaaksiroop",
"karvan cevitam",
"yoghurt",
"vloeibare margarine",
"becel voor op brood",
"vloeibare margarine",
"kookroom light",
"boter",
"macaroni",
"spaghetti",
"tomatenpuree",
"tomatenblokjes",
"chips",
"Kroket",
"Frikandel",
"Yildriz knoflooksaus",
"duyvis Aioli",
"currysaus",
"mosterd",
"rijst",
"ketjap",
"fajita kruiden",
"taco kruiden",
"sweet chilisaus",
"salsasaus",
"kerrie",
"uienpoeder",
"komijn",
"koriander",
"gemberpoeder",
"kurkuma",
"djinten",
"cacao",
"tijm",
"majoraan",
"paprikapoeder",
"picadillokruiden",
"peper",
"zout",
"kipkruiden",
"nootmuskaat",
"rookworst",
"vuilniszakken",
"allesreiniger",
"tauge",
"broccoli",
"witlof",
"sperziebonen",
"paprika",
"Komkommer",
"courgette",
"tomaat",
"rundergehakt",
"blikje ananas",
"blikje ananas",
"blikje mais",
"blikje doperwten",


"AH",
"albert heijn",
"gesneden ijsbergsla",
"ijsbergsla",
"fijngesneden andijvie",
"appel",
"kipfiletstukjes",
"kipstukjes",
"kipfilet",
"Smeerpathe",
"Pate",
"cheddar kaas",
"broodbeleg",
"Ovengebakken kip",
"tortilla",
"kokosmelk",
"kippenbouillon",
"kippenbouillonblokje",
"runderbouillonblokje",
"runderbouillon",
"pompoenpit knackebrod",
"pompoenpitknackebrod",


"markt",
"aardappels"


];

function sort() {
  let items = getSelectedText();
  Logger.log(JSON.stringify(items));
  items = selectedTextToArr(items);
  Logger.log(JSON.stringify(items));
  items = doTheSort(items);
  insertText(items.join("\n"));
}

function doTheSort(items) {
  items.sort(function (a, b) {
    a = a.toLocaleLowerCase();
    b = b.toLocaleLowerCase();
    let numa = 0;
    let numb = 0;
    order.forEach(function callback(value, index) {
      value = value.toLocaleLowerCase();
      if (a.includes(value)) {
        numa = index;
      }
      if (b.includes(value)) {
        numb = index;
      }
    });
    return numa - numb;
  });
  return items;
}

function selectedTextToArr(items) {
  items2 = [];
  items.forEach(function callback(value, index) {
    value.split(/[\r\n]+/).forEach(function callback(val) {
      items2.push(val);
    });
  });
  items = items2;
  return items;
}
function getSelectedText() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  const text = [];
  if (selection) {
    const elements = selection.getSelectedElements();
    for (let i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        const element = elements[i].getElement().asText();
        const startIndex = elements[i].getStartOffset();
        const endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        const element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          const elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText) {
            text.push(elementText);
          }
        }
      }
    }
  }
  if (!text.length) throw new Error('Please select some text.');
  return text;
}

function insertText(newText) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    let replaced = false;
    const elements = selection.getSelectedElements();
    if (elements.length === 1 && elements[0].getElement().getType() ===
      DocumentApp.ElementType.INLINE_IMAGE) {
      throw new Error('Can\'t insert text into an image.');
    }
    for (let i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        const element = elements[i].getElement().asText();
        const startIndex = elements[i].getStartOffset();
        const endIndex = elements[i].getEndOffsetInclusive();
        element.deleteText(startIndex, endIndex);
        if (!replaced) {
          element.insertText(startIndex, newText);
          replaced = true;
        } else {
          // This block handles a selection that ends with a partial element. We
          // want to copy this partial text to the previous element so we don't
          // have a line-break before the last partial.
          const parent = element.getParent();
          const remainingText = element.getText().substring(endIndex + 1);
          parent.getPreviousSibling().asText().appendText(remainingText);
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just remove the text within the last paragraph instead.
          if (parent.getNextSibling()) {
            parent.removeFromParent();
          } else {
            element.removeFromParent();
          }
        }
      } else {
        const element = elements[i].getElement();
        if (!replaced && element.editAsText) {
          // Only translate elements that can be edited as text, removing other
          // elements.
          element.clear();
          element.asText().setText(newText);
          replaced = true;
        } else {
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just clear the element.
          if (element.getNextSibling()) {
            element.removeFromParent();
          } else {
            element.clear();
          }
        }
      }
    }
  } else {
    const cursor = DocumentApp.getActiveDocument().getCursor();
    const surroundingText = cursor.getSurroundingText().getText();
    const surroundingTextOffset = cursor.getSurroundingTextOffset();

    // If the cursor follows or preceds a non-space character, insert a space
    // between the character and the translation. Otherwise, just insert the
    // translation.
    if (surroundingTextOffset > 0) {
      if (surroundingText.charAt(surroundingTextOffset - 1) !== ' ') {
        newText = ' ' + newText;
      }
    }
    if (surroundingTextOffset < surroundingText.length) {
      if (surroundingText.charAt(surroundingTextOffset) !== ' ') {
        newText += ' ';
      }
    }
    cursor.insertText(newText);
  }
}
