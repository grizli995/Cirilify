function onHomepageDocs(e) {
  return onHomepage(e, 'Docs');
}

function onHomepageSheets(e) {
  return onHomepage(e, 'Sheets');
}

function onHomepageSlides(e) {
  return onHomepage(e, 'Slides');
}

function onHomepage(e, app) {
  const builder = CardService.newCardBuilder();
  const submitAction = CardService.newAction()
      .setFunctionName('toggleSelectedTextIn' + app)
      .setLoadIndicator(CardService.LoadIndicator.SPINNER);
  const submitButton = CardService.newTextButton()
      .setText('Ćirilify!')
      .setOnClickAction(submitAction)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
  const optionsSection = CardService.newCardSection()
      .addWidget(submitButton);
  builder.addSection(optionsSection);
  return builder.build();
}

function toggleSelectedText() {
  var hostApp = ScriptApp.getHostApplication();
  if (hostApp == "docs") {
    toggleSelectedTextInDocs();
  } else if (hostApp == "sheets") {
    toggleSelectedTextInSheets();
  } else if (hostApp == "slides") {
    toggleSelectedTextInSlides();
  }
}

function toggleSelectedTextInDocs() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (selection) {
    const elements = selection.getRangeElements();
    for (let i = 0; i < elements.length; i++) {
      const element = elements[i];
      if (element.getElement().editAsText) {
        const text = element.getElement().asText();
        const txt = text.getText();
        text.setText(convertCyrillic(txt));
      }
    }
  }
}

function toggleSelectedTextInSheets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  const newValues = values.map(row => row.map(value => convertCyrillic(value.toString())));
  range.setValues(newValues);
}

function toggleSelectedTextInSlides() {
  const slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  const shapes = slide.getShapes();
  shapes.forEach(shape => {
    const textRange = shape.getText();
    const originalText = textRange.asString();
    textRange.setText(convertCyrillic(originalText));
  });
}

function convertCyrillic(input) {
  const mapping =  {
    'A': 'А', 'B': 'Б', 'C': 'Ц', 'Č': 'Ч', 'Ć': 'Ћ', 'D': 'Д', 'Dž': 'Џ', 'Đ': 'Ђ',
    'E': 'Е', 'F': 'Ф', 'G': 'Г', 'H': 'Х', 'I': 'И', 'J': 'Ј', 'K': 'К', 'L': 'Л',
    'Lj': 'Љ', 'M': 'М', 'N': 'Н', 'Nj': 'Њ', 'O': 'О', 'P': 'П', 'R': 'Р', 'S': 'С',
    'Š': 'Ш', 'T': 'Т', 'U': 'У', 'V': 'В', 'Z': 'З', 'Ž': 'Ж',
    'a': 'а', 'b': 'б', 'c': 'ц', 'č': 'ч', 'ć': 'ћ', 'd': 'д', 'dž': 'џ', 'đ': 'ђ',
    'e': 'е', 'f': 'ф', 'g': 'г', 'h': 'х', 'i': 'и', 'j': 'ј', 'k': 'к', 'l': 'л',
    'lj': 'љ', 'm': 'м', 'n': 'н', 'nj': 'њ', 'o': 'о', 'p': 'п', 'r': 'р', 's': 'с',
    'š': 'ш', 't': 'т', 'u': 'у', 'v': 'в', 'z': 'з', 'ž': 'ж', 
    'А': 'A', 'Б': 'B', 'Ц': 'C', 'Ч': 'Č', 'Ћ': 'Ć', 'Д': 'D', 'Џ': 'Dž', 'Ђ': 'Đ',
    'Е': 'E', 'Ф': 'F', 'Г': 'G', 'Х': 'H', 'И': 'I', 'Ј': 'J', 'К': 'K', 'Л': 'L',
    'Љ': 'Lj', 'М': 'M', 'Н': 'N', 'Њ': 'Nj', 'О': 'O', 'П': 'P', 'Р': 'R', 'С': 'S',
    'Ш': 'Š', 'Т': 'T', 'У': 'U', 'В': 'V', 'З': 'Z', 'Ж': 'Ž',
    'а': 'a', 'б': 'b', 'ц': 'c', 'ч': 'č', 'ћ': 'ć', 'д': 'd', 'џ': 'dž', 'ђ': 'đ',
    'е': 'e', 'ф': 'f', 'г': 'g', 'х': 'h', 'и': 'i', 'ј': 'j', 'к': 'k', 'л': 'l',
    'љ': 'lj', 'м': 'm', 'н': 'n', 'њ': 'nj', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's',
    'ш': 'š', 'т': 't', 'у': 'u', 'в': 'v', 'з': 'z', 'ж': 'ž'
  };

  let result = '';
  for (let i = 0; i < input.length; i++) {
    if (input[i] === 'N' && input[i + 1] === 'j') {
      result += mapping['Nj'];
      i++;
    } else if ( (input[i] === 'n' && input[i + 1] === 'j')) {
      result += mapping['nj'];
      i++;
    } else if ( (input[i] === 'L' && input[i + 1] === 'j')) {
      result += mapping['Lj'];
      i++;
    } else if ( (input[i] === 'l' && input[i + 1] === 'j')) {
      result += mapping['lj'];
      i++;
    } else if ((input[i] === 'D' && input[i + 1] === 'j')) {
      result += mapping['Đ'];
      i++;
    } else if ( (input[i] === 'd' && input[i + 1] === 'j')) {
      result += mapping['đ'];
      i++;
    } else if ( (input[i] === 'D' && input[i + 1] === 'ž')) {
      result += mapping['Dž'];
      i++;
    } else if ( (input[i] === 'd' && input[i + 1] === 'ž')) {
      result += mapping['dž'];
      i++;
    } else if (mapping[input[i]] === undefined) {
      result += input[i];
      continue; 
    } else {
      result += mapping[input[i]] || input[i];
    }
  }

  return result;
}
