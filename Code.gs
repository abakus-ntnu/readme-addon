/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Add bold style to italic text
 */
function bold2italic() {
  var body = DocumentApp.getActiveDocument().getBody();
  var text = body.editAsText();
  var startItalic = 0;
  var italic = false;
  for (var i = 0; i < text.getText().length; i++) {
    if (text.isItalic(i) && !italic) {
      startItalic = i;
      italic = true;
    }
    else if (!text.isItalic(i) && italic) {
      italic = false;
      text.setBold(startItalic, i-1, true);
    }
  }
  if (italic) {
    text.setBold(startItalic, text.getText().length - 1, true);
  }
}


/**
 * Replace "" with «»
 */
function replaceQuotes() {
  var body = DocumentApp.getActiveDocument().getBody();
  var text = body.editAsText();
  var textString = body.getText();
  for (var i = 0; i < textString.length; i++) {
    var char = textString[i];
    if (char === '“') {
      text.deleteText(i, i);
      text.insertText(i, '«');
    } else if (char === '”') {
      text.deleteText(i, i);
      text.insertText(i, '»');
    }
  }
}

/**
 * Convert " - " and " -- " to " – "
 */
function replaceHyphensWithDashes() {
  var body = DocumentApp.getActiveDocument().getBody();
  body.replaceText(' --? ', ' – ');
}

/**
 * Bold occurences of "readme". Lower-case only.
 */
function boldReadme() {
  var body = DocumentApp.getActiveDocument().getBody();
  var text = body.editAsText();
  var textString = body.getText();

  var offset = 0
  var readmeIndex = textString.indexOf('readme');
  while (readmeIndex > -1) {
    Logger.log(readmeIndex);
    offset = readmeIndex + 5;
    text.setBold(readmeIndex, offset, true);
    var nextIndex = textString.slice(offset).indexOf('readme');
    readmeIndex = nextIndex > -1 ? offset + nextIndex : -1;
  }
}

function doAllGeneral() {
  bold2italic();
  replaceQuotes();
  replaceHyphensWithDashes();
}

function doAllReadme() {
  boldReadme();
}

function doAll() {
  doAllGeneral();
  doAllReadme();
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Gjør alt', 'doAll')
    .addSeparator()
    .addItem('Gjør alt under Generelt', 'doAllGeneral')
    .addItem('Gjør kursiv tekst også fet', 'bold2italic')
    .addItem('Bytt hermetegn med anførselstegn', 'replaceQuotes')
    .addItem('Bytt bindestreker med tankestreker', 'replaceHyphensWithDashes')
    .addSeparator()
    .addItem('Gjør alt under readme', 'doAllReadme')
    .addItem('Gjør readme fet', 'boldReadme')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}
