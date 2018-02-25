/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('sidebar').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Linkify');
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Gets the text the user has selected. If there is no selection,
 * this function displays an error message.
 *
 * @return {Array.<string>} The selected text.
 */
function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var text = [];
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText != '') {
            text.push(elementText);
          }
        }
      }
    }
    if (text.length == 0) {
      throw 'Please select some text.';
    }
    return text;
  } else {
    throw 'Please select some text.';
  }
}

/**
 * Create a URL from the provided text, by searching the user's drive for
 * files containing the word in their title.
 * If found, the files are presented as radio buttons with the name "link" 
 * and the value is the URL - this allows the containing HTML page to listen 
 * to the radio button change event and use the right URL.
 * by default, the last item is checked.
 * @return {object} word, url, page. Page is the preview page to be presented.
 */
function createDriveUrls(text) {
  var word = text[0];
  var url;
  var q = "title contains '" + text[0]+"'";
  var files = DriveApp.searchFiles(q);
  Logger.log(files);
  Logger.log(files.hasNext());
  var filesList = '';
  while (files.hasNext()) {
    var file = files.next();      
    if (filesList === '') {      
      filesList = '<form action=""> ';
    }    
    filesList += '<input type="radio" name="link" value="' + file.getUrl() + '" checked=true>' + file.getName() + '<br>'
    url = file.getUrl();
  }
  if (filesList === '') {
    filesList = 'No matching results found :-('
  }
  else {
    filesList += '</form>';
  }
  Logger.log(filesList);
  return {word: word, url: url, page: filesList};
}

/**
 * Create a URL from the provided text, by linking to Wikipedia
 * 
 * @return {object} word, url, page. Page is result from fetching the URL, with links disabled
 */
function createWikiUrl(text, lang) {
  var word = text[0].trim();
  lang = lang ? lang : 'en';
  var url = 'http://' + lang + '.wikipedia.org/wiki/' + encodeURI(word);  
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
   var noLinks = response.getContentText().replace(/href="/g, 'nohref="');
  return {word: word, url: url, page: noLinks};
}

/**
 * Create a URL from the provided text, according to the requested method
 * 
 * @return {object} word, url, page. Page is the preview page to be presented.
 */
function linkify(method, lang) {
  var text = getSelectedText();  
  if (method == 'url') {
    return createWikiUrl(text, lang);
  }  
  if (method == 'searchDrive') {
    return createDriveUrls(text);
  }
}

/**
 * Set the link URL on the element and recursively on all its children, if it has children.
 */
function setLinkOnAllChildren(element, link) {
  if (!element.getNumChildren || element.getNumChildren() == 0) {
    element.setLinkUrl(link);
    return;
  }
  for (var i = 0; i < element.getNumChildren(); i++) {
    var child = element.getChild(i);
      setLinkOnAllChildren(child, link);
  }    
}

/**
 * Add the link to the text of the current selection, or
 * inserts text at the current cursor location. (There will always be either
 * a selection or a cursor.) If multiple elements are selected, only inserts the
 *  text in the first element that can contain text and removes the
 * other elements.
 *
 * @param {string} newText The text with which to replace the current selection.
 */
function insertLink(newText, link) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {    
    var elements = selection.getRangeElements();    
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i].getElement();
      Logger.log("Found Element " + i + " type: " + element.getType());      
      if (element.getType() === DocumentApp.ElementType.TEXT && 
          elements[i].isPartial()) {
        element = element.asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        element.setLinkUrl(startIndex, endIndex, link);
      }
      else {
        setLinkOnAllChildren(element, link);
      }   
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    if (!cursor) {
      throw ("To insert, select a location in the document");
    }
    var surroundingText = cursor.getSurroundingText().getText();
    var surroundingTextOffset = cursor.getSurroundingTextOffset();

    // If the cursor follows or preceds a non-space character, insert a space
    // between the character and the translation. Otherwise, just insert the
    // translation.
    if (surroundingTextOffset > 0) {
      if (surroundingText.charAt(surroundingTextOffset - 1) != ' ') {
        newText = ' ' + newText;
      }
    }
    if (surroundingTextOffset < surroundingText.length) {
      if (surroundingText.charAt(surroundingTextOffset) != ' ') {
        newText += ' ';
      }
    }
    
    cursor.insertText(newText).setLinkUrl(link);
  }
}

// Helper function that puts external JS / CSS into the HTML file.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
//Icon: http://www.xiconeditor.com/GetIcon.ashx?R=8acbc530-2686-460d-8e63-4b42e5a45b8a
