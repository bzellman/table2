var preferredFileName = null;
var baseTicketURL = null;
var basePeopleURL = null;
var currentDocName = null;
 
function getUserPreferences() {
    var userProperties = PropertiesService.getUserProperties();
    
    baseTicketURL = userProperties.getProperty('baseTicketURL')|| null;
    basePeopleURL = userProperties.getProperty('basePeopleURL')|| null;
    preferredFileName = userProperties.getProperty('preferredFileName')|| null;

    var preferences = {
      baseTicketURL: userProperties.getProperty('baseTicketURL')|| "",
      basePeopleURL: userProperties.getProperty('basePeopleURL')|| "",
      preferredFileName: userProperties.getProperty('preferredFileName')|| "",
    }
    return preferences;
}

function onOpen(e) {
  // Add a custom menu to the spreadsheet.
  DocumentApp.getUi()
      .createMenu('Tables2')
      .addItem('Options', 'showUserConfigForm')
      .addItem('New File As List', 'makeListDoc')
      .addItem('New File As JSON', 'makeJSONDoc')
      .addToUi();
}

function saveUserSettings(formData) {
    var userProperties = PropertiesService.getUserProperties();
    for (var key in formData) {
        if (formData.hasOwnProperty(key)) {
            userProperties.setProperty(key, formData[key]);
        }
    }
}

function showUserConfigForm() {
    var template = HtmlService.createTemplateFromFile('userPreferences');
    template.preferences = getUserPreferences();
    
    var html = template.evaluate()
        .setWidth(500) 
        .setHeight(500); 
    DocumentApp.getUi().showModalDialog(html, 'Tables 2');
}

function makeJSONDoc() {
  const jsonDataToUse = extractBetterStructuredData();
  const ui = DocumentApp.getUi();
  exportJsonData(jsonDataToUse, ui);
}

function makeListDoc() {
  let ui = DocumentApp.getUi();
  const jsonDataToUse = extractBetterStructuredData();
  createFormattedDocument(jsonDataToUse, ui);
}

function makeListDocDebug() {
  let ui = null
  const jsonDataToUse = extractBetterStructuredData();
  createFormattedDocument(jsonDataToUse, ui);
}

function exportJsonData(jsonDataToUse, ui) {
    if (jsonDataToUse === null) {
        showCompletionModal(false, ui, null);
    } else {
      var docName = ""
      if (preferredFileName != null && preferredFileName != "") {
        docName = preferredFileName;
      } else {
        let currentDocName = DocumentApp.getActiveDocument().getName();    
        docName = String(currentDocName + " as JSON.json");
      }
      var doc = DocumentApp.create(String(docName));  
      var body = doc.getBody();
      body.setText(jsonDataToUse);
      doc.saveAndClose();
      showCompletionModal(true, ui, doc.getUrl());
    }
}

function extractBetterStructuredData() {
  getUserPreferences();
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  var dataObjects = [];
  var paragraphs = body.getParagraphs();
  var currentHeader = '';
  var currentTable;
  
  paragraphs.forEach(function(paragraph) {

    if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
      // Store current heading as the section title
      currentHeader = paragraph.getText();
    } else if (paragraph.getText().trim() === '' && paragraph.getNextSibling()) {

      var nextElement = paragraph.getNextSibling();
      if (nextElement.getType() === DocumentApp.ElementType.TABLE) {
        currentTable = nextElement.asTable();
        
        // Parse the table into a JSON object
        var tableData = parseTable(currentTable);
        // Check if the table contains any non-empty row
        var hasNonEmptyRow = tableData.some(function(rowObject) {
          return Object.values(rowObject).some(function(cellValue) {

            if (typeof cellValue === "string") {
              return cellValue.trim() !== '';
            } else if (typeof cellValue === "object") {
              return cellValue.some(function(value) {
              return value.trim() !== '';
              });
            } else {
              return cellValue !== null;
            }
          });
        });

        if (hasNonEmptyRow) {
          var sectionData = {
            header: currentHeader,
            data: tableData
          };

          dataObjects.push(sectionData);
        }
      } 
      // else {
      //   // TBD NOT BUT important
      //   if (typeof nextElement.asText().getText() === "string" && nextElement.asText().getText() != "" ) {
      //     // get text and add to Other Section
      //   }        
      // }
    }
  });

  return JSON.stringify(dataObjects, null, 2);
}

function parseTable(table) {
  let headers = [];
  let tableData = [];

  var headerRow = table.getRow(0);
  for (var i = 0; i < headerRow.getNumCells(); i++) {
    headers.push(headerRow.getCell(i).getText());
  }
  for (var rowNum = 1; rowNum < table.getNumRows(); rowNum++) {
    let row = table.getRow(rowNum);
    let rowData = {};

    for (var cellNum = 0; cellNum < row.getNumCells(); cellNum++) {
      let cell = row.getCell(cellNum);

      // Break out Cell Text elemenets into Text split by "\n" and make markup links for evaluation
      let linesArrayToAdd = trimWhiteSpaceItemsFromArray(splitTextElementByNewLine(cell.editAsText()));


      if (linesArrayToAdd !== null) {
        if (headers[cellNum] == "Ticket") {
          if (sanitizeIssueKey(linesArrayToAdd.join(' ')).length > 0) {
            rowData[headers[cellNum]] = sanitizeIssueKey(sanitizeIssueKey(linesArrayToAdd.join(' ')));
          }        
        } else if (linesArrayToAdd.length !== 0) {
          if (linesArrayToAdd.length > 0) {
            rowData[headers[cellNum]] = linesArrayToAdd;
          }     
        }
      }

    }
    tableData.push(rowData);
  }
  return tableData;
}

function splitTextElementByNewLine(textElement) {
  let fullText = textElement.getText();
    // Split the full text into lines
    let lines = fullText.split("\n");
    let formattedLines = [];

    lines.forEach((line, lineIndex) => {
      // Find the start index of the current line in the full text
      let lineStartIndex = fullText.indexOf(line, lineIndex > 0 ? fullText.indexOf(lines[lineIndex - 1]) + lines[lineIndex - 1].length + 1 : 0);
      let lineEndIndex = lineStartIndex + line.length;

      // Initialize an empty string to accumulate the formatted line
      let formattedLine = "";
      let previousIndex = 0;

      // Iterate over attribute runs to find links within the current line
      textElement.getTextAttributeIndices().forEach((attrIndex) => {
        if (attrIndex >= lineStartIndex && attrIndex < lineEndIndex) {
          let linkUrl = textElement.getLinkUrl(attrIndex);
          if (linkUrl) {
            // Calculate the end index of the link text
            let nextAttrIndex = textElement.getTextAttributeIndices().find(index => index > attrIndex);
            let linkEndIndex = nextAttrIndex && nextAttrIndex < lineEndIndex ? nextAttrIndex : lineEndIndex;
            let linkText = fullText.substring(attrIndex, linkEndIndex);

            // Append text before the link (if any), then append the formatted link
            formattedLine += line.substring(previousIndex, attrIndex - lineStartIndex) + `[${linkText}](${linkUrl})` + "";
            previousIndex = linkEndIndex - lineStartIndex;
          }
        }
      });

      // Append any remaining text after the last link (or the whole line if no links)
      formattedLine += line.substring(previousIndex);
      formattedLines.push(formattedLine);
    });

  return formattedLines;
}

function trimWhiteSpaceItemsFromArray(array) {
    return array.filter(item => {
      return item.trim() !== "";
  }); 
}

// Helper Function 
function splitIntoBulletListItems(text) {
  const listItems = text.split("\n").filter(item => item.startsWith("• "));
  return listItems.map(item => item.substring(2).trim());
}

function createFormattedDocument(json, ui) {
  if (preferredFileName != null && preferredFileName != "") {
    var doc = DocumentApp.create(String(preferredFileName));  
  } else {
    var currentDocName = DocumentApp.getActiveDocument().getName();
    var doc = DocumentApp.create(String(currentDocName + " as List"));  
  } 

  let body = doc.getBody();

  JSON.parse(json).forEach(section => {    
    body.appendParagraph(section.header).setHeading(DocumentApp.ParagraphHeading.HEADING2); 
      section.data.forEach(dataItem => {
        processNestedData(dataItem, 0, body);
      });
  });
  doc.saveAndClose();
  if( ui != null) {
    showCompletionModal(true, ui, doc.getUrl());
  }
}

function processNestedData(data, nestingLevel, body) { 
  var nestingLevel = nestingLevel;
  if (typeof data === 'object' && data !== null) {
    for (const key in data) {
      var localNesting = nestingLevel;
      const value = data[key];
        if (keysHaveChildKeysWithValues(value) == false) {
          if (Array.isArray(value) && value.length > 1) {
            body.appendListItem(String(key + ":")).setNestingLevel(localNesting).setBold(true);
            for (subitem in value) {              
              let listItem = body.appendListItem(" ").setNestingLevel(localNesting+1).setBold(false);
              applyMarkupLinksToText(listItem, value[subitem]);              
            }            
          } else {
            if (key == "Ticket") {
              let ticketLink = body.appendListItem(String(value)).setNestingLevel(localNesting).setBold(true)
              nestingLevel = nestingLevel + 1;
              if (baseTicketURL != "" && baseTicketURL != null) {
                ticketLink.setLinkUrl(String(baseTicketURL+"/"+value));
              }
            } else {              
              
              let listItem = body.appendListItem(" ");
              applyMarkupLinksToText(listItem, String(value));
              // Logger.log(key);
              // Logger.log(value);
              // Logger.log(nestingLevel)
              if (nestingLevel == 0) {
                nestingLevel = 1;
              }
              listItem.editAsText().insertText(0, String(key+":"));
              listItem.setNestingLevel(localNesting);
            }                    
          }
        } else {
          nestingLevel = nestingLevel+1;
          processNestedData(value, nestingLevel, body);
        }
    }
  }
}

function applyMarkupLinksToText(listItem, textWithMarkup) {
    // Regular expression to identify [text](url) patterns
    const linkPattern = /\[([^\]]+)]\((http[s]?:\/\/[^\)]+)\)/g;
    
    let currentInsertionPoint = 0;
    
    let lastIndex = 0; 
    let match;
    
    while ((match = linkPattern.exec(textWithMarkup)) !== null) {
        const [fullMatch, linkText, linkUrl] = match;
        var textForLink = linkText;
        var linkURLForInsertion = linkUrl;
        // Append any text before the link
        if (match.index > lastIndex) {
            listItem.appendText(textWithMarkup.substring(lastIndex, match.index)).setBold(false);
            currentInsertionPoint += match.index - lastIndex;
        }

        if (basePeopleURL !== "" && basePeopleURL !== null) {
          if (linkUrl && linkUrl.startsWith(basePeopleURL)) {
            // listItem.appendText(String("@" + linkText));
            textForLink = String("@" + textForLink);
            linkURLForInsertion = null;
          }

        } 
        // else {
          listItem.appendText(textForLink);
          let textElement = listItem.editAsText();
          textElement.setLinkUrl(currentInsertionPoint+1, currentInsertionPoint + textForLink.length, linkURLForInsertion);
          currentInsertionPoint += textForLink.length;
          lastIndex = match.index + fullMatch.length;
        // }
        
    }
    
    // Append any remaining text after the last link
    if (lastIndex < textWithMarkup.length) {
        listItem.appendText(textWithMarkup.substring(lastIndex))
          .setBold(false);
        Logger.log(typeof currentInsertionPoint);
        let element = listItem.editAsText();
        element.setLinkUrl(currentInsertionPoint+1, listItem.getText().length-1, null);
    }
}

function keysHaveChildKeysWithValues(data) {
  for (const key in data) {
    const value = data[key];

    // Check if the value is an array
    if (Array.isArray(value)) {
      // Check if any element within the array is an object
      for (const element of value) {
        if (typeof element === 'object' && element !== null) {
          return true;
        }
      }
    }
  }
  return false;
}

function showCompletionModal(successBool, ui, url) {
  var template = null
  if (successBool === true && url != null && ui !== null) {
    template = HtmlService.createTemplateFromFile('completionModalSuccess');
    template.preferences = {
      url: url
    };
  } else {
    template = HtmlService.createTemplateFromFile('completionModalFail');
  }

  var html = template.evaluate()
    .setWidth(500) 
    .setHeight(200); 
  ui.showModalDialog(html, 'Tables 2');
}

function sanitizeIssueKey(text) {
  const issueKeyPattern = /([A-Z]+-\d+)/;

  const issueKeyMatch = text.match(issueKeyPattern);

  if (issueKeyMatch) {
    var issueKey = issueKeyMatch[0];
    return issueKey;
  }
  // If no issue key is found, return the original text
  return text;
}