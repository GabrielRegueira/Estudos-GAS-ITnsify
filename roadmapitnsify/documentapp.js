function main(){
    var docCreate = createDocInFolder();
    var editDoc = editDocument(docCreate);
}

function createDocInFolder() {
    // ID da pasta onde o documento será criado
    var folderId = "1BX7eaZB9ERUyxf8NxRud4o7L1iqcoz5N";
    // Nome do documento
    var docName = "Relatório 2.0";
    // Pasta onde o documento será criado
    var folder = DriveApp.getFolderById(folderId);
    // Cria o documento
    var doc = DocumentApp.create(docName);
    // ID do documento
    var docId = doc.getId();
    // Pega o arquivo do documento
    var docFile = DriveApp.getFileById(docId);
    // Move o arquivo para a pasta
    docFile.moveTo(folder);
    // retorna o id do documento
    return docId;
}

// function editDocument(docId){
    
//     Logger.log("Editando documento " + docId);
//     var body = DocumentApp.openById(docId).getBody();
    

//     var title = body.appendParagraph("Relatório 2.0");
//     title.setHeading(DocumentApp.ParagraphHeading.TITLE);

//     var par1 = body.appendParagraph("Item 1");
//     body.appendParagraph("Item 2");
//     body.appendParagraph("Item 3");
//     par1.setHeading(DocumentApp.ParagraphHeading.NORMAL);

//     var par2 = body.appendParagraph("Este é um relatório automatizado, gerado pelo Google Apps Script.");
//     par2.setHeading(DocumentApp.ParagraphHeading.HEADING6);
// }

// function bReplace(){ 
//     var body = DocumentApp.openById("1Rtnq-RjTy_38UHIbWUUKL6cFoJ8d8fUrOGqMFmy6HVI");
//     var stringReplace = body.getBlob().getDataAsString('UTF-8').replace('Item', 'Ponto');
    
//     return stringReplace;   
// }


// testando com id fixo
function editDocument() {
    try {
      const doc = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
      const body = doc.getBody();
      
      const title = body.insertParagraph(0, "Relatório 2.0");
      title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  
      const items = ["Item 1", "Item 2", "Item 3"];
      for (let i = 0; i < items.length; i++) {
        list = body.appendListItem(items[i]);
        list.setGlyphType(DocumentApp.GlyphType.BULLET);
      }
  
      const description = body.appendParagraph("Este é um relatório automatizado, gerado pelo Google Apps Script.");
      description.setHeading(DocumentApp.ParagraphHeading.HEADING6);
    } catch (error) {
      Logger.log("Erro ao editar o documento: " + error.message);
    }
  }
  

function Replace(){ 
    var body = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    var text = body.getBody().getText();
    console.log(text);
    // replace all instances of "Item" with "Ponto"
    var newText = text.replace(/Item/g, "Ponto");
    body.getBody().setText(newText);
}

function captureRelatorio(){
    var doc = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    var text = doc.getBody().getText();
    var searchResult = doc.getBody().findText("Relatório");
    var start = searchResult.getStartOffset();
    var a = text.substring(0,start);
    console.log(start);
    console.log(a);
}

function insetImage(){
    var body = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    // var paragraph = body.appendParagraph("A imagem entrará aqui");
    var image = DriveApp.getFileById('10cRwhFwPpUmGAw-ClM6Plt4eGnxxXPhH').getBlob();
    var posImage = paragraph.addPositionedImage(image)
    // .setTopOffset(10)
    // .setLeftOffset(20);
}

function table(){
    var body = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    var cells = [
        ['Coluna 1 , Linha 1', 'Coluna 2, Linha 1'],
        ['Coluna 1, Linha 2', 'Coluna 2 , Linha 2'],
        ['Coluna 1, Linha 3', 'Coluna 2 , Linha 3']
        ];
    body.appendTable(cells);
}

function searchLastElement(){
    var doc = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    Logger.log('Element type: ' + rangeElement.getElement().getType());
if (rangeElement.isPartial()) {
  Logger.log('The character range begins at ' + rangeElement.getStartOffset());
  Logger.log('The character range ends at ' + rangeElement.getEndOffsetInclusive());
} else {
  Logger.log('The entire range element is included.');
}

}

function getLastParagraphStyle(){
    var doc = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    var paragraphs = doc.getBody().getParagraphs();
    var lastParagraph = paragraphs[paragraphs.length - 1];
    var style = lastParagraph.getTextAlignment();
    
    if (style) {
      Logger.log(style);
      return style;
    } else {
      Logger.log('O estilo de alinhamento de texto não está definido para o último parágrafo.');
      return 'O estilo de alinhamento de texto não está definido para o último parágrafo.';
    }
}
  
function boldAllParagraphs(){
    var documentId = "1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4";
    var paragraphs = DocumentApp.openById(documentId).getBody().getParagraphs();
    paragraphs.forEach(function(paragraph) {
      paragraph.setBold(true);
    });
}

function removeBoldFromAllParagraphs(){
    var documentId = "1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4";
    var paragraphs = DocumentApp.openById(documentId).getBody().getParagraphs();
    paragraphs.forEach(function(paragraph) {
      paragraph.setBold(false);
    });
}
  
function createAndGetSpreadsheet(){
    var spreadsheet = SpreadsheetApp.create("Nova Planilha");
    var spreadsheetUrl = spreadsheet.getUrl();
    Logger.log("A URL da nova planilha é: " + spreadsheetUrl);
    return spreadsheetUrl;
}

function addDataToSpreadsheet(){
    var spreadsheet = SpreadsheetApp.openById("16Qet5wuCKu6d90DTzMklW7gseFT9be3-mHG-R5L5AJk");
    var sheet = spreadsheet.getActiveSheet();
    sheet.appendRow(["nome", "idade", "cidade"]);
}
  
function createForm(){
    var form = FormApp.create('Formulário de Captura de Dados');
    
    form.addTextItem()
        .setTitle('ID do Usuário')
        .setRequired(true);
    
    form.addTextItem()
        .setTitle('Nome')
        .setRequired(true);
    
    form.addTextItem()
        .setTitle('Email')
        .setRequired(true);
    
    form.addTextItem()
        .setTitle('CPF')
        .setRequired(true);
    
    var sheet = SpreadsheetApp.create('Planilha de Dados do Usuário');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
    
    Logger.log('Formulário criado com sucesso: ' + form.getEditUrl());
}
  

  
  





