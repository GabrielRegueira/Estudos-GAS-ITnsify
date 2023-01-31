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
function editDocument(){
    
    
    var body = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4").getBody();
    

    var title = body.appendParagraph("Relatório 2.0");
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);

    var par1 = body.appendParagraph("Item 1");
    body.appendParagraph("Item 2");
    body.appendParagraph("Item 3");
    par1.setHeading(DocumentApp.ParagraphHeading.NORMAL);

    var par2 = body.appendParagraph("Este é um relatório automatizado, gerado pelo Google Apps Script.");
    par2.setHeading(DocumentApp.ParagraphHeading.HEADING6);
}

function Replace(){ 
    var body = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    var text = body.getBody().getText();
    console.log(text);
    // replace all instances of "Item" with "Ponto"
    var newText = text.replace(/Item/g, "Ponto");
    body.getBody().setText(newText);
}

function captureRelatorio() {
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






