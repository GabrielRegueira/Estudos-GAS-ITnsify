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

function bReplace(){ 
    var body = DocumentApp.openById("1xPhcLo1O4aOGmLYwkKLRemaYSoIJseQdrXy3NKlJEX4");
    var stringReplace = body.getBlob().getDataAsString().replace('Item', 'Ponto');
    return stringReplace;

  
}

