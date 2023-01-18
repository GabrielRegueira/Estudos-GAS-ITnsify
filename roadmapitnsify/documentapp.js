function main(){
    var docCreate = createDocInFolder();
    var editDoc = editDocument(docCreate);
}

function createDocInFolder() {
    // ID da pasta onde o documento será criado
    var folderId = "1BX7eaZB9ERUyxf8NxRud4o7L1iqcoz5N";
    // Nome do documento
    var docName = "Relatório";
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

function editDocument(id){
    Logger.log("Editando documento " + id);
    var body = DocumentApp.openById('1udh_R40UyPcm0KQJncQVbAVU_rOkKT91hUB9-X5_rBU').getBody();

    var title = body.appendParagraph("Relatório");
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);

    var par1 = body.appendParagraph("Item 1");
    body.appendParagraph("Item 2");
    body.appendParagraph("Item 3");
    par1.setHeading(DocumentApp.ParagraphHeading.NORMAL);

    var par2 = body.appendParagraph("Este é um relatório automatizado, gerado pelo Google Apps Script.");
    par2.setHeading(DocumentApp.ParagraphHeading.HEADING6);
}
