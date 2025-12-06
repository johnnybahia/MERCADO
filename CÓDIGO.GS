function doGet(e) {
  try {
    var action = e.parameter.action;

    // Retorna TODOS OS DADOS da planilha
    if (action === "loadData") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("LISTA");

      if (!sheet || sheet.getLastRow() <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ "markets": [] })).setMimeType(ContentService.MimeType.JSON);
      }

      var data = sheet.getDataRange().getValues();
      var markets = {};

      // Pula o cabeçalho (linha 0)
      for (var i = 1; i < data.length; i++) {
        var marketName = data[i][0];  // Coluna A: Mercado
        var itemName = data[i][1];     // Coluna B: Item
        var price = data[i][2];        // Coluna C: Preço
        var itemId = data[i][4];       // Coluna E: ID

        if (!marketName || !itemName) continue;

        // Cria mercado se não existe
        if (!markets[marketName]) {
          markets[marketName] = {
            id: marketName.replace(/\s/g, '_').toLowerCase(),
            name: marketName,
            date: new Date().toLocaleDateString('pt-BR'),
            products: []
          };
        }

        // Adiciona produto
        markets[marketName].products.push({
          id: itemId || Date.now(),
          name: itemName,
          price: String(price).replace(',', '.'),
          img: null  // Não carrega imagens por performance
        });
      }

      var marketsList = Object.values(markets);
      return ContentService.createTextOutput(JSON.stringify({ "markets": marketsList })).setMimeType(ContentService.MimeType.JSON);
    }

    // Retorna chave API (comportamento padrão)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("API");
    var key = sheet ? sheet.getRange("A1").getValue() : "";
    return ContentService.createTextOutput(JSON.stringify({ "apiKey": key })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "error": err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Evita conflito se salvar dois itens muito rápido

  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LISTA");
    
    if (!sheet) sheet = ss.insertSheet("LISTA");

    // Configura cabeçalho se vazio (Agora com coluna ID oculta na coluna E)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["MERCADO", "ITEM", "PREÇO", "FOTO", "ID_SISTEMA"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9d9d9");
      sheet.setColumnWidth(4, 100); 
      // A coluna 5 (ID) fica visível mas é técnica. Você pode ocultar no Excel se quiser.
    }

    var action = data.action; // 'add' ou 'delete'

    // --- AÇÃO: ADICIONAR ---
    if (action === "add") {
      var prod = data.product;
      var priceBR = prod.price.toString().replace('.', ',');
      var nextRow = sheet.getLastRow() + 1;

      // Salva: Mercado, Nome, Preço, Foto, ID
      sheet.getRange(nextRow, 1, 1, 3).setValues([[data.marketName, prod.name, priceBR]]);
      sheet.getRange(nextRow, 5).setValue(prod.id); // Salva o ID único

      if (prod.img) {
        try {
          var base64Data = prod.img.split(',')[1];
          var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'produto.jpg');
          var img = SpreadsheetApp.newCellImage().setSourceUrl(blob).build();
          sheet.getRange(nextRow, 4).setValue(img);
        } catch (e) {
          Logger.log("Erro ao salvar imagem: " + e.message);
          // Tenta método alternativo
          try {
            var base64Data = prod.img.split(',')[1];
            var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'produto.jpg');
            sheet.insertImage(blob, 4, nextRow);
            sheet.getRange(nextRow, 4).setValue("IMG");
          } catch (e2) {
            sheet.getRange(nextRow, 4).setValue("Sem Foto");
          }
        }
      } else {
        sheet.getRange(nextRow, 4).setValue("-");
      }
      
      sheet.setRowHeight(nextRow, 80);
    }

    // --- AÇÃO: DELETAR ---
    else if (action === "delete") {
      var idToDelete = data.id;
      var dataRange = sheet.getDataRange().getValues();
      
      // Procura o ID na coluna E (índice 4) de baixo para cima (mais rápido para itens recentes)
      for (var i = dataRange.length - 1; i >= 1; i--) {
        if (String(dataRange[i][4]) === String(idToDelete)) {
          sheet.deleteRow(i + 1); // Deleta a linha encontrada
          break; // Para após encontrar
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ "status": "success" })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": e.message })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
