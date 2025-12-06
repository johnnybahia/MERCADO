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
        var imgData = data[i][5];      // Coluna F: IMG_DATA (base64)

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

        // Adiciona produto (COM FOTO!)
        markets[marketName].products.push({
          id: itemId || Date.now(),
          name: itemName,
          price: String(price).replace(',', '.'),
          img: imgData || null  // Carrega a imagem em base64
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

    // Configura cabeçalho se vazio (Agora com coluna ID oculta na coluna E e IMG_DATA na F)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["MERCADO", "ITEM", "PREÇO", "FOTO", "ID_SISTEMA", "IMG_DATA"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#d9d9d9");
      sheet.setColumnWidth(4, 100);
      sheet.hideColumns(6); // Esconde coluna F (IMG_DATA) pois é só para dados
    }

    var action = data.action; // 'add' ou 'delete'

    // --- AÇÃO: ADICIONAR ---
    if (action === "add") {
      var prod = data.product;
      var priceBR = prod.price.toString().replace('.', ',');
      var nextRow = sheet.getLastRow() + 1;

      // Salva: Mercado, Nome, Preço, Foto, ID, IMG_DATA
      sheet.getRange(nextRow, 1, 1, 3).setValues([[data.marketName, prod.name, priceBR]]);
      sheet.getRange(nextRow, 5).setValue(prod.id); // Salva o ID único

      if (prod.img) {
        // Salva o base64 completo na coluna F (para recuperar depois)
        sheet.getRange(nextRow, 6).setValue(prod.img);

        try {
          var base64Data = prod.img.split(',')[1];
          var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'produto.jpg');
          var img = SpreadsheetApp.newCellImage().setSourceUrl(blob).build();
          sheet.getRange(nextRow, 4).setValue(img);
        } catch (e) {
          Logger.log("Erro ao salvar imagem visual: " + e.message);
          // Tenta método alternativo
          try {
            var base64Data = prod.img.split(',')[1];
            var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'produto.jpg');
            sheet.insertImage(blob, 4, nextRow);
            sheet.getRange(nextRow, 4).setValue("IMG");
          } catch (e2) {
            sheet.getRange(nextRow, 4).setValue("✓");
          }
        }
      } else {
        sheet.getRange(nextRow, 4).setValue("-");
        sheet.getRange(nextRow, 6).setValue(""); // Sem foto
      }
      
      sheet.setRowHeight(nextRow, 80);
    }

    // --- AÇÃO: DELETAR ---
    else if (action === "delete") {
      var idToDelete = data.id;

      Logger.log("=== INICIANDO DELEÇÃO ===");
      Logger.log("ID recebido para deletar: " + idToDelete);
      Logger.log("Tipo do ID: " + typeof idToDelete);
      Logger.log("Payload completo recebido: " + JSON.stringify(data));

      var dataRange = sheet.getDataRange().getValues();
      Logger.log("Total de linhas na planilha: " + dataRange.length);

      var encontrado = false;

      // Procura o ID na coluna E (índice 4) de baixo para cima (mais rápido para itens recentes)
      for (var i = dataRange.length - 1; i >= 1; i--) {
        var idNaPlanilha = dataRange[i][4];

        // Log detalhado de cada comparação
        if (i <= 5) { // Loga apenas as últimas 5 linhas para não poluir
          Logger.log("Linha " + (i + 1) + " - ID na planilha: '" + idNaPlanilha + "' (tipo: " + typeof idNaPlanilha + ")");
          Logger.log("  Comparando com: '" + idToDelete + "'");
          Logger.log("  Match? " + (String(idNaPlanilha) === String(idToDelete)));
        }

        if (String(dataRange[i][4]) === String(idToDelete)) {
          encontrado = true;
          var rowNumber = i + 1;

          Logger.log("✅ ID ENCONTRADO na linha " + rowNumber + "!");
          Logger.log("Dados da linha: " + JSON.stringify(dataRange[i]));

          // PRIMEIRO: Remove todas as imagens dessa linha
          var images = sheet.getImages();
          Logger.log("Total de imagens na planilha: " + images.length);

          var imagensRemovidas = 0;
          for (var j = images.length - 1; j >= 0; j--) {
            var imgRow = images[j].getAnchorRow();
            Logger.log("Imagem " + j + " está na linha: " + imgRow);
            if (imgRow === rowNumber) {
              images[j].remove();
              imagensRemovidas++;
              Logger.log("  ✅ Imagem removida!");
            }
          }
          Logger.log("Total de imagens removidas: " + imagensRemovidas);

          // DEPOIS: Limpa todos os dados da linha
          sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).clearContent();
          Logger.log("✅ Conteúdo da linha limpo");

          // FINALMENTE: Deleta a linha vazia
          sheet.deleteRow(rowNumber);
          Logger.log("✅ Linha " + rowNumber + " deletada completamente!");

          break;
        }
      }

      if (!encontrado) {
        Logger.log("❌ ID NÃO ENCONTRADO na planilha!");
        Logger.log("IDs disponíveis na planilha:");
        for (var k = 1; k < dataRange.length; k++) {
          Logger.log("  Linha " + (k + 1) + ": ID = '" + dataRange[k][4] + "'");
        }
      }

      Logger.log("=== FIM DA DELEÇÃO ===");
    }

    return ContentService.createTextOutput(JSON.stringify({ "status": "success" })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": e.message })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
