// ==============================================================================
// PARTE 3: GRAVAÇÃO DO MORADOR + ALIMENTAÇÃO AUTOMÁTICA DE AVARIAS
// ==============================================================================
function salvarLeituraApp(pacote) {
  try {
    if (!pacote) {
      return { sucesso: false, erro: "Pacote de dados vazio." };
    }

    var listaMeses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    var mesPadrao = listaMeses[new Date().getMonth()];
    var mesAlvo = (pacote.mes || mesPadrao).toUpperCase().trim();
    var idPlanilha = PLANILHAS_MESES[mesAlvo];
    var ss = idPlanilha ? SpreadsheetApp.openById(idPlanilha) : SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("ESPELHO") || ss.getSheetByName("Espelho") || ss.getSheetByName("espelho") || ss.getSheets()[0];

    // Localiza a linha correta do morador
    var linha = parseInt(pacote.linha, 10);
    var nomesTabela = aba.getRange("C20:C200").getValues();
    var matriculasTabela = aba.getRange("G20:G200").getValues();
    var nomeBusca = (pacote.nome || "").toString().trim().toUpperCase();
    var matBusca = (pacote.matricula || "").toString().trim().toUpperCase();

    var precisaBuscar = isNaN(linha) || linha < 20 || linha > 200;
    if (!precisaBuscar) {
      var nomeNaLinha = (nomesTabela[linha - 20][0] || "").toString().trim().toUpperCase();
      if (nomeBusca && nomeNaLinha !== nomeBusca && nomeNaLinha.indexOf(nomeBusca) === -1 && nomeBusca.indexOf(nomeNaLinha) === -1) {
        precisaBuscar = true;
      }
    }

    if (precisaBuscar) {
      for (var k = 0; k < nomesTabela.length; k++) {
        var n = (nomesTabela[k][0] || "").toString().trim().toUpperCase();
        var m = (matriculasTabela[k][0] || "").toString().trim().toUpperCase();
        if ((matBusca && m === matBusca) || (nomeBusca && (n === nomeBusca || n.indexOf(nomeBusca) !== -1 || nomeBusca.indexOf(n) !== -1))) {
          linha = k + 20;
          break;
        }
      }
    }

    if (isNaN(linha) || linha < 20) linha = 20;
    var nomeRealMorador = (pacote.nome || aba.getRange("C" + linha).getValue() || "Morador").toString().trim();

    // 1. Grava o valor da leitura na Coluna H
    var celulaH = aba.getRange("H" + linha);
    if (pacote.leitura !== undefined && pacote.leitura !== null && pacote.leitura !== "") {
      var txtNum = pacote.leitura.toString().replace(/\s+/g, '').replace(',', '.');
      var num = parseFloat(txtNum);
      celulaH.setValue(isNaN(num) ? pacote.leitura : num);
    }

    // 2. Grava o status na Coluna J (SITUAÇÃO SISTEMA) da aba ESPELHO
    var celulaJ = aba.getRange("J" + linha);
    if (pacote.situacaoCavalete && pacote.situacaoCavalete !== 'NORMAL') {
      var textoStatus = pacote.situacaoCavalete + (pacote.detalhesReparo ? " - " + pacote.detalhesReparo : "");
      celulaJ.setValue(textoStatus);
      celulaH.setNote("SITUAÇÃO: " + textoStatus);
    } else {
      celulaJ.setValue("NORMAL");
    }

    var urlFoto = "";

    // 3. Salva a foto na pasta do morador no Drive e grava link na Coluna K
    if (pacote.fotoBase64 && pacote.fotoBase64.length > 50) {
      try {
        var pastaPrincipal = DriveApp.getFolderById(ID_PASTA_USUARIOS_ATIVOS);

        function obterOuCriar(pastaPai, nomePasta) {
          var busca = pastaPai.getFoldersByName(nomePasta);
          return busca.hasNext() ? busca.next() : pastaPai.createFolder(nomePasta);
        }

        var pastaMorador = obterOuCriar(pastaPrincipal, nomeRealMorador);
        var anoAtual = new Date().getFullYear();
        var pastaAno = obterOuCriar(pastaMorador, "FOTO HIDRÔMETRO " + anoAtual);
        var pastaMes = obterOuCriar(pastaAno, mesAlvo);

        var base64Dados = pacote.fotoBase64;
        var mimeType = "image/jpeg";
        if (base64Dados.indexOf("base64,") !== -1) {
          var partes = base64Dados.split("base64,");
          if (partes[0].indexOf(":") !== -1 && partes[0].indexOf(";") !== -1) {
            mimeType = partes[0].split(":")[1].split(";")[0];
          }
          base64Dados = partes[1];
        }

        base64Dados = base64Dados.replace(/\s/g, '');
        var bytes = Utilities.base64Decode(base64Dados);
        var nomeSanitizado = nomeRealMorador.replace(/[\/\\?%*:|"<>]/g, '_');
        var nomeArquivo = mesAlvo + "_" + anoAtual + "_" + nomeSanitizado + ".jpg";

        var blob = Utilities.newBlob(bytes, mimeType, nomeArquivo);
        var arquivoCriado = pastaMes.createFile(blob);
        urlFoto = arquivoCriado.getUrl();

        aba.getRange("K" + linha).setFormula('=HYPERLINK("' + urlFoto + '"; "IMAGEM")');

      } catch (errFoto) {
        Logger.log("Erro ao salvar foto do morador: " + errFoto.toString());
        aba.getRange("K" + linha).setValue("ERRO FOTO");
      }
    }

    // 4. REGISTRO AUTOMÁTICO NA ABA 'SITUACAO_SISTEMA' (PAINEL DE OCORRÊNCIAS)
    if (pacote.situacaoCavalete && pacote.situacaoCavalete !== 'NORMAL') {
      try {
        var abaAvarias = ss.getSheetByName("SITUACAO_SISTEMA");
        if (!abaAvarias) {
          abaAvarias = ss.insertSheet("SITUACAO_SISTEMA");
          abaAvarias.appendRow(["Data", "Matrícula", "Nome", "Problema", "Detalhes", "Foto Avaria", "Status", "Data Reparo", "Foto Reparo"]);
          abaAvarias.getRange("A1:I1").setFontWeight("bold").setBackground("#fef3c7");
        }

        var matriculaMorador = pacote.matricula || aba.getRange("G" + linha).getValue() || ("MAT-" + linha);
        var dataHoje = Utilities.formatDate(new Date(), "America/Fortaleza", "dd/MM/yyyy HH:mm");

        abaAvarias.appendRow([
          dataHoje,
          matriculaMorador,
          nomeRealMorador,
          pacote.situacaoCavalete,
          pacote.detalhesReparo || "",
          urlFoto ? '=HYPERLINK("' + urlFoto + '"; "VER FOTO")' : "",
          "PENDENTE",
          "",
          ""
        ]);
      } catch (errAvaria) {
        Logger.log("Erro ao alimentar aba SITUACAO_SISTEMA: " + errAvaria.toString());
      }
    }

    SpreadsheetApp.flush();
    return { sucesso: true, linkFoto: urlFoto };

  } catch (err) {
    Logger.log("Erro geral ao salvar leitura: " + err.toString());
    return { sucesso: false, erro: err.toString() };
  }
}
