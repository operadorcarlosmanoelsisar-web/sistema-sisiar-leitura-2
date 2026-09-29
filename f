// ==============================================================================
// 1. PONTO DE ENTRADA DO SISTEMA
// ==============================================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SaneVeneza SISAR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==============================================================================
// 2. BANCO DE DADOS E IDS GLOBAIS DO DRIVE
// ==============================================================================

// PASTA PRINCIPAL DOS MORADORES ("Usuários Ativos (SISAR)")
var ID_PASTA_USUARIOS_ATIVOS = "137LrCZeAfanepkHsa5NZ81qlCVFBYiEe";

// PASTAS DE FECHAMENTO (PASTA RAIZ: 1-FOLHAR DE LEITURA)
var ID_PASTA_FOLHA_PDF    = "1JahsVTSALFIl1L6kyDJ4Qt9vwRCY6PtH";  // 1-FOLHA DE LEITURAS 
var ID_PASTA_FOTO_MACRO   = "1KmsObUI49ZM8qB1nyib9g9oDBG_Cfypo";  // 2-LEITURA DO MACRO MEDIDOR 
var ID_PASTA_FOTO_ENERGIA = "1-wD6HdKloRACSvnIC2unGCpt3-q0WuTg";  // 3-FOTO DO MEDIDOR DE ENERGIA 

// MAPEAMENTO DAS 12 PLANILHAS MENSAIS (COM A ABA "ESPELHO")
var PLANILHAS_MESES = {
  "JANEIRO": "1Zj0sS9IzGBkr0z_4CWEiWdJ3RhX6gTNC8QKkxYBOTXQ",
  "FEVEREIRO": "1edBnG28JOpyBnV806TAqRuLUpP7fRWBWFSAPsb4wt2Y",
  "MARÇO": "1vGVi1jPLY7JYYD6kca95JuClyv4CRltCUjKVlLzYVHg",
  "ABRIL": "1LRkheFZSTVz2lTE8YKHt2SY_MRYw-MJA3-quYZuwKH8",
  "MAIO": "1k11DP_d9Og8Ucs2DvZ19O1asarTNvMyLfzDmZCfAJ8I",
  "JUNHO": "1Rqn-Vt9-kUE3Wkhr0QG_vxYU60YXiMl1NfTLJpIQ-8I",
  "JULHO": "1bCI7ux5kg89UHVyjpwFLEx_UrwIlhLjjsDMUhhiLweI",
  "AGOSTO": "1BuxvW2XMirwbQ-T-uMelXSz4uT1kxtoELKgQY9C8Gs0",
  "SETEMBRO": "1az-WRArVD1PGuODe0UOpUXVRkXyYy8_kcuFTLM3B4Rc",
  "OUTUBRO": "1aSfknpky1X-zlfgQaq2yGABbeyxoUAePq3dnmLPSsKY",
  "NOVEMBRO": "1eOohSbmAwBpw9mnu_4uuIY2UCULLs6HPCB4xeVVi7Eg",
  "DEZEMBRO": "1aCeKDctNrQxkcpUvFhPAy-iVidDT4Vp9eNoZh5vfxFM"
};

// ==============================================================================
// 3. CARREGAMENTO DOS MORADORES E LEITURA ANTERIOR (AO ABRIR O SITE)
// ==============================================================================
function obterDadosSistema(mesEscolhido) {
  try {
    var listaMeses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    var mesAlvo = (mesEscolhido || listaMeses[new Date().getMonth()]).toUpperCase().trim();
    var idPlanilhaAtual = PLANILHAS_MESES[mesAlvo];
   
    // Abre a planilha do mês escolhido no topo do site
    var ssAtual = idPlanilhaAtual ? SpreadsheetApp.openById(idPlanilhaAtual) : SpreadsheetApp.getActiveSpreadsheet();
    var abaAtual = ssAtual.getSheetByName("ESPELHO") || ssAtual.getSheetByName("Espelho") || ssAtual.getSheetByName("espelho") || ssAtual.getSheets()[0];

    // Descobre o mês anterior para resgatar a leitura anterior e calcular consumo
    var indiceAtual = listaMeses.indexOf(mesAlvo);
    var indiceAnterior = indiceAtual > 0 ? (indiceAtual - 1) : 11;
    var mesAnteriorNome = listaMeses[indiceAnterior];
    var idPlanilhaAnterior = PLANILHAS_MESES[mesAnteriorNome];

    var mapaLeiturasAnteriores = {};
    if (idPlanilhaAnterior) {
      try {
        var ssAnt = SpreadsheetApp.openById(idPlanilhaAnterior);
        var abaAnt = ssAnt.getSheetByName("ESPELHO") || ssAnt.getSheetByName("Espelho") || ssAnt.getSheetByName("espelho") || ssAnt.getSheets()[0];
        var dadosAnt = abaAnt.getRange("G20:H200").getValues();
        for (var a = 0; a < dadosAnt.length; a++) {
          var matAnt = dadosAnt[a][0] ? dadosAnt[a][0].toString().trim() : "";
          var valAnt = dadosAnt[a][1];
          if (matAnt && valAnt !== "" && valAnt !== null && !isNaN(valAnt)) {
            mapaLeiturasAnteriores[matAnt] = Number(valAnt);
          }
        }
      } catch (eAnt) {
        Logger.log("Aviso ao buscar mês anterior: " + eAnt.toString());
      }
    }

    // Lê a planilha atual: C: Nome, G: Matrícula, H: Leitura, I: Status, J: Foto
    var dados = abaAtual.getRange("C20:J200").getValues();
    var notas = abaAtual.getRange("H20:H200").getNotes();
    var lista = [];

    for (var i = 0; i < dados.length; i++) {
      var nome = dados[i][0] ? dados[i][0].toString().trim() : "";
      var matricula = dados[i][4] ? dados[i][4].toString().trim() : ("MAT-" + (i + 1));
      var leitura = (dados[i][5] !== "" && dados[i][5] !== null && dados[i][5] !== undefined) ? Number(dados[i][5]) : null;
      var status = dados[i][6] ? dados[i][6].toString().trim().toUpperCase() : "";
      var temFoto = (dados[i][7] !== "" && dados[i][7] !== null && dados[i][7] !== undefined);
      var notaCavalete = notas[i] ? (notas[i][0] || "") : "";

      // Ignora linha sem nome ou moradores cortados/inativos
      if (nome === "") continue;
      if (status.indexOf("CORTAD") !== -1 || status.indexOf("INATIV") !== -1) continue;

      var situacaoCavalete = "NORMAL";
      if (notaCavalete.indexOf("SITUAÇÃO:") !== -1) {
        situacaoCavalete = "AVARIA";
      }

      var leituraAnterior = mapaLeiturasAnteriores[matricula] !== undefined ? mapaLeiturasAnteriores[matricula] : 0;

      lista.push({
        linha: i + 20,
        matricula: matricula,
        nome: nome,
        leituraAnterior: leituraAnterior,
        leitura: leitura,
        status: (leitura !== null || notaCavalete !== "") ? "concluido" : "pendente",
        situacaoCavalete: situacaoCavalete,
        detalhesReparo: notaCavalete,
        foto: temFoto
      });
    }

    return {
      sucesso: true,
      mes: mesAlvo,
      moradores: lista
    };

  } catch (err) {
    Logger.log("Erro em obterDadosSistema: " + err.toString());
    return { sucesso: false, erro: err.toString(), moradores: [] };
  }
}

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

// ==============================================================================
// PARTE 4: FECHAMENTO GERAL COMPLETO E UNIFICADO
// ==============================================================================

/**
 * Função auxiliar que organiza os arquivos nas pastas do Google Drive:
 * - Macro: pasta "FOTO MACRO MEDIDOR [MÊS]" > [ANO]
 * - Folha e Energia: pasta "[MÊS]" > [ANO]
 */
function salvarArquivoFechamentoDrive(base64Data, nomeArquivo, idPastaRaiz, nomeMes, ehMacro) {
  var pastaRaiz = DriveApp.getFolderById(idPastaRaiz);

  function obterOuCriar(pastaPai, nomePasta) {
    var busca = pastaPai.getFoldersByName(nomePasta);
    return busca.hasNext() ? busca.next() : pastaPai.createFolder(nomePasta);
  }

  // 1º Nível: Pasta do Mês
  var nomePastaMes = ehMacro ? ("FOTO MACRO MEDIDOR " + nomeMes) : nomeMes;
  var pastaMes = obterOuCriar(pastaRaiz, nomePastaMes);

  // 2º Nível: Subpasta do Ano correspondente ao envio
  var anoUpload = new Date().getFullYear().toString();
  var pastaAno = obterOuCriar(pastaMes, anoUpload);

  // Decodifica a imagem/PDF em Base64
  var base64Limpo = base64Data.indexOf("base64,") !== -1 ? base64Data.split(',')[1] : base64Data;
  var contentType = "image/jpeg";
  if (base64Data.indexOf(":") !== -1 && base64Data.indexOf(";") !== -1) {
    contentType = base64Data.split(',')[0].match(/:(.*?);/)[1];
  }

  var decoded = Utilities.base64Decode(base64Limpo.replace(/\s/g, ''));
  var blob = Utilities.newBlob(decoded, contentType, nomeArquivo);
  var arquivoCriado = pastaAno.createFile(blob);

  return arquivoCriado.getUrl();
}

/**
 * Função principal do Fechamento: salva fotos, leituras e cabeçalhos em um só fluxo
 */
function processarFechamentoFolhaLeitura(dadosEnvio) {
  if (!dadosEnvio) {
    throw new Error("Nenhum dado foi enviado pela interface.");
  }

  var listaMeses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  var mesAlvo = (dadosEnvio.nomeMes || dadosEnvio.mesAno || listaMeses[new Date().getMonth()]).toUpperCase().trim().split(" ")[0];
  var anoAtual = new Date().getFullYear().toString();

  var idPlanilha = PLANILHAS_MESES[mesAlvo];
  var ss = idPlanilha ? SpreadsheetApp.openById(idPlanilha) : SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("ESPELHO") || ss.getSheetByName("Espelho") || ss.getSheetByName("espelho") || ss.getSheets()[0];

  var enviouFolha = dadosEnvio.fotoFolhaLeitura && dadosEnvio.fotoFolhaLeitura.length > 50;
  var enviouEnergia = dadosEnvio.fotoMedidorEnergia && dadosEnvio.fotoMedidorEnergia.length > 50;
  var enviouMacro = dadosEnvio.fotoMacroMedidor && dadosEnvio.fotoMacroMedidor.length > 50;

  // 1. Folha de Leitura -> Pasta Folha > [MÊS] > [ANO] -> Link em D9
  if (enviouFolha) {
    var extFolha = dadosEnvio.fotoFolhaLeitura.indexOf("application/pdf") !== -1 ? ".pdf" : ".jpg";
    var nomeArqFolha = "Folha_Leitura_" + mesAlvo + "_" + anoAtual + extFolha;
    var linkFolha = salvarArquivoFechamentoDrive(dadosEnvio.fotoFolhaLeitura, nomeArqFolha, ID_SUBPASTA_FOLHA_LEITURA, mesAlvo, false);
    aba.getRange("D9").setFormula('=HYPERLINK("' + linkFolha + '"; "CLIQUE AQUI")');
  }

  // 2. Foto Medidor de Energia -> Pasta Energia > [MÊS] > [ANO] -> Link em D13
  if (enviouEnergia) {
    var nomeArqEnergia = "Medidor_Energia_" + mesAlvo + "_" + anoAtual + ".jpg";
    var linkEnergia = salvarArquivoFechamentoDrive(dadosEnvio.fotoMedidorEnergia, nomeArqEnergia, ID_SUBPASTA_FOTO_ENERGIA, mesAlvo, false);
    aba.getRange("D13").setFormula('=HYPERLINK("' + linkEnergia + '"; "CLIQUE AQUI")');
  }

  // 3. Foto Macro Medidor -> Pasta Macro > FOTO MACRO MEDIDOR [MÊS] > [ANO] -> Link em H13
  if (enviouMacro) {
    var nomeArqMacro = "Macro_Medidor_" + mesAlvo + "_" + anoAtual + ".jpg";
    var linkMacro = salvarArquivoFechamentoDrive(dadosEnvio.fotoMacroMedidor, nomeArqMacro, ID_SUBPASTA_FOTO_MACRO, mesAlvo, true);
    aba.getRange("H13").setFormula('=HYPERLINK("' + linkMacro + '"; "CLIQUE AQUI")');
  }

  // 4. Leituras Numéricas -> F15 (Energia) e I15 (Macro)
  if (dadosEnvio.leituraEnergia !== undefined && dadosEnvio.leituraEnergia !== null && dadosEnvio.leituraEnergia !== "") {
    aba.getRange("F15").setValue(dadosEnvio.leituraEnergia);
  }
  if (dadosEnvio.leituraMacro !== undefined && dadosEnvio.leituraMacro !== null && dadosEnvio.leituraMacro !== "") {
    aba.getRange("I15").setValue(dadosEnvio.leituraMacro);
  }

  // 5. Carimbo de Fechamento: Título em A1 e Data em B3
  if (enviouFolha || enviouMacro) {
    var agora = new Date();
    var dataApenas = Utilities.formatDate(agora, "America/Fortaleza", "dd/MM/yyyy");
    var novoTitulo = "LEITURA DO MÊS DE " + mesAlvo + "\n" + dataApenas;
    aba.getRange("A1").setValue(novoTitulo);
    aba.getRange("B3").setValue(dataApenas);
  }

  SpreadsheetApp.flush();
  return obterFechamentoMes(mesAlvo);
}

// 6. Consulta de status para preencher o card verde e lembrar os dados salvos
function obterFechamentoMes(mesEscolhido) {
  try {
    var listaMeses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    var mesAlvo = (mesEscolhido || listaMeses[new Date().getMonth()]).toUpperCase().trim().split(" ")[0];
    var idPlanilha = PLANILHAS_MESES[mesAlvo];
    var ss = idPlanilha ? SpreadsheetApp.openById(idPlanilha) : SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("ESPELHO") || ss.getSheetByName("Espelho") || ss.getSheetByName("espelho") || ss.getSheets()[0];

    var valEnergia = aba.getRange("F15").getValue();
    var valMacro = aba.getRange("I15").getValue();
    var formFolha = aba.getRange("D9").getFormula() || aba.getRange("D9").getValue();
    var formEnergia = aba.getRange("D13").getFormula() || aba.getRange("D13").getValue();
    var formMacro = aba.getRange("H13").getFormula() || aba.getRange("H13").getValue();
    var dataFechamento = aba.getRange("B3").getValue();

    var temDados = (valEnergia !== "" && valEnergia !== null) ||
                   (valMacro !== "" && valMacro !== null) ||
                   (formFolha !== "" && formFolha !== null) ||
                   (formEnergia !== "" && formEnergia !== null) ||
                   (formMacro !== "" && formMacro !== null);

    return {
      sucesso: true,
      mes: mesAlvo,
      temFechamento: temDados,
      energia: valEnergia || "",
      macro: valMacro || "",
      temFolha: !!formFolha,
      temFotoEnergia: !!formEnergia,
      temFotoMacro: !!formMacro,
      dataHora: dataFechamento ? dataFechamento.toString() : ""
    };
  } catch(e) {
    return { sucesso: false, erro: e.toString() };
  }
}
