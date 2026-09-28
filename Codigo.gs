/**
 * EXEMPLO - MOTOR DE SIMULACAO DE CONTRATACAO
 * Apps Script vinculado a planilha. Gera o PDF "Estudo de Custo de Contratacao"
 * a partir da aba PDF_LAYOUT (que ja traz as formulas ligadas a aba SIMULADOR).
 *
 * INSTALACAO:
 * 1. Na planilha Google, abra Extensoes > Apps Script.
 * 2. Apague o conteudo padrao de Code.gs e cole este arquivo inteiro.
 * 3. Salve (icone de disquete). Volte para a planilha e recarregue a pagina.
 * 4. Um novo menu "Exemplo" aparecera na barra superior.
 * 5. Preencha a aba SIMULADOR e clique em Exemplo > Gerar PDF da Simulacao.
 * 6. Na primeira execucao o Google vai pedir autorizacao (normal, e o proprio
 *    script agindo na planilha e no Drive do usuario) - clique em Avancar >
 *    Permitir.
 */

var NOME_ABA_SIMULADOR = 'SIMULADOR';
var NOME_ABA_LAYOUT = 'PDF_LAYOUT';
var NOME_PASTA_DRIVE = 'Simulacoes de Contratacao - Exemplo';

/**
 * Cria o menu customizado quando a planilha e aberta.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Exemplo')
    .addItem('Gerar PDF da Simulacao', 'gerarPdfSimulacao')
    .addItem('Recalcular agora', 'recalcularAgora')
    .addToUi();
}

/**
 * Forca o recalculo da planilha (util antes de gerar o PDF, caso o usuario
 * tenha acabado de alterar algum campo do SIMULADOR).
 */
function recalcularAgora() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Planilha recalculada.');
}

/**
 * Funcao principal: exporta a aba PDF_LAYOUT como PDF, salva no Drive em uma
 * pasta dedicada e mostra ao usuario um link para abrir/baixar o arquivo.
 */
function gerarPdfSimulacao() {
  var ui = SpreadsheetApp.getUi();
  var planilha = SpreadsheetApp.getActiveSpreadsheet();

  var abaSimulador = planilha.getSheetByName(NOME_ABA_SIMULADOR);
  var abaLayout = planilha.getSheetByName(NOME_ABA_LAYOUT);

  if (!abaSimulador || !abaLayout) {
    ui.alert('Erro: nao encontrei as abas "' + NOME_ABA_SIMULADOR + '" e/ou "' +
             NOME_ABA_LAYOUT + '". Verifique se os nomes das abas nao foram alterados.');
    return;
  }

  // Validacoes minimas antes de gerar o PDF para o cliente.
  var empresa = abaSimulador.getRange('C4').getValue();
  var cargo = abaSimulador.getRange('C5').getValue();
  var custoTotal = abaSimulador.getRange('C65').getValue();

  if (!empresa) {
    ui.alert('Selecione uma empresa no campo "Empresa" (SIMULADOR!C4) antes de gerar o PDF.');
    return;
  }
  if (!custoTotal || isNaN(custoTotal) || custoTotal <= 0) {
    ui.alert('O custo total calculado esta zerado ou invalido. Confira os campos preenchidos na aba SIMULADOR antes de gerar o PDF.');
    return;
  }

  // Garante que todas as formulas estao atualizadas antes de exportar.
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  try {
    var blobPdf = exportarAbaComoPdf_(planilha, abaLayout);

    var dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Bahia', 'dd-MM-yyyy');
    var nomeArquivo = 'Simulacao de Custo - ' + empresa + ' - ' + (cargo || 'Sem cargo') + ' - ' + dataHoje + '.pdf';
    nomeArquivo = nomeArquivo.replace(/[\\/:*?"<>|]/g, '-');
    blobPdf.setName(nomeArquivo);

    var pasta = obterOuCriarPasta_(NOME_PASTA_DRIVE);
    var arquivo = pasta.createFile(blobPdf);

    ui.showModalDialog(
      HtmlService.createHtmlOutput(
        '<div style="font-family:Arial, sans-serif; padding: 8px;">' +
        '<p><b>PDF gerado com sucesso!</b></p>' +
        '<p>Empresa: ' + empresa + '<br>Cargo: ' + cargo + '</p>' +
        '<p><a href="' + arquivo.getUrl() + '" target="_blank">Abrir / baixar o PDF</a></p>' +
        '</div>'
      ).setWidth(380).setHeight(160),
      'Simulacao gerada'
    );
  } catch (erro) {
    ui.alert('Nao foi possivel gerar o PDF. Detalhe do erro: ' + erro.message);
    throw erro;
  }
}

/**
 * Exporta uma unica aba da planilha como PDF, usando a URL de exportacao do
 * Google Sheets (mesma rota usada pelo menu Arquivo > Fazer download > PDF),
 * mas restrita ao gid da aba desejada.
 */
function exportarAbaComoPdf_(planilha, aba) {
  var idPlanilha = planilha.getId();
  var gid = aba.getSheetId();

  var url = 'https://docs.google.com/spreadsheets/d/' + idPlanilha + '/export' +
    '?format=pdf' +
    '&gid=' + gid +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&top_margin=0.3' +
    '&bottom_margin=0.3' +
    '&left_margin=0.3' +
    '&right_margin=0.3' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenumbers=false' +
    '&horizontal_alignment=CENTER';

  var resposta = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (resposta.getResponseCode() !== 200) {
    throw new Error('Falha ao exportar PDF (HTTP ' + resposta.getResponseCode() + ').');
  }

  return resposta.getBlob();
}

/**
 * Retorna a pasta de destino no Drive do usuario, criando-a caso nao exista
 * ainda (evita espalhar PDFs soltos na raiz do Drive).
 */
function obterOuCriarPasta_(nomePasta) {
  var pastas = DriveApp.getFoldersByName(nomePasta);
  if (pastas.hasNext()) {
    return pastas.next();
  }
  return DriveApp.createFolder(nomePasta);
}