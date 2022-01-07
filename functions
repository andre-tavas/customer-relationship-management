const PLANILHA = 'https://docs.google.com/spreadsheets/d/1A6Nsx0i5MMr0T_onWKrKKVIuALbBt9U75lcNTzBsk-M/edit#gid=462904825';
const FORM = 'https://docs.google.com/forms/d/e/1FAIpQLSfSrbOGqwGYEzI_v1zez6rBznD3y00orEM_iql79nloHTAw-w/viewform';

function sendPromoSegmentada(){
  var info = getInfoCampanha();
  var clientes = getClientesPromo(info.categoria,info.marca);

  clientes.map(function(row){
    response = GmailApp.sendEmail(row[2],info.assunto, info.texto)
  })
}

function SendFeedback(){
  var worksheet = SpreadsheetApp.openByUrl(PLANILHA);
  
  if(feedbackAtivado(worksheet)){
    var sheet = worksheet.getSheetByName('Respostas do formulário');
    var dados = sheet.getDataRange().getValues().slice(1);
    var infoFeedback = getInfoFeedback();
    var tempoFeedback = toDays(infoFeedback.tempo, infoFeedback.tempoType);
    // Data atual
    var date1 = new Date();

    var enviar = dados.filter(
      function(row){
        return difference(row[0],date1) == tempoFeedback
      }
    );
    
    enviar.map(function(row){
      response = GmailApp.sendEmail(row[2], infoFeedback.assunto , infoFeedback.texto)
      console.log(row[2])
    })
  }
}

function updateForm(){
  var form = FormApp.openByUrl(FORM);

  var categoria_question = form.getItemById(
      form.getItems()[4].getId()
    ).asCheckboxItem();

  var marcas_question = form.getItemById(
      form.getItems()[5].getId()
    ).asCheckboxItem();

  var categorias = getCategorias().map(function(x){return x[0];});
  var marcas = getMarcas().map(function(x){ return x[0];});

  categoria_question
      .setChoiceValues(categorias);

  marcas_question
      .setChoiceValues(marcas);
};


function confirmarEnvio(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
  'Confirmar envio?',
  'Enviar esta campanha para todos os clientes que colocaram ter interesse nesta categoria de produtos e nesta marca.', 
  ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
  sendPromoSegmentada();
  }
  else {}
};

function getInfoFeedback(){
  var sheet = SpreadsheetApp.openByUrl(PLANILHA).getSheetByName('Feedback');
  var tempo = sheet.getRange('C3').getValue();
  var tempoType = sheet.getRange('D3').getValue();
  var assunto = sheet.getRange('C5').getValue();
  var texto = sheet.getRange('C7').getValue();

  return {
    "tempo": tempo,
    "tempoType": tempoType,
    "assunto": assunto,
    "texto": texto
  }
}

function getCategorias(){
  var sheet = SpreadsheetApp.openByUrl(PLANILHA).getSheetByName('Configurações');
  var categorias = sheet.getRange('B4').activate().getDataRegion().getValues().slice(1);
  
  return categorias;
};

function getMarcas(){
  var sheet = SpreadsheetApp.openByUrl(PLANILHA).getSheetByName('Configurações');
  var marcas = sheet.getRange('D4').activate().getDataRegion().getValues().slice(1);
  
  return marcas;
};

function getInfoCampanha(){
  var sheet = SpreadsheetApp.openByUrl(PLANILHA).getSheetByName('Campanhas segmentadas');
  var categoria = sheet.getRange('C3').getValue();
  var marca = sheet.getRange('C5').getValue();
  var assunto = sheet.getRange('C7').getValue();
  var texto = sheet.getRange('C9').getValue();

  return {
    "categoria": categoria,
    "marca": marca,
    "assunto": assunto,
    "texto": texto
  }
};

function getClientesPromo(categoria,marca) {
  var sheet = SpreadsheetApp.openByUrl(PLANILHA).getSheetByName('Respostas do formulário');
  var dados = sheet.getDataRange().getValues().slice(1);
  var enviar = dados.filter(
    function(row){
      return row[5].search(categoria) >= 0 && row[6].search(marca) >= 0 
    }
  );
  return enviar;
};

/**
 * Converte para dias uma quantidade de semanas ou meses
 */
function toDays(tempo,tipo){
  if(tipo == 'Dias'){return tempo}
  if(tipo == 'Semanas'){return tempo * 7}
  else{return tempo * 30}
}

/**
 * Retorna a diferenca em dias entre as duas datas
 */
function difference(date1, date2) {  
  const date1utc = Date.UTC(date1.getFullYear(), date1.getMonth(), date1.getDate());
  const date2utc = Date.UTC(date2.getFullYear(), date2.getMonth(), date2.getDate());
  day = 1000*60*60*24;
  
  return Math.abs((date2utc - date1utc)/day)
}

/**
 * Confere se esta ativado para enviar o feedback
 */
function feedbackAtivado(worksheet){
  var sheet = worksheet.getSheetByName('Feedback');
  var status = sheet.getRange('I3').getValue();
  
  return status == 'Ativado'
}

