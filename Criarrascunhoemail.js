function Criarrascunho() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Deseja criar um e-mail rascunho para pré-visualizar?", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Planilha1');
    var startRow = 7; // início da linha de dados
    var numRows = 1; // número total de linhas de dados (apenas a primeira linha)
    var dataRange = sheet.getRange(startRow, 1, numRows, 12); // intervalo de dados (colunas A até L, começando da linha 7)
    var data = dataRange.getValues(); // valores das células do intervalo de dados
    var subject = 'COBRANÇA EMPENHOS';
    var body = '<html><body>Olá, %s !<br><br>Identificamos empenho(s) emitido(s) pela Pró-reitoria de Pesquisa (CNPJ: 28.523.215/0033-93) com saldo(s) em aberto e gostaríamos de confirmar com vocês a situação deste(s).<br><br>Segue abaixo quadro resumo deste(s) empenho(s) para que vocês nos informem, caso ainda haja interesse em atender nosso pedido, uma programação de entrega do mesmo:<br><br>EMPENHO: %s<br>VALOR: R$ %s<br>PROGRAMAÇÃO DE ENTREGA (data)<br><br>Observação: No caso de materiais entregues/ serviços prestados, solicitamos o envio da nota fiscal correspondente e um canhoto com assinatura do recebedor, para que possamos verificar a situação e prosseguir com os trâmites.</body></html>';
    
    var row = data[0];
    var nomeFornecedor = row[3]; // valor da coluna D
    var empenho = row[2]; // valor da coluna C
    var valor = row[10]; // valor da coluna k
    var emailAddress = row[11]; // valor da coluna L

    // Verifica se as células estão vazias antes de criar o rascunho do e-mail
    if (empenho && valor && emailAddress) {
      var message = body.replace('%s', nomeFornecedor).replace('%s', empenho).replace('%s', valor); // substitui os valores no corpo do e-mail
      
      // Cria um rascunho do e-mail em vez de enviá-lo
      GmailApp.createDraft(emailAddress, subject, '', {htmlBody: message});
      
      ui.alert("Rascunho de e-mail criado com sucesso!");
    } else {
      ui.alert("Não foi possível criar o rascunho do e-mail porque faltam informações na planilha.");
    }
  }
}
