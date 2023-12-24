function alertadeContrato() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📚Registros'); // Substitua 'NOME_DA_PLANILHA' pelo nome da sua planilha
  var ultimaLinha = planilha.getLastRow();
  
  contratosAVencer= []
  // Começando pela segunda linha, já que a primeira linha geralmente contém cabeçalhos
  for (var i = 2; i <= ultimaLinha; i++) {
    var linha = planilha.getRange(i, 1, 1, planilha.getLastColumn()).getValues()[0];
    
    // Armazenando o conteúdo de cada célula em variáveis diferentes
    var id = linha[0]; // Conteúdo da primeira célula da linha
    var n_Os = linha[1];
    var carimboDataHora = linha[2];
    var processo = linha[3];
    var numeroIdentificador = linha[4];
    var contratada = linha[5]
    var numeroDaModalidadedeLicitacao = linha[5];
    var objeto = linha[6];
    var gerencia = linha[7]
    var valorContratado = linha[11];
    var iniciodaVigencia = linha[7];
    var prazoContratual = linha[8];
    var prazodeExecucaodoObjeto = linha[9];
    var gestor = linha[15];
    var ambito = linha[11];
    var situacaodoObjeto = linha[12];
    var arquivodeAcompanhamento = linha[13]; // Conteúdo da segunda célula da linha
    var email = linha[19];
    var valorExecutadoporcentagem = linha[15];
    var naturezadoContrato = linha[16];
    var codigoeDescricaodoInvest = linha[17];
    var codigoeDescricaodoCusteio = linha[18];
    var valorVigente = linha[19];
    var valorExecutado = linha[20];
    var porcentagemFinanceira = linha[21];
    var dataProposta = linha[22];
    var dataEncerramento = linha[23];
    var diasVigentes = linha[29];
    var situacao = linha[25]; // Conteúdo da segunda célula da linha
   
    //console.log(diasVigentes)
    if(parseInt(diasVigentes)>0&&parseInt(diasVigentes)<60){
      dadosdoContrato = [gerencia, gestor, processo, contratada, objeto, valorContratado, diasVigentes, email]
      contratosAVencer.push(dadosdoContrato)
    }
    // Adicione mais variáveis conforme necessário para cada célula da linha
    
    // Faça o que desejar com o conteúdo das células (por exemplo, log no Console)
    //Logger.log("Conteúdo da célula 1: " + linha);
    //Logger.log("Conteúdo da célula 1: " + processo);
    //Logger.log("Conteúdo da célula 2: " + diasVigentes);
    //console.log(contratosAVencer)
    // Adicione mais logs ou operações conforme necessário
  }
  var emailsAgrupados = {};

  // Iterar sobre o array original
  for (var i = 0; i < contratosAVencer.length; i++) {
    var elemento = contratosAVencer[i];
    var email = elemento.pop(); // Remover e obter o email do final do array
    
    if (!emailsAgrupados[email]) {
      // Se o email ainda não existir no objeto, criar um novo array com esse email como chave
      emailsAgrupados[email] = [elemento];
    } else {
      // Se o email já existir, adicionar o elemento ao array correspondente ao email
      emailsAgrupados[email].push(elemento);
    }
  }

  // Converter o objeto para um array de arrays

  for (var key in emailsAgrupados) {
    if (emailsAgrupados.hasOwnProperty(key)) {
      notificar(key, emailsAgrupados[key]);
    }
  }
}

function notificar(email, dados) {
  
  dados.sort(function(a, b) {
  return a[a.length - 1] - b[b.length - 1];
  });

  var email = email
  console.log(email)
  var assunto = 'SISTEMA DE AVISOS DE CONTRATOS - DOM';
  var mensagem = '<b>FIQUE ATENTO: EXISTEM CONTRATOS A VENCER CADASTRADOS SOB ESSE EMAIL.</b><br><br>' +
    '<table style="width: 100%; border-collapse: collapse; text-align: center;">' +
    '<tr><th style="border: 1px solid #000;">GERÊNCIA</th><th style="border: 1px solid #000;">GESTOR</th><th style="border: 1px solid #000;">PROCESSO</th><th style="border: 1px solid #000;">CONTRATADA</th><th style="border: 1px solid #000;">OBJETO</th><th style="border: 1px solid #000;">VALOR CONTRATADO (R$)</th><th style="border: 1px solid #000;">DIAS VIGENTES</th></tr>';

  // Loop para percorrer o array e construir a tabela
  for (var i = 0; i < dados.length; i++) {
    mensagem += '<tr style="border: 1px solid #000;">'; // Início da linha

    for (var j = 0; j < dados[i].length; j++) {
      mensagem += '<td style="border: 1px solid #000;">' + dados[i][j] + '</td>'; // Adicionando cada valor como célula da tabela
    }

    mensagem += '</tr>'; // Fim da linha
  }

  // Fim da string HTML
  mensagem += '</table><br>' +
    'Link da planilha: <a href="https://docs.google.com/spreadsheets/d/18QKTxTshjcrT-ZUv2npd26klfkcIk1VEFoj-B9vvspQ/edit#gid=1930517805">DOM - Cadastro de Contratos</a>' +
    '<br><br>Esse email é reenviado todos os dias quando um contrato possui menos de 60 dias de vigência';




  console.log(mensagem); // Saída: a string HTML gerada

  // var libera = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AVISOS-CONTRATOS').getRange('B5').getValue(); // verificador se o email deve ser enviado
  var libera = 'S';

  if (libera === "S") {
    MailApp.sendEmail({
      to: email, // pega os dados da variável email
      subject: assunto, // pega os dados da variável assunto
      htmlBody: mensagem, // a mensagem pode ser feita em HTML ou em texto comum dentro da célula da planilha. Esta opção da função permite incorporar HTML no e-mail
      name: "SISTEMA DE AVISOS DE VENCIMENTOS DE CONTRATOS" // nome do remetente de envio do e-mail
    });
  }
}
