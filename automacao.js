const diferencaDeDiasPrimeiroParaSegundoEmail = 5
const diferencaDeDiasSegundoParaTerceiroEmail = 5


const personalizarEmail  = (corpoEmail,funcionario = '',cargo = '' , empresa = '' )=>{

    corpoEmail  =  corpoEmail.replace('funcionario',funcionario)
    corpoEmail  =  corpoEmail.replace('cargo',cargo)
    corpoEmail  =  corpoEmail.replace('empresa',empresa)
    return corpoEmail
    }

function personalizaEEnviaEmail(email,assunto,corpo,nome,cargo,empresa){
      corpo = personalizarEmail(corpo,funcionario = nome ,cargo,empresa)
    
      MailApp.sendEmail({
      to:email,
      subject: assunto,
      htmlBody: corpo
    })
}

function diasPassados(data) {
    // Garantir que a data seja no formato ISO "YYYY-MM-DD"
    let dataAtual = new Date();
    let dataInput = new Date(data); // A data fornecida deve estar no formato "YYYY-MM-DD"
    
    // Garantir que ambas as datas estejam no formato de "YYYY-MM-DD"
    let dataAtualFormatada = dataAtual.toISOString().split('T')[0]; // Ex: "2025-01-16"
    let dataInputFormatada = dataInput.toISOString().split('T')[0]; // Ex: "2025-01-10"
    
    // Criar novas instâncias de Date com as datas formatadas corretamente
    let dataAtualFinal = new Date(dataAtualFormatada);
    let dataInputFinal = new Date(dataInputFormatada);

    // Zerar as horas, minutos, segundos e milissegundos para garantir comparação apenas de dias
    dataAtualFinal.setHours(0, 0, 0, 0);
    dataInputFinal.setHours(0, 0, 0, 0);
    
    // Calcular a diferença em milissegundos
    let diferencaMilissegundos = dataAtualFinal - dataInputFinal;
    
    // Converter de milissegundos para dias
    let diferencaDias = Math.floor(diferencaMilissegundos / (1000 * 60 * 60 * 24));
    return diferencaDias;
}

function sendEmails() {
  const sheet  =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListaDeContatosEconodata');
  const data = sheet.getDataRange().getValues()
  const headers = data.shift()
  const assuntoPrimeiro = 'Benefício Gratúito para sua Equipe'
  const assuntoSegundo  = 'assunto para o segundo e-mail'
  const assuntoTerceiro  =  'assunto para o terceiro e-mail'
  let   uniqueEmails = []
  let quantidadeDePrimeiroEmailEnviado  =  0
  let quantidadeDeSegundoEmailEnviado   =  0
  let quantidadeDeTerceiroEmailEnviado   =  0
  let terminouComSucesso                =  true
  let rowIndex = 0  // essa variável quero usar para mudar o valor da celula desejada.
  data.every((row)=>{
    const email = row[headers.indexOf('Email')];
    const nome = row[headers.indexOf('Nome')];
    const cargo = row[headers.indexOf('Cargo')];
    const empresa = row[row.indexOf('RAZÃO SOCIAL')]
    let dataAtual = new Date()
    rowIndex += 1

    // falta implementar a personalização dos e-mails

    if(email){
      if(!uniqueEmails.includes(email)){
      if(!row[headers.indexOf('primeiro E-MAIL ENVIADO?')]){ // não enviou o primeiro e-mail
      const corpo = row[headers.indexOf('Corpo primeiro e-mail')]
      try{
        personalizaEEnviaEmail(email,assuntoPrimeiro,corpo,nome,cargo,empresa)
        uniqueEmails.push(email)
        //row[headers.indexOf('segundo email enviado?')] = dataAtual
        sheet.getRange(rowIndex, headers.indexOf('primeiro E-MAIL ENVIADO?') + 1).setValue(dataAtual);
        quantidadeDePrimeiroEmailEnviado += 1

        return true
      }catch (e) {
        console.error("Erro ao enviar e-mail:", e);
        terminouComSucesso = false;
        return false;
      }

      } else if(row[headers.indexOf('primeiro E-MAIL ENVIADO?')] &&
              !row[headers.indexOf('segundo email enviado?')]  && 
              diasPassados(row[headers.indexOf('primeiro E-MAIL ENVIADO?')]) >= diferencaDeDiasPrimeiroParaSegundoEmail){ // não enviou o segundo e-mail e ja passou o tempo para mandar
        const corpo = row[headers.indexOf('Corpo segundo e-mail')]
        try{
          personalizaEEnviaEmail(email,assuntoSegundo,corpo,nome,cargo,empresa)
          uniqueEmails.push(email)
          //row[headers.indexOf('segundo email enviado?')] = dataAtual
          sheet.getRange(rowIndex, headers.indexOf('segundo email enviado?') + 1).setValue(dataAtual);
        
          quantidadeDeSegundoEmailEnviado  +=  1
          return true
        }catch (e) {
          console.error("Erro ao enviar e-mail:", e);
          terminouComSucesso = false;
          return false;
        }
      }else if(row[headers.indexOf('primeiro E-MAIL ENVIADO?')] &&
             row[headers.indexOf('segundo email enviado?')]   && 
             !row[headers.indexOf('terceiro email enviado?')] && 
             diasPassados(row[headers.indexOf('segundo email enviado?')]) >= diferencaDeDiasSegundoParaTerceiroEmail){ // não enviou o terceiro e-mail e ja passou o tempo para mandar
        const corpo = row[headers.indexOf('Corpo terceiro e-mail')]
        try{
          personalizaEEnviaEmail(email,assuntoTerceiro,corpo,nome,cargo,empresa)
          uniqueEmails.push(email)
          //row[headers.indexOf('terceiro email enviado?')] = dataAtual
          sheet.getRange(rowIndex, headers.indexOf('terceiro email enviado?') + 1).setValue(dataAtual);
        
          quantidadeDeTerceiroEmailEnviado  +=  1
          return true
        }catch (e) {
          console.error("Erro ao enviar e-mail:", e);
          terminouComSucesso = false;
          return false;
        }
      }

      if(uniqueEmails.includes(email)){
        sheet.getRange(rowIndex, headers.indexOf('primeiro E-MAIL ENVIADO?') + 1).setValue("contato repetido");
        sheet.getRange(rowIndex, headers.indexOf('segundo email enviado?') + 1).setValue("contato repetido");
        sheet.getRange(rowIndex, headers.indexOf('terceiro email enviado?') + 1).setValue("contato repetido");
      }

    } else{
      sheet.getRange(rowIndex, headers.indexOf('primeiro E-MAIL ENVIADO?') + 1).setValue("não há e-mail de contato");
      sheet.getRange(rowIndex, headers.indexOf('segundo email enviado?') + 1).setValue("não há e-mail de contato");
      sheet.getRange(rowIndex, headers.indexOf('terceiro email enviado?') + 1).setValue("não há e-mail de contato");
        
    }
    

  }
  
  
}
