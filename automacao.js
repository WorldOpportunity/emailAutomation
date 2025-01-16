const diferencaDeDiasPrimeiroParaSegundoEmail = 5;
const diferencaDeDiasSegundoParaTerceiroEmail = 5;

const personalizarEmail = (corpoEmail, funcionario = '', cargo = '', empresa = '') => {
    corpoEmail = corpoEmail.replace('funcionario', funcionario);
    corpoEmail = corpoEmail.replace('cargo', cargo || '');  // Substituir por uma string vazia se o cargo estiver faltando
    corpoEmail = corpoEmail.replace('empresa', empresa || '');  // Substituir por uma string vazia se a empresa estiver faltando
    return corpoEmail;
};

function personalizaEEnviaEmail(email, assunto, corpo, nome, cargo, empresa) {
    corpo = personalizarEmail(corpo, funcionario = nome, cargo, empresa);

    MailApp.sendEmail({
        to: email,
        subject: assunto,
        htmlBody: corpo
    });
}

function diasPassados(data) {
    let dataAtual = new Date();
    // Garantir que a data fornecida esteja no formato correto
    let dataInput = new Date(data);
    if (isNaN(dataInput)) {
        throw new Error("Formato de data inválido: " + data);
    }
    let diferencaMilissegundos = dataAtual - dataInput;
    let diferencaDias = Math.floor(diferencaMilissegundos / (1000 * 60 * 60 * 24));
    return diferencaDias;
}

function sendEmails() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListaDeContatosEconodata');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const assuntoPrimeiro = 'Benefício Gratúito para sua Equipe';
    const assuntoSegundo = 'Assunto para o segundo e-mail';
    const assuntoTerceiro = 'Assunto para o terceiro e-mail';
    let uniqueEmails = [];
    let quantidadeDePrimeiroEmailEnviado = 0;
    let quantidadeDeSegundoEmailEnviado = 0;
    let quantidadeDeTerceiroEmailEnviado = 0;
    let terminouComSucesso = true;
    let rowIndex = 0;  // Para controlar o índice da linha

    data.every((row) => {
        const email = row[headers.indexOf('Email')];
        const nome = row[headers.indexOf('Nome')];
        const cargo = row[headers.indexOf('Cargo')];
        const empresa = row[headers.indexOf('RAZÃO SOCIAL')];
        let dataAtual = new Date();
        rowIndex += 1;

        // Verificar se o e-mail está presente
        if (!email) {
            console.log("E-mail faltando para a linha:", rowIndex);
            sheet.getRange(rowIndex, headers.indexOf('primeiro E-MAIL ENVIADO?') + 1).setValue("E-mail faltando");
            return true;  // Pula o envio para essa linha
        }

        // Verificar o estado dos e-mails
        const primeiroEmailEnviado = row[headers.indexOf('primeiro E-MAIL ENVIADO?')];
        const segundoEmailEnviado = row[headers.indexOf('segundo email enviado?')];
        const terceiroEmailEnviado = row[headers.indexOf('terceiro email enviado?')];

        // Enviar o primeiro e-mail, se ainda não enviado
        if (!primeiroEmailEnviado) {
            const corpo = row[headers.indexOf('Corpo primeiro e-mail')];
            try {
                personalizaEEnviaEmail(email, assuntoPrimeiro, corpo, nome, cargo, empresa);
                uniqueEmails.push(email);
                sheet.getRange(rowIndex, headers.indexOf('primeiro E-MAIL ENVIADO?') + 1).setValue(dataAtual);
                quantidadeDePrimeiroEmailEnviado += 1;
            } catch (e) {
                console.error("Erro ao enviar o primeiro e-mail para " + email + " na linha " + rowIndex + ": " + e.message);
                sheet.getRange(rowIndex, headers.indexOf('primeiro E-MAIL ENVIADO?') + 1).setValue("Erro: " + e.message + " (linha " + rowIndex + ")");
                terminouComSucesso = false;
                return false;
            }
        }

        // Enviar o segundo e-mail, se o primeiro foi enviado e o tempo para o envio já passou
        else if (primeiroEmailEnviado && !segundoEmailEnviado && diasPassados(primeiroEmailEnviado) >= diferencaDeDiasPrimeiroParaSegundoEmail) {
            const corpo = row[headers.indexOf('Corpo segundo e-mail')];
            try {
                personalizaEEnviaEmail(email, assuntoSegundo, corpo, nome, cargo, empresa);
                uniqueEmails.push(email);
                sheet.getRange(rowIndex, headers.indexOf('segundo email enviado?') + 1).setValue(dataAtual);
                quantidadeDeSegundoEmailEnviado += 1;
            } catch (e) {
                console.error("Erro ao enviar o segundo e-mail para " + email + " na linha " + rowIndex + ": " + e.message);
                sheet.getRange(rowIndex, headers.indexOf('segundo email enviado?') + 1).setValue("Erro: " + e.message + " (linha " + rowIndex + ")");
                terminouComSucesso = false;
                return false;
            }
        }

        // Enviar o terceiro e-mail, se o segundo foi enviado e o tempo para o envio já passou
        else if (primeiroEmailEnviado && segundoEmailEnviado && !terceiroEmailEnviado && diasPassados(segundoEmailEnviado) >= diferencaDeDiasSegundoParaTerceiroEmail) {
            const corpo = row[headers.indexOf('Corpo terceiro e-mail')];
            try {
                personalizaEEnviaEmail(email, assuntoTerceiro, corpo, nome, cargo, empresa);
                uniqueEmails.push(email);
                sheet.getRange(rowIndex, headers.indexOf('terceiro email enviado?') + 1).setValue(dataAtual);
                quantidadeDeTerceiroEmailEnviado += 1;
            } catch (e) {
                console.error("Erro ao enviar o terceiro e-mail para " + email + " na linha " + rowIndex + ": " + e.message);
                sheet.getRange(rowIndex, headers.indexOf('terceiro email enviado?') + 1).setValue("Erro: " + e.message + " (linha " + rowIndex + ")");
                terminouComSucesso = false;
                return false;
            }
        }

        return true;
    });

    // Logs finais para o total de e-mails enviados
    console.log("Total de primeiros e-mails enviados: ", quantidadeDePrimeiroEmailEnviado);
    console.log("Total de segundos e-mails enviados: ", quantidadeDeSegundoEmailEnviado);
    console.log("Total de terceiros e-mails enviados: ", quantidadeDeTerceiroEmailEnviado);
    console.log("Processo terminado com sucesso? ", terminouComSucesso ? "Sim" : "Não");
}
