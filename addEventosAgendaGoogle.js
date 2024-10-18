function agendarEventos() {
  var spreadsheetId = '1UOmt7mXbPxPBDwHnoahWFHtZgtiaIeR2qgQfc38S5x0';
  var sheetName = 'agendamento';

  // Abre a planilha
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("A aba especificada não foi encontrada.");
    return;
  }

  // Obtem todos os dados
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var calendarId = 'example@gmail.com'; // E-mail fictício

  var calendario = CalendarApp.getDefaultCalendar();
  var calendarioExtra = CalendarApp.getCalendarById(calendarId);

  if (!calendarioExtra) {
    Logger.log("O calendário com o ID "+calendarId+" não foi encontrado.");
    //return;
  }

  var dataHoraAtual = new Date(); // Data e hora atual
  // Ignora a primeira linha de cabeçalho e começa da segunda linha
  try {
    for (var i = 1; i < data.length; i++) {
      var [ , email, nomeContratante, dataHoraEmbarqueStr, dataHoraRetornoStr, numeroContrato, contato, ciaAerea, localizador, observacao, statusIntegracao] = data[i];

      var dataHoraEmbarque = new Date(dataHoraEmbarqueStr);
      var dataHoraRetorno = new Date(dataHoraRetornoStr);

      if (statusIntegracao === 'Integrado') 
        continue; // Pula se já integrado

      if (dataHoraEmbarque <= dataHoraAtual && dataHoraRetorno <= dataHoraAtual) {
        console.log(`O registro com o localizador ${localizador} foi pulado porque a data de embarque é igual ou anterior à data atual.`);
        continue; 
      }
      
      // Calcula os dias restantes
      var { diasAteEmbarque, diasAteRetorno } = calcularDiasRestantes(dataHoraEmbarque, dataHoraRetorno);

      // Obtém as cores baseadas na companhia aérea e tempo restante
      var colorIdEmbarque = getColorId(diasAteEmbarque, ciaAerea);
      var colorIdRetorno = getColorId(diasAteRetorno, ciaAerea);

      // Criação de eventos
      var eventos = [
        { tipo: "Check - IN", data: getCheckDate(dataHoraEmbarque, ciaAerea), colorId: colorIdEmbarque },
        { tipo: "Embarque", data: dataHoraEmbarque, colorId: colorIdEmbarque },
        { tipo: "Check - OUT", data: getCheckDate(dataHoraRetorno, ciaAerea), colorId: colorIdRetorno },
        { tipo: "Retorno", data: dataHoraRetorno, colorId: colorIdRetorno }
      ];

      var sucessoIntegracao = true;

      for (var evento of eventos) {
        
        var titulo = `${evento.tipo}: ${nomeContratante}: ${ciaAerea}: ${localizador}`;
    
        if (!eventoJaExisteNoDia(calendario, titulo, evento.data)) {
          try {
            var criado = criarEventoAgenda(calendario, email, nomeContratante, evento.data, numeroContrato, contato, ciaAerea, localizador, evento.colorId, evento.tipo, observacao);
            
            Utilities.sleep(1200); // Pequeno atraso entre criações para evitar sobrecarga
            
            if (!criado) {
              sucessoIntegracao = false;
              break; // Se algum evento não for criado, interrompa o processo para este contrato
            }

          } catch (error) {
            Logger.log(`Erro ao criar evento: ${titulo} - ${error}`);
            sucessoIntegracao = false;
            break; // Interrompe o loop se houver um erro
          }
        } else {
          Logger.log(`Evento já existe: ${titulo}`);
        }

        if (!eventoJaExisteNoDia(calendarioExtra, titulo, evento.data)) {
          try { 
            var criadoExtra = criarEventoAgenda(calendarioExtra, email, nomeContratante, evento.data, numeroContrato, contato, ciaAerea, localizador, evento.colorId, evento.tipo, observacao);
          } catch (error) {
            Logger.log(`Erro ao criar evento na agenda Extra: ${titulo} - ${error}`);
          } 
        } else {
          Logger.log(`Evento já existe na agenda extra: ${titulo}`);
        }
      }

      if (sucessoIntegracao) {
        sheet.getRange(i + 1, 11).setValue('Integrado'); // Atualiza o status de integração
      }
    }
  } catch (error) {
    enviarEmail(evento);
  }
}

function calcularDiasRestantes(dataEmbarque, dataRetorno) {
  var hoje = new Date();
  var diasAteEmbarque = Math.ceil((dataEmbarque - hoje) / (1000 * 60 * 60 * 24));
  var diasAteRetorno = Math.ceil((dataRetorno - hoje) / (1000 * 60 * 60 * 24));
  return { diasAteEmbarque, diasAteRetorno };
}

function getColorId(diasAteEvento, ciaAerea) {
  if (ciaAerea === 'Azul' && diasAteEvento >= 3) {
    return 9; // Azul (72h ou mais)
  } else if ((ciaAerea === 'Gol' || ciaAerea === 'LATAM') && diasAteEvento >= 2) {
    return 5; // Amarelo (48h ou mais)
  } else {
    return 11; // Vermelho (menos de 72h ou 48h, dependendo da cia aérea)
  }
}

function getCheckDate(dataHora, ciaAerea) {
  var checkDate = new Date(dataHora);
  var offset = (ciaAerea === 'Azul') ? 3 : 2; // 3 dias para Azul, 2 dias para outras
  checkDate.setDate(checkDate.getDate() - offset);
  return checkDate;
}

function criarEventoAgenda(calendario, email, nomeContratante, dataHoraEvento, numeroContrato, contato, ciaAerea, localizador, colorId, tipoEvento, observacao) {
  var descricao = getDescricao(nomeContratante, numeroContrato, contato, ciaAerea, localizador, observacao, dataHoraEvento);
  var titulo = `${tipoEvento}: ${nomeContratante}: ${ciaAerea}: ${localizador}`;

  try {
    var evento = calendario.createEvent(titulo, dataHoraEvento, new Date(dataHoraEvento), {
      description: descricao,
      guests: email,
      sendInvites: true
    });

    // Adiciona lembretes baseados na companhia aérea
    if (ciaAerea === 'Azul') {
      evento.addPopupReminder(4320); // 72 horas antes
      evento.addEmailReminder(4320);
    } else if (ciaAerea === 'Gol' || ciaAerea === 'LATAM') {
      evento.addPopupReminder(2880); // 48 horas antes
      evento.addEmailReminder(2880);
    } 

    if (colorId) {
      evento.setColor(colorId);
    }

    Logger.log(`${tipoEvento} evento criado: ${titulo} com cor: ${colorId}`);
    return evento;
  } catch (error) {
    Logger.log(`Erro ao criar evento: ${titulo} - ${error}`);
    return null;
  }
}

function getDescricao(nomeContratante, numeroContrato, contato, ciaAerea, localizador, observacao, dataHoraEvento) {
  return `Nome do contratante: ${nomeContratante}
          Número do Contrato: ${numeroContrato}
          Contato: ${contato}
          Cia Aérea: ${ciaAerea}
          Localizador: ${localizador}
          Data do Evento: ${dataHoraEvento}
          Observação: ${observacao}`;
}

function eventoJaExisteNoDia(calendario, titulo, data) {
  var dataInicio = new Date(data);
  dataInicio.setHours(0, 0, 0, 0);
  var dataFim = new Date(data);
  dataFim.setHours(23, 59, 59, 999);

  var eventosNoDia = calendario.getEvents(dataInicio, dataFim);

  for (var i = 0; i < eventosNoDia.length; i++) {
    if (eventosNoDia[i].getTitle() === titulo) {
      return true;
    }
  }
  return false;
}

function enviarEmail(evento) {  
  // Envio do e-mail
  try {
      MailApp.sendEmail({
        to: "example@gmail.com", // E-mail fictício
        bcc: "examplebcc@gmail.com", // E-mail fictício
        subject: `Lembrete do evento: ${evento.getTitle()}`,
        body: `Lembre-se que o evento "${evento.getTitle()}" ocorrerá em breve. \n\nDetalhes do Evento:\n"${evento.getDescription()}"`
      });
    Logger.log(`E-mail de lembrete enviado para ${evento.getTitle()}`);
  } catch (error) {
      Logger.log(`Erro ao enviar e-mail: ${error}`);
  }
}


function atualizarCoresEventos() {
  const calendario = CalendarApp.getDefaultCalendar();
  const calendarioExtra = CalendarApp.getCalendarById('example@gmail.com'); // E-mail fictício

  // Obtém eventos da próxima semana
  const eventos = calendario.getEvents(new Date(), new Date(new Date().getTime() + 7 * 24 * 60 * 60 * 1000));
  const eventosExtra = calendarioExtra.getEvents(new Date(), new Date(new Date().getTime() + 7 * 24 * 60 * 60 * 1000));

  // Atualiza cores e envia e-mails para eventos do calendário padrão
  eventos.forEach(evento => {
    const diasAteEvento = DateDiff(evento.getStartTime(), new Date());
    const colorId = getColorId(diasAteEvento);

    evento.setColor(colorId); // Define a cor do evento
    Logger.log(`Cor do evento atualizada: ${evento.getTitle()} - ${colorId}`);

    // Verifica a cor e o tempo até o evento para enviar o e-mail
    if (evento.getTitle().toLowerCase().includes('azul') && diasAteEvento <= 3) { // Azul e faltam 72h
      enviarEmail(evento);
    } else if ((diasAteEvento <= 2) || // Vermelho e faltam 48h
               (evento.getTitle().toLowerCase().includes('latan') && diasAteEvento <= 2) || // Latan e faltam 48h
               (evento.getTitle().toLowerCase().includes('gol') && diasAteEvento <= 2)) { // Gol e faltam 48h
      enviarEmail(evento);
    }
  });

  // Atualiza cores e envia e-mails para eventos do calendário extra
  eventosExtra.forEach(evento => {
    const diasAteEvento = DateDiff(evento.getStartTime(), new Date());
    const colorId = getColorId(diasAteEvento);

    evento.setColor(colorId); // Define a cor do evento
    Logger.log(`Cor do eventosExtra atualizada: ${evento.getTitle()} - ${colorId}`);

    // Verifica a cor e o tempo até o evento para enviar o e-mail
    if (evento.getTitle().toLowerCase().includes('azul') && diasAteEvento <= 3) { // Azul e faltam 72h
      enviarEmail(evento);
    } else if ((diasAteEvento <= 2) || // Vermelho e faltam 48h
               (evento.getTitle().toLowerCase().includes('latan') && diasAteEvento <= 2) || // Latan e faltam 48h
               (evento.getTitle().toLowerCase().includes('gol') && diasAteEvento <= 2)) { // Gol e faltam 48h
      enviarEmail(evento);
    }
  });
}

function DateDiff(date1, date2) {
  const oneDay = 24 * 60 * 60 * 1000; // Horas * Minutos * Segundos * Milissegundos
  return Math.round(Math.abs((date1 - date2) / oneDay));
}

function getColorId(diasAteEvento) {
  if (diasAteEvento <= 3) {
    return CalendarApp.EventColor.RED; // Vermelho para eventos em até 3 dias
  } else if (diasAteEvento <= 6) {
    return CalendarApp.EventColor.YELLOW; // Amarelo para eventos em até 6 dias
  } else if (diasAteEvento <= 12) {
    return CalendarApp.EventColor.GREEN; // Verde para eventos em até 12 dias
  } else {
    return null; 
  }
}

function enviarEmail(evento) {  
  // Envio do e-mail
  try {
    MailApp.sendEmail({
      to: "example1@gmail.com; example2@gmail.com; example3@gmail.com; example4@gmail.com", // E-mails fictícios
      bcc: "examplebcc@gmail.com", // E-mail fictício
      subject: `Lembrete do evento: ${evento.getTitle()}`,
      body: `Lembre-se que o evento "${evento.getTitle()}" ocorrerá em breve. \n\nDetalhes do Evento:\n"${evento.getDescription()}"`
    });
    Logger.log('E-mail enviado para ' + "example1@gmail.com; example2@gmail.com; example3@gmail.com; example4@gmail.com");
  } catch (error) {
    Logger.log('Erro ao enviar e-mail: ' + error.toString());
  }
}