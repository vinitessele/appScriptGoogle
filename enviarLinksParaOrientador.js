function enviarLinksParaOrientador() {
  var spreadsheetId = '1Q-j8Qc_pCfcAxzXhRz1dMLKC03jZj-rn2bmb7Wu3Wsw'; // Substitua com o ID da sua planilha
  var sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
  
  // Obter todos os dados da planilha
  var data = sheet.getDataRange().getValues();
  
  const orientadores = {
    "Professor A": "email_professor_a@exemplo.com", // Substitua por identificadores genéricos
    "Professor B": "email_professor_b@exemplo.com",
    "Professor C": "email_professor_c@exemplo.com",
    // Adicione mais orientadores conforme necessário
};

function encontrarEmailOrientador(nomeOrientador) {
    for (const [nome, email] of Object.entries(orientadores)) {
        if (nome.includes(nomeOrientador)) {
            return email; // Retorna o e-mail se encontrar uma correspondência parcial
        }
    }
    return "email_padrao@exemplo.com"; // E-mail padrão se não encontrar
}
  
  // Iterar sobre os dados e enviar e-mail para cada artigo validado
  for (var i = 1; i < data.length; i++) { // Começar em 1 para ignorar o cabeçalho
    var statusIntegracao = data[i][9]; 
    if (statusIntegracao === 'SIM') 
        continue; 

    var validado = data[i][6]; 
    var linkArtigo = data[i][7];
    var linkBanner = data[i][8];
    var alunoNome = data[i][3]; 
    var orientadorNome = data[i][4].split(' ')[0];

    // Verificar se o artigo foi validado e o link do artigo não está vazio
    if (validado === "SIM" && linkArtigo) {
      
      var professorEmail = encontrarEmailOrientador(orientadorNome);
      var mensagem = "Olá Professor,\n\nO artigo corrigido do aluno " + alunoNome + " está disponível no seguinte link:\n" + linkArtigo + 
      "\nLink do Banner se houver: " + linkBanner +
      "\n\nAtenciosamente,\nBiopark Educação";

      // Enviar o e-mail
      MailApp.sendEmail({
        to: professorEmail,
        subject: "Artigo Corrigido: " + data[i][2], // Título do artigo
        body: mensagem
      });
      
      // Marcar a coluna "Integrado" (10ª coluna) como "SIM"
      sheet.getRange(i + 1, 10).setValue("SIM"); // i + 1 porque as linhas no Google Sheets começam em 1
      Logger.log("E-mail enviado para o professor com o link do artigo: " + data[i][2]);
    }
  }
}
