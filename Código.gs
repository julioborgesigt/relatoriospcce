/**
 * PONTO DE ENTRADA PARA SITES EXTERNOS (API)
 * Recebe as requisições do seu formulário hospedado fora do Google.
 */
function doPost(e) {  
  var requestData;
  try {
    requestData = JSON.parse(e.postData.contents);
  } catch (f) {
    return ContentService.createTextOutput(JSON.stringify({sucesso: false, msg: "Dados inválidos"})).setMimeType(ContentService.MimeType.JSON);
  }

  var action = requestData.action;
  var result;

  try {
    if (action === "buscarServidor") {
        result = buscarServidorPelaMatricula(requestData.matricula);
      } 
      else if (action === "obterDelegacias") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Apoio_Delegacias");
      // Pega da linha 2 até a última linha com dados na coluna A
      var lastRow = sheet.getLastRow();
      var delegacias = [];
      
      if (lastRow > 1) {
        // Lê a coluna A (índice 1) da linha 2 até o fim
        var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        // Transforma matriz [[a],[b]] em array [a,b] e remove vazios
        delegacias = values.map(function(r){ return r[0]; }).filter(function(d){ return d !== ""; });
      }
      
      return ContentService.createTextOutput(JSON.stringify(delegacias)).setMimeType(ContentService.MimeType.JSON);
    }

    else if (action === "enviarToken") {
      result = { sucesso: enviarTokenAcesso(requestData.email, requestData.nome, requestData.matricula) };
    } 
    else if (action === "validarToken") {
      result = validarTokenNoServidor(requestData.email, requestData.token);
    } 
    else if (action === "obterListaServidores") {
      result = obterListaServidoresCompleta();
    } 
    else if (action === "salvarRelatorio") {
      // Passamos o idRascunho como 4º parâmetro para a função processarEnvioProdutividade
      result = processarEnvioProdutividade(requestData.quant, requestData.quali, requestData.user, requestData.idRascunho);
      
      // Se era um rascunho (R-), apaga. Se era retificação (FT-), a limpeza já foi feita dentro da função.
      if (result.sucesso && requestData.idRascunho && requestData.idRascunho.startsWith("R-")) {
        apagarRascunho(requestData.idRascunho);
      }
      // Atualiza a base do Dashboard após envio definitivo
      sincronizarBaseDashboard(); 
    }
    else if (action === "obterDadosDashboard") {
      sincronizarBaseDashboard(); // Força a atualização antes de retornar os dados
      var dados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Dashboard").getDataRange().getValues();
      return ContentService.createTextOutput(JSON.stringify(dados)).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "salvarRascunho") {
      result = gerenciarSalvarRascunho(requestData.dados);
      // Atualiza a base do Dashboard após salvar rascunho para mostrar dados temporários
      sincronizarBaseDashboard(); 
    }
    else if (action === "carregarRascunho") {
      result = carregarRascunhoExistente(requestData.idRascunho);
    }
    else if (action === "carregarRelatorioFinalizado") {
      result = carregarRelatorioFinalizado(requestData.idProtocolo);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({sucesso: false, msg: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * PONTO DE ENTRADA PARA O RELATÓRIO
 * Permite que o link de impressão abra normalmente.
 */
function doGet(e) {
  var output;

  if (e.parameter.id) {
    var template = HtmlService.createTemplateFromFile('Relatorio');
    template.dados = obterDadosParaRelatorio(e.parameter.id); 
    
    // BUSCA OFICIAL DO ASSINANTE NA ABA SERVIDORES
    if (e.parameter.asn) {
      const nomeParaBusca = decodeURIComponent(e.parameter.asn);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetServ = ss.getSheetByName("Servidores");
      const dadosServ = sheetServ.getDataRange().getValues();
      const serv = dadosServ.find(r => r[0].toString().trim() === nomeParaBusca.trim());
      
      if (serv) {
        let cargoExtenso = serv[1] === "OIP" ? "Oficial Investigador de Polícia" : 
                           serv[1] === "DPC" ? "Delegado de Polícia Civil" : serv[1];
        template.assinante = {
          nome: serv[0],
          cargo: cargoExtenso,
          matricula: serv[2],
          telefone: serv[3],
          lotacao: serv[6]
        };
      } else {
        template.assinante = null;
      }
    }
    output = template.evaluate().setTitle('Relatório de Plantão - ' + e.parameter.id);
  } else {
    // Se acessar a URL sem ID, apenas avisa que o backend está ativo
    return ContentService.createTextOutput("Backend DPI-SUL Ativo. Use o site externo para acessar o formulário.");
  }

  return output
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ==========================================================
   FUNÇÕES ORIGINAIS DE NEGÓCIO (MANTIDAS)
   ========================================================== */

function buscarServidorPelaMatricula(matriculaOriginal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetServ = ss.getSheetByName("Servidores");
  const sheetApoio = ss.getSheetByName("Apoio_Delegacias");
  
  const dadosServ = sheetServ.getDataRange().getValues();
  const dadosApoio = sheetApoio.getDataRange().getValues();
  
  const matriculaLimpa = matriculaOriginal.replace(/[^0-9a-zA-Z]/g, "").toUpperCase();
  
  // Captura data atual (apenas dia, mês e ano)
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  // 1. LOCALIZAÇÃO DO SERVIDOR
  let servidorEncontrado = null;
  for (let i = 1; i < dadosServ.length; i++) {
    // Coluna Z = índice 25
    if (dadosServ[i][25].toString() === matriculaLimpa) {
      servidorEncontrado = {
        nome: dadosServ[i][0],
        email: dadosServ[i][18].toString().split(/[\s,;]+/)[0], // Primeiro e-mail da célula
        lotacao: dadosServ[i][6].toString().trim(), // Coluna G = índice 6
        matricula: matriculaLimpa
      };
      break;
    }
  }

  if (!servidorEncontrado) {
    return { sucesso: false, msg: "Matrícula não localizada." };
  }

  // 2. REGRA ESPECIAL: ACESSO DIRETO DPI SUL / JUAZEIRO
  const lotacaoServ = servidorEncontrado.lotacao.toUpperCase();
  if (lotacaoServ.includes("DEPARTAMENTO DE POLICIA DO INTERIOR SUL")) {
     return { sucesso: true, ...servidorEncontrado };
  }

  // 3. VALIDAÇÃO DE PLANTÃO (ABA APOIO_DELEGACIAS)
  let unidadeValidada = false;
  let msgErro = "Você não faz parte de uma Delegacia plantonista.";

  for (let j = 1; j < dadosApoio.length; j++) {
    const nomeUnidadeApoio = dadosApoio[j][0].toString().trim(); // Coluna A
    const condicaoPlantonista = dadosApoio[j][3].toString().toUpperCase().trim(); // Coluna D
    const dataVencimento = dadosApoio[j][4]; // Coluna E (Objeto Date)

    if (nomeUnidadeApoio === servidorEncontrado.lotacao) {
      if (condicaoPlantonista === "SIM") {
        unidadeValidada = true;
      } 
      else if (condicaoPlantonista === "TEMPORÁRIO" || condicaoPlantonista === "TEMPORARIO") {
        if (dataVencimento instanceof Date) {
          // Verifica se ainda está no prazo
          if (hoje <= dataVencimento) {
            unidadeValidada = true;
          } else {
            msgErro = "Sua permissão temporária de acesso expirou em " + Utilities.formatDate(dataVencimento, "GMT-3", "dd/MM/yyyy") + ".";
          }
        } else {
          msgErro = "Erro na data de vencimento da unidade. Contate a gestão.";
        }
      }
      break; // Encontrou a unidade, encerra o loop de busca
    }
  }

  if (unidadeValidada) {
    return { sucesso: true, ...servidorEncontrado };
  } else {
    return { sucesso: false, msg: msgErro };
  }
}

function enviarTokenAcesso(email, nome, matricula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTokens = ss.getSheetByName("Tokens_Acesso");
  const token = Math.floor(100000 + Math.random() * 900000).toString();
  const dataExpiracao = new Date(new Date().getTime() + 15 * 60000);
  sheetTokens.appendRow([email.toLowerCase(), token, dataExpiracao, "Pendente"]);
  
  const assunto = "Código de Acesso - PLANTÃO - DPI SUL";
  const corpoHtml = `<div style="font-family: sans-serif; background-color: #0a192f; color: white; padding: 20px; border-radius: 10px;">
    <h2 style="color: #c5a059;">Plantão - DPI SUL</h2>
    <p>Olá, <strong>${nome}</strong>.</p>
    <p>Seu código é: <span style="background-color: #c5a059; color: #0a192f; padding: 5px; font-weight: bold;">${token}</span></p>
  </div>`;
  
  MailApp.sendEmail({ to: email, subject: assunto, htmlBody: corpoHtml });
  return true;
}

function validarTokenNoServidor(email, tokenDigitado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tokens_Acesso");
  const dados = sheet.getDataRange().getValues();
  const agora = new Date();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().toLowerCase() === email.toLowerCase() && dados[i][1].toString() === tokenDigitado) {
      if (dados[i][3] === "Pendente" && agora <= new Date(dados[i][2])) {
        sheet.getRange(i + 1, 4).setValue("Usado");
        return { sucesso: true };
      }
    }
  }
  return { sucesso: false, msg: "Código incorreto ou expirado." };
}

function processarEnvioProdutividade(quant, quali, user, idExistente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetQuant = ss.getSheetByName("Base_Quantitativos");
  const sheetQuali = ss.getSheetByName("Base_Qualitativos");
  const sheetEquipe = ss.getSheetByName("Base_Equipe");
  const timezone = ss.getSpreadsheetTimeZone();
  
  // 1. Definições Iniciais e Verificação de Retificação
  const ehRetificacao = !!(idExistente && idExistente.startsWith("FT-"));
  const agora = new Date();
  const dataRegistroBR = Utilities.formatDate(agora, timezone, "dd/MM/yyyy HH:mm:ss");
  
  const idTransacao = ehRetificacao ? idExistente : "FT-" + Math.random().toString(36).substr(2, 9).toUpperCase();
  
  // 2. Define o conteúdo do Status (Coluna S) e Observações (Coluna N)
  const statusConteudo = ehRetificacao 
    ? `[RETIFICADO EM: ${dataRegistroBR} POR: ${user.nome.toUpperCase()}]` 
    : "ORIGINAL";
    
  const obsFinal = (quant.observacoes || "").toUpperCase().trim();

  try {
    // 3. Se for retificação, remove os registros antigos antes de inserir os novos
    if (ehRetificacao) {
      removerDadosAntigosID(idExistente);
    }

    // 4. GRAVAÇÃO NA BASE_QUANTITATIVOS (Mapeamento A até S - 19 Colunas)
    sheetQuant.insertRowBefore(2);
    sheetQuant.getRange(2, 1, 1, 19).setValues([[
      idTransacao,                   // A: Protocolo
      dataRegistroBR,                // B: Data Registro
      "'" + user.matricula,          // C: Matrícula do Operador
      quant.delegacia,               // D: Unidade/Delegacia
      quant.bos,                     // E: B.O.
      quant.guias,                   // F: Guias
      quant.apreensoes,              // G: Apreensões
      quant.presos,                  // H: Presos
      quant.medidas,                 // I: Medidas Protetivas
      quant.outros,                  // J: Outros
      "'" + quant.entrada,           // K: Início Plantão
      "'" + quant.saida,             // L: Fim Plantão
      quant.equipeResumo,            // M: Equipe (Nomes)
      obsFinal,                      // N: Observações (Limpa e em CAIXA ALTA)
      quant.ipFlagrante || 0,        // O: Contagem IP-Flagrante
      quant.ipPortaria || 0,         // P: Contagem IP-Portaria
      quant.tco || 0,                // Q: Contagem TCO
      quant.aiBoc || 0,              // R: Contagem AI/BOC
      statusConteudo                 // S: Carimbo de Retificação
    ]]);

    // ... restante da função (Base_Equipe e Base_Qualitativos) permanece igual

    // 2. GRAVAÇÃO NA BASE_EQUIPE (Individualizado)
    // quant.equipeDetalhada é o array de objetos que virá do frontend
    if (quant.equipeDetalhada && quant.equipeDetalhada.length > 0) {
      const cidadeAtuacao = extrairCidade(quant.delegacia);
      
      quant.equipeDetalhada.forEach(pol => {
        sheetEquipe.insertRowBefore(2);
        sheetEquipe.getRange(2, 1, 1, 12).setValues([[
          idTransacao,
          pol.nome,
          cidadeAtuacao,
          "Plantão",
          pol.tipoEscala, // "Normal" ou "Extraordinária"
          "'" + pol.dataEnt,
          "'" + pol.horaEnt,
          "'" + pol.dataSai,
          "'" + pol.horaSai,
          pol.totalHoras, // Valor numérico para cálculos
          "'" + pol.matricula,
          pol.cargo
        ]]);
      });
    }

    // 3. GRAVAÇÃO NA BASE_QUALITATIVOS (Procedimentos)
    quali.reverse().forEach(item => {
      sheetQuali.insertRowBefore(2);
      sheetQuali.getRange(2, 1, 1, 9).setValues([[
        idTransacao, 
        item.tipo, 
        item.numero, 
        "'" + item.numero.replace(/\D/g, ""), 
        sanitizarTexto(item.crime), 
        sanitizarTexto(item.nomePessoa), 
        item.papel === 'V' ? 'VITIMA' : 'INFRATOR', 
        Utilities.formatDate(agora, timezone, "dd/MM/yyyy"), 
        Utilities.formatDate(agora, timezone, "HH:mm")
      ]]);
    });
    
    return { sucesso: true, idTransacao: idTransacao };
  } catch (e) {
    return { sucesso: false, msg: e.toString() };
  }
}

function obterListaServidoresCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Servidores");
  const dados = sheet.getRange("A2:G" + sheet.getLastRow()).getValues();
  return dados.map(r => ({ nome: r[0], cargo: r[1], classe: r[5], matricula: r[2], telefone: r[3], lotacao: r[6] })).filter(r => r.nome !== "");
}

function obterDadosParaRelatorio(idTransacao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registro = ss.getSheetByName("Base_Quantitativos").getDataRange().getValues().find(r => r[0] === idTransacao);
  const detalhesBrutos = ss.getSheetByName("Base_Qualitativos").getDataRange().getValues().filter(r => r[0] === idTransacao);
  const servidores = ss.getSheetByName("Servidores").getDataRange().getValues();
  
  if (!registro) return null;

  const equipe = registro[12].split(' / ').map(nome => {
    const s = servidores.find(row => row[0].toString().trim() === nome.trim());
    return { nome: nome, cargo: s?s[1]:"", classe: s?s[5]:"", matricula: s?s[2]:"", telefone: s?s[3]:"", lotacao: s?s[6]:"" };
  }).sort((a, b) => (a.cargo === "DPC" ? -1 : b.cargo === "DPC" ? 1 : a.nome.localeCompare(b.nome)));

  const procedimentosMap = {};
  
  // AJUSTE DE SEGURANÇA: Usamos Number() e || 0 para garantir compatibilidade com registros antigos
  // registro[14] = Coluna O, registro[15] = Coluna P, etc.
  const cont = { 
    IPF: (registro.length > 14) ? (Number(registro[14]) || 0) : 0, 
    IPP: (registro.length > 15) ? (Number(registro[15]) || 0) : 0, 
    TCO: (registro.length > 16) ? (Number(registro[16]) || 0) : 0, 
    BOC: (registro.length > 17) ? (Number(registro[17]) || 0) : 0 
  };

  detalhesBrutos.forEach(r => {
    if (!procedimentosMap[r[2]]) {
      procedimentosMap[r[2]] = { tipo: r[1], numero: r[2], crime: r[4], vitimas: [], infratores: [] };
    }
    if (r[6] === "VITIMA") procedimentosMap[r[2]].vitimas.push(r[5]);
    else procedimentosMap[r[2]].infratores.push(r[5]);
  });

  return {
    id: idTransacao, 
    delegacia: registro[3], 
    entrada: String(registro[10]), 
    saida: String(registro[11]),
    equipe: equipe, 
    resumo: cont, 
    observacoes: registro[13],
    quant: { 
      bos: registro[4], guias: registro[5], apre: registro[6], 
      presos: registro[7], medi: registro[8], outr: registro[9]
    },
    detalhes: Object.values(procedimentosMap)
  };
}

function sanitizarTexto(texto) {
  if (!texto) return "";
  return texto.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

/* ==========================================================
   FUNÇÕES DE RASCUNHO (SISTEMA DE RETOMADA)
   ========================================================== */

function gerenciarSalvarRascunho(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Rascunhos");
  const agora = new Date();
  
  // 1. Definição do ID e Linha (Lógica já existente)
  let id = dados.idRascunho;
  let linhaExistente = -1;
  if (id) {
    const idsExistentes = sheet.getRange("A:A").getValues();
    for (let i = 0; i < idsExistentes.length; i++) {
      if (idsExistentes[i][0] === id) { linhaExistente = i + 1; break; }
    }
  }
  if (linhaExistente === -1) {
    id = "R-" + Math.random().toString(36).substr(2, 6).toUpperCase();
  }

  // 2. Preparação dos dados para a planilha
  let resumoCrimes = dados.quali.map(q => q.tipo + ": " + q.crime).join(" | ");
  const linhaParaSalvar = [
    id,
    dados.user.matricula,
    dados.quant.delegacia,
    Utilities.formatDate(agora, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss"),
    JSON.stringify(dados.quant),
    resumoCrimes,
    JSON.stringify(dados)
  ];

  if (linhaExistente !== -1) {
    sheet.getRange(linhaExistente, 1, 1, 7).setValues([linhaParaSalvar]);
  } else {
    sheet.appendRow(linhaParaSalvar);
  }

  // 3. ENVIO DE E-MAIL COM O CÓDIGO
  enviarEmailRascunho(dados.user.email, id, dados.quant.delegacia, dados.quant.entrada);

  return { sucesso: true, idRascunho: id };
}

// FUNÇÃO AUXILIAR DE E-MAIL
function enviarEmailRascunho(email, idRascunho, unidade, dataPlantao) {
  const assunto = "CÓDIGO DE RASCUNHO - PLANTÃO DPI SUL";
  const corpoHtml = `
    <div style="font-family: sans-serif; background-color: #0a192f; color: white; padding: 20px; border-radius: 10px; border: 1px solid #c5a059;">
      <h2 style="color: #c5a059;">Rascunho de Plantão Salvo</h2>
      <p>Unidade: <strong>${unidade}</strong></p>
      <p>Início do Plantão: <strong>${dataPlantao}</strong></p>
      <hr style="border: 0; border-top: 1px solid #c5a059;">
      <p style="font-size: 1.2em;">Seu código de recuperação é: <span style="background-color: #c5a059; color: #0a192f; padding: 5px 10px; font-weight: bold; border-radius: 5px;">${idRascunho}</span></p>
      <p style="color: #ffc107; font-size: 0.9em;">⚠️ Este rascunho expirará automaticamente em 36 horas se não for finalizado.</p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: email,
    subject: assunto,
    htmlBody: corpoHtml
  });
}

function carregarRascunhoExistente(idRascunho) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Rascunhos");
  const dados = sheet.getDataRange().getValues();
  const agora = new Date().getTime();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === idRascunho) {
      // Verifica se o rascunho tem mais de 36 horas
      const dataCriacao = new Date(dados[i][3]).getTime();
      const horasDecorridas = (agora - dataCriacao) / (1000 * 60 * 60);

      if (horasDecorridas > 36) {
        return { sucesso: false, msg: "Este rascunho expirou (limite de 36h)." };
      }

      return { sucesso: true, payload: JSON.parse(dados[i][6]) };
    }
  }
  return { sucesso: false, msg: "ID de Rascunho não localizado." };
}

function apagarRascunho(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rascunhos");
  const dados = sheet.getRange("A:A").getValues();
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

/**
 * Função de limpeza automática (Rodar via Gatilho de Tempo)
 */
function limpezaAutomaticaRascunhos() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rascunhos");
  if (!sheet) return;
  const dados = sheet.getDataRange().getValues();
  const agora = new Date().getTime();
  let linhasDeletadas = 0;

  // Percorre de baixo para cima para não errar o índice ao deletar
  for (let i = dados.length - 1; i >= 1; i--) {
    const dataCriacao = new Date(dados[i][3]).getTime();
    if (isNaN(dataCriacao)) continue;
    
    const horasDecorridas = (agora - dataCriacao) / (1000 * 60 * 60);
    if (horasDecorridas > 36) {
      sheet.deleteRow(i + 1);
      linhasDeletadas++;
    }
  }
  console.log("Limpeza concluída. Rascunhos expirados removidos: " + linhasDeletadas);
}

//**Dashboard **/

function obterDadosDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaQuant = ss.getSheetByName("Base_Quantitativos");
  const abaQuali = ss.getSheetByName("Base_Qualitativos");
  const abaRasc = ss.getSheetByName("Rascunhos");

  let dadosSincronizados = [];

  const vQuant = abaQuant.getDataRange().getValues();
  const vQuali = abaQuali.getDataRange().getValues();
  
  // 1. PROCESSAR DEFINITIVOS
  for (let i = 1; i < vQuant.length; i++) {
    let idTransacao = vQuant[i][0];
    // Filtra todos os crimes ligados a este ID específico
    let listaQualiParaID = vQuali.filter(q => q[0] === idTransacao);
    
    dadosSincronizados.push({
      unidade: String(vQuant[i][3]),
      bo: Number(vQuant[i][4]) || 0,
      guia: Number(vQuant[i][5]) || 0,
      apre: Number(vQuant[i][6]) || 0,
      preso: Number(vQuant[i][7]) || 0,
      medida: Number(vQuant[i][8]) || 0, // Coluna I
      outros: Number(vQuant[i][9]) || 0,
      plantao_ini: vQuant[i][10], // Início do plantão
      plantao_fim: vQuant[i][11], // Fim do plantão
      // Detalhamento Qualitativo para a Tabela
      lista_procedimentos: listaQualiParaID.map(q => ({
        tipo: q[1],
        numero: q[2],
        crime: q[4]
      })),
      data_ref: vQuant[i][1] // Data do registro para o gráfico
    });
  }

  // 2. PROCESSAR RASCUNHOS
  const vRasc = abaRasc.getDataRange().getValues();
  for (let i = 1; i < vRasc.length; i++) {
    try {
      const json = JSON.parse(vRasc[i][6]);
      if (json.quant) {
        dadosSincronizados.push({
          unidade: json.quant.delegacia,
          bo: Number(json.quant.bos) || 0,
          guia: Number(json.quant.guias) || 0,
          apre: Number(json.quant.apreensoes) || 0,
          preso: Number(json.quant.presos) || 0,
          medida: Number(json.quant.medidas) || 0,
          outros: Number(json.quant.outros) || 0,
          plantao_ini: json.quant.entrada,
          plantao_fim: json.quant.saida,
          lista_procedimentos: json.quali ? json.quali.map(q => ({ tipo: q.tipo, numero: q.numero, crime: q.crime })) : [],
          data_ref: new Date(),
          status: "RASCUNHO"
        });
      }
    } catch (e) {}
  }
  return dadosSincronizados;
}

function sincronizarBaseDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDash = ss.getSheetByName("Base_Dashboard");
  const abaQuant = ss.getSheetByName("Base_Quantitativos");
  const abaQuali = ss.getSheetByName("Base_Qualitativos");
  const abaRasc = ss.getSheetByName("Rascunhos");

  // 1. Limpa tudo exceto o cabeçalho
  if (abaDash.getLastRow() > 1) {
    abaDash.getRange(2, 1, abaDash.getLastRow() - 1, 15).clearContent();
  }

  let novasLinhas = [];

  // 2. PROCESSAR DADOS DEFINITIVOS (Cruzando Quant + Quali com De-duplicação)
  const vQuant = abaQuant.getDataRange().getValues();
  const vQuali = abaQuali.getDataRange().getValues();

  for (let i = 1; i < vQuant.length; i++) {
    let id = vQuant[i][0];
    if (!id) continue;

    let dadosPlantao = [
      vQuant[i][0], vQuant[i][3], vQuant[i][1], vQuant[i][10], vQuant[i][11],
      vQuant[i][4], vQuant[i][5], vQuant[i][6], vQuant[i][7], vQuant[i][8], vQuant[i][9]
    ];

    // Busca procedimentos qualitativos para este ID
    let procedimentos = vQuali.filter(q => q[0] === id);

    if (procedimentos.length > 0) {
      // MAPA PARA DE-DUPLICAR: Garante que cada número de procedimento apareça apenas uma vez por ID
      let procedimentosUnicos = {};
      procedimentos.forEach(p => {
        let numProc = p[2]; 
        if (!procedimentosUnicos[numProc]) {
          procedimentosUnicos[numProc] = p; // Guarda a primeira ocorrência do procedimento
        }
      });

      // Transforma o mapa de volta em linhas para a planilha
      Object.values(procedimentosUnicos).forEach(p => {
        novasLinhas.push([...dadosPlantao, p[1], p[2], p[4], "DEFINITIVO"]);
      });
    } else {
      novasLinhas.push([...dadosPlantao, "", "", "", "DEFINITIVO"]);
    }
  }

  // 3. PROCESSAR RASCUNHOS (Com a mesma lógica de De-duplicação)
  const vRasc = abaRasc.getDataRange().getValues();
  for (let i = 1; i < vRasc.length; i++) {
    try {
      const json = JSON.parse(vRasc[i][6]);
      if (json.quant) {
        let dadosPlantaoRasc = [
          vRasc[i][0], json.quant.delegacia, vRasc[i][3], json.quant.entrada, json.quant.saida,
          json.quant.bos, json.quant.guias, json.quant.apreensoes, json.quant.presos, json.quant.medidas, json.quant.outros
        ];

        if (json.quali && json.quali.length > 0) {
          // DE-DUPLICAÇÃO NOS RASCUNHOS
          let rascUnicos = {};
          json.quali.forEach(q => {
            if (!rascUnicos[q.numero]) {
              rascUnicos[q.numero] = q;
            }
          });

          Object.values(rascUnicos).forEach(q => {
            novasLinhas.push([...dadosPlantaoRasc, q.tipo, q.numero, q.crime, "TEMPORARIO"]);
          });
        } else {
          novasLinhas.push([...dadosPlantaoRasc, "", "", "", "TEMPORARIO"]);
        }
      }
    } catch (e) { /* Pula rascunhos corrompidos */ }
  }

  // 4. Grava tudo na aba de uma vez só (Performance)
  if (novasLinhas.length > 0) {
    abaDash.getRange(2, 1, novasLinhas.length, 15).setValues(novasLinhas);
  }
  SpreadsheetApp.flush();
  return { sucesso: true, total: novasLinhas.length };
}

function extrairCidade(unidade) {
  // 1. Regras das Seccionais (Exceções Fixas)
  if (unidade.includes("1ª Seccional")) return "Russas";
  if (unidade.includes("2ª Seccional")) return "Juazeiro do Norte";
  if (unidade.includes("3ª Seccional")) return "Quixadá";
  if (unidade.includes("4ª Seccional")) return "Iguatu";
  if (unidade.includes("5ª Seccional")) return "Tauá";

  // 2. Regra Geral: Extrai após o segundo " de "
  // Usamos uma expressão regular para garantir que pegue o "de" como palavra isolada
  var partes = unidade.split(/\sde\s/i); 
  if (partes.length >= 3) {
    // Pega tudo após o segundo " de ", remove parênteses e limpa espaços
    var cidadeRaw = partes.slice(2).join(" de ");
    return cidadeRaw.replace(/[\(\)]/g, "").trim();
  }
  
  return unidade.trim(); 
}

/**
 * Busca um relatório já finalizado e converte para o formato JSON do formulário
 * CORREÇÃO: Agora lê a Base_Equipe para recuperar Tipo_Escala e horários individuais
 */
function carregarRelatorioFinalizado(idProtocolo) {
  try {
    const dadosRelatorio = obterDadosParaRelatorio(idProtocolo);
    if (!dadosRelatorio) return { sucesso: false, msg: "Protocolo não localizado." };

    // Busca os dados detalhados da equipe na Base_Equipe
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEquipe = ss.getSheetByName("Base_Equipe");
    const dadosEquipe = sheetEquipe.getDataRange().getValues();

    // Filtra apenas as linhas deste protocolo
    const equipeDoProtocolo = [];
    for (var i = 1; i < dadosEquipe.length; i++) {
      if (dadosEquipe[i][0] === idProtocolo) {
        equipeDoProtocolo.push({
          nome: String(dadosEquipe[i][1]).trim(),              // B: Nome_Policial
          tipoEscala: String(dadosEquipe[i][4]).trim(),        // E: Tipo_Escala
          dataEnt: String(dadosEquipe[i][5]).trim(),            // F: Data_Entrada_Individual
          horaEnt: String(dadosEquipe[i][6]).trim(),            // G: Hora_Entrada_Individual
          dataSai: String(dadosEquipe[i][7]).trim(),            // H: Data_Saida_Individual
          horaSai: String(dadosEquipe[i][8]).trim(),            // I: Hora_Saida_Individual
          totalHoras: dadosEquipe[i][9],                        // J: Total_Horas
          matricula: String(dadosEquipe[i][10]).trim(),         // K: Matricula
          cargo: String(dadosEquipe[i][11]).trim()              // L: Cargo
        });
      }
    }

    // Extrai entrada e saída gerais do plantão para comparar com individuais
    var entradaGeral = dadosRelatorio.entrada; // Ex: "24/02/2026 08:00"
    var saidaGeral = dadosRelatorio.saida;
    var dEntGeral = entradaGeral.split(" ")[0];
    var hEntGeral = entradaGeral.split(" ")[1];
    var dSaiGeral = saidaGeral.split(" ")[0];
    var hSaiGeral = saidaGeral.split(" ")[1];

    // Monta a equipe usando os dados da Base_Equipe (ou fallback para dados básicos)
    var equipeFormatada;
    if (equipeDoProtocolo.length > 0) {
      equipeFormatada = equipeDoProtocolo.map(function(pol) {
        // Remove apóstrofos que o Sheets pode adicionar
        var pDataEnt = pol.dataEnt.replace(/^'/, "");
        var pHoraEnt = pol.horaEnt.replace(/^'/, "");
        var pDataSai = pol.dataSai.replace(/^'/, "");
        var pHoraSai = pol.horaSai.replace(/^'/, "");

        // Determina se está sincronizado comparando todos os 4 campos
        var isSynced = (pDataEnt === dEntGeral && pHoraEnt === hEntGeral && pDataSai === dSaiGeral && pHoraSai === hSaiGeral);

        return {
          nome: pol.nome,
          matricula: pol.matricula.replace(/^'/, ""),
          cargo: pol.cargo,
          escala: pol.tipoEscala || "Normal",
          sync: isSynced,
          dEnt: pDataEnt,
          hEnt: pHoraEnt,
          dSai: pDataSai,
          hSai: pHoraSai,
          totalH: String(pol.totalHoras) + "h"
        };
      });
    } else {
      // Fallback: se Base_Equipe não tiver dados, usa o resumo básico
      equipeFormatada = dadosRelatorio.equipe.map(function(p) {
        return {
          nome: p.nome,
          matricula: p.matricula,
          cargo: p.cargo,
          escala: "Normal",
          sync: true,
          dEnt: dEntGeral,
          hEnt: hEntGeral,
          dSai: dSaiGeral,
          hSai: hSaiGeral,
          totalH: "24h"
        };
      });
    }

    // Tradução do objeto do relatório para o formato que a função reconstruirFormulario espera
    const payload = {
      idRascunho: idProtocolo,
      quant: {
        delegacia: dadosRelatorio.delegacia,
        bos: dadosRelatorio.quant.bos,
        guias: dadosRelatorio.quant.guias,
        apreensoes: dadosRelatorio.quant.apre,
        presos: dadosRelatorio.quant.presos,
        medidas: dadosRelatorio.quant.medi,
        outros: dadosRelatorio.quant.outr,
        entrada: dadosRelatorio.entrada,
        saida: dadosRelatorio.saida,
        equipeResumo: dadosRelatorio.equipe.map(p => p.nome).join(" / "),
        observacoes: dadosRelatorio.observacoes,
        equipe: equipeFormatada
      },
      quali: dadosRelatorio.detalhes.flatMap(proc => 
        proc.vitimas.map(v => ({ tipo: proc.tipo, numero: proc.numero, crime: proc.crime, nomePessoa: v, papel: 'V' }))
        .concat(proc.infratores.map(i => ({ tipo: proc.tipo, numero: proc.numero, crime: proc.crime, nomePessoa: i, papel: 'I' })))
      )
    };

    return { sucesso: true, payload: payload };
  } catch (e) {
    return { sucesso: false, msg: e.toString() };
  }
}

/**
 * Remove todas as linhas associadas a um ID nas abas de Base para evitar duplicidade na retificação
 */
function removerDadosAntigosID(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ["Base_Quantitativos", "Base_Qualitativos", "Base_Equipe"];
  
  abas.forEach(nomeAba => {
    const sheet = ss.getSheetByName(nomeAba);
    const dados = sheet.getDataRange().getValues();
    // Percorre de baixo para cima para manter os índices íntegros durante a deleção
    for (let i = dados.length - 1; i >= 1; i--) {
      if (dados[i][0] === id) {
        sheet.deleteRow(i + 1);
      }
    }
  });
}