/*******************************************************
 * DGF - SISTEMA DE FÉRIAS
 * Google Apps Script + Google Sheets
 *
 * BANCO DE DADOS:
 * Cada cargo terá sua própria aba.
 *
 * ABAS DO SISTEMA:
 * CONFIG
 * SOLICITACOES
 * CARGO - NOME DO CARGO
 *
 * IMPORTANTE:
 * Execute CONFIGURAR_SISTEMA() uma vez antes de publicar.
 *******************************************************/

const CONFIG = {
  NOME_SISTEMA: 'Sistema de Férias - DGF',

  // TROQUE ESTA SENHA ANTES DE PUBLICAR
  SENHA_ADMIN: '123456',

  CARGOS_INICIAIS: [
    'ADMINISTRATIVO',
    'FINANCEIRO',
    'CONTABILIDADE',
    'ORÇAMENTO',
    'TESOURARIA',
    'PESSOAL',
    'OUTROS'
  ],

  MESES: [
    'Janeiro',
    'Fevereiro',
    'Março',
    'Abril',
    'Maio',
    'Junho',
    'Julho',
    'Agosto',
    'Setembro',
    'Outubro',
    'Novembro',
    'Dezembro'
  ],

  PREFIXO_CARGO: 'CARGO - ',

  ABA_CONFIG: 'CONFIG',
  ABA_SOLICITACOES: 'SOLICITACOES',

  CABECALHO_FUNCIONARIOS: [
    'ID',
    'DATA_CADASTRO',
    'NOME_COMPLETO',
    'CPF',
    'MATRICULA',
    'DATA_ADMISSAO',
    'DATA_NASCIMENTO',
    'CARGO',
    'MES_1',
    'MES_2',
    'VENDE_10_DIAS',
    'BLOQUEADO',
    'ULTIMA_ALTERACAO',
    'ALTERADO_POR',
    'MES_ULTIMO_BENEFICIO'
  ],

  CABECALHO_SOLICITACOES: [
    'ID_SOLICITACAO',
    'DATA_SOLICITACAO',
    'ID_FUNCIONARIO',
    'NOME_FUNCIONARIO',
    'CPF',
    'MATRICULA',
    'TIPO',
    'DESCRICAO',
    'STATUS',
    'DATA_ANALISE',
    'ADMIN_ANALISE',
    'OBSERVACAO_ADMIN'
  ]
};


/**
 * Abre o sistema.
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.NOME_SISTEMA)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * CONFIGURAÇÃO INICIAL
 */
function CONFIGURAR_SISTEMA() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbaConfig_(ss);
  criarAbaSolicitacoes_(ss);

  CONFIG.CARGOS_INICIAIS.forEach(function(cargo) {
    criarAbaCargo_(ss, cargo);
  });

  return 'Sistema configurado com sucesso.';
}


/**
 * Cria aba CONFIG.
 */
function criarAbaConfig_(ss) {

  let sheet = ss.getSheetByName(CONFIG.ABA_CONFIG);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ABA_CONFIG);
  }

  sheet.clear();

  sheet.getRange(1, 1).setValue('DGF - CONFIGURAÇÃO DO SISTEMA');

  sheet.getRange(3, 1).setValue('CARGOS');

  CONFIG.CARGOS_INICIAIS.forEach(function(cargo, index) {
    sheet.getRange(index + 4, 1).setValue(cargo);
  });

  sheet.getRange(3, 1).setFontWeight('bold');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.autoResizeColumn(1);
}


/**
 * Cria aba de solicitações.
 */
function criarAbaSolicitacoes_(ss) {

  let sheet = ss.getSheetByName(CONFIG.ABA_SOLICITACOES);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ABA_SOLICITACOES);
  }

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, CONFIG.CABECALHO_SOLICITACOES.length)
      .setValues([CONFIG.CABECALHO_SOLICITACOES]);

    formatarCabecalho_(sheet);
  }
}


/**
 * Cria uma aba para determinado cargo.
 * Também garante que a coluna do último benefício exista
 * nas abas antigas, sem apagar os dados existentes.
 */
function criarAbaCargo_(ss, cargo) {

  cargo = String(cargo || '').trim();

  if (!cargo) {
    throw new Error('Cargo inválido.');
  }

  const nomeAba = nomeAbaCargo_(cargo);

  let sheet = ss.getSheetByName(nomeAba);

  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);

    sheet
      .getRange(1, 1, 1, CONFIG.CABECALHO_FUNCIONARIOS.length)
      .setValues([CONFIG.CABECALHO_FUNCIONARIOS]);

    formatarCabecalho_(sheet);

    sheet.setFrozenRows(1);
  }

  const colunaBeneficio = CONFIG.CABECALHO_FUNCIONARIOS.length;

  const nomeColunaBeneficio =
    CONFIG.CABECALHO_FUNCIONARIOS[colunaBeneficio - 1];

  if (
    sheet.getRange(1, colunaBeneficio).getValue() !==
    nomeColunaBeneficio
  ) {
    sheet.getRange(1, colunaBeneficio)
      .setValue(nomeColunaBeneficio)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff');
  }

  return sheet;
}


/**
 * Formata cabeçalho.
 */
function formatarCabecalho_(sheet) {

  const ultimaColuna = sheet.getLastColumn();

  if (ultimaColuna < 1) return;

  sheet
    .getRange(1, 1, 1, ultimaColuna)
    .setFontWeight('bold')
    .setBackground('#0f172a')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);
}


/**
 * Normaliza nome de aba.
 */
function nomeAbaCargo_(cargo) {

  let nome = String(cargo)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();

  nome = nome.replace(/[\\\/\?\*\[\]\:]/g, '');
  nome = nome.substring(0, 70);

  return CONFIG.PREFIXO_CARGO + nome;
}


/**
 * DADOS GERAIS
 */
function obterDadosIniciais() {

  return {
    sistema: CONFIG.NOME_SISTEMA,
    meses: CONFIG.MESES,
    cargos: obterCargos_(),
    hoje: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'America/Sao_Paulo',
      'yyyy-MM-dd'
    )
  };
}


/**
 * Obtém cargos da aba CONFIG.
 */
function obterCargos_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ABA_CONFIG);

  if (!sheet) {
    return CONFIG.CARGOS_INICIAIS;
  }

  const ultimaLinha = sheet.getLastRow();

  if (ultimaLinha < 4) {
    return CONFIG.CARGOS_INICIAIS;
  }

  const valores = sheet
    .getRange(4, 1, ultimaLinha - 3, 1)
    .getValues()
    .flat()
    .map(String)
    .map(v => v.trim())
    .filter(Boolean);

  return [...new Set(valores)];
}


/**
 * CADASTRO DE FUNCIONÁRIO
 */
function cadastrarFuncionario(dados) {

  const lock = LockService.getScriptLock();

  lock.waitLock(20000);

  try {

    validarDadosCadastro_(dados);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const cpf = somenteNumeros_(dados.cpf);
    const matricula = String(dados.matricula).trim();

    if (funcionarioExiste_(cpf, matricula)) {
      throw new Error(
        'Já existe um funcionário cadastrado com este CPF ou matrícula.'
      );
    }

    const cargo = String(dados.cargo).trim();

    criarAbaCargo_(ss, cargo);

    const sheet = ss.getSheetByName(nomeAbaCargo_(cargo));

    const id = gerarId_();

    const agora = new Date();

    const linha = [
      id,
      agora,
      String(dados.nome).trim().toUpperCase(),
      cpf,
      matricula,
      converterData_(dados.dataAdmissao),
      converterData_(dados.dataNascimento),
      cargo,
      String(dados.mes1),
      String(dados.mes2),
      dados.vende10Dias === true || dados.vende10Dias === 'true'
        ? 'SIM'
        : 'NÃO',
      'SIM',
      agora,
      'CADASTRO',
      String(dados.mesUltimoBeneficio || '')
    ];

    sheet.appendRow(linha);

    return {
      sucesso: true,
      id: id,
      mensagem:
        'Cadastro realizado com sucesso. Seus dados agora estão bloqueados para edição.'
    };

  } finally {

    lock.releaseLock();
  }
}


/**
 * Validação do cadastro.
 */
function validarDadosCadastro_(dados) {

  if (!dados) {
    throw new Error('Dados não recebidos.');
  }

  const nome = String(dados.nome || '').trim();

  if (nome.length < 5) {
    throw new Error('Informe o nome completo.');
  }

  const cpf = somenteNumeros_(dados.cpf);

  if (cpf.length !== 11) {
    throw new Error('CPF inválido.');
  }

  const matricula = String(dados.matricula || '').trim();

  if (!matricula) {
    throw new Error('Informe a matrícula.');
  }

  if (!dados.dataAdmissao) {
    throw new Error('Informe a data de admissão.');
  }

  if (!dados.dataNascimento) {
    throw new Error('Informe a data de nascimento.');
  }

  if (!dados.cargo) {
    throw new Error('Selecione o cargo.');
  }

  if (!dados.mes1 || !dados.mes2) {
    throw new Error('Escolha os dois meses.');
  }

  if (String(dados.mes1) === String(dados.mes2)) {
    throw new Error('Os dois meses precisam ser diferentes.');
  }
}


/**
 * Verifica duplicidade.
 */
function funcionarioExiste_(cpf, matricula) {

  const funcionarios = obterTodosFuncionarios_();

  return funcionarios.some(function(f) {

    const cpfIgual =
      somenteNumeros_(f.cpf) === somenteNumeros_(cpf);

    const matriculaIgual =
      String(f.matricula).trim().toLowerCase() ===
      String(matricula).trim().toLowerCase();

    return cpfIgual || matriculaIgual;
  });
}


/**
 * CONSULTA DO FUNCIONÁRIO
 */
function consultarFuncionario(cpf, matricula) {

  cpf = somenteNumeros_(cpf);
  matricula = String(matricula || '').trim();

  if (cpf.length !== 11 || !matricula) {
    throw new Error('Informe CPF e matrícula.');
  }

  const funcionarios = obterTodosFuncionarios_();

  const encontrado = funcionarios.find(function(f) {

    return somenteNumeros_(f.cpf) === cpf &&
      String(f.matricula).trim().toLowerCase() ===
      matricula.toLowerCase();

  });

  if (!encontrado) {
    throw new Error('Funcionário não encontrado.');
  }

  return prepararFuncionarioPublico_(encontrado);
}


/**
 * Prepara funcionário para tela pública.
 */
function prepararFuncionarioPublico_(f) {

  return {
    id: f.id,
    nome: f.nome,
    cpf: mascararCPF_(f.cpf),
    matricula: f.matricula,
    dataAdmissao: formatarData_(f.dataAdmissao),
    dataNascimento: formatarData_(f.dataNascimento),
    cargo: f.cargo,
    mes1: f.mes1,
    mes2: f.mes2,
    mesUltimoBeneficio: f.mesUltimoBeneficio || '',
    vende10Dias: f.vende10Dias === 'SIM',
    mesesServico: calcularMesesServico_(f.dataAdmissao),
    statusPrazo: calcularStatusPrazo_(f.dataAdmissao)
  };
}


/**
 * SOLICITAÇÃO DE ALTERAÇÃO
 */
function solicitarAlteracao(dados) {

  if (!dados || !dados.id || !dados.descricao) {
    throw new Error('Informe o motivo da solicitação.');
  }

  const funcionario = localizarFuncionarioPorId_(dados.id);

  if (!funcionario) {
    throw new Error('Funcionário não encontrado.');
  }

  let sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.ABA_SOLICITACOES);

  if (!sheet) {
    criarAbaSolicitacoes_(
      SpreadsheetApp.getActiveSpreadsheet()
    );

    sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.ABA_SOLICITACOES);
  }

  const idSolicitacao = gerarId_();

  const linha = [
    idSolicitacao,
    new Date(),
    funcionario.id,
    funcionario.nome,
    funcionario.cpf,
    funcionario.matricula,
    String(dados.tipo || 'ALTERAÇÃO'),
    String(dados.descricao).trim(),
    'PENDENTE',
    '',
    '',
    ''
  ];

  sheet.appendRow(linha);

  return {
    sucesso: true,
    protocolo: idSolicitacao,
    mensagem: 'Solicitação enviada para a administração.'
  };
}


/**
 * LOGIN ADMINISTRATIVO
 */
function loginAdministrador(senha) {

  if (String(senha || '') !== CONFIG.SENHA_ADMIN) {
    throw new Error('Senha administrativa incorreta.');
  }

  const token = Utilities.getUuid();

  CacheService
    .getScriptCache()
    .put(
      'ADMIN_' + token,
      'AUTORIZADO',
      1800
    );

  return {
    sucesso: true,
    token: token,
    mensagem: 'Acesso administrativo autorizado.'
  };
}


/**
 * Valida token administrativo.
 */
function validarAdmin_(token) {

  if (!token) {
    throw new Error('Sessão administrativa inválida.');
  }

  const autorizado =
    CacheService
      .getScriptCache()
      .get('ADMIN_' + token);

  if (autorizado !== 'AUTORIZADO') {
    throw new Error(
      'Sua sessão administrativa expirou. Entre novamente.'
    );
  }

  return true;
}


/**
 * Painel administrativo.
 */
function obterPainelAdmin(token) {

  validarAdmin_(token);

  const funcionarios = obterTodosFuncionarios_();

  const solicitacoes = obterSolicitacoes_();

  return {
    funcionarios: funcionarios.map(function(f) {
      return prepararFuncionarioAdmin_(f);
    }),
    solicitacoes: solicitacoes
  };
}


/**
 * Dados administrativos.
 */
function prepararFuncionarioAdmin_(f) {

  return {
    id: f.id,
    nome: f.nome,
    cpf: f.cpf,
    matricula: f.matricula,
    dataAdmissao: formatarDataInput_(f.dataAdmissao),
    dataNascimento: formatarDataInput_(f.dataNascimento),
    cargo: f.cargo,
    mes1: f.mes1,
    mes2: f.mes2,
    mesUltimoBeneficio: f.mesUltimoBeneficio || '',
    vende10Dias: f.vende10Dias === 'SIM',
    mesesServico: calcularMesesServico_(f.dataAdmissao),
    statusPrazo: calcularStatusPrazo_(f.dataAdmissao)
  };
}


/**
 * Salva edição administrativa.
 */
function editarFuncionario(token, dados) {

  validarAdmin_(token);

  if (!dados || !dados.id) {
    throw new Error('Funcionário inválido.');
  }

  const localizado = localizarFuncionarioPorId_(dados.id);

  if (!localizado) {
    throw new Error('Funcionário não encontrado.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const origem = localizado.sheet;

  const linhaNumero = localizado.row;

  const cargoNovo = String(dados.cargo || '').trim();

  if (!cargoNovo) {
    throw new Error('Informe o cargo.');
  }

  criarAbaCargo_(ss, cargoNovo);

  const dadosLinha = [
    localizado.id,
    localizado.dataCadastro || new Date(),
    String(dados.nome).trim().toUpperCase(),
    somenteNumeros_(dados.cpf),
    String(dados.matricula).trim(),
    converterData_(dados.dataAdmissao),
    converterData_(dados.dataNascimento),
    cargoNovo,
    String(dados.mes1),
    String(dados.mes2),
    dados.vende10Dias === true || dados.vende10Dias === 'true'
      ? 'SIM'
      : 'NÃO',
    'SIM',
    new Date(),
    obterEmailAdmin_(),
    String(dados.mesUltimoBeneficio || '')
  ];

  if (dados.mes1 === dados.mes2) {
    throw new Error('Os dois meses precisam ser diferentes.');
  }

  const destino = ss.getSheetByName(nomeAbaCargo_(cargoNovo));

  const cargoMudou =
    nomeAbaCargo_(localizado.cargo) !==
    nomeAbaCargo_(cargoNovo);

  if (cargoMudou) {

    destino.appendRow(dadosLinha);

    origem.deleteRow(linhaNumero);

  } else {

    origem
      .getRange(
        linhaNumero,
        1,
        1,
        CONFIG.CABECALHO_FUNCIONARIOS.length
      )
      .setValues([dadosLinha]);

  }

  return {
    sucesso: true,
    mensagem: 'Funcionário atualizado com sucesso.'
  };
}


/**
 * Exclui funcionário.
 */
function excluirFuncionario(token, id) {

  validarAdmin_(token);

  const localizado = localizarFuncionarioPorId_(id);

  if (!localizado) {
    throw new Error('Funcionário não encontrado.');
  }

  localizado.sheet.deleteRow(localizado.row);

  return {
    sucesso: true,
    mensagem: 'Funcionário excluído.'
  };
}


/**
 * SOLICITAÇÕES ADMINISTRATIVAS
 */
function analisarSolicitacao(
  token,
  idSolicitacao,
  novoStatus,
  observacao
) {

  validarAdmin_(token);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(CONFIG.ABA_SOLICITACOES);

  if (!sheet) {
    throw new Error('Aba de solicitações não encontrada.');
  }

  const valores = sheet.getDataRange().getValues();

  for (let i = 1; i < valores.length; i++) {

    if (String(valores[i][0]) === String(idSolicitacao)) {

      sheet.getRange(i + 1, 9).setValue(novoStatus);
      sheet.getRange(i + 1, 10).setValue(new Date());
      sheet.getRange(i + 1, 11).setValue(obterEmailAdmin_());
      sheet.getRange(i + 1, 12).setValue(observacao || '');

      return {
        sucesso: true,
        mensagem: 'Solicitação atualizada.'
      };
    }
  }

  throw new Error('Solicitação não encontrada.');
}


/**
 * RANKING
 */
function obterRanking(cargoFiltro) {

  const funcionarios = obterTodosFuncionarios_();

  let lista = funcionarios;

  if (cargoFiltro && cargoFiltro !== 'TODOS') {

    lista = lista.filter(function(f) {
      return String(f.cargo) === String(cargoFiltro);
    });
  }

  const hoje = new Date();
  const mesAtual = hoje.getMonth() + 1;

  lista.sort(function(a, b) {

    const distanciaA1 =
      distanciaMes_(mesAtual, obterNumeroMes_(a.mes1));

    const distanciaB1 =
      distanciaMes_(mesAtual, obterNumeroMes_(b.mes1));

    if (distanciaA1 !== distanciaB1) {
      return distanciaA1 - distanciaB1;
    }

    const distanciaA2 =
      distanciaMes_(mesAtual, obterNumeroMes_(a.mes2));

    const distanciaB2 =
      distanciaMes_(mesAtual, obterNumeroMes_(b.mes2));

    if (distanciaA2 !== distanciaB2) {
      return distanciaA2 - distanciaB2;
    }

    const idadeA =
      calcularIdade_(a.dataNascimento);

    const idadeB =
      calcularIdade_(b.dataNascimento);

    return idadeB - idadeA;
  });

  return lista.map(function(f, index) {

    return {
      posicao: index + 1,
      id: f.id,
      nome: f.nome,
      cargo: f.cargo,
      mes1: f.mes1,
      mes2: f.mes2,
      vende10Dias: f.vende10Dias === 'SIM',
      idade: calcularIdade_(f.dataNascimento),
      mesesServico: calcularMesesServico_(f.dataAdmissao),
      statusPrazo: calcularStatusPrazo_(f.dataAdmissao)
    };

  });
}


/**
 * Distância cíclica entre meses.
 */
function distanciaMes_(atual, escolhido) {

  let distancia = escolhido - atual;

  if (distancia < 0) {
    distancia += 12;
  }

  return distancia;
}


/**
 * Obtém número do mês.
 */
function obterNumeroMes_(nome) {

  const index =
    CONFIG.MESES
      .map(m => m.toLowerCase())
      .indexOf(String(nome).toLowerCase());

  return index >= 0 ? index + 1 : 99;
}


/**
 * DASHBOARD
 */
function obterDashboard() {

  const funcionarios = obterTodosFuncionarios_();

  const resumo = {
    total: funcionarios.length,
    atencao: 0,
    alerta: 0,
    urgente: 0,
    vencidas: 0,
    normais: 0,
    abono: 0
  };

  funcionarios.forEach(function(f) {

    const status =
      calcularStatusPrazo_(f.dataAdmissao);

    if (status.codigo === 'ATENCAO') {
      resumo.atencao++;
    } else if (status.codigo === 'ALERTA') {
      resumo.alerta++;
    } else if (status.codigo === 'URGENTE') {
      resumo.urgente++;
    } else if (status.codigo === 'VENCIDA') {
      resumo.vencidas++;
    } else {
      resumo.normais++;
    }

    if (f.vende10Dias === 'SIM') {
      resumo.abono++;
    }
  });

  return resumo;
}


/**
 * FUNÇÕES INTERNAS
 */

/**
 * Obtém todos os funcionários de todas as abas de cargo.
 */
function obterTodosFuncionarios_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ss.getSheets();

  const lista = [];

  sheets.forEach(function(sheet) {

    const nome = sheet.getName();

    if (!nome.startsWith(CONFIG.PREFIXO_CARGO)) {
      return;
    }

    const ultimaLinha = sheet.getLastRow();

    if (ultimaLinha < 2) {
      return;
    }

    const valores =
      sheet
        .getRange(
          2,
          1,
          ultimaLinha - 1,
          CONFIG.CABECALHO_FUNCIONARIOS.length
        )
        .getValues();

    valores.forEach(function(row, index) {

      if (!row[0]) return;

      lista.push({
        id: String(row[0]),
        dataCadastro: row[1],
        nome: String(row[2] || ''),
        cpf: String(row[3] || ''),
        matricula: String(row[4] || ''),
        dataAdmissao: row[5],
        dataNascimento: row[6],
        cargo: String(row[7] || ''),
        mes1: String(row[8] || ''),
        mes2: String(row[9] || ''),
        vende10Dias: String(row[10] || 'NÃO'),
        bloqueado: String(row[11] || 'SIM'),
        ultimaAlteracao: row[12],
        alteradoPor: String(row[13] || ''),
        mesUltimoBeneficio: String(row[14] || ''),
        sheet: sheet,
        row: index + 2
      });

    });

  });

  return lista;
}


/**
 * Localiza funcionário por ID.
 */
function localizarFuncionarioPorId_(id) {

  const lista = obterTodosFuncionarios_();

  return lista.find(function(f) {
    return String(f.id) === String(id);
  });
}


/**
 * Obtém solicitações.
 */
function obterSolicitacoes_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CONFIG.ABA_SOLICITACOES);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const valores =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        CONFIG.CABECALHO_SOLICITACOES.length
      )
      .getValues();

  return valores
    .filter(row => row[0])
    .map(function(row) {

      return {
        id: String(row[0]),
        data: formatarDataHora_(row[1]),
        idFuncionario: String(row[2]),
        nome: String(row[3]),
        cpf: String(row[4]),
        matricula: String(row[5]),
        tipo: String(row[6]),
        descricao: String(row[7]),
        status: String(row[8]),
        dataAnalise: formatarDataHora_(row[9]),
        admin: String(row[10]),
        observacao: String(row[11])
      };

    })
    .reverse();
}


/**
 * STATUS DE PRAZO
 */
function calcularMesesServico_(dataAdmissao) {

  if (!dataAdmissao) {
    return 0;
  }

  const inicio =
    dataAdmissao instanceof Date
      ? new Date(dataAdmissao)
      : new Date(dataAdmissao);

  const hoje = new Date();

  let meses =
    (hoje.getFullYear() - inicio.getFullYear()) * 12 +
    (hoje.getMonth() - inicio.getMonth());

  if (hoje.getDate() < inicio.getDate()) {
    meses--;
  }

  return Math.max(0, meses);
}


/**
 * Status administrativo.
 */
function calcularStatusPrazo_(dataAdmissao) {

  const meses = calcularMesesServico_(dataAdmissao);

  if (meses < 12) {

    return {
      codigo: 'NORMAL',
      nome: 'Ainda não completou 12 meses',
      descricao: 'Período aquisitivo em andamento.',
      classe: 'normal',
      meses: meses
    };

  }

  if (meses >= 12 && meses <= 15) {

    return {
      codigo: 'ATENCAO',
      nome: 'ATENÇÃO',
      descricao: 'Direito adquirido; planejar o agendamento.',
      classe: 'atencao',
      meses: meses
    };

  }

  if (meses >= 16 && meses <= 19) {

    return {
      codigo: 'ALERTA',
      nome: 'ALERTA',
      descricao: 'Recomenda-se definir o período de férias.',
      classe: 'alerta',
      meses: meses
    };

  }

  if (meses >= 20 && meses <= 23) {

    return {
      codigo: 'URGENTE',
      nome: 'URGENTE',
      descricao: 'Prioridade máxima para análise e agendamento.',
      classe: 'urgente',
      meses: meses
    };

  }

  return {
    codigo: 'VENCIDA',
    nome: 'VERIFICAR FÉRIAS VENCIDAS',
    descricao: 'Prazo de 24 meses ou mais. Conferir imediatamente a situação.',
    classe: 'vencida',
    meses: meses
  };
}


/**
 * UTILITÁRIOS
 */
function gerarId_() {

  return Utilities
    .getUuid()
    .replace(/-/g, '')
    .substring(0, 16)
    .toUpperCase();
}


function somenteNumeros_(valor) {

  return String(valor || '')
    .replace(/\D/g, '');
}


function converterData_(valor) {

  if (!valor) {
    throw new Error('Data inválida.');
  }

  if (valor instanceof Date) {
    return valor;
  }

  const partes =
    String(valor).split('-');

  if (partes.length === 3) {

    return new Date(
      Number(partes[0]),
      Number(partes[1]) - 1,
      Number(partes[2])
    );
  }

  const partesBR =
    String(valor).split('/');

  if (partesBR.length === 3) {

    return new Date(
      Number(partesBR[2]),
      Number(partesBR[1]) - 1,
      Number(partesBR[0])
    );
  }

  throw new Error('Data inválida.');
}


function formatarData_(data) {

  if (!data) return '';

  const d =
    data instanceof Date
      ? data
      : new Date(data);

  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || 'America/Sao_Paulo',
    'dd/MM/yyyy'
  );
}


function formatarDataInput_(data) {

  if (!data) return '';

  const d =
    data instanceof Date
      ? data
      : new Date(data);

  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || 'America/Sao_Paulo',
    'yyyy-MM-dd'
  );
}


function formatarDataHora_(data) {

  if (!data) return '';

  const d =
    data instanceof Date
      ? data
      : new Date(data);

  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || 'America/Sao_Paulo',
    'dd/MM/yyyy HH:mm'
  );
}


function mascararCPF_(cpf) {

  const n = somenteNumeros_(cpf);

  if (n.length !== 11) {
    return '***.***.***-**';
  }

  return (
    n.substring(0, 3) +
    '.***.***-' +
    n.substring(9)
  );
}


function calcularIdade_(dataNascimento) {

  if (!dataNascimento) return 0;

  const nascimento =
    dataNascimento instanceof Date
      ? new Date(dataNascimento)
      : new Date(dataNascimento);

  const hoje = new Date();

  let idade =
    hoje.getFullYear() -
    nascimento.getFullYear();

  const mes =
    hoje.getMonth() -
    nascimento.getMonth();

  if (
    mes < 0 ||
    (mes === 0 && hoje.getDate() < nascimento.getDate())
  ) {
    idade--;
  }

  return idade;
}


function obterEmailAdmin_() {

  try {

    const email =
      Session.getActiveUser().getEmail();

    return email || 'ADMIN';

  } catch (e) {

    return 'ADMIN';
  }
}


/**
 * Atualiza o cabeçalho do último benefício nas abas existentes.
 *
 * Execute esta função uma vez após substituir o Código.gs.
 * Ela não apaga nem altera os registros existentes.
 */
function ATUALIZAR_COLUNA_ULTIMO_BENEFICIO() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const col = CONFIG.CABECALHO_FUNCIONARIOS.length;

  const titulo = CONFIG.CABECALHO_FUNCIONARIOS[col - 1];

  ss.getSheets().forEach(function(sheet) {

    if (sheet.getName().startsWith(CONFIG.PREFIXO_CARGO)) {

      sheet.getRange(1, col)
        .setValue(titulo)
        .setFontWeight('bold')
        .setBackground('#0f172a')
        .setFontColor('#ffffff');

      sheet.setFrozenRows(1);
    }

  });

  return 'Coluna de último benefício atualizada.';
}
