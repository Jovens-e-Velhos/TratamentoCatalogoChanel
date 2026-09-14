'use strict';

/* =========================================================================
 *  CABEÇALHOS DAS ABAS DE SAÍDA (idênticos ao script original)
 * ========================================================================= */

const HEADERS_PRODUTOS = [
  'Código Interno', 'NCM', 'Denominação', 'Descrição', 'Modalidade',
  'Prioridade', 'Observação', 'Detalhamento Complementar', 'Ativo',
  'Mensagem Importação', 'REMOVER', 'Backlog', 'Novo Código Interno',
];

const HEADERS_IDENTIFICADORES = [
  'Produto(Código Interno)', 'Identificador', 'Mensagem Importação', 'REMOVER',
];

const HEADERS_FABRICANTES = [
  'Produto(Código Interno)', 'País', 'Operador Estrangeiro(Código Interno)',
  'CNPJ', 'Mensagem Importação', 'REMOVER',
];

const HEADERS_UNIDADES_NEGOCIO = [
  'Produto(Código Interno)', 'Unidade Negócio(Código Interno)',
  'Mensagem Importação', 'REMOVER',
];

/* =========================================================================
 *  UTILITÁRIOS
 * ========================================================================= */

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Espelha o comportamento do pandas: célula vazia/ausente vira NaN, que ao
// entrar em um f-string do Python é renderizado literalmente como "nan".
// Isso é usado propositalmente pelas regras de limpeza a seguir.
function pyStr(value) {
  if (value === null || value === undefined) return 'nan';
  return String(value);
}

/* =========================================================================
 *  FUNÇÕES EQUIVALENTES ÀS DO SCRIPT PYTHON
 * ========================================================================= */

// extracao_denominacao(self)
function extracaoDenominacao(texto) {
  if (typeof texto !== 'string') return texto;
  const marca = /\bMARCA\b/i.exec(texto);
  if (!marca) return texto.trim();
  const posMarca = marca.index;
  const antes = texto.slice(0, posMarca);
  const posSeparador = antes.lastIndexOf(' - ');
  if (posSeparador === -1) return texto.slice(0, posMarca).trim();
  return texto.slice(0, posSeparador).trim();
}

// codigo_fabricante(self, planilhaFabricantes, sub_department, origem)
function codigoFabricante(nome, linhasFabricantes, subDepartment, origem) {
  if (typeof nome !== 'string' || nome.trim() === '') return null;

  let nomeBusca = nome.trim().toUpperCase();
  const subDeptFormatado = (subDepartment ?? '').toString().trim().toUpperCase();
  const origemFormatado = (origem ?? '').toString().trim().toUpperCase();

  if (origemFormatado === 'IT' || ['IT', 'CHANEL COORDINATION', 'LUXXOTICA.'].includes(nomeBusca)) {
    nomeBusca = (subDeptFormatado === 'SUNGLASSES') ? 'LUXXOTICA.' : 'CHANEL COORDINATION';
  }

  const encontrado = linhasFabricantes.find(
    (r) => (r['Nome'] ?? '').toString().trim().toUpperCase() === nomeBusca
  );

  return encontrado ? (encontrado['Código Interno'] ?? null) : null;
}

// remocao_campo(texto, campo)
function remocaoCampo(texto, campo) {
  // Réplica literal do f-string original: rf'\s+{campo}+:+\s*\S+'
  const pattern = new RegExp(`\\s+${campo}+:+\\s*\\S+`, 'gi');
  return texto.replace(pattern, '');
}

// limpeza_tratamento(self)
function limpezaTratamento(texto) {
  if (typeof texto !== 'string') return texto;

  let atual = texto;
  const upper = atual.toUpperCase();

  for (const [campo, categorias] of Object.entries(REGRAS_LIMPEZA)) {
    if (categorias.some((cat) => upper.includes(cat))) {
      atual = remocaoCampo(atual, campo);
    }
  }

  let tratamento = atual.replace(/\s+\w+:+\s*nan\b/g, '');
  tratamento = tratamento.replace(/(\s*-\s*)+/g, ' - ');

  const padraoPattern = new RegExp(PADRAO_REMOCAO.map(escapeRegExp).join('|'), 'gi');
  tratamento = tratamento.replace(padraoPattern, '');

  tratamento = tratamento.replace(/^[ \-]+|[ \-]+$/g, '').trim();
  return tratamento;
}

// traducao(self)
const TERMOS_ORDENADOS = Object.keys(DICIONARIO_TRADUCAO).sort((a, b) => b.length - a.length);

function traducao(texto) {
  if (typeof texto !== 'string') return texto;
  let resposta = texto;
  for (const termo of TERMOS_ORDENADOS) {
    const regex = new RegExp(`\\b${escapeRegExp(termo)}\\b`, 'gi');
    resposta = resposta.replace(regex, DICIONARIO_TRADUCAO[termo]);
  }
  return resposta;
}

// dedup por 'Item', mantendo a primeira ocorrência (equivalente a drop_duplicates)
function dropDuplicatesByItem(linhas) {
  const vistos = new Set();
  const saida = [];
  for (const linha of linhas) {
    const chave = linha['Item'];
    if (!vistos.has(chave)) {
      vistos.add(chave);
      saida.push(linha);
    }
  }
  return saida;
}

// criacao_planilha_catalogo(planilhaOg)
function criacaoPlanilhaCatalogo(planilhaBrazil, planilhaFabricantes) {
  const produtos = [];
  const identificadores = [];
  const fabricantes = [];
  const unidadeNegocio = [];

  const mapaNome = {};
  for (const r of planilhaFabricantes) {
    mapaNome[r['País']] = r['Nome'];
  }

  for (const row of planilhaBrazil) {
    const codigoInterno = row['Item'];
    const origem = row['Origin'];
    const detalhamentoComplementar = row['Traducao'];
    const nomeFabricantes = Object.prototype.hasOwnProperty.call(mapaNome, origem)
      ? mapaNome[origem]
      : 'Não encontrado';
    const subDept = row['Sub Department Description'] ?? '';

    const produtoRow = {
      'Código Interno': codigoInterno,
      'NCM': null,
      'Denominação': extracaoDenominacao(detalhamentoComplementar),
      'Descrição': detalhamentoComplementar,
      'Modalidade': 'Importação',
      'Prioridade': 30,
      'Observação': null,
      'Detalhamento Complementar': detalhamentoComplementar,
      'Ativo': 'Sim',
      'Mensagem Importação': null,
      'REMOVER': null,
      'Backlog': null,
      'Novo Código Interno': null,
    };

    const identificadoresRow = {
      'Produto(Código Interno)': codigoInterno,
      'Identificador': codigoInterno,
      'Mensagem Importação': null,
      'REMOVER': null,
    };

    unidadeNegocio.push(Object.fromEntries(HEADERS_UNIDADES_NEGOCIO.map((h) => [h, null])));

    const codigoFabricanteValor = codigoFabricante(nomeFabricantes, planilhaFabricantes, subDept, origem);

    if (codigoFabricanteValor !== null && codigoFabricanteValor !== undefined) {
      produtos.push(produtoRow);
      identificadores.push(identificadoresRow);
      fabricantes.push({
        'Produto(Código Interno)': codigoInterno,
        'País': origem,
        'Operador Estrangeiro(Código Interno)': codigoFabricanteValor,
        'CNPJ': null,
        'Mensagem Importação': null,
        'REMOVER': null,
      });
    }
  }

  return { produtos, identificadores, fabricantes, unidadeNegocio, totalLinhas: planilhaBrazil.length };
}

/* =========================================================================
 *  LEITURA DE ARQUIVOS (.xlsx)
 * ========================================================================= */

function readFirstSheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('Falha ao ler o arquivo.'));
    reader.readAsArrayBuffer(file);
  });
}

/* =========================================================================
 *  FLUXO PRINCIPAL (equivalente a principal())
 * ========================================================================= */

async function processar(fileBrazil, fileFabricantes, log) {
  log('Lendo planilha E-invoice FA BRAZIL…', 'step');
  let planilhaBrazil = await readFirstSheet(fileBrazil);
  log(`${planilhaBrazil.length} linha(s) lida(s).`, 'ok');

  log('Lendo planilha de Operadores/Fabricantes…', 'step');
  const planilhaFabricantes = await readFirstSheet(fileFabricantes);
  log(`${planilhaFabricantes.length} linha(s) lida(s).`, 'ok');

  log('Removendo itens duplicados (coluna "Item")…', 'step');
  planilhaBrazil = dropDuplicatesByItem(planilhaBrazil);
  log(`${planilhaBrazil.length} item(ns) único(s) restante(s).`, 'ok');

  log('Traduzindo "Item Description"…', 'step');
  for (const row of planilhaBrazil) {
    row['Item Description'] = traducao(row['Item Description']);
  }

  log('Montando SKU consolidado ("RES")…', 'step');
  for (const row of planilhaBrazil) {
    row['RES'] = `SKU: ${pyStr(row['Item'])}-${pyStr(row['Class Description'])} - MARCA CHANEL - ${pyStr(row['Item Composition'])}`;
  }

  log('Aplicando regras de limpeza ("Tratamento")…', 'step');
  for (const row of planilhaBrazil) {
    row['Tratamento'] = limpezaTratamento(row['RES']);
  }

  log('Traduzindo texto tratado ("Traducao")…', 'step');
  for (const row of planilhaBrazil) {
    row['Traducao'] = traducao(row['Tratamento']);
  }

  log('Gerando catálogo de produtos e vínculos de fabricante…', 'step');
  const resultado = criacaoPlanilhaCatalogo(planilhaBrazil, planilhaFabricantes);
  log(`${resultado.produtos.length} produto(s) gerado(s).`, 'ok');

  return resultado;
}

function gerarWorkbook(resultado) {
  const wb = XLSX.utils.book_new();

  const wsProdutos = XLSX.utils.json_to_sheet(resultado.produtos, { header: HEADERS_PRODUTOS });
  const wsIdentificadores = XLSX.utils.json_to_sheet(resultado.identificadores, { header: HEADERS_IDENTIFICADORES });
  const wsFabricantes = XLSX.utils.json_to_sheet(resultado.fabricantes, { header: HEADERS_FABRICANTES });
  const wsUnidadeNegocio = XLSX.utils.json_to_sheet(resultado.unidadeNegocio, { header: HEADERS_UNIDADES_NEGOCIO });

  XLSX.utils.book_append_sheet(wb, wsProdutos, 'Produtos');
  XLSX.utils.book_append_sheet(wb, wsIdentificadores, 'Identificadores');
  XLSX.utils.book_append_sheet(wb, wsFabricantes, 'Fabricantes');
  XLSX.utils.book_append_sheet(wb, wsUnidadeNegocio, 'UnidadesNegocio');

  const arrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Blob([arrayBuffer], { type: 'application/octet-stream' });
}

/* =========================================================================
 *  UI
 * ========================================================================= */

(function initUI() {
  const dzBrazil = document.getElementById('dz-brazil');
  const dzFabricantes = document.getElementById('dz-fabricantes');
  const inputBrazil = document.getElementById('file-brazil');
  const inputFabricantes = document.getElementById('file-fabricantes');
  const btnProcessar = document.getElementById('btn-processar');
  const btnDownload = document.getElementById('btn-download');
  const logEl = document.getElementById('log');
  const summaryBlock = document.getElementById('summary-block');
  const statsEl = document.getElementById('stats');

  let downloadUrl = null;

  function log(message, kind) {
    if (logEl.querySelector('.log-idle')) logEl.innerHTML = '';
    const line = document.createElement('span');
    line.className = 'log-line' + (kind === 'step' ? ' log-step' : kind === 'ok' ? ' log-ok' : kind === 'err' ? ' log-err' : '');
    line.textContent = message;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  }

  function resetLog() {
    logEl.innerHTML = '';
  }

  function updateDropzoneLabel(dropzone, file) {
    const nameEl = dropzone.querySelector('[data-role="filename"]');
    if (file) {
      nameEl.textContent = file.name;
      dropzone.classList.add('has-file');
      dropzone.classList.remove('dz-error');
    } else {
      nameEl.textContent = 'Nenhum arquivo selecionado';
      dropzone.classList.remove('has-file');
    }
  }

  function refreshButtonState() {
    btnProcessar.disabled = !(inputBrazil.files[0] && inputFabricantes.files[0]);
  }

  inputBrazil.addEventListener('change', () => {
    updateDropzoneLabel(dzBrazil, inputBrazil.files[0]);
    refreshButtonState();
  });

  inputFabricantes.addEventListener('change', () => {
    updateDropzoneLabel(dzFabricantes, inputFabricantes.files[0]);
    refreshButtonState();
  });

  // Arrastar e soltar: aceita apenas .xlsx/.xls, atribui ao <input> correspondente
  // e reaproveita o mesmo caminho do evento "change" acima.
  function isPlanilha(file) {
    return file && /\.(xlsx|xls)$/i.test(file.name);
  }

  function setupDragAndDrop(dropzone, input) {
    let dragCounter = 0;

    dropzone.addEventListener('dragenter', (e) => {
      e.preventDefault();
      dragCounter++;
      dropzone.classList.add('dz-dragover');
    });

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    });

    dropzone.addEventListener('dragleave', () => {
      dragCounter = Math.max(0, dragCounter - 1);
      if (dragCounter === 0) dropzone.classList.remove('dz-dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dragCounter = 0;
      dropzone.classList.remove('dz-dragover');

      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (!file) return;

      if (!isPlanilha(file)) {
        dropzone.classList.add('dz-error');
        updateDropzoneLabel(dropzone, null);
        const nameEl = dropzone.querySelector('[data-role="filename"]');
        nameEl.textContent = 'Formato inválido — use .xlsx ou .xls';
        return;
      }

      const dt = new DataTransfer();
      dt.items.add(file);
      input.files = dt.files;
      input.dispatchEvent(new Event('change'));
    });
  }

  setupDragAndDrop(dzBrazil, inputBrazil);
  setupDragAndDrop(dzFabricantes, inputFabricantes);

  // Evita que o navegador abra o arquivo se o usuário soltar fora das zonas de upload.
  ['dragover', 'drop'].forEach((evt) => {
    window.addEventListener(evt, (e) => {
      if (!e.target.closest('.dropzone')) e.preventDefault();
    });
  });

  function renderStats(resultado) {
    const semFabricante = resultado.totalLinhas - resultado.produtos.length;
    const stats = [
      { label: 'Produtos gerados', value: resultado.produtos.length },
      { label: 'Identificadores', value: resultado.identificadores.length },
      { label: 'Vínculos de fabricante', value: resultado.fabricantes.length },
      { label: 'Sem fabricante correspondente', value: semFabricante, warn: semFabricante > 0 },
    ];

    statsEl.innerHTML = '';
    for (const s of stats) {
      const wrap = document.createElement('div');
      const dt = document.createElement('p');
      dt.className = 'stat-label';
      dt.textContent = s.label;
      const dd = document.createElement('p');
      dd.className = 'stat-value' + (s.warn ? ' warn' : '');
      dd.textContent = s.value;
      wrap.appendChild(dt);
      wrap.appendChild(dd);
      statsEl.appendChild(wrap);
    }
    summaryBlock.hidden = false;
  }

  btnProcessar.addEventListener('click', async () => {
    resetLog();
    summaryBlock.hidden = true;
    btnDownload.hidden = true;
    btnProcessar.disabled = true;

    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      downloadUrl = null;
    }

    try {
      const resultado = await processar(inputBrazil.files[0], inputFabricantes.files[0], log);

      log('Montando ModeloProduto.xlsx…', 'step');
      const blob = gerarWorkbook(resultado);
      downloadUrl = URL.createObjectURL(blob);
      btnDownload.hidden = false;
      log('Concluído.', 'ok');

      renderStats(resultado);
    } catch (err) {
      console.error(err);
      log(`Erro: ${err.message || err}`, 'err');
    } finally {
      refreshButtonState();
    }
  });

  btnDownload.addEventListener('click', () => {
    if (!downloadUrl) return;
    const a = document.createElement('a');
    a.href = downloadUrl;
    a.download = 'ModeloProduto.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  });
})();
