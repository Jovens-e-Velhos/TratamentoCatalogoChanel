// Planilha Cliente: itens do E-invoice cujo CÓDIGO PRINCIPAL (coluna "Item")
// NÃO está na primeira coluna do catálogo de produtos (equivalente a um XLOOKUP sem retorno).
// Independente do app.js.

(function () {
  const inputBrazil   = document.getElementById('file-brazil');
  const inputCatalogo = document.getElementById('file-catalogo');
  const dzCatalogo    = document.getElementById('dz-catalogo');
  const btnGerar      = document.getElementById('btn-cliente');
  const btnDownload   = document.getElementById('btn-download-cliente');
  const logEl         = document.getElementById('log-cliente');
  let blobCliente = null;

  // ── Log ──
  function limparLog() { logEl.innerHTML = ''; }
  function log(msg, tipo) {
    const s = document.createElement('span');
    s.className = 'log-line ' + (tipo === 'step' ? 'log-step' : tipo === 'ok' ? 'log-ok' : tipo === 'err' ? 'log-err' : '');
    s.textContent = msg;
    logEl.appendChild(s);
    logEl.scrollTop = logEl.scrollHeight;
  }

  // ── Utilidades ──
  const vazio = v => v === null || v === undefined || v === '';
  const norm  = v => String(v).trim().toUpperCase();

  // Busca a coluna pelo nome do cabeçalho (sem diferenciar maiúsculas/espaços)
  function getCol(row, nome) {
    const alvo = nome.trim().toLowerCase();
    for (const k of Object.keys(row)) {
      if (k.trim().toLowerCase() === alvo) return row[k];
    }
    return null;
  }

  async function lerPlanilha(file, opts) {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, opts);
  }

  // ── Habilitar botão e nome do arquivo ──
  function atualizarEstado() {
    const f = inputCatalogo.files && inputCatalogo.files[0];
    dzCatalogo.querySelector('[data-role="filename"]').textContent = f ? f.name : 'Nenhum arquivo selecionado';
    dzCatalogo.classList.toggle('has-file', !!f);
    btnGerar.disabled = !f;
  }
  inputCatalogo.addEventListener('change', atualizarEstado);

  // Arrastar e soltar no campo do catálogo
  dzCatalogo.addEventListener('dragover', e => { e.preventDefault(); dzCatalogo.classList.add('dz-dragover'); });
  dzCatalogo.addEventListener('dragleave', () => dzCatalogo.classList.remove('dz-dragover'));
  dzCatalogo.addEventListener('drop', e => {
    e.preventDefault();
    dzCatalogo.classList.remove('dz-dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const dt = new DataTransfer();
      dt.items.add(e.dataTransfer.files[0]);
      inputCatalogo.files = dt.files;
      atualizarEstado();
    }
  });

  // ── Montagem das linhas ──
  function montarLinha(row) {
    return {
      'CÓDIGO PRINCIPAL': getCol(row, 'Item'),
      'CODIGO SECUNDARIO': null,
      'DESCRICAO RESUMIDO DO PRODUTO': null,
      'UNIDADE MEDIDA': null,
      'NCM': null,
      'QUANTIDADE': getCol(row, 'Quantity'),
      'PESO LIQUIDO UNITÁRIO': getCol(row, 'Unit Net Weight'),
      'PESO BRUTO UNITÁRIO': null,
      'CÓDIGO DO EXPORTADOR': null,
      'DESCRICAO PORTUGUES': null,
      'VALOR UNITARIO ITEM': getCol(row, 'Unit Price'),
      'DESTAQUE DE ANUENCIA': null,
      'CÓDIGO DO FABRICANTE': null,
      'PAÍS DE ORIGEM': getCol(row, 'Origin'),
      'NVE': null,
      'DESCRICAO DO PRODUTO PARA NOTA FISCAL': null,
      'FATURA': getCol(row, 'Invoice Number'),
      'QUEBRA AUX': null,
      'Reservado': null,
      'Moeda': getCol(row, 'Currency'),
      'AREA': null,
      'NECESSITA LI': null,
      'ORDEM DE COMPRA': getCol(row, 'Boutique Code'),
      'CODIGO(GTIN / EAN)': getCol(row, 'Bar Code'),
    };
  }

  // ── Execução ──
  async function gerar() {
    blobCliente = null;
    btnDownload.hidden = true;
    limparLog();
    btnGerar.disabled = true;

    try {
      const fBrazil = inputBrazil.files && inputBrazil.files[0];
      const fCat    = inputCatalogo.files && inputCatalogo.files[0];

      if (!fBrazil) { log('Selecione primeiro o E-invoice FA BRAZIL no campo 01.', 'err'); return; }
      if (!fCat)    { log('Selecione o catálogo de produtos.', 'err'); return; }

      log('Lendo catálogo de produtos…', 'step');
      const catRows = await lerPlanilha(fCat, { header: 1, defval: null });
      const codigosCatalogo = new Set();
      for (const r of catRows) {
        if (r && !vazio(r[0])) codigosCatalogo.add(norm(r[0]));
      }
      log(`Catálogo: ${codigosCatalogo.size} códigos únicos na primeira coluna.`);

      log('Lendo E-invoice…', 'step');
      const dados = await lerPlanilha(fBrazil, { defval: null });
      log(`E-invoice: ${dados.length} linhas.`);

      log('Cruzando CÓDIGO PRINCIPAL x catálogo…', 'step');
      const linhas = [];
      let jaCadastrados = 0, semItem = 0;

      for (const row of dados) {
        const item = getCol(row, 'Item');
        if (vazio(item)) { semItem++; continue; }
        if (codigosCatalogo.has(norm(item))) { jaCadastrados++; continue; }
        linhas.push(montarLinha(row));
      }

      log(`Já cadastrados no catálogo (ignorados): ${jaCadastrados}`);
      if (semItem) log(`Linhas sem "Item" (ignoradas): ${semItem}`);

      if (!linhas.length) {
        log('Nenhum item novo: todos os códigos já estão no catálogo.', 'ok');
        return;
      }

      const out = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(out, XLSX.utils.json_to_sheet(linhas), 'Itens');
      const buf = XLSX.write(out, { bookType: 'xlsx', type: 'array' });
      blobCliente = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      btnDownload.hidden = false;
      log(`✓ Planilha Cliente: ${linhas.length} linhas (itens não cadastrados).`, 'ok');
    } catch (e) {
      console.error(e);
      log('ERRO: ' + e.message, 'err');
    } finally {
      btnGerar.disabled = !(inputCatalogo.files && inputCatalogo.files[0]);
    }
  }

  btnGerar.addEventListener('click', gerar);

  btnDownload.addEventListener('click', () => {
    if (!blobCliente) return;
    const url = URL.createObjectURL(blobCliente);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Planilha Cliente.xlsx';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
})();
