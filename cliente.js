// Nova saída: planilha do cliente (layout de importação).
// Independente do app.js: lê o mesmo arquivo do campo #file-brazil.

(function () {
  const btnProcessar = document.getElementById('btn-processar');
  const btnDownload  = document.getElementById('btn-download-cliente');
  const inputBrazil  = document.getElementById('file-brazil');
  let blobCliente = null;

  // Busca a coluna pelo nome do cabeçalho (sem diferenciar maiúsculas/espaços)
  function getCol(row, nome) {
    const alvo = nome.trim().toLowerCase();
    for (const k of Object.keys(row)) {
      if (k.trim().toLowerCase() === alvo) return row[k];
    }
    return null;
  }

  const vazio = v => v === null || v === undefined || v === '';

  function montarLinhas(origem) {
    const linhas = [];
    for (const row of origem) {
      const item    = getCol(row, 'Item');
      const qtd     = getCol(row, 'Quantity');
      const pesoLiq = getCol(row, 'Unit Net Weight');
      const preco   = getCol(row, 'Unit Price');
      const pais    = getCol(row, 'Origin');
      const fatura  = getCol(row, 'Invoice Number');
      const moeda   = getCol(row, 'Currency');
      const oc      = getCol(row, 'Boutique Code');
      const ean     = getCol(row, 'Bar Code');

      if ([item, qtd, pesoLiq, preco, pais, fatura, moeda, oc, ean].every(vazio)) continue;

      linhas.push({
        'CÓDIGO PRINCIPAL': item,
        'CODIGO SECUNDARIO': null,
        'DESCRICAO RESUMIDO DO PRODUTO': null,
        'UNIDADE MEDIDA': null,
        'NCM': null,
        'QUANTIDADE': qtd,
        'PESO LIQUIDO UNITÁRIO': pesoLiq,
        'PESO BRUTO UNITÁRIO': null,
        'CÓDIGO DO EXPORTADOR': null,
        'DESCRICAO PORTUGUES': null,
        'VALOR UNITARIO ITEM': preco,
        'DESTAQUE DE ANUENCIA': null,
        'CÓDIGO DO FABRICANTE': null,
        'PAÍS DE ORIGEM': pais,
        'NVE': null,
        'DESCRICAO DO PRODUTO PARA NOTA FISCAL': null,
        'FATURA': fatura,
        'QUEBRA AUX': null,
        'Reservado': null,
        'Moeda': moeda,
        'AREA': null,
        'NECESSITA LI': null,
        'ORDEM DE COMPRA': oc,
        'CODIGO(GTIN / EAN)': ean,
      });
    }
    return linhas;
  }

  async function gerar() {
    blobCliente = null;
    btnDownload.hidden = true;
    const file = inputBrazil.files && inputBrazil.files[0];
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const dados = XLSX.utils.sheet_to_json(ws, { defval: null });
      const linhas = montarLinhas(dados);
      if (!linhas.length) { console.warn('Planilha Cliente: nenhuma linha gerada.'); return; }

      const out = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(out, XLSX.utils.json_to_sheet(linhas), 'Itens');
      const buf = XLSX.write(out, { bookType: 'xlsx', type: 'array' });
      blobCliente = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      btnDownload.hidden = false;
    } catch (e) {
      console.error('Erro na Planilha Cliente:', e);
    }
  }

  // Roda junto com o processamento principal (sem interferir nele)
  btnProcessar.addEventListener('click', gerar);

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
