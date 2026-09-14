# Gerador de Catálogo de Importação

Versão web (100% client-side) do script Python de geração do `ModeloProduto.xlsx`
a partir da planilha de E-invoice (Brazil) e da planilha de Operadores/Fabricantes.

Nenhum arquivo é enviado a um servidor — tudo roda no navegador do usuário,
usando a biblioteca [SheetJS](https://sheetjs.com) para ler e escrever `.xlsx`.

## Estrutura

- `index.html` — interface (dois campos de upload + botão de processar/baixar)
- `style.css` — estilo visual
- `dicionario.js` — dicionário de tradução e regras de limpeza (copiados do script original)
- `app.js` — lógica de transformação (equivalente às funções do `main.py`) + interação da UI

## Como publicar no GitHub Pages

1. Crie um repositório novo no GitHub (ou use um existente) e envie estes 4 arquivos
   (`index.html`, `style.css`, `dicionario.js`, `app.js`) para a raiz do repositório
   (ou para a pasta `/docs`, se preferir).
2. No GitHub, vá em **Settings → Pages**.
3. Em **Source**, selecione a branch (ex.: `main`) e a pasta (`/root` ou `/docs`).
4. Salve. Em alguns minutos o GitHub fornecerá uma URL do tipo
   `https://seu-usuario.github.io/nome-do-repositorio/`.

Também funciona perfeitamente abrindo o `index.html` direto no navegador, sem hospedagem alguma.

## O que a ferramenta faz

1. Lê a planilha de E-invoice (equivalente a `planilhaBrazil`) e a de Operadores
   (equivalente a `planilhaFabricantes`) — sempre a primeira aba de cada arquivo.
2. Remove itens duplicados pela coluna `Item`.
3. Traduz `Item Description` usando o dicionário de termos.
4. Monta o campo `RES` (`SKU: ... - MARCA CHANEL - ...`).
5. Aplica as regras de limpeza (`Tratamento`), incluindo remoção de campos como
   `COMPRIMENTO:`, `FORRO:`, `MANGA:`, `MALHA:`, `SOLA:` conforme a categoria do item,
   remoção de resíduos `nan`, colapso de hifens repetidos e remoção de marcadores
   como `(CANADA ONLY)` / `(FOR FILLING ONLY)`.
6. Traduz o texto tratado (`Traducao`).
7. Para cada item, localiza o fabricante pelo país de origem (`Origin`) na planilha
   de Operadores (com regra especial para Itália/Sunglasses → `LUXXOTICA.` /
   `CHANEL COORDINATION`) e só inclui o item no catálogo final se um fabricante
   correspondente for encontrado.
8. Gera `ModeloProduto.xlsx` com 4 abas: `Produtos`, `Identificadores`,
   `Fabricantes`, `UnidadesNegocio` — mesmos cabeçalhos do script original.

## Observações de fidelidade ao script original

- Onde o Python deixaria um valor ausente virar `NaN` (e, ao entrar num f-string,
  aparecer como o texto `nan`), a versão web replica esse mesmo comportamento —
  isso importa porque a rotina de limpeza remove especificamente padrões
  `campo: nan`.
- A extração de fabricante (`codigo_fabricante`) e a limpeza de texto
  (`limpeza_tratamento`, `traducao`, `extracao_denominacao`) foram portadas
  função a função, mantendo a mesma ordem de aplicação das regras.
