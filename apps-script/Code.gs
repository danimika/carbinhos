/**
 * Carbinhos — caderninho de receitas.
 *
 * Abas usadas (setup() cria as que faltam, sem apagar nada):
 *   Receitas              uma linha por receita
 *   Ingredientes_Receita  uma linha por ingrediente (coluna "grupo" vira coluna na tela)
 *   Preparo               uma linha por passo
 *   Unidades              conversor de volume (tudo em ml)
 *   Tags                  lista oficial de tags (de, para, tipo)
 *   Ingredients           catálogo de ingredientes (já existia; não é alterado)
 */
const SHEET_ID = '1GUX1zS-ZkvD4A6MVwc5X2D7mIFSyxNnvrMONBJgq9BQ';

const ABAS = {
  receitas: { nome: 'Receitas', cab: ['id', 'nome', 'categoria', 'rendimento', 'tempo_min', 'link_insta', 'foto_url', 'tags', 'obs'] },
  itens: { nome: 'Ingredientes_Receita', cab: ['receita_id', 'grupo', 'ordem', 'qtd', 'unidade', 'ingrediente', 'obs', 'a_gosto'] },
  preparo: { nome: 'Preparo', cab: ['receita_id', 'grupo', 'ordem', 'passo'] },
  unidades: { nome: 'Unidades', cab: ['codigo', 'nome', 'ml'] },
  tags: { nome: 'Tags', cab: ['de', 'para', 'tipo'] },
  catalogo: { nome: 'Ingredients', cab: ['Name', 'Keto', 'Carnivore', 'Obs'] },
};

const UNIDADES_PADRAO = [
  ['xicara', 'xícara', 240],
  ['colher_sopa', 'colher de sopa', 15],
  ['colher_sobremesa', 'colher de sobremesa', 10],
  ['colher_cha', 'colher de chá', 5],
  ['colher_cafe', 'colher de café', 2.5],
  ['ml', 'ml', 1],
  ['l', 'litro', 1000],
  ['un', 'unidade', ''],
  ['g', 'g', ''],
  ['kg', 'kg', ''],
  ['pitada', 'pitada', ''],
  ['a_gosto', 'a gosto', ''],
];

function ss_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Carbinhos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Pode rodar quantas vezes quiser: só cria o que falta. */
function setup() {
  const ss = ss_();
  Object.keys(ABAS).forEach(function (k) {
    const a = ABAS[k];
    const sh = ss.getSheetByName(a.nome) || ss.insertSheet(a.nome);
    garantirCabecalho_(sh, a.cab);
  });

  const un = ss.getSheetByName(ABAS.unidades.nome);
  if (un.getLastRow() < 2) un.getRange(2, 1, UNIDADES_PADRAO.length, 3).setValues(UNIDADES_PADRAO);

  if (ss.getSheetByName(ABAS.receitas.nome).getLastRow() < 2 &&
      ss.getSheetByName(ABAS.itens.nome).getLastRow() < 2) {
    exemplo_(ss);
  }

  validacoes_(ss);
}

function garantirCabecalho_(sh, cab) {
  const atual = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  let n = atual.length;
  while (n && !atual[n - 1].trim()) n--;
  const falta = cab.filter(function (c) { return atual.indexOf(c) === -1; });
  if (falta.length) sh.getRange(1, n + 1, 1, falta.length).setValues([falta]);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function coluna_(sh, nome) {
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  return cab.indexOf(nome) + 1;
}

function lista_(sh, col, origem, permitirOutros) {
  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInRange(origem, true)
    .setAllowInvalid(permitirOutros)
    .build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(regra);
}

function validacoes_(ss) {
  const rec = ss.getSheetByName(ABAS.receitas.nome);
  const ids = rec.getRange(2, coluna_(rec, 'id'), rec.getMaxRows() - 1, 1);

  const un = ss.getSheetByName(ABAS.unidades.nome);
  const codigos = un.getRange(2, coluna_(un, 'codigo'), un.getMaxRows() - 1, 1);

  const cat = ss.getSheetByName(ABAS.catalogo.nome);
  const nomesCat = cat.getRange(2, coluna_(cat, 'Name'), cat.getMaxRows() - 1, 1);

  const it = ss.getSheetByName(ABAS.itens.nome);
  lista_(it, coluna_(it, 'receita_id'), ids, false);
  lista_(it, coluna_(it, 'unidade'), codigos, false);
  // Ingrediente fora do catálogo só ganha um aviso (triângulo), não é bloqueado.
  lista_(it, coluna_(it, 'ingrediente'), nomesCat, true);
  // Marcado: o ingrediente é "a gosto" e qtd/unidade viram só uma sugestão.
  it.getRange(2, coluna_(it, 'a_gosto'), it.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

  const pr = ss.getSheetByName(ABAS.preparo.nome);
  lista_(pr, coluna_(pr, 'receita_id'), ids, false);
}

function exemplo_(ss) {
  const id = 'bolo-caneca';
  ss.getSheetByName(ABAS.receitas.nome).getRange(2, 1, 1, 9).setValues([[
    id, 'Bolo de caneca (exemplo)', 'Bolos', '1 caneca', 5, '', '', 'doce, sobremesa', 'Receita de exemplo — pode apagar.',
  ]]);
  const itens = [
    [id, 'Secos', 1, 4, 'colher_sopa', 'farinha de amêndoas', '', false],
    [id, 'Secos', 2, 1, 'colher_sopa', 'cacau em pó', '', false],
    [id, 'Secos', 3, 1, 'colher_sopa', 'adoçante', '', true],
    [id, 'Secos', 4, 1, 'colher_cha', 'fermento em pó', '', false],
    [id, 'Úmidos', 1, 1, 'un', 'ovo', '', false],
    [id, 'Úmidos', 2, 1, 'colher_sopa', 'manteiga', 'derretida', false],
    [id, 'Úmidos', 3, 2, 'colher_sopa', 'creme de leite', '', false],
  ];
  ss.getSheetByName(ABAS.itens.nome).getRange(2, 1, itens.length, 8).setValues(itens);
  const passos = [
    [id, '', 1, 'Misture os secos na caneca.'],
    [id, '', 2, 'Junte os úmidos e mexa até ficar homogêneo.'],
    [id, '', 3, 'Leve ao micro-ondas por 1min30.'],
  ];
  ss.getSheetByName(ABAS.preparo.nome).getRange(2, 1, passos.length, 4).setValues(passos);
}

/** Lê uma aba como lista de objetos {cabecalho: valor}. */
function ler_(ss, nome) {
  const sh = ss.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  const v = sh.getDataRange().getValues();
  const cab = v.shift().map(function (c) { return String(c).trim(); });
  return v
    // Checkbox desmarcado (false) conta como vazio, senão toda linha com caixinha vira "preenchida".
    .filter(function (l) { return l.some(function (c) { return c !== '' && c !== false; }); })
    .map(function (l) {
      const o = {};
      cab.forEach(function (c, i) {
        if (!c) return;
        const x = l[i];
        o[c] = x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : x;
      });
      return o;
    });
}

function txt_(v) {
  return String(v == null ? '' : v).trim();
}

/** Junta linhas filhas (ingredientes, passos) em grupos dentro de cada receita. */
function agrupar_(linhas, porId, campo, mapear) {
  linhas.forEach(function (l) {
    const r = porId[txt_(l.receita_id)];
    if (!r) return;
    const nomeG = txt_(l.grupo);
    let g = r[campo].filter(function (x) { return x.nome === nomeG; })[0];
    if (!g) {
      g = { nome: nomeG, itens: [] };
      r[campo].push(g);
    }
    g.itens.push(mapear(l));
  });
  Object.keys(porId).forEach(function (id) {
    porId[id][campo].forEach(function (g) {
      g.itens.sort(function (a, b) { return (Number(a.ordem) || 0) - (Number(b.ordem) || 0); });
    });
  });
}

/** Chamado pela tela. */
function getData() {
  const ss = ss_();
  const porId = {};
  const receitas = ler_(ss, ABAS.receitas.nome)
    .filter(function (r) { return txt_(r.id); })
    .map(function (r) {
      const o = {
        id: txt_(r.id),
        nome: txt_(r.nome),
        categoria: txt_(r.categoria),
        rendimento: txt_(r.rendimento),
        tempo: txt_(r.tempo_min),
        insta: txt_(r.link_insta),
        foto: txt_(r.foto_url),
        obs: txt_(r.obs),
        tags: txt_(r.tags).split(',').map(function (t) { return t.trim(); }).filter(String),
        grupos: [],
        passos: [],
      };
      porId[o.id] = o;
      return o;
    });

  agrupar_(ler_(ss, ABAS.itens.nome), porId, 'grupos', function (l) {
    return { ordem: l.ordem, qtd: l.qtd, unidade: txt_(l.unidade), ingrediente: txt_(l.ingrediente), obs: txt_(l.obs),
      aGosto: l.a_gosto === true || /^(sim|s|x|true|verdadeiro)$/i.test(txt_(l.a_gosto)) };
  });
  agrupar_(ler_(ss, ABAS.preparo.nome), porId, 'passos', function (l) {
    return { ordem: l.ordem, passo: txt_(l.passo) };
  });

  const tags = ler_(ss, ABAS.tags.nome)
    .filter(function (t) { return txt_(t.de); })
    .map(function (t) { return { tag: txt_(t.de), en: txt_(t.para), tipo: txt_(t.tipo) }; });

  const oficiais = {};
  tags.forEach(function (t) { oficiais[t.tag.toLowerCase()] = 1; });
  const desconhecidas = {};
  receitas.forEach(function (r) {
    r.tags.forEach(function (t) { if (!oficiais[t.toLowerCase()]) desconhecidas[t] = 1; });
  });

  const unidades = ler_(ss, ABAS.unidades.nome)
    .filter(function (u) { return txt_(u.codigo); })
    .map(function (u) { return { codigo: txt_(u.codigo), nome: txt_(u.nome), ml: Number(u.ml) || 0 }; });

  return {
    receitas: receitas,
    tags: tags,
    tagsDesconhecidas: Object.keys(desconhecidas),
    unidades: unidades,
  };
}
