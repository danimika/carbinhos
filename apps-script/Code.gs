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
  unidades: { nome: 'Unidades', cab: ['codigo', 'nome', 'ml', 'abrev'] },
  tags: { nome: 'Tags', cab: ['de', 'para', 'tipo'] },
  catalogo: { nome: 'Ingredients', cab: ['Name', 'Keto', 'Carnivore', 'Obs'] },
};

// codigo, nome, ml (para converter volume), abreviação (usada no modo de preparo)
const UNIDADES_PADRAO = [
  ['xicara', 'xícara', 240, 'xíc.'],
  ['colher_sopa', 'colher de sopa', 15, 'cs'],
  ['colher_sobremesa', 'colher de sobremesa', 10, 'csb'],
  ['colher_cha', 'colher de chá', 5, 'cc'],
  ['colher_cafe', 'colher de café', 2.5, 'ccafé'],
  ['ml', 'ml', 1, 'ml'],
  ['l', 'litro', 1000, 'l'],
  ['un', 'unidade', '', 'unid.'],
  ['g', 'g', '', 'g'],
  ['kg', 'kg', '', 'kg'],
  ['pitada', 'pitada', '', 'pitada'],
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
  if (un.getLastRow() < 2) un.getRange(2, 1, UNIDADES_PADRAO.length, 4).setValues(UNIDADES_PADRAO);
  else migrarUnidades_(un);

  if (ss.getSheetByName(ABAS.receitas.nome).getLastRow() < 2 &&
      ss.getSheetByName(ABAS.itens.nome).getLastRow() < 2) {
    exemplo_(ss);
  }

  validacoes_(ss);
}

/** Preenche abreviações que faltam e tira "a gosto" das unidades (agora é a caixinha). */
function migrarUnidades_(sh) {
  const cab = cabecalho_(sh);
  const cCod = cab.indexOf('codigo'), cAb = cab.indexOf('abrev');
  const n = ultimaLinha_(sh);
  if (n < 2) return;
  const v = sh.getRange(2, 1, n - 1, cab.length).getValues();
  const padrao = {};
  UNIDADES_PADRAO.forEach(function (u) { padrao[u[0]] = u[3]; });
  for (let i = v.length - 1; i >= 0; i--) {
    const cod = txt_(v[i][cCod]);
    if (cod === 'a_gosto') { sh.deleteRow(i + 2); continue; }
    if (!txt_(v[i][cAb]) && padrao[cod]) sh.getRange(i + 2, cAb + 1).setValue(padrao[cod]);
  }
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
    .map(function (u) { return { codigo: txt_(u.codigo), nome: txt_(u.nome), ml: Number(u.ml) || 0, abrev: txt_(u.abrev) }; });

  // Sugestões para o campo "ingrediente" do formulário: catálogo + já usados.
  const nomesIng = {};
  ler_(ss, ABAS.catalogo.nome).forEach(function (c) { if (txt_(c.Name)) nomesIng[txt_(c.Name)] = 1; });
  receitas.forEach(function (r) {
    r.grupos.forEach(function (g) {
      g.itens.forEach(function (it) { if (it.ingrediente) nomesIng[it.ingrediente] = 1; });
    });
  });

  return {
    receitas: receitas,
    tags: tags,
    tagsDesconhecidas: Object.keys(desconhecidas),
    unidades: unidades,
    ingredientes: Object.keys(nomesIng).sort(),
  };
}

/* ---------- cadastro e edição (chamado pelo formulário) ---------- */

function slug_(s) {
  return txt_(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}

/**
 * "0,5", "1/2" e "1 1/2" viram número (a tela mostra como ½, 1½...).
 * Precisa: se gravar o texto "1/2", a planilha transforma em data (1º de fevereiro).
 * Qualquer outro texto vai com apóstrofo para ficar como texto mesmo.
 */
function qtd_(v) {
  const s = txt_(v);
  if (!s) return '';
  if (/^\d+([.,]\d+)?$/.test(s)) return Number(s.replace(',', '.'));
  let m = s.match(/^(\d+)\s+(\d+)\/(\d+)$/);
  if (m) return Number(m[1]) + Number(m[2]) / Number(m[3]);
  m = s.match(/^(\d+)\/(\d+)$/);
  if (m) return Number(m[1]) / Number(m[2]);
  return "'" + s;
}

function cabecalho_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (c) { return String(c).trim(); });
}

/** Última linha com algo na coluna A (getLastRow conta as caixinhas vazias). */
function ultimaLinha_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const v = sh.getRange(1, 1, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (txt_(v[i][0])) return i + 1;
  return 1;
}

/** Escreve objetos {cabecalho: valor} logo depois da última linha preenchida. */
function anexar_(sh, objs) {
  if (!objs.length) return;
  const cab = cabecalho_(sh);
  const linhas = objs.map(function (o) {
    return cab.map(function (c) { return o[c] === undefined ? '' : o[c]; });
  });
  const ini = ultimaLinha_(sh) + 1;
  const falta = ini + linhas.length - 1 - sh.getMaxRows();
  if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta);
  sh.getRange(ini, 1, linhas.length, cab.length).setValues(linhas);
}

function apagarLinhas_(sh, colNome, id) {
  const n = ultimaLinha_(sh);
  if (n < 2) return;
  const col = cabecalho_(sh).indexOf(colNome) + 1;
  const v = sh.getRange(2, col, n - 1, 1).getValues();
  // Apaga de baixo para cima, em blocos de linhas seguidas (cada chamada é lenta).
  let i = v.length - 1;
  while (i >= 0) {
    if (txt_(v[i][0]) !== id) { i--; continue; }
    let fim = i;
    while (i - 1 >= 0 && txt_(v[i - 1][0]) === id) i--;
    sh.deleteRows(i + 2, fim - i + 1);
    i--;
  }
}

/**
 * Cria ou atualiza uma receita. Na edição (r.idOriginal), as linhas antigas
 * das 3 abas são trocadas pelas novas. Tags novas entram na aba Tags.
 */
function salvarReceita(r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const nome = txt_(r.nome);
    if (!nome) throw new Error('A receita precisa de um nome.');
    const id = slug_(r.id || nome);
    if (!id) throw new Error('Não consegui gerar um identificador para a receita.');
    const original = txt_(r.idOriginal);

    const ids = ler_(ss, ABAS.receitas.nome).map(function (x) { return txt_(x.id); });
    if (id !== original && ids.indexOf(id) > -1) {
      throw new Error('Já existe uma receita com o identificador "' + id + '".');
    }

    const shR = ss.getSheetByName(ABAS.receitas.nome);
    const shI = ss.getSheetByName(ABAS.itens.nome);
    const shP = ss.getSheetByName(ABAS.preparo.nome);
    if (original) {
      apagarLinhas_(shR, 'id', original);
      apagarLinhas_(shI, 'receita_id', original);
      apagarLinhas_(shP, 'receita_id', original);
    }

    const tags = (r.tags || []).map(txt_).filter(String);
    anexar_(shR, [{
      id: id, nome: nome, categoria: txt_(r.categoria), rendimento: txt_(r.rendimento),
      tempo_min: qtd_(r.tempo), link_insta: txt_(r.insta), foto_url: txt_(r.foto),
      tags: tags.join(', '), obs: txt_(r.obs),
    }]);

    // Unidade digitada que não existe (ex.: "fatia") vira unidade nova na aba Unidades.
    const unidades = ler_(ss, ABAS.unidades.nome);
    const resolverUnidade = function (valor) {
      const v = txt_(valor);
      if (!v) return '';
      const k = v.toLowerCase();
      const achada = unidades.filter(function (u) {
        return [u.codigo, u.nome, u.abrev].some(function (x) { return txt_(x).toLowerCase() === k; });
      })[0];
      if (achada) return txt_(achada.codigo);
      const nova = { codigo: slug_(v).replace(/-/g, '_'), nome: v, ml: '', abrev: v };
      unidades.push(nova);
      anexar_(ss.getSheetByName(ABAS.unidades.nome), [nova]);
      return nova.codigo;
    };

    const itens = [];
    (r.grupos || []).forEach(function (g) {
      let ordem = 0;
      (g.itens || []).forEach(function (it) {
        if (!txt_(it.ingrediente)) return;
        itens.push({
          receita_id: id, grupo: txt_(g.nome), ordem: ++ordem, qtd: qtd_(it.qtd), unidade: resolverUnidade(it.unidade),
          ingrediente: txt_(it.ingrediente), obs: txt_(it.obs), a_gosto: !!it.aGosto,
        });
      });
    });
    anexar_(shI, itens);

    const passos = [];
    (r.passos || []).forEach(function (g) {
      let ordem = 0;
      (g.itens || []).forEach(function (p) {
        if (!txt_(p.passo)) return;
        passos.push({ receita_id: id, grupo: txt_(g.nome), ordem: ++ordem, passo: txt_(p.passo) });
      });
    });
    anexar_(shP, passos);

    const oficiais = ler_(ss, ABAS.tags.nome).map(function (t) { return txt_(t.de).toLowerCase(); });
    const novas = tags.filter(function (t) { return oficiais.indexOf(t.toLowerCase()) < 0; });
    anexar_(ss.getSheetByName(ABAS.tags.nome), novas.map(function (t) { return { de: t }; }));

    SpreadsheetApp.flush();
    return { id: id, dados: getData() };
  } finally {
    lock.releaseLock();
  }
}
