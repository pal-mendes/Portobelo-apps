
// =========================
// File: ShAnuncios.gs - projeto ligado ao Google Sheets "Anúncios de semanas.gsheet"
// =========================

/***********************************************
* Fórmulas em uso no spreadsheet
 **********************************************/

const ANUNCIOS_SHEET = "Anúncios";
const TITULARES_SHEET = "Titulares"; 

const ANUNCIOS_HEADER = 4;
const TITULARES_HEADER = 2;

const COLS_TITULARES = {
  CT_SEMANAS: "Semanas",
  CT_TEL: "Telemóvel",
  CT_EMAIL: "e-mail",
  CT_PAGO: "Pago",
  CT_SALDO: "Saldo"
};

const COLS_ANUNCIOS = {
  CA_TEL: "Telemóvel",
  CA_EMAIL: "e-mail",
  CA_PAGO: "Pago",
  CA_SALDO: "Saldo",
  CA_SEMS: "Semanas"
};

/**
 * Função principal que atualiza os dados de uma linha específica no separador Anúncios.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetAnuncios
 * @param {number} row Linha a atualizar
 * @param {Object} maps Objeto contendo o mapeamento das colunas de ambos os sheets
 */
function updateAnuncioRowData_(sheetAnuncios, row, maps) {
  const ss = sheetAnuncios.getParent();
  const sheetTitulares = ss.getSheetByName(TITULARES_SHEET);
  
  // Obter dados da linha do anúncio
  const anuncioData = sheetAnuncios.getRange(row, 1, 1, sheetAnuncios.getLastColumn()).getValues()[0];
  // Guardamos uma cópia imutável do que foi lido da folha
  const dadosLidos = [...anuncioData];

  const emailBusca = anuncioData[maps.anuncios.CA_EMAIL - 1];
  const telBusca = anuncioData[maps.anuncios.CA_TEL - 1];
  //console.log("anuncioData=" + anuncioData + ", emailBusca=" + emailBusca + ", telBusca=" + telBusca);

  // Se não houver e-mail nem telemóvel, não há o que procurar
  if (!emailBusca && !telBusca) {
    return { linha: row, status: "Ignorado - Sem Contactos", lido: dadosLidos };
  } 

  // Obter dados dos Titulares usando a nova constante (ex: começa na linha 3 se header for 2)
  const linhaInicioTitulares = TITULARES_HEADER + 1;
  const numLinhasTitulares = sheetTitulares.getLastRow() - TITULARES_HEADER;
  const dataTitulares = sheetTitulares.getRange(linhaInicioTitulares, 1, numLinhasTitulares, sheetTitulares.getLastColumn()).getValues();

  let somaPago = 0;
  let somaSaldo = 0;
  let semanasLista = [];
  let encontrouCorrespondencia = false;
  
  const searchKey = emailBusca ? emailBusca.toString().toLowerCase().trim() : "";
  const searchTel = telBusca ? telBusca.toString().replace(/\D/g, '') : "";

  dataTitulares.forEach(r => {
    const ctEmail = r[maps.titulares.CT_EMAIL - 1] ? r[maps.titulares.CT_EMAIL - 1].toString().toLowerCase().trim() : "";
    const ctTelRaw = r[maps.titulares.CT_TEL - 1] ? r[maps.titulares.CT_TEL - 1].toString() : "";
    const ctPago = r[maps.titulares.CT_PAGO - 1] || 0;
    const ctSaldo = r[maps.titulares.CT_SALDO - 1] || 0;
    const ctSemanas = r[maps.titulares.CT_SEMANAS - 1] || "";

    let match = false;

    // DICA DE DEBUG: Se na tabela Titulares tiver mais de um e-mail na mesma célula separados por ;
    // a igualdade estrita (ctEmail === searchKey) irá falhar. Se for esse o caso, 
    // substitua a linha abaixo por: if (searchKey && ctEmail.includes(searchKey))
    if (searchKey && ctEmail.includes(searchKey)) {
      match = true;
    } else if (!searchKey && searchTel) {
      // Lógica de telemóvel para anúncios antigos ou sem e-mail
      const telefonesTitular = ctTelRaw.split(';').map(t => t.replace(/\D/g, ''));
      if (telefonesTitular.some(t => t.endsWith(searchTel) || searchTel.endsWith(t))) {
        match = true;
      }
    }

    if (match) {
      encontrouCorrespondencia = true;
      somaPago += Number(ctPago);
      somaSaldo += Number(ctSaldo);
      if (ctSemanas) semanasLista.push(ctSemanas);
    }
  });

  // Escrever resultados em CA_SALDO e CA_SEMS (colunas adjacentes)
  const concatSemanas = semanasLista.join('; ');
  // CORREÇÃO: Arredondar para 2 casas decimais (ex: 0.00)
  somaPago = Math.round(somaPago * 100) / 100;
  somaSaldo = Math.round(somaSaldo * 100) / 100;
  // Escrever na folha
  sheetAnuncios.getRange(row, maps.anuncios.CA_PAGO, 1, 3).setValues([[somaPago, somaSaldo, concatSemanas]]);

  // Atualizar a matriz original para mostrar como ficou
  anuncioData[maps.anuncios.CA_PAGO - 1] = somaPago;
  anuncioData[maps.anuncios.CA_SALDO - 1] = somaSaldo;
  anuncioData[maps.anuncios.CA_SEMS - 1] = concatSemanas;

  // Retornar o resultado detalhado para o Logger
  return {
    linha: row,
    status: encontrouCorrespondencia ? "Sucesso" : "Falha - Sem correspondência nos Titulares",
    emailBuscado: searchKey,
    telBuscado: searchTel,
    lido: dadosLidos,
    escrito: anuncioData
  };
}

/**
 * Gatilho automático para novas entradas ou edições no e-mail.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== ANUNCIOS_SHEET) return;

  const maps = getAllMaps_(sheet);
  const colEditada = e.range.getColumn();

  // Reage se a alteração for na coluna de e-mail
  if (colEditada === maps.anuncios.CA_EMAIL) {
    const row = e.range.getRow();
    //console.log("onEdit: row " + row);
    if (row <= ANUNCIOS_HEADER) return; // Ignora cabeçalhos
    updateAnuncioRowData_(sheet, row, maps);
  }
}

/**
 * Função para processar todo o histórico (Há que a correr antes de cada análise, para ter a certeza dos saldos serem os atuais).
 */
function processarHistoricoAnuncios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAnuncios = ss.getSheetByName(ANUNCIOS_SHEET);
  const maps = getAllMaps_(sheetAnuncios);
  
  const startRow = ANUNCIOS_HEADER + 1; // tblAnuncios começa em A4
  const lastRow = sheetAnuncios.getLastRow();
  if (lastRow < startRow) return;

  // Carrega todos os dados de uma vez para ser mais rápido a encontrar os vazios
  const data = sheetAnuncios.getRange(startRow, 1, lastRow - startRow + 1, sheetAnuncios.getLastColumn()).getValues();
  
  let contadorAtualizados = 0;
  
  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const rowNum = startRow + i; // i começa em 0, por isso startRow + 0 = primeira linha de dados
    
    const pagoAtual = rowData[maps.anuncios.CA_PAGO - 1];
    const saldoAtual = rowData[maps.anuncios.CA_SALDO - 1];
    const semsAtual = rowData[maps.anuncios.CA_SEMS - 1];
    
    const logResult = updateAnuncioRowData_(sheetAnuncios, rowNum, maps);
    console.log(`Pendente na linha ${rowNum} atualizado. Status: ${logResult.status}`);
    contadorAtualizados++;
  }
  
  // Opcional: Avisa o utilizador de quantos foram atualizados
  const ui = SpreadsheetApp.getUi();
  ui.alert(`Concluído! ${contadorAtualizados} anúncio(s) atualizado(s).`);
}

/**
 * Utilitário para mapear índices de colunas dinamicamente.
 */
function getAllMaps_(sheetAnuncios) {
  const ss = sheetAnuncios.getParent();
  const sheetTitulares = ss.getSheetByName(TITULARES_SHEET);
  
  return {
    anuncios: getColMap_(sheetAnuncios, COLS_ANUNCIOS, ANUNCIOS_HEADER), 
    titulares: getColMap_(sheetTitulares, COLS_TITULARES, TITULARES_HEADER) 
  };
}

function getColMap_(sheet, config, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (const [key, label] of Object.entries(config)) {
    const idx = headers.indexOf(label);
    if (idx !== -1) map[key] = idx + 1;
  }
  return map;
}