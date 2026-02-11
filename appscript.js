/**
 * LOKO-VOŽNJE API
 * Google Apps Script za automatsko evidentiranje vožnji
 */

// ============================================================================
// KONFIGURACIJA
// ============================================================================

const CONFIG = {
  TEMPLATE_SHEET: 'TEMPLATE',
  DATA_START_ROW: 16,
  
  // Defaultne vrijednosti
  DEFAULTS: {
    marka: 'BMW',
    regBroj: 'RI1479P',
    relacija: 'RIJEKA'
  },
  
  // DODAJ OVO:
  STOPA_NADOKNADE: 0.5, // EUR po kilometru
  
  // Opcije za izvješće (dropdown)
  IZVJESCE_OPCIJE: [
    'prijevoz do ureda',
    'posjeta klijentu',
    'službeni put',
    'prijevoz na sastanak',
    'inspekcijski nadzor',
    'servisiranje vozila'
  ],
  
  // Mjeseci na hrvatskom
  MJESECI: [
    'sijecanj', 'veljaca', 'ozujak', 'travanj', 'svibanj', 'lipanj',
    'srpanj', 'kolovoz', 'rujan', 'listopad', 'studeni', 'prosinac'
  ]
};

function testAPI() {
  // Simuliraj POST request
  const testData = {
    prijedeniKm: 15,
    izvjesce: 'službeni put',
    datum: '2025-02-10',
    relacija: 'RIJEKA'
  };
  
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(mockEvent);
  Logger.log('Result: ' + result.getContent());
}

// ============================================================================
// GLAVNI API ENDPOINT
// ============================================================================

function doPost(e) {
  try {
    let data;
    
    // Pokušaj parsirati iz postData.contents (fetch API)
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } 
    // Ili iz form parametra (form submission)
    else if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    }
    else {
      return createResponse(400, 'Nema podataka');
    }
    
    // Validacija
    if (!data.prijedeniKm || data.prijedeniKm <= 0) {
      return createResponse(400, 'Prijeđeni km je obavezan i mora biti > 0');
    }
    
    if (!data.izvjesce) {
      return createResponse(400, 'Izvješće je obavezno');
    }
    
    // Datum procesiranje
    const datum = data.datum ? new Date(data.datum) : new Date();
    const mjesec = CONFIG.MJESECI[datum.getMonth()];
    const godina = datum.getFullYear();
    const sheetName = `${mjesec}-${godina}`;
    
    // Osiguraj da sheet postoji
    const sheet = getOrCreateSheet(sheetName, mjesec, godina);
    
    // Dohvati zadnje završno stanje
    const pocetnoStanje = getLastZavrsnoStanje(sheet);
    const zavrsnoStanje = pocetnoStanje + data.prijedeniKm;
    
    // DEBUG - dodaj log
    Logger.log('DEBUG: pocetnoStanje = ' + pocetnoStanje);
    Logger.log('DEBUG: prijedeniKm = ' + data.prijedeniKm);
    Logger.log('DEBUG: zavrsnoStanje = ' + zavrsnoStanje);
    
    // Formatiraj vrijeme (zaokruženo na 10 minuta)
    const vrijeme = formatVrijemeZaokruzeno(datum);
    Logger.log('DEBUG: vrijeme = ' + vrijeme);
    
    // Pripremi podatke za redak
    const rowData = [
      '', '', // A i B kolone su prazne
      formatDatum(datum),                    // C (3) - Datum
      CONFIG.DEFAULTS.marka,                 // D (4) - Marka vozila
      CONFIG.DEFAULTS.regBroj,               // E (5) - Registracijski broj
      pocetnoStanje,                         // F (6) - Početno stanje
      zavrsnoStanje,                         // G (7) - Završno stanje
      data.relacija || CONFIG.DEFAULTS.relacija, // H (8) - Relacija
      vrijeme,                               // I (9) - Vrijeme
      data.prijedeniKm,                      // J (10) - Prijeđeni km
      '',                                    // K (11) - Nadoknada (prazno)
      data.izvjesce                          // L (12) - Izvješće
    ];
    
    Logger.log('DEBUG: rowData = ' + JSON.stringify(rowData));
    
    // Dodaj redak
    addRow(sheet, rowData);
    
    // Ažuriraj ukupno
    updateUkupno(sheet);
    
    return createResponse(200, 'Vožnja uspješno evidentirana', {
      sheetName: sheetName,
      datum: formatDatum(datum),
      pocetnoStanje: pocetnoStanje,
      zavrsnoStanje: zavrsnoStanje,
      prijedeniKm: data.prijedeniKm
    });
    
  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    Logger.log('ERROR Stack: ' + error.stack);
    return createResponse(500, 'Greška: ' + error.message);
  }
}

// Test GET endpoint
function doGet(e) {
  // Ako ima submission parametri - tretirati kao POST
  if (e.parameter && e.parameter.prijedeniKm) {
    try {
      const data = {
        prijedeniKm: parseInt(e.parameter.prijedeniKm),
        izvjesce: e.parameter.izvjesce,
        relacija: e.parameter.relacija || CONFIG.DEFAULTS.relacija
      };
      
      // Reuse doPost logiku
      const datum = new Date();
      const mjesec = CONFIG.MJESECI[datum.getMonth()];
      const godina = datum.getFullYear();
      const sheetName = `${mjesec}-${godina}`;
      
      const sheet = getOrCreateSheet(sheetName, mjesec, godina);
      const pocetnoStanje = getLastZavrsnoStanje(sheet);
      const zavrsnoStanje = pocetnoStanje + data.prijedeniKm;
      const vrijeme = formatVrijemeZaokruzeno(datum);
      
      const rowData = [
        '', '',
        formatDatum(datum),
        CONFIG.DEFAULTS.marka,
        CONFIG.DEFAULTS.regBroj,
        pocetnoStanje,
        zavrsnoStanje,
        data.relacija,
        vrijeme,
        data.prijedeniKm,
        '',
        data.izvjesce
      ];
      
      addRow(sheet, rowData);
      updateUkupno(sheet);
      
      return createResponse(200, 'Vožnja uspješno evidentirana', {
        sheetName: sheetName,
        datum: formatDatum(datum),
        pocetnoStanje: pocetnoStanje,
        zavrsnoStanje: zavrsnoStanje,
        prijedeniKm: data.prijedeniKm
      });
      
    } catch (error) {
      return createResponse(500, 'Greška: ' + error.message);
    }
  }
  
  // Default status response
  return createResponse(200, 'Loko-vožnje API je aktivan', {
    version: '1.0',
    endpoints: {
      GET: 'Dodaj vožnju sa parametrima',
      POST: 'Dodaj novu vožnju'
    },
    requiredFields: ['prijedeniKm', 'izvjesce'],
    optionalFields: ['relacija']
  });
}

// ============================================================================
// SHEET MANAGEMENT
// ============================================================================

function getOrCreateSheet(sheetName, mjesec, godina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log(`Sheet ${sheetName} ne postoji, kreiram novi...`);
    sheet = createSheetFromTemplate(sheetName, mjesec, godina);
  }
  
  return sheet;
}

function createSheetFromTemplate(sheetName, mjesec, godina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  
  if (!template) {
    throw new Error('TEMPLATE sheet ne postoji!');
  }
  
  // Kopiraj template
  const newSheet = template.copyTo(ss);
  newSheet.setName(sheetName);
  
  // Ažuriraj mjesec u headeru
  const mjesecCapitalized = mjesec.charAt(0).toUpperCase() + mjesec.slice(1);
  
  // Pokušaj naći ćeliju koja sadrži "mjesec" i ažuriraj susjednu
  const range = newSheet.getRange('A1:Z10');
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j].toString().toLowerCase().includes('mjesec')) {
        newSheet.getRange(i + 1, j + 2).setValue(mjesecCapitalized);
        break;
      }
    }
  }
  
  // Pronađi red sa "ukupno"
  const lastRow = newSheet.getLastRow();
  let ukupnoRow = null;
  
  const searchRange = newSheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 12);
  const searchValues = searchRange.getValues();
  
  for (let i = 0; i < searchValues.length; i++) {
    for (let j = 0; j < searchValues[i].length; j++) {
      if (searchValues[i][j].toString().toLowerCase().includes('ukupno')) {
        ukupnoRow = CONFIG.DATA_START_ROW + i;
        Logger.log('Pronađen ukupno red: ' + ukupnoRow);
        break;
      }
    }
    if (ukupnoRow) break;
  }
  
  // Obriši podatke samo između DATA_START_ROW i ukupnoRow (ne diraj ukupno red!)
  if (ukupnoRow && ukupnoRow > CONFIG.DATA_START_ROW) {
    const rowsToClear = ukupnoRow - CONFIG.DATA_START_ROW;
    if (rowsToClear > 0) {
      newSheet.getRange(CONFIG.DATA_START_ROW, 1, rowsToClear, 15).clearContent();
      Logger.log(`Obrisano ${rowsToClear} redova podataka, zadržan ukupno red ${ukupnoRow}`);
    }
  } else if (lastRow >= CONFIG.DATA_START_ROW) {
    // Ako nema ukupno reda, obriši sve
    newSheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 15).clearContent();
  }
  
  // Postavi sheet kao aktivan i premjesti na početak
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);
  
  Logger.log(`Kreiran novi sheet: ${sheetName}`);
  return newSheet;
}

function addRow(sheet, rowData) {
  // Pronađi prvi prazan red od DATA_START_ROW nadalje
  const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 3, 100, 1); // KOLONA C (3) - datum
  const values = dataRange.getValues();
  
  let newRow = CONFIG.DATA_START_ROW;
  
  // Pronađi prvi prazan red u koloni C (datum)
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0] || values[i][0] === '') {
      newRow = CONFIG.DATA_START_ROW + i;
      break;
    }
  }
  
  Logger.log('DEBUG addRow: newRow nakon search = ' + newRow);
  
  // Pronađi gdje je "ukupno" red
  const lastRow = sheet.getLastRow();
  let ukupnoRow = null;
  
  const searchRange = sheet.getRange(newRow, 1, lastRow - newRow + 1, 12);
  const searchValues = searchRange.getValues();
  
  for (let i = 0; i < searchValues.length; i++) {
    for (let j = 0; j < searchValues[i].length; j++) {
      if (searchValues[i][j].toString().toLowerCase().includes('ukupno')) {
        ukupnoRow = newRow + i;
        break;
      }
    }
    if (ukupnoRow) break;
  }
  
  Logger.log('DEBUG addRow: ukupnoRow = ' + ukupnoRow);
  
  // Ako smo stigli do "ukupno" reda, pomakni ga dolje za 1 red
  if (ukupnoRow && newRow >= ukupnoRow) {
    Logger.log(`Novi red ${newRow} bi pregazio ukupno red ${ukupnoRow}, pomičem ukupno red dolje...`);
    sheet.insertRowBefore(ukupnoRow);
    ukupnoRow++; // Ukupno red je sada jedan red niže
    Logger.log(`Ukupno red premješten na red ${ukupnoRow}`);
  }
  
  Logger.log('DEBUG addRow: finalni newRow = ' + newRow);
  Logger.log('DEBUG addRow: CONFIG.STOPA_NADOKNADE = ' + CONFIG.STOPA_NADOKNADE);
  
  // Dodaj podatke (od kolone A, ali A i B su prazne)
  const range = sheet.getRange(newRow, 1, 1, rowData.length);
  range.setValues([rowData]);
  
  // Dodaj FORMULU za nadoknadu u kolonu K (11)
  const formulaCell = sheet.getRange(newRow, 11);
  const formula = `=J${newRow}*${CONFIG.STOPA_NADOKNADE}`;
  Logger.log('DEBUG addRow: formula = ' + formula);
  
  formulaCell.setFormula(formula);
  formulaCell.setNumberFormat('0.00 " EUR"');
  
  // Kopiraj SAMO FORMAT (ne i formule) iz template reda
  const templateRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, 1, rowData.length);
  const newRowRange = sheet.getRange(newRow, 1, 1, rowData.length);
  
  // Kopiraj samo vizualni format
  templateRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  
  Logger.log(`Dodan red ${newRow} u sheet ${sheet.getName()}`);
}

function getLastZavrsnoStanje(sheet) {
  // KOLONA G (7) je završno stanje
  const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 7, 100, 1);
  const values = dataRange.getValues();
  
  let lastValue = null;
  
  // Traži unazad prvi neprazan red U TRENUTNOM SHEET-U
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] && !isNaN(values[i][0]) && values[i][0] !== '') {
      lastValue = Number(values[i][0]);
      break;
    }
  }
  
  // Ako u trenutnom sheet-u nema podataka, traži u prethodnom mjesecu
  if (!lastValue) {
    Logger.log('Nema podataka u trenutnom sheet-u, tražim prethodni mjesec...');
    lastValue = getLastZavrsnoStanjeFromPreviousMonth(sheet);
  }
  
  return lastValue;
}

function getLastZavrsnoStanjeFromPreviousMonth(currentSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheetName = currentSheet.getName(); // npr. "ozujak-2025"
  
  // Parsiraj trenutni mjesec i godinu
  const parts = currentSheetName.split('-');
  if (parts.length !== 2) {
    Logger.log('Ne mogu parsirati naziv sheet-a, vraćam default 215000');
    return 215000;
  }
  
  const currentMjesecIndex = CONFIG.MJESECI.indexOf(parts[0].toLowerCase());
  const currentGodina = parseInt(parts[1]);
  
  if (currentMjesecIndex === -1) {
    Logger.log('Nepoznat mjesec, vraćam default 215000');
    return 215000;
  }
  
  // Izračunaj prethodni mjesec
  let previousMjesecIndex = currentMjesecIndex - 1;
  let previousGodina = currentGodina;
  
  if (previousMjesecIndex < 0) {
    // Ako je siječanj, idi na prosinac prošle godine
    previousMjesecIndex = 11;
    previousGodina--;
  }
  
  const previousMjesec = CONFIG.MJESECI[previousMjesecIndex];
  const previousSheetName = `${previousMjesec}-${previousGodina}`;
  
  Logger.log(`Tražim prethodni sheet: ${previousSheetName}`);
  
  const previousSheet = ss.getSheetByName(previousSheetName);
  
  if (!previousSheet) {
    Logger.log(`Prethodni sheet ${previousSheetName} ne postoji, vraćam default 215000`);
    return 215000;
  }
  
  // Dohvati zadnje završno stanje iz prethodnog mjeseca
  const dataRange = previousSheet.getRange(CONFIG.DATA_START_ROW, 7, 100, 1);
  const values = dataRange.getValues();
  
  let lastValue = null;
  
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] && !isNaN(values[i][0]) && values[i][0] !== '') {
      lastValue = Number(values[i][0]);
      Logger.log(`Pronađeno završno stanje iz ${previousSheetName}: ${lastValue} km`);
      break;
    }
  }
  
  // Ako ni prethodni mjesec nema podataka, vraćamo default
  if (!lastValue) {
    Logger.log('Ni prethodni mjesec nema podataka, vraćam default 215000');
    return 215000;
  }
  
  return lastValue;
}

function updateUkupno(sheet) {
  const lastRow = sheet.getLastRow();
  const searchRange = sheet.getRange(lastRow - 5, 1, 6, 12);
  const values = searchRange.getValues();
  
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j].toString().toLowerCase().includes('ukupno')) {
        const ukupnoRow = lastRow - 5 + i + 1;
        
        // Izračunaj sumu nadoknada (kolona K - pozicija 11)
        const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 11, ukupnoRow - CONFIG.DATA_START_ROW, 1);
        const nadoknade = dataRange.getValues().flat().filter(v => v && v !== '');
        
        const ukupnaNadoknada = nadoknade.reduce((sum, val) => {
          const num = parseFloat(val.toString().replace(' EUR', '').replace(',', '.'));
          return sum + (isNaN(num) ? 0 : num);
        }, 0);
        
        // Ažuriraj ukupno u koloni K (11)
        sheet.getRange(ukupnoRow, 11).setValue(ukupnaNadoknada.toFixed(2) + ' EUR');
        Logger.log(`Ažurirano ukupno: ${ukupnaNadoknada.toFixed(2)} EUR`);
        return;
      }
    }
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function formatDatum(date) {
  const dan = date.getDate();
  const mjesec = date.getMonth() + 1;
  const godina = date.getFullYear();
  return `${dan}.${mjesec}.${godina}`;
}

function formatVrijemeZaokruzeno(date) {
  const sati = date.getHours();
  let minute = date.getMinutes();
  
  // Zaokruži minute na najbližih 10 minuta prema dolje
  minute = Math.floor(minute / 10) * 10;
  
  return `${String(sati).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function formatVrijeme(date) {
  const sati = String(date.getHours()).padStart(2, '0');
  const minute = String(date.getMinutes()).padStart(2, '0');
  return `${sati}:${minute}`;
}

function createResponse(status, message, data = null) {
  const response = {
    status: status,
    message: message,
    timestamp: new Date().toISOString()
  };
  
  if (data) {
    response.data = data;
  }
  
  // Google Apps Script automatski dozvoljava CORS za Web Apps deployane kao "Anyone"
  return ContentService
    .createTextOutput(JSON.stringify(response, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// TEST FUNKCIJA - za testiranje više vožnji odjednom
// ============================================================================

function testMultipleTrips() {
  const izvjesceOpcije = [
    'prijevoz do ureda',
    'posjeta klijentu',
    'službeni put',
    'prijevoz na sastanak',
    'inspekcijski nadzor',
    'servisiranje vozila'
  ];
  
  const relacije = ['RIJEKA', 'ZAGREB', 'SPLIT', 'RIJEKA-OPATIJA', 'PULA'];
  
  // Dodaj 20 vožnji
  for (let i = 0; i < 20; i++) {
    const randomKm = Math.floor(Math.random() * 50) + 5; // 5-55 km
    const randomIzvjesce = izvjesceOpcije[Math.floor(Math.random() * izvjesceOpcije.length)];
    const randomRelacija = relacije[Math.floor(Math.random() * relacije.length)];
    
    // Random datum u veljači 2025
    const dan = Math.floor(Math.random() * 28) + 1;
    const datum = `2025-02-${String(dan).padStart(2, '0')}`;
    
    const testData = {
      prijedeniKm: randomKm,
      izvjesce: randomIzvjesce,
      datum: datum,
      relacija: randomRelacija
    };
    
    const mockEvent = {
      postData: {
        contents: JSON.stringify(testData)
      }
    };
    
    Logger.log(`=== Vožnja ${i + 1}/20 ===`);
    const result = doPost(mockEvent);
    const response = JSON.parse(result.getContent());
    
    if (response.status === 200) {
      Logger.log(`✓ ${randomKm} km - ${randomIzvjesce}`);
    } else {
      Logger.log(`✗ ERROR: ${response.message}`);
    }
    
    // Kratka pauza da se ne overload-a
    Utilities.sleep(100);
  }
  
  Logger.log('=== TEST ZAVRŠEN ===');
  Logger.log('Dodano 20 vožnji u veljaca-2025 sheet');
}