/**
 * LOKO-VOŽNJE 
 * Google Apps Script za automatsko generiranje mjesečnih vožnji
 * 
 * KAKO KORISTITI:
 * 1. Otvori Apps Script editor
 * 2. Pokreni funkciju: generateCurrentMonth() - za trenutni mjesec
 * 3. Ili: generateMonth(mjesecIndex, godina) - za specifični mjesec (0=siječanj, 11=prosinac)
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
    relacija: 'RIJEKA',
    pocetnoStanje: 213519 // Početno stanje za prvi mjesec
  },
  
  STOPA_NADOKNADE: 0.5, // EUR po kilometru
  
  // Ciljevi generiranja vožnji
  GENERIRANJE: {
    minEuro: 200,         // Minimalni iznos u EUR
    maxEuro: 250,         // Maksimalni iznos u EUR
    minKm: 400,           // Minimalni ukupni kilometri (200 EUR / 0.5)
    maxKm: 500,           // Maksimalni ukupni kilometri (250 EUR / 0.5)
    minVoznjiDnevno: 2,   // Minimum vožnji po danu
    maxVoznjiDnevno: 3,   // Maximum vožnji po danu
    minUredPostotak: 60,  // Minimum % vožnji "prijevoz do ureda"
    maxUredPostotak: 80,  // Maximum % vožnji "prijevoz do ureda"
  },
  
  // Samo 3 tipa vožnji
  IZVJESCE_OPCIJE: [
    {
      naziv: 'prijevoz do ureda',
      fiksnoKm: 5,
      randomKm: false
    },
    {
      naziv: 'posjeta klijentu',
      minKm: 5,
      maxKm: 30,
      randomKm: true
    },
    {
      naziv: 'prijevoz na sastanak',
      minKm: 5,
      maxKm: 30,
      randomKm: true
    }
  ],
  
  // Mjeseci na hrvatskom
  MJESECI: [
    'sijecanj', 'veljaca', 'ozujak', 'travanj', 'svibanj', 'lipanj',
    'srpanj', 'kolovoz', 'rujan', 'listopad', 'studeni', 'prosinac'
  ]
};

// ============================================================================
// GLAVNE FUNKCIJE ZA GENERIRANJE
// ============================================================================

/**
 * Generiraj vožnje za trenutni mjesec
 */
function generateCurrentMonth() {
  const danas = new Date();
  const mjesec = danas.getMonth(); // 0-11
  const godina = danas.getFullYear();
  
  generateMonth(mjesec, godina);
}

/**
 * Helper funkcije za brzo pokretanje
 */
function generateVeljaca2026() {
  generateMonth(1, 2026); // Veljača = indeks 1
}

function generateOzujak2026() {
  generateMonth(2, 2026); // Ožujak = indeks 2
}

function generateTravanj2026() {
  generateMonth(3, 2026); // Travanj = indeks 3
}

/**
 * Test funkcija - generiraj veljaču 2026
 */
function testGenerate() {
  Logger.log('=== POKRETANJE TEST FUNKCIJE ===');
  generateMonth(1, 2026); // Veljača 2026
  Logger.log('=== TEST ZAVRŠEN ===');
}

/**
 * AUTOMATSKO POKRETANJE - Generiraj trenutni mjesec bez alert-a
 * Ovu funkciju postavi kao trigger da se pokreće prvog dana u mjesecu
 * 
 * KAKO POSTAVITI TRIGGER:
 * 1. U Apps Script editoru, klikni na sat ikonu (Triggers) u lijevom meniju
 * 2. Klikni "+ Add Trigger" (dolje desno)
 * 3. Odaberi funkciju: generateMonthlyAutomatic
 * 4. Event source: Time-driven
 * 5. Type of time based trigger: Month timer
 * 6. Select day of month: 1
 * 7. Select time of day: 1am to 2am (ili bilo koje doba)
 * 8. Snimi
 */
function generateMonthlyAutomatic() {
  const danas = new Date();
  const mjesec = danas.getMonth(); // 0-11
  const godina = danas.getFullYear();
  
  Logger.log(`=== AUTOMATSKO GENERIRANJE: ${CONFIG.MJESECI[mjesec]} ${godina} ===`);
  
  try {
    // Pozovi generateMonth BEZ alert dijaloga
    const rezultat = generateMonth(mjesec, godina, false);
    
    Logger.log('✓ Automatsko generiranje uspješno!');
    Logger.log(`Sheet: ${rezultat.sheetName}`);
    Logger.log(`Vožnji: ${rezultat.voznji}`);
    Logger.log(`Ukupno: ${rezultat.ukupnoEur.toFixed(2)} EUR`);
    
  } catch (error) {
    Logger.log('✗ GREŠKA pri automatskom generiranju:');
    Logger.log(error.message);
    Logger.log(error.stack);
    
    // Opcionalno: pošalji email notifikaciju o greški
    // MailApp.sendEmail('tvoj-email@example.com', 
    //   'Greška u automatskom generiranju vožnji',
    //   error.message);
  }
}

/**
 * Generiraj vožnje za specifični mjesec
 * @param {number} mjesecIndex - Indeks mjeseca (0-11)
 * @param {number} godina - Godina
 * @param {boolean} showAlert - Prikaži li alert dialog (default: true)
 */
function generateMonth(mjesecIndex, godina, showAlert = true) {
  Logger.log('========================================');
  Logger.log(`GENERIRANJE VOŽNJI ZA: ${CONFIG.MJESECI[mjesecIndex].toUpperCase()} ${godina}`);
  Logger.log('========================================');
  
  const mjesecNaziv = CONFIG.MJESECI[mjesecIndex];
  const sheetName = `${mjesecNaziv}-${godina}`;
  
  // Provjeri da li sheet postoji i obriši ga
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName(sheetName);
  
  if (existingSheet) {
    Logger.log(`Sheet ${sheetName} već postoji. Brišem ga...`);
    ss.deleteSheet(existingSheet);
  }
  
  // Kreiraj novi sheet
  const sheet = createSheetFromTemplate(sheetName, mjesecNaziv, godina);
  
  // Dohvati početno stanje
  const pocetnoStanje = getInitialStartState(mjesecNaziv, godina);
  Logger.log(`Početno stanje kilometara: ${pocetnoStanje} km`);
  
  // Generiraj vožnje
  const voznje = generateTripsForMonth(mjesecIndex, godina, pocetnoStanje);
  
  Logger.log(`\n=== GENERIRANJE ===`);
  Logger.log(`Ukupno vožnji: ${voznje.length}`);
  
  // Pronađi UKUPNO red i osigurai da ima dovoljno prostora
  const ukupnoRow = findUkupnoRow(sheet);
  const potrebnoRedova = voznje.length;
  const dostupnoRedova = ukupnoRow - CONFIG.DATA_START_ROW;
  
  if (potrebnoRedova > dostupnoRedova) {
    // Treba nam više redova - ubaci ih prije UKUPNO reda
    const nedostaje = potrebnoRedova - dostupnoRedova;
    Logger.log(`Dodajem ${nedostaje} redova prije UKUPNO reda...`);
    sheet.insertRowsBefore(ukupnoRow, nedostaje);
  }
  
  // Dodaj vožnje u sheet
  let currentKm = pocetnoStanje;
  
  voznje.forEach((voznja, index) => {
    const zavrsnoStanje = currentKm + voznja.km;
    
    const rowData = [
      '', '', // A i B kolone prazne
      formatDatum(voznja.datum),           // C - Datum
      CONFIG.DEFAULTS.marka,               // D - Marka
      CONFIG.DEFAULTS.regBroj,             // E - Reg. broj
      currentKm,                           // F - Početno stanje
      zavrsnoStanje,                       // G - Završno stanje
      CONFIG.DEFAULTS.relacija,            // H - Relacija
      voznja.vrijeme,                      // I - Vrijeme
      voznja.km,                           // J - Prijeđeni km
      '',                                  // K - Nadoknada (formula)
      voznja.izvjesce                      // L - Izvješće
    ];
    
    addRowDirect(sheet, rowData, CONFIG.DATA_START_ROW + index);
    
    currentKm = zavrsnoStanje;
  });
  
  // Ažuriraj ukupno
  updateUkupno(sheet);
  
  // Finalni log
  const ukupnoKm = currentKm - pocetnoStanje;
  const ukupnoEur = ukupnoKm * CONFIG.STOPA_NADOKNADE;
  
  Logger.log(`\n=== REZULTAT ===`);
  Logger.log(`Početno stanje: ${pocetnoStanje.toFixed(0)} km`);
  Logger.log(`Završno stanje: ${currentKm.toFixed(0)} km`);
  Logger.log(`Ukupno km: ${ukupnoKm.toFixed(0)} km`);
  Logger.log(`Ukupno EUR: ${ukupnoEur.toFixed(2)} EUR`);
  Logger.log('========================================');
  
  // Prikaz poruke korisniku (samo ako je showAlert = true i UI je dostupan)
  if (showAlert) {
    try {
      SpreadsheetApp.getUi().alert(
        'Vožnje generirane!',
        `Sheet: ${sheetName}\n` +
        `Vožnji: ${voznje.length}\n` +
        `Ukupno km: ${ukupnoKm.toFixed(0)} km\n` +
        `Ukupno: ${ukupnoEur.toFixed(2)} EUR`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e) {
      // UI nije dostupan (npr. kod trigera) - to je OK, preskoči alert
      Logger.log('UI nije dostupan, preskaćem alert dialog.');
    }
  }
  
  return {
    sheetName: sheetName,
    voznji: voznje.length,
    ukupnoKm: ukupnoKm,
    ukupnoEur: ukupnoEur
  };
}

/**
 * Generiraj listu vožnji za mjesec
 */
function generateTripsForMonth(mjesecIndex, godina, pocetnoStanje) {
  // Izračunaj ciljni broj kilometara (između 400-500)
  const targetKm = randomBetween(CONFIG.GENERIRANJE.minKm, CONFIG.GENERIRANJE.maxKm);
  Logger.log(`Ciljni broj kilometara: ${targetKm} km (${targetKm * CONFIG.STOPA_NADOKNADE} EUR)`);
  
  // Izračunaj postotak vožnji "prijevoz do ureda" (60-80%)
  const uredPostotak = randomBetween(CONFIG.GENERIRANJE.minUredPostotak, CONFIG.GENERIRANJE.maxUredPostotak);
  Logger.log(`Postotak "prijevoz do ureda": ${uredPostotak}%`);
  
  // Izračunaj koliko km treba biti "prijevoz do ureda"
  const uredKmTarget = targetKm * (uredPostotak / 100);
  const ostaloKmTarget = targetKm - uredKmTarget;
  
  Logger.log(`- Prijevoz do ureda: ${uredKmTarget.toFixed(0)} km`);
  Logger.log(`- Ostale vožnje: ${ostaloKmTarget.toFixed(0)} km`);
  
  // PRVO: Generiraj potreban broj vožnji svake vrste
  const sveVoznje = [];
  
  // 1. Dodaj "prijevoz do ureda" vožnje (fiksno 5km svaka)
  let uredKmCurrent = 0;
  while (uredKmCurrent < uredKmTarget) {
    sveVoznje.push({
      km: 5,
      izvjesce: 'prijevoz do ureda'
    });
    uredKmCurrent += 5;
  }
  
  // 2. Dodaj ostale vožnje (random 5-30km)
  let ostaloKmCurrent = 0;
  while (ostaloKmCurrent < ostaloKmTarget) {
    // Random odabir između "posjeta klijentu" i "prijevoz na sastanak"
    const tipIndex = Math.random() < 0.5 ? 1 : 2;
    const tipVoznje = CONFIG.IZVJESCE_OPCIJE[tipIndex];
    
    // Random km ali pazi da ne prekoračiš target previše
    const preostaloKm = ostaloKmTarget - ostaloKmCurrent;
    const maxKm = Math.min(tipVoznje.maxKm, preostaloKm + 10); // Dozvoli malo preko
    const km = randomBetween(tipVoznje.minKm, maxKm);
    
    sveVoznje.push({
      km: km,
      izvjesce: tipVoznje.naziv
    });
    ostaloKmCurrent += km;
  }
  
  Logger.log(`Generirano ${sveVoznje.length} vožnji za raspored`);
  
  // DRUGO: Rasporedi vožnje po danima
  const dani = getDaysInMonth(mjesecIndex, godina);
  const voznje = [];
  let voznjaIndex = 0;
  
  // Shuffle vožnje da budu random raspoređene
  shuffleArray(sveVoznje);
  
  for (let danIndex = 0; danIndex < dani.length; danIndex++) {
    const dan = dani[danIndex];
    const brVoznji = randomBetween(CONFIG.GENERIRANJE.minVoznjiDnevno, CONFIG.GENERIRANJE.maxVoznjiDnevno);
    
    // Dodaj vožnje za ovaj dan
    for (let i = 0; i < brVoznji && voznjaIndex < sveVoznje.length; i++) {
      const voznja = sveVoznje[voznjaIndex++];
      
      // Generiraj vrijeme
      const sat = randomBetween(7, 17);
      const minuta = randomBetween(0, 5) * 10;
      const vrijeme = `${String(sat).padStart(2, '0')}:${String(minuta).padStart(2, '0')}`;
      
      voznje.push({
        datum: dan,
        vrijeme: vrijeme,
        km: voznja.km,
        izvjesce: voznja.izvjesce
      });
    }
    
    // Ako smo potrošili sve vožnje, izađi
    if (voznjaIndex >= sveVoznje.length) {
      break;
    }
  }
  
  // Ako ima još neraspoređenih vožnji, dodaj ih na preostale dane
  while (voznjaIndex < sveVoznje.length && dani.length > 0) {
    const randomDanIndex = randomBetween(0, dani.length - 1);
    const voznja = sveVoznje[voznjaIndex++];
    
    const sat = randomBetween(7, 17);
    const minuta = randomBetween(0, 5) * 10;
    const vrijeme = `${String(sat).padStart(2, '0')}:${String(minuta).padStart(2, '0')}`;
    
    voznje.push({
      datum: dani[randomDanIndex],
      vrijeme: vrijeme,
      km: voznja.km,
      izvjesce: voznja.izvjesce
    });
  }
  
  // Sortiraj vožnje po datumu i vremenu
  voznje.sort((a, b) => {
    if (a.datum.getTime() !== b.datum.getTime()) {
      return a.datum - b.datum;
    }
    return a.vrijeme.localeCompare(b.vrijeme);
  });
  
  const ukupnoKm = voznje.reduce((sum, v) => sum + v.km, 0);
  const uredCount = voznje.filter(v => v.izvjesce === 'prijevoz do ureda').length;
  const finalniPostotak = (uredCount / voznje.length * 100).toFixed(1);
  
  Logger.log(`\n=== STATISTIKA ===`);
  Logger.log(`Ukupno vožnji: ${voznje.length}`);
  Logger.log(`- Prijevoz do ureda: ${uredCount} (${finalniPostotak}%)`);
  Logger.log(`- Posjeta klijentu: ${voznje.filter(v => v.izvjesce === 'posjeta klijentu').length}`);
  Logger.log(`- Prijevoz na sastanak: ${voznje.filter(v => v.izvjesce === 'prijevoz na sastanak').length}`);
  Logger.log(`Ukupno km: ${ukupnoKm} km (${(ukupnoKm * CONFIG.STOPA_NADOKNADE).toFixed(2)} EUR)`);
  
  return voznje;
}

/**
 * Shuffle array (Fisher-Yates algorithm)
 */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}

/**
 * Pronađi red sa "ukupno" u sheet-u
 */
function findUkupnoRow(sheet) {
  const lastRow = sheet.getLastRow();
  const searchRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, 12);
  const values = searchRange.getValues();
  
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] && values[i][j].toString().toLowerCase().includes('ukupno')) {
        const ukupnoRow = CONFIG.DATA_START_ROW + i;
        Logger.log(`Pronađen UKUPNO red na poziciji: ${ukupnoRow}`);
        return ukupnoRow;
      }
    }
  }
  
  // Ako ne nađemo UKUPNO red, pretpostavi da je na kraju
  Logger.log('UKUPNO red nije pronađen, koristim lastRow');
  return lastRow;
}

/**
 * Dohvati sve dane u mjesecu osim nedjelja
 */
function getDaysInMonth(mjesecIndex, godina) {
  const dani = [];
  const brojDana = new Date(godina, mjesecIndex + 1, 0).getDate();
  
  for (let dan = 1; dan <= brojDana; dan++) {
    const datum = new Date(godina, mjesecIndex, dan);
    
    // Preskoči nedjelje (0 = nedjelja)
    if (datum.getDay() === 0) {
      continue;
    }
    
    dani.push(datum);
  }
  
  Logger.log(`Broj radnih dana (bez nedjelja): ${dani.length}`);
  return dani;
}

/**
 * Random broj između min i max (inclusive)
 */
function randomBetween(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

/**
 * Dohvati početno stanje za mjesec
 */
function getInitialStartState(mjesecNaziv, godina) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Pronađi indeks trenutnog mjeseca
  const currentMjesecIndex = CONFIG.MJESECI.indexOf(mjesecNaziv.toLowerCase());
  
  if (currentMjesecIndex === -1) {
    Logger.log('Nepoznat mjesec, vraćam default');
    return CONFIG.DEFAULTS.pocetnoStanje;
  }
  
  // Izračunaj prethodni mjesec
  let previousMjesecIndex = currentMjesecIndex - 1;
  let previousGodina = godina;
  
  if (previousMjesecIndex < 0) {
    previousMjesecIndex = 11;
    previousGodina--;
  }
  
  const previousMjesec = CONFIG.MJESECI[previousMjesecIndex];
  const previousSheetName = `${previousMjesec}-${previousGodina}`;
  
  const previousSheet = ss.getSheetByName(previousSheetName);
  
  if (!previousSheet) {
    Logger.log(`Prethodni sheet ${previousSheetName} ne postoji, koristim default: ${CONFIG.DEFAULTS.pocetnoStanje}`);
    return CONFIG.DEFAULTS.pocetnoStanje;
  }
  
  // Dohvati zadnje završno stanje iz prethodnog mjeseca (kolona G = 7)
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
  
  if (!lastValue) {
    Logger.log('Prethodni sheet nema podataka, koristim default');
    return CONFIG.DEFAULTS.pocetnoStanje;
  }
  
  return lastValue;
}

/**
 * Dodaj red direktno na specifičnu poziciju
 */
function addRowDirect(sheet, rowData, rowNumber) {
  // Dodaj podatke
  const range = sheet.getRange(rowNumber, 1, 1, rowData.length);
  range.setValues([rowData]);
  
  // Dodaj formulu za nadoknadu u kolonu K (11)
  const formulaCell = sheet.getRange(rowNumber, 11);
  const formula = `=J${rowNumber}*${CONFIG.STOPA_NADOKNADE}`;
  formulaCell.setFormula(formula);
  formulaCell.setNumberFormat('0.00 " EUR"');
  
  // Kopiraj format iz template reda
  const templateRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, 1, rowData.length);
  const newRowRange = sheet.getRange(rowNumber, 1, 1, rowData.length);
  templateRange.copyTo(newRowRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

// ============================================================================
// SHEET MANAGEMENT FUNKCIJE
// ============================================================================

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

function updateUkupno(sheet) {
  // Pronađi UKUPNO red
  const lastRow = sheet.getLastRow();
  let ukupnoRow = null;
  
  // Traži kroz sve redove od DATA_START_ROW do kraja
  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    const cellValue = sheet.getRange(row, 1, 1, 12).getValues()[0];
    
    for (let col = 0; col < cellValue.length; col++) {
      if (cellValue[col] && cellValue[col].toString().toLowerCase().includes('ukupno')) {
        ukupnoRow = row;
        Logger.log(`Pronađen UKUPNO red na poziciji: ${ukupnoRow}`);
        break;
      }
    }
    
    if (ukupnoRow) break;
  }
  
  if (!ukupnoRow) {
    Logger.log('UKUPNO red nije pronađen!');
    return;
  }
  
  // Izračunaj sumu SVIH nadoknada u koloni K (od DATA_START_ROW do ukupnoRow-1)
  let ukupnaNadoknada = 0;
  
  for (let row = CONFIG.DATA_START_ROW; row < ukupnoRow; row++) {
    const cell = sheet.getRange(row, 11); // Kolona K (11)
    const value = cell.getValue();
    
    if (value && value !== '') {
      // Ako je string sa "EUR", očisti ga
      if (typeof value === 'string') {
        const num = parseFloat(value.replace(' EUR', '').replace(',', '.').trim());
        if (!isNaN(num)) {
          ukupnaNadoknada += num;
        }
      } else if (typeof value === 'number') {
        ukupnaNadoknada += value;
      }
    }
  }
  
  // Ažuriraj ukupno u koloni K
  sheet.getRange(ukupnoRow, 11).setValue(ukupnaNadoknada.toFixed(2) + ' EUR');
  Logger.log(`Ažurirano ukupno: ${ukupnaNadoknada.toFixed(2)} EUR (${ukupnoRow - CONFIG.DATA_START_ROW} redova)`);
}

// ============================================================================
// HELPER FUNKCIJE
// ============================================================================

function formatDatum(date) {
  const dan = date.getDate();
  const mjesec = date.getMonth() + 1;
  const godina = date.getFullYear();
  return `${dan}.${mjesec}.${godina}`;
}
