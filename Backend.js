/*************************************************
 * 0. WEB APP ENTRY
 *************************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAP Data Automation')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*************************************************
 * 1. CONFIGURATION RULES
 *************************************************/
const IGNORE_LISTS = {
  gls: ['40010010', '40010020', '40010060', '40300000', '40300400', '40300500', '40300600', '40300700', '40302000', '40302010'],
  glRanges: [],

  materials: ['1000000041', '1000000061', '1000000064', '10000000612', '6000000005'],
  materialRanges: [
    {from: 1000000051, to: 1000000058}, {from: 1000000071, to: 1000000075}, {from: 1000000091, to: 1000000094}, {from: 1000000111, to: 1000000114},
    {from: 1000000401, to: 1000000403}, {from: 1000000601, to: 1000000614}, {from: 1100000004, to: 1100000747}, {from: 2000000008, to: 2000000018},
    {from: 2100000020, to: 2100001185}, {from: 3000000001, to: 3000000024}, {from: 3100000001, to: 3100001213}, {from: 4000000001, to: 4000021371},
    {from: 5000000001, to: 5000002563}, {from: 7000000001, to: 7000000032}, {from: 7100000001, to: 7100000006}, {from: 7110000001, to: 7110000069},
    {from: 7120000001, to: 7120000066}, {from: 7130000001, to: 7130000067}, {from: 8000000001, to: 8000000235}, {from: 9000000001, to: 9000000003}
  ]
};

const SECTION_RULES = {
  'MONTHLY PRODUCTION': {
    type: 'STANDARD',
    unit: 'MT',
    byCategory: {
      'Billet Production': {
        plants: ['1001', '1003'],
        movements: ['101', '102'],
        singleMaterials: ['2000000001', '2000000002'] 
      },
      'TMT Production (Excluding Chilli & Random)': {
        plants: ['1002', '1004'],
        movements: ['101', '102'],
        excludePOSeries: ['5200', '5300'],
        materialRanges: [
          {from: 1000000001, to: 1000000040},
          // {from: 1000000018, to: 1000000018},
          // {from: 1000000021, to: 1000000024},
          // {from: 1000000031, to: 1000000038},
          // {from: 1000000161, to: 1000000188},
          {from: 1000000121, to: 1000000188},
          {from: 1000000321, to: 1000000322},
          {from: 1000000501, to: 1000000510},
          {from: 1000000615, to: 1000000628}
        ]
      },
      'Wire Rod Production': {
        plants: ['1004'],
        movements: ['101'],
        materialRanges: [
          {from: 1000000323, to: 1000000331}
        ]        
      }
    }
  },


  'MONTHLY SALES': {
    type: 'UNIT_SUM',
    unit: 'MT',
    glAccounts: {
      byCategory: {
        'TMT Sales Qty': {
          gls: ['40000000', '40300300', '40600000'],
          excludeMaterials: ['1000000075']
        },
        'TMT Random': { 
          gls: ['40010070'],
          sharedGLs: ['40000000'],
          includeMaterials: ['1000000075']
        },
        'Wire Rod Sales Qty': { gls: ['40010090'], excludeMaterials: [] },
        'Billet Sales Qty':   { gls: ['40010000', '40302020'], excludeMaterials: [] },
        'TMT Chilli':         { gls: ['40020030'], excludeMaterials: [] },
        'Binding Wire':       { gls: ['40020040'], excludeMaterials: [] }
      }
    }
  },

  'SALE PRICE': {
    type: 'CURRENCY_RATIO',
    unit: 'MT',
    glAccounts: {
      byCategory: {
        'TMT': {
          gls: ['40000000', '40300300', '40600000'],
          excludeMaterials: ['1000000075']
        },
        'TMT Random': { 
          gls: ['40010070'],
          sharedGLs: ['40000000'],
          includeMaterials: ['1000000075']
        },
        'WRD':          { gls: ['40010090'], excludeMaterials: [] },
        'Billet':       { gls: ['40010000', '40302020'], excludeMaterials: [] },
        'TMT Chilli':   { gls: ['40020030'], excludeMaterials: [] },
        'Binding Wire': { gls: ['40020040'], excludeMaterials: [] }
      }
    }
  }
};

/*************************************************
 * 2. USER CONFIG STORAGE
 *************************************************/
function getCurrentFY_() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();
  const startYear = month >= 3 ? year : year - 1;
  const endShort = String((startYear + 1) % 100).padStart(2, '0');
  return startYear + '-' + endShort;
}

function saveUserConfig(config) {
  const fy = getCurrentFY_();
  PropertiesService.getUserProperties().setProperty('CFG_' + fy, JSON.stringify(config || {}));
  return { success: true };
}

function getUserConfig() {
  const fy = getCurrentFY_();
  const raw = PropertiesService.getUserProperties().getProperty('CFG_' + fy);
  return raw ? JSON.parse(raw) : null;
}

function clearUserConfig() {
  const fy = getCurrentFY_();
  PropertiesService.getUserProperties().deleteProperty('CFG_' + fy);
  return { success: true };
}

/*************************************************
 * 3. SHEET VALIDATION
 *************************************************/
function extractSheetId(url) {
  if (!url || typeof url !== 'string') return url;
  const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : url;
}

function validateAndGetSheets(sheetUrl) {
  try {
    const ss = SpreadsheetApp.openById(extractSheetId(sheetUrl));
    return {
      success: true,
      message: 'Connected: ' + ss.getName(),
      sheets: ss.getSheets().map(s => s.getName())
    };
  } catch (err) {
    throw new Error('Invalid URL or Permission Denied.');
  }
}

/*************************************************
 * 4. DROPDOWN DATA
 *************************************************/
function getSections() {
  return Object.keys(SECTION_RULES).map(k => 
    k.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, s => s.toUpperCase())
  );
}

function getCategoriesForSection(section) {
  const key = section.toUpperCase();
  let rule = SECTION_RULES[key];
  if (!rule) {
    const k = Object.keys(SECTION_RULES).find(x => x === key.replace(' ', '_'));
    rule = SECTION_RULES[k];
  }
  
  if (rule) {
    if (rule.byCategory) return Object.keys(rule.byCategory);
    if (rule.glAccounts && rule.glAccounts.byCategory) return Object.keys(rule.glAccounts.byCategory);
  }
  return [];
}

function getYears() {
  const today = new Date();
  const m = today.getMonth();
  const calYear = today.getFullYear();
  const currentFYStart = m >= 3 ? calYear : calYear - 1;
  const years = [];
  for (let y = currentFYStart; y >= 2021; y--) {
    years.push(y + '-' + String((y + 1) % 100).slice(-2));
  }
  return years;
}

function getMonths() {
  return ['April','May','June','July','August','September','October','November','December','January','February','March'];
}

/*************************************************
 * 5. MAIN UPLOAD HANDLER (With File Validation)
 *************************************************/
function handleSingleFile(base64Data, fileName, sectionFromUI, categoriesFromUI, monthFromUI, yearFromUI, sheetUrl, sheetName) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), getMimeType(fileName), fileName);
    const sapData = parseFile(blob, fileName);
    
    if (!sapData || !sapData.length) return { status: 'ERROR', message: 'No data found in file.' };

    const effectiveSection = sectionFromUI.toString().trim().toUpperCase();
    const rule = SECTION_RULES[effectiveSection];
    if (!rule) return { status: 'ERROR', message: 'No rules for Section "' + effectiveSection + '".' };

    let headerRowIndex = -1;
    let headers = [];

    for (let i = 0; i < 100 && i < sapData.length; i++) {
      const rowString = sapData[i].map(cell => cell ? cell.toString().toUpperCase() : '').join(' ');
      if (rowString.includes('QUANTITY') || rowString.includes('QTY')) {
        headerRowIndex = i;
        headers = sapData[i].map(h => h ? h.toString().toUpperCase().trim() : '');
        break;
      }
    }

    if (headerRowIndex === -1) {
      return { status: 'ERROR', message: 'Invalid SAP File: Could not find a Quantity column.' };
    }

    const hasGL = headers.some(h => h.includes('G/L ACCOUNT') || h.includes('GL ACCOUNT') || h === 'G/L' || h === 'ACCOUNT');
    const hasMovement = headers.some(h => h.includes('MOVEMENT') || h === 'MVT' || h === 'BWART');
    const hasMaterial = headers.some(h => h.includes('MATERIAL') || h === 'MATNR');

    const isProduction = effectiveSection.includes('PRODUCTION');
    const isSalesOrPrice = effectiveSection.includes('SALES') || effectiveSection.includes('PRICE');

    if (isProduction && !hasMovement && !hasMaterial) {
      return { 
        status: 'ERROR', 
        message: 'FILE MISMATCH: You selected Production, but the uploaded file is missing Material or Movement columns. Did you upload the Sales file by mistake?' 
      };
    }

    if (isSalesOrPrice && !hasGL) {
      return { 
        status: 'ERROR', 
        message: 'FILE MISMATCH: You selected Sales/Price, but the uploaded file is missing the G/L Account column. Did you upload the Production file by mistake?' 
      };
    }

    const sheetId = extractSheetId(sheetUrl);
    let allUnknownItems = [];
    let unknownType = ''; 

    for (let i = 0; i < categoriesFromUI.length; i++) {
      const cat = categoriesFromUI[i];
      const result = calculateSAPDataFromMatrix_(sapData, rule, cat, effectiveSection);

      if (result.unknowns && result.unknowns.length > 0) {
        unknownType = result.unknownType;
        result.unknowns.forEach(u => {
          if (allUnknownItems.indexOf(u) === -1) allUnknownItems.push(u);
        });
      }

      const target = getTargetLocation(sheetId, sheetName, monthFromUI, cat, yearFromUI, effectiveSection);
      writeTotalToGoogleSheet(sheetId, sheetName, target.row, target.column, result.value);
    }

    const cfg = getUserConfig() || {};
    cfg.sheetUrl = sheetUrl;
    cfg.sheetTab = sheetName;
    cfg.financialYear = getCurrentFY_();
    cfg.validated = true;
    saveUserConfig(cfg);

    if (allUnknownItems.length > 0) {
      return {
        status: 'WARNING', 
        message: 'Data uploaded, but undefined codes were found.',
        unknownItems: allUnknownItems,
        unknownType: unknownType
      };
    }

    return {
      status: 'OK',
      message: 'Processed successfully.'
    };

  } catch (e) {
    return { status: 'ERROR', message: 'Mapping Error: ' + e.message };
  }
}

/*************************************************
 * 6. CALCULATION LOGIC (Shared G/L Material Fix)
 *************************************************/
function calculateSAPDataFromMatrix_(sapData, rule, categoryName, sectionName) {
  let headerRowIndex = -1;
  let headers = [];

  for (let i = 0; i < 7 && i < sapData.length; i++) {
    const rowString = sapData[i].join(' ').toUpperCase();
    if (rowString.includes('QTY') || rowString.includes('QUANTITY')) {
      headerRowIndex = i;
      headers = sapData[i];
      break;
    }
  }
  if (headerRowIndex === -1) throw new Error('Header row not found.');

  const findCol = (keywords) => {
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i] ? headers[i].toString().toUpperCase().trim() : '';
      if (keywords.some(k => h === k)) return i;
    }
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i] ? headers[i].toString().toUpperCase().trim() : '';
      if (keywords.some(k => h.includes(k))) return i;
    }
    return -1;
  };

  const colQty     = findCol(['QUANTITY', 'QTY', 'BILLED QUANTITY']);
  const colCurr    = findCol(['NET VALUE','NET AMOUNT','CURRENCY VALUE','COMPANY CODE CURRENCY VALUE','AMOUNT','VALUE']);
  const colUnit    = findCol(['UNIT OF MEASURE', 'UNIT', 'UOM']);
  const colMvmt    = findCol(['MOVEMENT', 'MVT']);
  const colPlant   = findCol(['PLANT']);
  const colGL      = findCol(['G/L ACCOUNT', 'GL ACCOUNT', 'G/L']);
  const colMatCode = findCol(['MATERIAL', 'MATERIAL CODE', 'MATNR']);
  const colOrder   = findCol([' PURCHASE ORDER', 'ORDER NUMBER', 'PO NUMBER', 'AUFNR', 'PURCHASING DOCUMENT']);

  const parseNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let s = val.toString().trim();
    if (s.endsWith('-')) {
      s = '-' + s.slice(0, -1);
    }
    s = s.replace(/,/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  const isProduction   = sectionName.includes('PRODUCTION');
  const isSalesOrPrice = sectionName.includes('SALES') || sectionName.includes('PRICE');

  // --- PRODUCTION CONFIG ---
  let allowedPlants = null, allowedMvts = null;
  let materialRanges = [], singleMaterials = [];
  let excludePOSeries = [];
  
  if (isProduction && rule.byCategory && rule.byCategory[categoryName]) {
    const cfg = rule.byCategory[categoryName];
    if (cfg.plants)          allowedPlants   = cfg.plants.map(String);
    if (cfg.movements)       allowedMvts     = cfg.movements.map(String);
    if (cfg.materialRanges)  materialRanges  = cfg.materialRanges;
    if (cfg.singleMaterials) singleMaterials = cfg.singleMaterials.map(x => String(x).trim());
    if (cfg.excludePOSeries) excludePOSeries = cfg.excludePOSeries.map(x => String(x).trim().replace(/^0+/, ''));
  }

  // --- SALES & PRICE CONFIG ---
  let allowedGLs = [];
  let sharedGLs = []; // NEW: Array for borrowed G/Ls
  let excludeMats = [];
  let includeMats = [];
  const globalValidGLs = new Set();

  if (isSalesOrPrice && rule.glAccounts && rule.glAccounts.byCategory) {
    let catName = categoryName.trim();
    let catConfig = rule.glAccounts.byCategory[catName];
    
    if (!catConfig) {
      const alt1 = catName + ' Sales Qty';
      const alt2 = catName.replace(/ Sales Qty$/i, '').trim();
      catConfig = rule.glAccounts.byCategory[alt1] || rule.glAccounts.byCategory[alt2];
    }

    if (catConfig) {
      if (typeof catConfig === 'object' && !Array.isArray(catConfig)) {
        if (catConfig.gls) allowedGLs = catConfig.gls.map(x => String(x).trim());
        if (catConfig.sharedGLs) sharedGLs = catConfig.sharedGLs.map(x => String(x).trim()); // Loads borrowed G/Ls
        
        if (catConfig.excludeMaterials) {
          excludeMats = catConfig.excludeMaterials.map(x => String(x).trim().replace(/^0+/, ''));
        }
        if (catConfig.includeMaterials) {
          includeMats = catConfig.includeMaterials.map(x => String(x).trim().replace(/^0+/, ''));
        }
      } else if (Array.isArray(catConfig)) {
        allowedGLs = catConfig.map(x => String(x).trim());
      }
    }

    Object.values(rule.glAccounts.byCategory).forEach(cfg => {
      if (Array.isArray(cfg)) {
        cfg.forEach(g => globalValidGLs.add(String(g).trim().replace(/^0+/, '')));
      } else {
        if (cfg.gls) cfg.gls.forEach(g => globalValidGLs.add(String(g).trim().replace(/^0+/, '')));
        if (cfg.sharedGLs) cfg.sharedGLs.forEach(g => globalValidGLs.add(String(g).trim().replace(/^0+/, '')));
      }
    });
  }

  const priceSpecialCats = new Set(['TMT', 'WRD', 'Billet', 'TMT Sales Qty']);
  const isPriceSpecialCat = isSalesOrPrice && rule.type === 'CURRENCY_RATIO' && 
                            (priceSpecialCats.has(categoryName) || priceSpecialCats.has(categoryName + ' Sales Qty'));

  function isMatAllowed(matKey) {
    if (!matKey) return false;
    const clean = matKey.toString().replace(/^0+/, '');
    if (singleMaterials.indexOf(clean) !== -1) return true;
    const num = parseInt(clean, 10);
    if (!isNaN(num)) {
      for (let i = 0; i < materialRanges.length; i++) {
        if (num >= materialRanges[i].from && num <= materialRanges[i].to) return true;
      }
    }
    return false;
  }

  function isGloballyValidMaterial(matKey) {
    if (!matKey) return false;
    const clean = matKey.toString().replace(/^0+/, '');
    const num = parseInt(clean, 10);

    if (rule.byCategory) {
      for (const catKey in rule.byCategory) {
        const catCfg = rule.byCategory[catKey];

        if (catCfg.singleMaterials) {
          const singles = catCfg.singleMaterials.map(x => String(x).trim().replace(/^0+/, ''));
          if (singles.includes(clean)) return true;
        }
        
        if (!isNaN(num) && catCfg.materialRanges) {
          for (let i = 0; i < catCfg.materialRanges.length; i++) {
            if (num >= catCfg.materialRanges[i].from && num <= catCfg.materialRanges[i].to) return true;
          }
        }
      }
    }
    return false;
  }

  function isIgnored(rawCode, exactList, rangeList) {
    if (!rawCode) return false;
    const clean = rawCode.toString().replace(/^0+/, '');
    if (IGNORE_LISTS[exactList] && IGNORE_LISTS[exactList].includes(clean)) return true;
    const num = parseInt(clean, 10);
    if (!isNaN(num) && IGNORE_LISTS[rangeList]) {
      for (let i = 0; i < IGNORE_LISTS[rangeList].length; i++) {
        const r = IGNORE_LISTS[rangeList][i];
        if (num >= r.from && num <= r.to) return true;
      }
    }
    return false;
  }

  let unknowns = [];
  let totalQty = 0;
  let totalCurrency = 0;

  for (let i = headerRowIndex + 1; i < sapData.length; i++) {
    const row = sapData[i];
    const c0 = (row[0] || '').toString().toUpperCase();
    const c1 = (row[1] || '').toString().toUpperCase();
    if (c0.includes('TOTAL') || c0.includes('*') || c0.includes('RESULT') ||
        c1.includes('TOTAL') || c1.includes('RESULT')) {
      continue;
    }

    // === PRODUCTION LOGIC ===
    if (rule.type === 'STANDARD') {
      if (colQty === -1) continue;
      
      const mvmt  = colMvmt  !== -1 ? (row[colMvmt]  || '').toString().trim() : '';
      const plant = colPlant !== -1 ? (row[colPlant] || '').toString().trim() : '';
 
      if (allowedMvts && !allowedMvts.includes(mvmt)) continue;
      if (allowedPlants && !allowedPlants.includes(plant)) continue;

      if (excludePOSeries.length > 0 && colOrder !== -1) {
        const orderNum = (row[colOrder] || '').toString().trim().replace(/^0+/, ''); // Strips leading zeros
        let isExcludedPO = false;
        
        for (let p = 0; p < excludePOSeries.length; p++) {
          if (orderNum.startsWith(excludePOSeries[p])) {
            isExcludedPO = true;
            break;
          }
        }
        if (isExcludedPO) continue;
      }

      if (materialRanges.length || singleMaterials.length) {
        const matRaw = colMatCode !== -1 ? row[colMatCode] : '';
        const mat = (matRaw || '').toString().trim();

        if (mat && !isMatAllowed(mat)) {

          if (!isGloballyValidMaterial(mat)) {
            
            if (!isIgnored(mat, 'materials', 'materialRanges')) {
              if (!unknowns.includes(mat)) unknowns.push(mat);
            }
          }
          continue;
        }
      }
      totalQty += parseNum(row[colQty]);

    // === SALES / PRICE LOGIC ===
    } else if (isSalesOrPrice && (rule.type === 'UNIT_SUM' || rule.type === 'CURRENCY_RATIO')) {
      if (colQty === -1) continue;

      if (allowedGLs.length === 0 && sharedGLs.length === 0) {
        continue;
      }

      const glRaw = colGL !== -1 ? row[colGL] : '';
      const gl = (glRaw || '').toString().trim().replace(/^0+/, '');
      
      const isMainGL = allowedGLs.includes(gl);
      const isSharedGL = sharedGLs.includes(gl);
      
      // 1. G/L FILTER
      if (!isMainGL && !isSharedGL) {
        if (gl && !globalValidGLs.has(gl)) {
          if (!isIgnored(gl, 'gls', 'glRanges')) {
            if (!unknowns.includes(gl)) unknowns.push(gl);
          }
        }
        continue; 
      }

      const matRaw = colMatCode !== -1 ? row[colMatCode] : '';
      const mat = (matRaw || '').toString().trim().replace(/^0+/, ''); 

      // 2. EXCLUDE FILTER (e.g. For Main TMT dropping '1000000075')
      if (excludeMats.length > 0 && mat && excludeMats.includes(mat)) {
        continue; 
      }

      // 3. SHARED G/L MATERIAL CHECK (The Magic Fix)
      // If this row came from a borrowed G/L, it MUST match the included material
      if (isSharedGL) {
        if (includeMats.length > 0 && !includeMats.includes(mat)) {
          continue; // Drop the row. It's a shared G/L but the wrong material code.
        }
      }

      // 4. UNIT NORMALIZATION
      const rawUnit = colUnit !== -1
        ? (row[colUnit] || '').toString().toUpperCase().trim()
        : '';
      const unitMap = { 'TON': 'MT', 'TO': 'MT', 'T': 'MT', 'MT': 'MT' };
      const unit = unitMap[rawUnit] || rawUnit;

      const qtyVal  = parseNum(row[colQty]);
      const currVal = colCurr !== -1 ? parseNum(row[colCurr]) : 0;

      // 5. FINAL SUMMATION
      if (isPriceSpecialCat) {
        if (rule.type === 'CURRENCY_RATIO') {
          totalCurrency += currVal;
        }
        if (unit === (rule.unit || 'MT').toUpperCase()) {
          totalQty += qtyVal;
        }
      } else {
        if (rule.unit && unit !== rule.unit.toUpperCase()) continue;

        totalQty += qtyVal;
        if (rule.type === 'CURRENCY_RATIO') {
          totalCurrency += currVal;
        }
      }
    }
  }

  let finalValue = 0;
  if (rule.type === 'CURRENCY_RATIO') {
    if (totalQty !== 0) {
      finalValue = Math.abs(totalCurrency / totalQty);
    }
  } else {
    finalValue = Math.abs(totalQty);
  }

  return {
    value: finalValue,
    unknowns: unknowns,
    unknownType: isProduction ? 'Material Codes' : 'G/L Accounts'
  };
}

/*************************************************
 * 7. TARGET LOCATION HELPERS
 *************************************************/
function getTargetLocation(sheetId, sheetName, targetMonth, targetCategory, targetYear, targetSection) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();

  const normalize = (t) => (t ? t.toString().replace(/\s+/g, '').trim().toUpperCase() : '');

  const cleanSection = normalize(targetSection);
  const cleanYear = normalize(targetYear);
  const shortYear = cleanYear.length >= 6 ? cleanYear.substring(2) : cleanYear; 
  const cleanCategory = normalize(targetCategory);
  const mPrefix = normalize(targetMonth).substring(0, 3);

  let expectedSectionKey = 'MONTHLYSALES';
  if (cleanSection.includes('PRODUCTION')) expectedSectionKey = 'MONTHLYPRODUCTION';
  else if (cleanSection.includes('PRICE')) expectedSectionKey = 'SALEPRICE';

  let activeSection = '';
  let activeYear = '';
  
  let catCol = -1;
  let headerRow = -1;
  let monthCol = -1;

  for (let r = 0; r < values.length; r++) {

    for (let c = 0; c < values[r].length; c++) {
      if (normalize(values[r][c]).includes(expectedSectionKey)) {
        activeSection = expectedSectionKey;
        let startRow = Math.max(0, r - 2);
        let endRow = Math.min(values.length - 1, r + 1);
  
        for (let scanRow = startRow; scanRow <= endRow; scanRow++) {
          for (let scanCol = 0; scanCol < values[scanRow].length; scanCol++) {
            let cell = normalize(values[scanRow][scanCol]);
            if (cell.includes(cleanYear) || (shortYear.length > 3 && cell.includes(shortYear))) {
              activeYear = cleanYear;
              break;
            }
          }
          if (activeYear === cleanYear) break;
        }
        break; 
      }
    }

    if (activeSection === expectedSectionKey) {
      for (let c = 0; c < values[r].length; c++) {
        let cell = normalize(values[r][c]);
        if (cell.includes(cleanYear) || (shortYear.length > 3 && cell.includes(shortYear))) {
          activeYear = cleanYear;
          break;
        }
      }
    }

    if (activeSection === expectedSectionKey && activeYear === cleanYear) {
      for (let c = 0; c < values[r].length; c++) {
        let cell = normalize(values[r][c]);

        let isMatch = false;
        if (cell === cleanCategory) {
          isMatch = true;
        } else if (cell.includes(cleanCategory)) {
          if ((cleanCategory === 'TMT' || cleanCategory === 'TMTSALESQTY') && 
              (cell.includes('RANDOM') || cell.includes('CHILLI'))) {
            isMatch = false;
          } else {
            isMatch = true;
          }
        }

        if (isMatch) {
          catCol = c;
          headerRow = r;

          for (let mc = 0; mc <= c; mc++) {
            if (normalize(values[r][mc]).includes('MONTH')) {
              monthCol = mc;
              break;
            }
          }
          if (monthCol === -1) monthCol = 0;
          
          break;
        }
      }
    }

    if (catCol > -1) break;
  }

  if (catCol === -1) {
    throw new Error('Could not find "' + targetCategory + '" under "' + targetSection + '" for year ' + targetYear + '.');
  }

  let targetRow = -1;
  for (let r = headerRow + 1; r < values.length; r++) {
    const monthCell = normalize(values[r][monthCol]);

    if (monthCell.includes('TOTAL') || monthCell.match(/\d{4}-\d{2}/) || monthCell.match(/\d{2}-\d{2}/)) {
      break;
    }

    if (monthCell.startsWith(mPrefix)) {
      targetRow = r;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error('Could not find Month "' + targetMonth + '" under ' + targetCategory);
  }

  return { row: targetRow + 1, column: catCol + 1 };
}

/*************************************************
 * 8. WRITE TO SHEET
 *************************************************/
function writeTotalToGoogleSheet(sheetId, sheetName, row, column, totalQuantity) {
  SpreadsheetApp.openById(sheetId).getSheetByName(sheetName).getRange(row, column).setValue(totalQuantity);
}

/*************************************************
 * 9. FILE PARSING
 *************************************************/
function parseFile(blob, fileName) {
  const ext = String(fileName).split('.').pop().toLowerCase();
  if (ext === 'csv') return Utilities.parseCsv(blob.getDataAsString());
  if (ext === 'xlsx' || ext === 'xls') return parseExcel(blob);
  throw new Error('Unsupported format');
}

function parseExcel(blob) {
  let tempId;
  try {
    const f = Drive.Files.insert({ title: 'TEMP', mimeType: MimeType.GOOGLE_SHEETS }, blob);
    tempId = f.id;
    return SpreadsheetApp.openById(tempId).getSheets()[0].getDataRange().getValues();
  } catch(e) { throw e; }
  finally { if(tempId) try{ Drive.Files.remove(tempId); } catch(e){} }
}

function getMimeType(fileName) {
  return fileName.endsWith('.csv') ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
}