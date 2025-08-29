/**
 * Service Cards Dashboard App Script
 * Location: MedSpa Rankings 2025 Google Workbook Sheet
 * 
 * This App Script powers the service cards visualization on the dashboard of the web app.
 * It processes and returns data for both normal and emerging medical spa services,
 * showing their search volumes, trends, and market metrics.
 * 
 * The script works with a spreadsheet containing service-specific keyword data across
 * different geographical locations (US & Canada) and calculates:
 * - Search volume trends
 * - Competition metrics
 * - Cost per click (CPC) averages
 * - Market share percentages
 * 
 * SUPPORTED LOCATION PARAMETERS:
 * - "USA" → Uses "US (Whole)" sheet, filters by Country column
 * - "Canada" → Uses "Canada (Whole)" sheet, filters by Country column  
 * - "City, State" (e.g., "New York, NY") → Uses "City (US & CA)" sheet, filters by City column
 * - "State/Province" (e.g., "Alabama", "Ontario") → Uses "State & Province" sheet, filters by State column
 * 
 * For other dashboard functionalities (Keywords, Backlinks, GeoGrid Maps),
 * please reference Client Workbook App Script.
 */

// --- Configuration ---
// Enable verbose logging while we diagnose state filtering issues
var DEBUG = true;
if (!DEBUG && typeof console !== 'undefined' && typeof console.log === 'function') {
  console.log = function() {};
}
const NORMAL_SERVICES = [
  'Botox', 'Lip Filler', 'Laser Facial', 'Semaglutide', 'HydraFacial', 
  'Laser hair Removal', 'Body Contouring', 'Skin Tighenting', 'IV Therapy', 
  'Dermal Fillers', 'Microneedling', 'Chemical Peel', 'Red Light Therapy', 
  'Kybella', 'Emsella', 'RF Microneedling'
];

const NEW_SERVICES = [
  'Polynucleotides (Salmon Sperm Facial)', 'Cryotherapy Facial', 'Stem Cell Therapy', 
  'Exosome Therapy', 'Platelet-Rich Fibrin (PRF)', 'Sofwave', 'Oxygen Facial', 
  'BioRePeel', 'SkinVive', 'NAD'
];


function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const location = e.parameter.location || 'USA'; // Default to USA if no location is specified
    
    console.log(`=== DEBUGGING doGet function ===`);
    console.log(`Incoming request - location parameter: "${location}"`);
    console.log(`Request object:`, JSON.stringify(e.parameter));

    // If listing cities for selector
    if (action === 'listCities') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let citySheet = ss.getSheetByName('City (US & CA)') || ss.getSheetByName('City (US & Canada)');
      if (!citySheet) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'City sheet not found', cities: [], _debug: { availableSheets: ss.getSheets().map(function(s){ return s.getName(); }) } })).setMimeType(ContentService.MimeType.JSON);
      }
      const lastRow = citySheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ cities: [] })).setMimeType(ContentService.MimeType.JSON);
      }
      // City, State, Country columns are 3,4,5
      const values = citySheet.getRange(2, 3, lastRow - 1, 3).getValues();
      const seen = {};
      const items = [];
      for (var i = 0; i < values.length; i++) {
        var city = values[i][0];
        var state = values[i][1];
        var country = values[i][2];
        if (!city || !state) continue;
        var id = String(city).trim() + ', ' + String(state).trim();
        var key = id + '||' + String(country || '').trim();
        if (seen[key]) continue;
        seen[key] = true;
        items.push({ id: id, city: String(city).trim(), state: String(state).trim(), country: String(country || '').trim() });
      }
      items.sort(function(a,b){ return (a.id.toLowerCase() < b.id.toLowerCase()) ? -1 : 1; });
      return ContentService.createTextOutput(JSON.stringify({ cities: items })).setMimeType(ContentService.MimeType.JSON);
    }

    // Determine which sheet to use based on the location parameter
    let sheetName;
    let locationColumnIndex; // 2 for City, 3 for State, 4 for Country
    let locationFilterValue = location;
    let secondaryToken = null; // used for City route (state/province token)

    if (location === 'USA') {
        sheetName = 'US (Whole)';
        locationColumnIndex = 4; // Country column
        console.log(`USA route selected`);
    } else if (location === 'Canada') {
        sheetName = 'Canada (Whole)';
        locationColumnIndex = 4; // Country column
        console.log(`Canada route selected`);
    } else if (location.indexOf(',') !== -1) {
        // Disambiguate between "City, ST" and "State, Country"
        var parts = location.split(',');
        var left = (parts[0] || '').trim();
        var rightRaw = (parts[1] || '').trim();
        var right = rightRaw.toLowerCase();

        // Known country tokens that indicate State/Province route
        var countryTokens = { 'usa':1, 'united states':1, 'canada':1, 'ca':1, 'us':1 };

        // Known state/province abbreviations (US + CA common) – used to detect City, ST
        var stateAbbrevsArr = ['al','ak','az','ar','ca','co','ct','de','fl','ga','hi','id','il','in','ia','ks','ky','la','me','md','ma','mi','mn','ms','mo','mt','ne','nv','nh','nj','nm','ny','nc','nd','oh','ok','or','pa','ri','sc','sd','tn','tx','ut','vt','va','wa','wv','wi','wy','ab','bc','mb','nb','nl','ns','nt','nu','on','pe','qc','sk','yt'];
        var stateAbbrevs = {};
        for (var i=0;i<stateAbbrevsArr.length;i++){ stateAbbrevs[stateAbbrevsArr[i]] = 1; }

        if (countryTokens[right]) {
          // e.g., "Alabama, USA" → State & Province, filter = "Alabama"
          sheetName = 'State & Province';
          locationColumnIndex = 3; // State column
          locationFilterValue = left;
          console.log('STATE/PROVINCE route selected (State, Country pattern) for:', location);
        } else if (rightRaw.length === 2 && stateAbbrevs[right]) {
          // e.g., "New York, NY" → City route
          sheetName = 'City (US & CA)';
          locationColumnIndex = 2; // City column
          locationFilterValue = left;
          secondaryToken = rightRaw; // state/province abbr
          console.log('City route selected (City, ST) for:', location);
        } else {
          // Fallback heuristic: If the left side looks like a multi-word city name (contains space)
          // treat as city, otherwise treat as state/province
          var isLikelyCity = /\s/.test(left);
          if (isLikelyCity) {
            sheetName = 'City (US & CA)';
            locationColumnIndex = 2; // City column
            locationFilterValue = left;
            secondaryToken = rightRaw || null; // could be full state/province name
            console.log('City route selected (fallback heuristic) for:', location);
          } else {
            sheetName = 'State & Province';
            locationColumnIndex = 3; // State column
            locationFilterValue = left;
            console.log('STATE/PROVINCE route selected (fallback heuristic) for:', location);
          }
        }
    } else { 
        // Single token – assume it's a State or Province name
        sheetName = 'State & Province';
        locationColumnIndex = 3; // State column
        console.log(`STATE/PROVINCE route selected for: ${location}`);
    }

    console.log(`ROUTE DECISION: location="${location}", sheetName="${sheetName}", columnIndex=${locationColumnIndex}, filterValue="${locationFilterValue}", secondaryToken="${secondaryToken || ''}"`);
    
    // Get all available sheets for debugging
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const availableSheetNames = allSheets.map(function(s){ return s.getName(); });
    console.log('All available sheets in workbook:', availableSheetNames);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet && sheetName === 'City (US & CA)') {
      var fallback = ss.getSheetByName('City (US & Canada)');
      if (fallback) {
        sheet = fallback;
        console.log('Using fallback city sheet name: City (US & Canada)');
      }
    }
    if (!sheet) {
      console.error('CRITICAL ERROR: Sheet "' + sheetName + '" not found!');
      console.error('Available sheets:', availableSheetNames.join(', '));
      throw new Error('Sheet "' + sheetName + '" not found. Available sheets: ' + availableSheetNames.join(', '));
    }
    
    console.log('✅ Successfully found sheet:', sheetName);
    console.log('Sheet dimensions:', sheet.getLastRow(), 'rows x', sheet.getLastColumn(), 'columns');
    
    // Let's also check the first few rows of the sheet to understand the structure
    if (sheet.getLastRow() > 0) {
      const headerRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
      console.log('Sheet headers (first 10):', headerRow);
      
      if (sheet.getLastRow() > 1) {
        const sampleRow = sheet.getRange(2, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
        console.log('Sample data row:', sampleRow);
      }
    }

    const data = aggregateServiceData(sheet, locationColumnIndex, locationFilterValue, sheetName, secondaryToken);

    console.log('\n=== FINAL RETURN FROM doGet ===');
    console.log('Returning data:', JSON.stringify(data, null, 2));

    // Add debug information to the response for browser console
    let targetColumnHeader = 'N/A';
    if (sheet && sheet.getLastColumn() > locationColumnIndex) {
      try {
        targetColumnHeader = sheet.getRange(1, locationColumnIndex + 1).getValue() || 'Empty';
      } catch (e) {
        targetColumnHeader = 'Error: ' + e.message;
      }
    }
    
    // Collect sample values from the target column (for debugging state matching)
    var stateSamples = [];
    var uniqueStatesSample = [];
    try {
      var totalRows = sheet.getLastRow();
      if (totalRows > 1) {
        var colValues = sheet
          .getRange(2, locationColumnIndex + 1, Math.max(0, totalRows - 1), 1)
          .getValues()
          .map(function(r){ return (r[0] !== null && r[0] !== undefined) ? r[0].toString() : ''; });
        stateSamples = colValues.slice(0, 20);
        var seen = {};
        for (var i = 0; i < colValues.length && uniqueStatesSample.length < 50; i++) {
          var v = colValues[i].toString().trim();
          if (v && !seen[v]) { seen[v] = true; uniqueStatesSample.push(v); }
        }
      }
    } catch (e) {
      stateSamples = ['Error collecting samples: ' + e.message];
    }

    const debugInfo = {
      requestLocation: location,
      selectedSheet: sheetName,
      targetColumn: locationColumnIndex,
      targetColumnHeader: targetColumnHeader,
      sheetRows: sheet ? sheet.getLastRow() : 0,
      sheetColumns: sheet ? sheet.getLastColumn() : 0,
      dataRowsFound: data.topServices.length + data.newServices.length,
      availableSheets: SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function(s){ return s.getName(); }),
      hasSheet: !!sheet,
      // expose month headers we detected and a few unique states to help debug
      monthHeaders: (function(){
        try {
          const hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          return hdrs.filter(function(h){ return /\b(?:january|february|march|april|may|june|july|august|september|october|november|december)\b/i.test(String(h)); });
        } catch (e) { return []; }
      })(),
      stateSamples: stateSamples,
      uniqueStatesSample: uniqueStatesSample,
      targetPresentExact: uniqueStatesSample.indexOf(locationFilterValue) !== -1,
      targetPresentCaseInsensitive: (function(){
        var lower = locationFilterValue.toLowerCase();
        return uniqueStatesSample.some(function(v){ return v.toLowerCase() === lower; });
      })()
    };
    
    // Include debug info in response
    const responseData = {
      ...data,
      _debug: debugInfo
    };

    // Support for JSONP
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + JSON.stringify(responseData) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    console.error(`\n=== ERROR IN doGet ===`);
    console.error(`Error message:`, error.message);
    console.error(`Error stack:`, error.stack);
    console.error(`Request parameters:`, JSON.stringify(e.parameter));
    
    // Include debug info in error response too
    const errorDebugInfo = {
      requestLocation: e.parameter.location || 'not provided',
      availableSheets: 'Error occurred before sheet access',
      errorType: error.name,
      errorMessage: error.message
    };
    
    try {
      errorDebugInfo.availableSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function(s){ return s.getName(); });
    } catch (sheetError) {
      errorDebugInfo.availableSheets = 'Sheet access error: ' + sheetError.message;
    }
    
    const errorResponse = { 
      error: error.message, 
      stack: error.stack,
      _debug: errorDebugInfo,
      topServices: [],
      newServices: []
    };
     if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + JSON.stringify(errorResponse) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(JSON.stringify(errorResponse)).setMimeType(ContentService.MimeType.JSON);
    }
  }
}

// Custom parser to handle "Month YYYY" format
function parseDateHeader(header) {
    // Accept Google Sheets Date objects directly
    if (header instanceof Date && !isNaN(header.getTime())) {
        return header;
    }
    // Accept numeric serials (in case values are numeric). Google Sheets usually returns Date,
    // but this is a safe fallback. Serial is days since 1899-12-30 in Sheets.
    if (typeof header === 'number' && isFinite(header)) {
        var serialEpoch = new Date(1899, 11, 30);
        var millis = serialEpoch.getTime() + Math.round(header * 24 * 60 * 60 * 1000);
        var d = new Date(millis);
        if (!isNaN(d.getTime())) return d;
    }
    // Accept strings like "May 2025"
    if (typeof header === 'string') {
        var months = {
            'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
            'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11
        };
        var str = header.trim();
        var parts = str.split(/\s+/);
        if (parts.length === 2) {
            var m = months[parts[0].toLowerCase()];
            var y = parseInt(parts[1], 10);
            if (m !== undefined && !isNaN(y)) {
                return new Date(y, m, 1);
            }
        }
        // Final fallback: let JS try to parse
        var parsed = new Date(str);
        if (!isNaN(parsed.getTime())) return parsed;
    }
    return null;
}

function aggregateServiceData(sheet, locationColumnIndex, locationFilterValue, sheetName, secondaryToken) {
  console.log(`\n=== AGGREGATE SERVICE DATA FUNCTION ===`);
  console.log(`Input parameters: sheet=${sheet ? 'exists' : 'null'}, locationColumnIndex=${locationColumnIndex}, locationFilterValue="${locationFilterValue}", sheetName="${sheetName}", secondaryToken="${secondaryToken || ''}"`);
  
  if (!sheet) {
    console.error(`ERROR: No sheet provided to aggregateServiceData`);
    return { topServices: [], newServices: [] };
  }
  
  const allData = sheet.getDataRange().getValues();
  console.log(`Raw data retrieved: ${allData.length} total rows (including header)`);
  
  const headers = allData.shift(); // Get and remove header row
  console.log(`After removing header: ${allData.length} data rows remaining`);
  
  console.log(`=== SHEET ANALYSIS ===`);
  console.log(`Sheet name: ${sheetName}`);
  console.log(`Headers (${headers.length} columns):`, headers);
  console.log(`Target column index: ${locationColumnIndex}`);
  console.log(`Target column header: "${headers[locationColumnIndex]}"`);
  console.log(`Looking for location: "${locationFilterValue}"`);
  
  if (allData.length === 0) {
    console.error(`ERROR: No data rows found in sheet after removing header!`);
    return { topServices: [], newServices: [] };
  }

  // For City route, optionally enforce state match using secondaryToken
  let data = allData;
  if (sheetName === 'City (US & CA)') {
    const cityCol = 2 - 1 + 1; // we will use row[2] below; keeping log simple
    const stateColIndex = 3; // State column is index 3 (0-based)
    const cityFilter = String(locationFilterValue || '').trim().toLowerCase();
    const secondary = secondaryToken ? String(secondaryToken).trim().toLowerCase() : '';

    data = allData.filter(function(row){
      var cityVal = (row[2] || '').toString().trim().toLowerCase();
      if (!cityVal) return false;
      if (cityVal !== cityFilter) return false;
      if (!secondary) return true; // no state constraint provided
      var stateVal = (row[stateColIndex] || '').toString().trim().toLowerCase();
      if (!stateVal) return true; // allow if state missing
      // Accept either abbr or full name containment in either direction for robustness
      return stateVal === secondary || stateVal.indexOf(secondary) !== -1 || secondary.indexOf(stateVal) !== -1;
    });
  } else {
    // Existing filtering for Country/State paths
    data = allData.filter(function(row){
      const rowLocation = row[locationColumnIndex];
      if (!rowLocation) return false;
      const rowLocationStr = rowLocation.toString().trim().toLowerCase();
      const filterValueStr = locationFilterValue.trim().toLowerCase();
      if (rowLocationStr === filterValueStr) return true;
      if (sheetName === 'State & Province') {
        if (rowLocationStr.indexOf(filterValueStr) !== -1 || filterValueStr.indexOf(rowLocationStr) !== -1) return true;
        const cleanRow = rowLocationStr.replace(/\b(state|province)\b/g, '').trim();
        const cleanFilter = filterValueStr.replace(/\b(state|province)\b/g, '').trim();
        if (cleanRow === cleanFilter) return true;
      }
      return false;
    });
  }

  if (data.length === 0) {
    console.warn('No matching rows after initial filtering for', sheetName, 'with location:', locationFilterValue, 'and secondaryToken:', secondaryToken);
    return { topServices: [], newServices: [] };
  }

  // Find the indices of all month columns using the robust date parser
  const monthColumns = headers.reduce(function(acc, header, index) {
    var date = parseDateHeader(header);
    if (date) {
      acc.push({ header: header, index: index, date: date });
    }
    return acc;
  }, []);
  monthColumns.sort(function(a, b) { return b.date - a.date; });

  // Compute aggregates (unchanged)
  const serviceCol = headers.indexOf('Service');
  const keywordCol = headers.indexOf('Keyword');
  const compCol = headers.indexOf('Competition Index');
  const cpcCol = headers.indexOf('CPC');

  if (serviceCol === -1 || keywordCol === -1 || monthColumns.length === 0) {
    console.error('Essential columns (Service, Keyword, Months) not found');
    return { topServices: [], newServices: [] };
  }

  const currentMonth = monthColumns[0];
  const previousMonth = monthColumns.length > 1 ? monthColumns[1] : null;

  const aggregatedData = {};
  data.forEach(function(row){
    const serviceName = row[serviceCol];
    const keyword = row[keywordCol];
    if (!serviceName || !keyword) return;
    if (!aggregatedData[serviceName]) {
      aggregatedData[serviceName] = {
        name: serviceName,
        totalCompetition: 0,
        totalCpc: 0,
        totalCurrentVolume: 0,
        totalPreviousVolume: 0,
        keywords: [],
        keywordCount: 0
      };
    }
    const service = aggregatedData[serviceName];
    const currentVol = parseInt(row[currentMonth.index], 10) || 0;
    const prevVol = previousMonth ? (parseInt(row[previousMonth.index], 10) || 0) : 0;
    let keywordTrend = 0; if (currentVol > prevVol) keywordTrend = 1; else if (currentVol < prevVol) keywordTrend = -1;
    service.keywords.push({ name: keyword, volume: currentVol, trend: keywordTrend });
    const competition = (compCol !== -1 && row[compCol] !== undefined) ? parseFloat(row[compCol]) : 0;
    const cpc = (cpcCol !== -1 && row[cpcCol] !== undefined) ? parseFloat(row[cpcCol]) : 0;
    if(!isNaN(competition) && competition > 0) service.totalCompetition += competition;
    if(!isNaN(cpc) && cpc > 0) service.totalCpc += cpc;
    service.totalCurrentVolume += currentVol;
    service.totalPreviousVolume += prevVol;
    service.keywordCount++;
  });

  const allServices = Object.values(aggregatedData);
  const allServicesTotalVolume = allServices.reduce(function(sum, s){ return sum + s.totalCurrentVolume; }, 0);
  const formattedServices = allServices.map(function(service){
    const numKeywords = service.keywordCount; if (numKeywords === 0) return null;
    let serviceTrend = 0; if (service.totalCurrentVolume > service.totalPreviousVolume) serviceTrend = 1; else if (service.totalCurrentVolume < service.totalPreviousVolume) serviceTrend = -1;
    return {
      name: service.name,
      totalVolume: service.totalCurrentVolume,
      volumePercentage: allServicesTotalVolume > 0 ? ((service.totalCurrentVolume / allServicesTotalVolume) * 100).toFixed(1) : '0.0',
      avgCompetition: (service.totalCompetition / numKeywords).toFixed(2),
      avgCpc: (service.totalCpc / numKeywords).toFixed(2),
      trend: serviceTrend,
      keywords: service.keywords.sort(function(a,b){ return b.volume - a.volume; })
    };
  }).filter(function(s){ return s !== null && s.totalVolume > 0; });

  const topServices = formattedServices
      .filter(function(s){ return NORMAL_SERVICES.indexOf(s.name) !== -1; })
      .sort(function(a,b){ return b.totalVolume - a.totalVolume; });

  const newServices = formattedServices
      .filter(function(s){ return NEW_SERVICES.indexOf(s.name) !== -1; })
      .sort(function(a,b){ return b.totalVolume - a.totalVolume; });
  
  return { topServices: topServices, newServices: newServices };
}

// --- DEPRECATED ---
// The function below is the old implementation and is no longer used.
// It is kept for reference but will be removed in a future version.
function processSheetData(sheet, locationColumnIndex, locationFilterValue) {
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();

    // --- ROBUST DATE COLUMN FINDER ---
    const dateColumns = [];
    header.forEach((cell, index) => {
        // Try our custom parser first for "Month YYYY"
        let date = parseDateHeader(cell);
        
        // Fallback for standard date formats if the custom parser fails
        if (!date && cell) {
            let parsed = new Date(cell);
            if (!isNaN(parsed.getTime())) {
                date = parsed;
            }
        }
        
        if (date) {
            dateColumns.push({ index: index, date: date });
        }
    });

    // Sort date columns to find the most recent two
    dateColumns.sort((a, b) => b.date - a.date);

    if (dateColumns.length === 0) {
        throw new Error("No valid date columns found in the header.");
    }
    
    const lastDateColIndex = dateColumns[0].index;
    const prevDateColIndex = dateColumns.length > 1 ? dateColumns[1].index : -1;
    
    const services = {};

    // Start from row 1 to skip header
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowLocation = row[locationColumnIndex];
        
        // Filter rows by the specified location
        if (rowLocation && rowLocation.toString().trim().toLowerCase() === locationFilterValue.trim().toLowerCase()) {
            const serviceName = row[0];
            const keyword = row[1];
            const competition = parseFloat(row[5]);
            const cpc = parseFloat(row[6]);
            const currentVolume = parseInt(row[lastDateColIndex], 10) || 0;
            const prevVolume = prevDateColIndex !== -1 ? (parseInt(row[prevDateColIndex], 10) || 0) : 0;

            if (!services[serviceName]) {
                services[serviceName] = {
                    totalVolume: 0,
                    totalPrevVolume: 0,
                    competitionSum: 0,
                    cpcSum: 0,
                    keywordCount: 0,
                    keywords: []
                };
            }

            services[serviceName].totalVolume += currentVolume;
            services[serviceName].totalPrevVolume += prevVolume;
            if (!isNaN(competition)) services[serviceName].competitionSum += competition;
            if (!isNaN(cpc)) services[serviceName].cpcSum += cpc;
            services[serviceName].keywordCount++;
            
            let trend = 0; // 0 for no change, 1 for up, -1 for down
            if (currentVolume > prevVolume) trend = 1;
            if (currentVolume < prevVolume) trend = -1;

            services[serviceName].keywords.push({
                name: keyword,
                competition: competition,
                cpc: cpc,
                volume: currentVolume,
                trend: trend
            });
        }
    }
    
    const allServicesTotalVolume = Object.values(services).reduce((acc, s) => acc + s.totalVolume, 0);

    const formattedServices = Object.keys(services).map(serviceName => {
        const service = services[serviceName];
        const overallTrend = service.totalVolume > service.totalPrevVolume ? 1 : (service.totalVolume < service.totalPrevVolume ? -1 : 0);

        return {
            name: serviceName,
            percentage: allServicesTotalVolume > 0 ? ((service.totalVolume / allServicesTotalVolume) * 100).toFixed(1) + '%' : '0.0%',
            avgCompetition: service.keywordCount > 0 ? (service.competitionSum / service.keywordCount).toFixed(2) : '0.00',
            avgCpc: service.keywordCount > 0 ? (service.cpcSum / service.keywordCount).toFixed(2) : '0.00',
            trend: overallTrend,
            keywords: service.keywords,
            totalVolume: service.totalVolume // Add totalVolume for sorting
        };
    });

    const topServices = formattedServices
        .filter(s => NORMAL_SERVICES.includes(s.name))
        .sort((a, b) => b.totalVolume - a.totalVolume);

    const newServices = formattedServices
        .filter(s => NEW_SERVICES.includes(s.name))
        .sort((a, b) => b.totalVolume - a.totalVolume);
        
    return { topServices, newServices };
} 