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
// Disable verbose logging in production to reduce execution overhead
var DEBUG = false;
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
    const location = e.parameter.location || 'USA'; // Default to USA if no location is specified
    
    console.log(`=== DEBUGGING doGet function ===`);
    console.log(`Incoming request - location parameter: "${location}"`);
    console.log(`Request object:`, JSON.stringify(e.parameter));
    
    // Determine which sheet to use based on the location parameter
    let sheetName;
    let locationColumnIndex; // 2 for City, 3 for State, 4 for Country
    let locationFilterValue = location;

    if (location === 'USA') {
        sheetName = 'US (Whole)';
        locationColumnIndex = 4; // Country column
        console.log(`USA route selected`);
    } else if (location === 'Canada') {
        sheetName = 'Canada (Whole)';
        locationColumnIndex = 4; // Country column
        console.log(`Canada route selected`);
    } else if (location.includes(',')) {
        // Disambiguate between "City, ST" and "State, Country"
        const parts = location.split(',').map(p => p.trim());
        const left = parts[0] || '';
        const right = (parts[1] || '').toLowerCase();

        // Known country tokens that indicate State/Province route
        const countryTokens = new Set(['usa', 'united states', 'canada', 'ca', 'us']);

        // Known state/province abbreviations (US + CA common) – used to detect City, ST
        const stateAbbrevs = new Set([
          'al','ak','az','ar','ca','co','ct','de','fl','ga','hi','id','il','in','ia','ks','ky','la','me','md','ma','mi','mn','ms','mo','mt','ne','nv','nh','nj','nm','ny','nc','nd','oh','ok','or','pa','ri','sc','sd','tn','tx','ut','vt','va','wa','wv','wi','wy',
          'ab','bc','mb','nb','nl','ns','nt','nu','on','pe','qc','sk','yt'
        ]);

        if (countryTokens.has(right)) {
          // e.g., "Alabama, USA" → State & Province, filter = "Alabama"
          sheetName = 'State & Province';
          locationColumnIndex = 3; // State column
          locationFilterValue = left;
          console.log(`STATE/PROVINCE route selected (State, Country pattern) for: ${location}`);
        } else if (right.length === 2 && stateAbbrevs.has(right)) {
          // e.g., "New York, NY" → City route
          sheetName = 'City (US & CA)';
          locationColumnIndex = 2; // City column
          console.log(`City route selected (City, ST) for: ${location}`);
        } else {
          // Fallback heuristic: If the left side looks like a multi-word city name (contains space)
          // treat as city, otherwise treat as state/province
          const isLikelyCity = /\s/.test(left);
          if (isLikelyCity) {
            sheetName = 'City (US & CA)';
            locationColumnIndex = 2; // City column
            console.log(`City route selected (fallback heuristic) for: ${location}`);
          } else {
            sheetName = 'State & Province';
            locationColumnIndex = 3; // State column
            locationFilterValue = left;
            console.log(`STATE/PROVINCE route selected (fallback heuristic) for: ${location}`);
          }
        }
    } else { 
        // Single token – assume it's a State or Province name
        sheetName = 'State & Province';
        locationColumnIndex = 3; // State column
        console.log(`STATE/PROVINCE route selected for: ${location}`);
    }

    console.log(`ROUTE DECISION: location="${location}", sheetName="${sheetName}", columnIndex=${locationColumnIndex}, filterValue="${locationFilterValue}"`);
    
    // Get all available sheets for debugging
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const availableSheetNames = allSheets.map(s => s.getName());
    console.log(`All available sheets in workbook:`, availableSheetNames);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      console.error(`CRITICAL ERROR: Sheet "${sheetName}" not found!`);
      console.error(`Available sheets: ${availableSheetNames.join(', ')}`);
      throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${availableSheetNames.join(', ')}`);
    }
    
    console.log(`✅ Successfully found sheet: ${sheetName}`);
    console.log(`Sheet dimensions: ${sheet.getLastRow()} rows x ${sheet.getLastColumn()} columns`);
    
    // Let's also check the first few rows of the sheet to understand the structure
    if (sheet.getLastRow() > 0) {
      const headerRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
      console.log(`Sheet headers (first 10):`, headerRow);
      
      if (sheet.getLastRow() > 1) {
        const sampleRow = sheet.getRange(2, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
        console.log(`Sample data row:`, sampleRow);
      }
    }

    const data = aggregateServiceData(sheet, locationColumnIndex, locationFilterValue, sheetName);

    console.log(`\n=== FINAL RETURN FROM doGet ===`);
    console.log(`Returning data:`, JSON.stringify(data, null, 2));

    // Add debug information to the response for browser console
    let targetColumnHeader = 'N/A';
    if (sheet && sheet.getLastColumn() > locationColumnIndex) {
      try {
        targetColumnHeader = sheet.getRange(1, locationColumnIndex + 1).getValue() || 'Empty';
      } catch (e) {
        targetColumnHeader = `Error: ${e.message}`;
      }
    }
    
    const debugInfo = {
      requestLocation: location,
      selectedSheet: sheetName,
      targetColumn: locationColumnIndex,
      targetColumnHeader: targetColumnHeader,
      sheetRows: sheet ? sheet.getLastRow() : 0,
      sheetColumns: sheet ? sheet.getLastColumn() : 0,
      dataRowsFound: data.topServices.length + data.newServices.length,
      availableSheets: SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()),
      hasSheet: !!sheet
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
      errorDebugInfo.availableSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
    } catch (sheetError) {
      errorDebugInfo.availableSheets = `Sheet access error: ${sheetError.message}`;
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
    if (!header || typeof header !== 'string') return null;
    const months = {
        'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
        'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11
    };
    const parts = header.trim().split(/\s+/);
    if (parts.length !== 2) return null;
    
    const month = months[parts[0].toLowerCase()];
    const year = parseInt(parts[1], 10);

    if (month !== undefined && !isNaN(year)) {
        return new Date(year, month, 1);
    }
    return null;
}

function aggregateServiceData(sheet, locationColumnIndex, locationFilterValue, sheetName) {
  console.log(`\n=== AGGREGATE SERVICE DATA FUNCTION ===`);
  console.log(`Input parameters: sheet=${sheet ? 'exists' : 'null'}, locationColumnIndex=${locationColumnIndex}, locationFilterValue="${locationFilterValue}", sheetName="${sheetName}"`);
  
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
  
  // Enhanced sample row logging
  if (allData.length > 0) {
    console.log(`\n=== DATA SAMPLE ANALYSIS ===`);
    for (let i = 0; i < Math.min(3, allData.length); i++) {
      const row = allData[i];
      console.log(`Row ${i}:`);
      console.log(`  Full row:`, row);
      console.log(`  Service: "${row[0]}"`);
      console.log(`  Keyword: "${row[1]}"`);
      console.log(`  City (col 2): "${row[2]}"`);
      console.log(`  State (col 3): "${row[3]}"`);  
      console.log(`  Country (col 4): "${row[4]}"`);
      console.log(`  Target column value: "${row[locationColumnIndex]}"`);
      console.log(`---`);
    }
    
    // Check if the target column has any data at all
    const targetColumnValues = allData.map(row => row[locationColumnIndex]).filter(val => val && val.toString().trim());
    console.log(`Target column has ${targetColumnValues.length} non-empty values out of ${allData.length} total rows`);
    console.log(`First 10 target column values:`, targetColumnValues.slice(0, 10));
  } else {
    console.error(`ERROR: No data rows found in sheet after removing header!`);
    return { topServices: [], newServices: [] };
  }

  // Find the indices of all month columns using the robust date parser
  const monthColumns = headers.reduce((acc, header, index) => {
    const date = parseDateHeader(header);
    if (date) {
      acc.push({ header, index, date });
    }
    return acc;
  }, []);

  // Sort by date to find the latest two months
  monthColumns.sort((a, b) => b.date - a.date);

  // Before filtering, let's count how many rows match our target location
  const locationCounts = {};
  const rawLocationSamples = []; // Collect raw samples for debugging
  
  allData.forEach((row, index) => {
    const loc = row[locationColumnIndex];
    if (loc) {
      const locStr = loc.toString().trim().toLowerCase();
      locationCounts[locStr] = (locationCounts[locStr] || 0) + 1;
      
      // Collect first 20 raw samples for debugging
      if (rawLocationSamples.length < 20) {
        rawLocationSamples.push({
          rowIndex: index,
          rawValue: loc,
          stringValue: loc.toString(),
          trimmedValue: loc.toString().trim(),
          lowercaseValue: locStr
        });
      }
    }
  });
  
  // Enhanced debugging for states
  if (sheetName === 'State & Province') {
    console.log(`\n=== RAW STATE DATA SAMPLES ===`);
    rawLocationSamples.forEach((sample, idx) => {
      console.log(`Sample ${idx + 1}: Raw="${sample.rawValue}" | String="${sample.stringValue}" | Trimmed="${sample.trimmedValue}" | Lowercase="${sample.lowercaseValue}"`);
    });
  }
  
  console.log(`\n=== LOCATION MATCHING ANALYSIS ===`);
  console.log(`Location counts for column ${locationColumnIndex}:`, locationCounts);
  console.log(`Looking for "${locationFilterValue.toLowerCase()}" - found ${locationCounts[locationFilterValue.toLowerCase()] || 0} matches`);
  
  // Show all unique location values for comparison
  const uniqueLocations = [...new Set(Object.keys(locationCounts))].sort();
  console.log(`All unique locations in data (${uniqueLocations.length} total):`, uniqueLocations);
  
  // Try to find close matches to help debug
  const targetLower = locationFilterValue.toLowerCase();
  const closeMatches = uniqueLocations.filter(loc => 
    loc.includes(targetLower) || targetLower.includes(loc) ||
    loc.startsWith(targetLower) || targetLower.startsWith(loc)
  );
  console.log(`Close matches to "${targetLower}":`, closeMatches);
  
  // For states, show more detailed analysis
  if (sheetName === 'State & Province') {
    console.log(`\n=== STATE MATCHING ANALYSIS ===`);
    console.log(`Target state: "${targetLower}"`);
    console.log(`Total unique states found: ${uniqueLocations.length}`);
    console.log(`First 10 unique states:`, uniqueLocations.slice(0, 10));
    console.log(`States containing "${targetLower}":`, uniqueLocations.filter(s => s.includes(targetLower)));
    console.log(`States that "${targetLower}" contains:`, uniqueLocations.filter(s => targetLower.includes(s)));
    
    // Check for common state abbreviations
    const stateAbbrevMap = {'arkansas': 'ar', 'delaware': 'de', 'alabama': 'al', 'nebraska': 'ne'};
    const abbrev = stateAbbrevMap[targetLower];
    if (abbrev) {
      console.log(`Looking for abbreviation "${abbrev}":`, uniqueLocations.filter(s => s === abbrev));
    }
  }
  
  console.log(`\n=== STARTING FIRST-PASS FILTERING ===`);
  
  // Filter rows for the selected location
  let data = allData.filter(row => {
      const rowLocation = row[locationColumnIndex];
      if (!rowLocation) return false;
      
      const rowLocationStr = rowLocation.toString().trim().toLowerCase();
      const filterValueStr = locationFilterValue.trim().toLowerCase();
      
      // Enhanced debugging for state filtering
      if (sheetName === 'State & Province') {
          const rowIndex = allData.indexOf(row);
          // Log first 10 comparisons and any that are close matches
          if (rowIndex < 10 || rowLocationStr.includes(filterValueStr) || filterValueStr.includes(rowLocationStr)) {
              console.log(`Row ${rowIndex}: "${rowLocationStr}" vs "${filterValueStr}" | Exact: ${rowLocationStr === filterValueStr} | Contains: ${rowLocationStr.includes(filterValueStr)} | Service: "${row[0]}"`);
          }
      }
      
      // First try exact match
      if (rowLocationStr === filterValueStr) return true;
      
      // For states, also try partial matching (in case of abbreviations or slight differences)
      if (sheetName === 'State & Province') {
          // Check if either contains the other (for cases like "NY" vs "New York")
          if (rowLocationStr.includes(filterValueStr) || filterValueStr.includes(rowLocationStr)) {
              return true;
          }
          // Also check without common state suffixes/prefixes
          const cleanRow = rowLocationStr.replace(/\b(state|province)\b/g, '').trim();
          const cleanFilter = filterValueStr.replace(/\b(state|province)\b/g, '').trim();
          if (cleanRow === cleanFilter) return true;
          
          // Additional check for common state abbreviations
          const stateAbbreviations = {
              'alabama': 'al', 'alaska': 'ak', 'arizona': 'az', 'arkansas': 'ar', 'california': 'ca',
              'colorado': 'co', 'connecticut': 'ct', 'delaware': 'de', 'florida': 'fl', 'georgia': 'ga',
              'hawaii': 'hi', 'idaho': 'id', 'illinois': 'il', 'indiana': 'in', 'iowa': 'ia',
              'kansas': 'ks', 'kentucky': 'ky', 'louisiana': 'la', 'maine': 'me', 'maryland': 'md',
              'massachusetts': 'ma', 'michigan': 'mi', 'minnesota': 'mn', 'mississippi': 'ms',
              'missouri': 'mo', 'montana': 'mt', 'nebraska': 'ne', 'nevada': 'nv', 'new hampshire': 'nh',
              'new jersey': 'nj', 'new mexico': 'nm', 'new york': 'ny', 'north carolina': 'nc',
              'north dakota': 'nd', 'ohio': 'oh', 'oklahoma': 'ok', 'oregon': 'or', 'pennsylvania': 'pa',
              'rhode island': 'ri', 'south carolina': 'sc', 'south dakota': 'sd', 'tennessee': 'tn',
              'texas': 'tx', 'utah': 'ut', 'vermont': 'vt', 'virginia': 'va', 'washington': 'wa',
              'west virginia': 'wv', 'wisconsin': 'wi', 'wyoming': 'wy'
          };
          
          // Check if one is the abbreviation of the other
          if (stateAbbreviations[rowLocationStr] === filterValueStr || 
              stateAbbreviations[filterValueStr] === rowLocationStr) {
              return true;
          }
          
          // Try reverse lookup - if filter is full name, check if row is abbreviation
          for (const [fullName, abbrev] of Object.entries(stateAbbreviations)) {
              if (fullName === filterValueStr && abbrev === rowLocationStr) {
                  return true;
              }
              if (fullName === rowLocationStr && abbrev === filterValueStr) {
                  return true;
              }
          }
      }
      
      return false;
  });

  /* ------------------------------------------------------------------ */
  /* 🔄  SECOND-PASS MATCHING – handle tricky edge-cases                 */
  /* ------------------------------------------------------------------ */
  if (data.length === 0 && sheetName === 'State & Province') {
    // Occasionally the state value in the spreadsheet may contain extra whitespace,
    // punctuation (e.g. "Alabama (USA)"), or different capitalisation. If the first
    // pass returns nothing, make a more permissive pass that strips all
    // non-alphabetic characters before comparison.

    const sanitize = s => s.toString().replace(/[^a-z]/gi, '').toLowerCase();
    const target = sanitize(locationFilterValue);

    data = allData.filter(row => {
      const loc = row[locationColumnIndex];
      if (!loc) return false;
      return sanitize(loc) === target;
    });

    console.log(`Second-pass state matching found ${data.length} rows for "${locationFilterValue}" after sanitisation.`);
  }

  /* ------------------------------------------------------------------ */
  /* 🔄  THIRD-PASS MATCHING – very aggressive matching                  */
  /* ------------------------------------------------------------------ */
  if (data.length === 0 && sheetName === 'State & Province') {
    console.log(`Attempting third-pass matching for "${locationFilterValue}"`);
    
    // Collect all unique state values to understand what's actually in the data
    const uniqueStates = [...new Set(allData.map(row => row[locationColumnIndex]).filter(val => val && val.toString().trim()))];
    console.log(`Unique states in data (first 50):`, uniqueStates.slice(0, 50));
    
    // More aggressive matching - try contains, starts with, ends with
    const filterLower = locationFilterValue.toLowerCase().trim();
    
    data = allData.filter(row => {
      const loc = row[locationColumnIndex];
      if (!loc) return false;
      
      const locLower = loc.toString().toLowerCase().trim();
      
      // Try various matching strategies
      if (locLower.includes(filterLower) || filterLower.includes(locLower)) return true;
      if (locLower.startsWith(filterLower) || filterLower.startsWith(locLower)) return true;
      if (locLower.endsWith(filterLower) || filterLower.endsWith(locLower)) return true;
      
      // Try removing common words and punctuation
      const cleanLoc = locLower.replace(/[^\w\s]/g, '').replace(/\b(state|province|of|the)\b/g, '').trim();
      const cleanFilter = filterLower.replace(/[^\w\s]/g, '').replace(/\b(state|province|of|the)\b/g, '').trim();
      
      if (cleanLoc === cleanFilter) return true;
      if (cleanLoc.includes(cleanFilter) || cleanFilter.includes(cleanLoc)) return true;
      
      return false;
    });

    console.log(`Third-pass state matching found ${data.length} rows for "${locationFilterValue}"`);
    
    // If still no matches, log some examples of what we're comparing
    if (data.length === 0 && uniqueStates.length > 0) {
      console.log(`Still no matches found. Comparing "${filterLower}" against samples:`, uniqueStates.slice(0, 10));
    }
  }

  /* ------------------------------------------------------------------ */

  console.log(`\n=== FINAL FILTERING RESULTS ===`);
  console.log(`After all filtering passes: Found ${data.length} matching rows for location "${locationFilterValue}"`);
  
  if (data.length > 0) {
    console.log(`✅ SUCCESS: Found matching data!`);
    console.log(`Sample of matched rows (first 3):`);
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const row = data[i];
      console.log(`  Row ${i}: Service="${row[0]}", Location="${row[locationColumnIndex]}"`);
    }
  }
  
  if (data.length === 0) {
    // Enhanced debugging: Log more sample values and show unique values
    const sampleValues = allData.slice(0, 10).map(row => row[locationColumnIndex]).filter(val => val);
    const uniqueValues = [...new Set(allData.map(row => row[locationColumnIndex]).filter(val => val))].slice(0, 50);
    console.log(`No matching data found. Sample values in column ${locationColumnIndex}:`, sampleValues);
    console.log(`All unique values in column ${locationColumnIndex}:`, uniqueValues);
    console.log(`Looking for exact match: "${locationFilterValue.trim().toLowerCase()}"`);
    console.log(`Sheet name: "${sheetName}", Column index: ${locationColumnIndex}`);
    console.log(`Total rows in sheet: ${allData.length}`);
    
    // Additional debugging for headers
    console.log(`Headers:`, headers);
    console.log(`Header at column ${locationColumnIndex}:`, headers[locationColumnIndex]);
    
    // If this is a state request, also show some sample full rows to understand data structure
    if (sheetName === 'State & Province' && allData.length > 0) {
      console.log('Sample rows from State & Province sheet:');
      for (let i = 0; i < Math.min(3, allData.length); i++) {
        const row = allData[i];
        console.log(`Row ${i}: [${row.slice(0, 8).map(cell => `"${cell}"`).join(', ')}...]`);
      }
    }
    
    return { topServices: [], newServices: [] };
  }

  // Find the last two months that actually contain volume data
  let currentMonth = null;
  let previousMonth = null;
  for (const monthCol of monthColumns) {
      const hasVolume = data.some(row => {
          const vol = row[monthCol.index];
          // Check for non-empty, non-dash, and numeric values
          return vol !== '' && vol !== '-' && !isNaN(parseInt(vol, 10));
      });

      if (hasVolume) {
          if (!currentMonth) {
              currentMonth = monthCol;
          } else {
              previousMonth = monthCol;
              break; // We have the two most recent months with data
          }
      }
  }

  // Get indices of other required columns
  const serviceCol = headers.indexOf('Service');
  const keywordCol = headers.indexOf('Keyword');
  const compCol = headers.indexOf('Competition Index');
  const cpcCol = headers.indexOf('CPC');

  // Log column indices for debugging
  console.log(`Column indices: Service=${serviceCol}, Keyword=${keywordCol}, Competition=${compCol}, CPC=${cpcCol}`);
  console.log(`Current month: ${currentMonth ? currentMonth.header : 'None'}, Previous month: ${previousMonth ? previousMonth.header : 'None'}`);

  // Exit if essential columns are not found
  if (serviceCol === -1 || keywordCol === -1 || !currentMonth) {
    console.error('Essential columns (Service, Keyword) or current month data not found for location: ' + locationFilterValue);
    console.error(`Headers found: ${headers}`);
    return { topServices: [], newServices: [] };
  }

  // Warn if optional columns are missing but continue processing
  if (compCol === -1) console.warn('Competition Index column not found - will use 0 for competition values');
  if (cpcCol === -1) console.warn('CPC column not found - will use 0 for CPC values');

  const aggregatedData = {};

  data.forEach(row => {
    const serviceName = row[serviceCol];
    const keyword = row[keywordCol];
    if (!serviceName || !keyword) return; // Skip empty rows

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
    
    let keywordTrend = 0;
    if (currentVol > prevVol) keywordTrend = 1;
    else if (currentVol < prevVol) keywordTrend = -1;

    // Add keyword-level data
    service.keywords.push({
      name: keyword,
      volume: currentVol,
      trend: keywordTrend
    });
    
    // Update service-level aggregates - handle missing columns safely
    const competition = (compCol !== -1 && row[compCol] !== undefined) ? parseFloat(row[compCol]) : 0;
    const cpc = (cpcCol !== -1 && row[cpcCol] !== undefined) ? parseFloat(row[cpcCol]) : 0;
    
    if(!isNaN(competition) && competition > 0) service.totalCompetition += competition;
    if(!isNaN(cpc) && cpc > 0) service.totalCpc += cpc;
    service.totalCurrentVolume += currentVol;
    service.totalPreviousVolume += prevVol;
    service.keywordCount++;
  });

  const allServices = Object.values(aggregatedData);
  const allServicesTotalVolume = allServices.reduce((sum, s) => sum + s.totalCurrentVolume, 0);

  // Convert the aggregated object into the final array format
  const formattedServices = allServices.map(service => {
    const numKeywords = service.keywordCount;
    if (numKeywords === 0) return null;

    let serviceTrend = 0;
    if (service.totalCurrentVolume > service.totalPreviousVolume) serviceTrend = 1;
    else if (service.totalCurrentVolume < service.totalPreviousVolume) serviceTrend = -1;

    return {
      name: service.name,
      totalVolume: service.totalCurrentVolume,
      volumePercentage: allServicesTotalVolume > 0 ? ((service.totalCurrentVolume / allServicesTotalVolume) * 100).toFixed(1) : '0.0',
      avgCompetition: (service.totalCompetition / numKeywords).toFixed(2),
      avgCpc: (service.totalCpc / numKeywords).toFixed(2),
      trend: serviceTrend,
      keywords: service.keywords.sort((a, b) => b.volume - a.volume) // Sort keywords by volume
    };
  }).filter(s => s !== null && s.totalVolume > 0); // Filter out services with no volume

  const topServices = formattedServices
      .filter(s => NORMAL_SERVICES.includes(s.name))
      .sort((a, b) => b.totalVolume - a.totalVolume);

  const newServices = formattedServices
      .filter(s => NEW_SERVICES.includes(s.name))
      .sort((a, b) => b.totalVolume - a.totalVolume);
      
  console.log(`\n=== FUNCTION RETURN ===`);
  console.log(`Returning ${topServices.length} top services and ${newServices.length} new services`);
  console.log(`Top services:`, topServices.map(s => `${s.name} (${s.totalVolume})`));
  console.log(`New services:`, newServices.map(s => `${s.name} (${s.totalVolume})`));
      
  return { topServices, newServices };
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