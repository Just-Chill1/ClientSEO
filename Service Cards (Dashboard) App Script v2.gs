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
    
    // Determine which sheet to use based on the location parameter
    let sheetName;
    let locationColumnIndex; // 2 for City, 3 for State, 4 for Country
    let locationFilterValue = location;

    if (location === 'USA') {
        sheetName = 'US (Whole)';
        locationColumnIndex = 4; // Country column
    } else if (location === 'Canada') {
        sheetName = 'Canada (Whole)';
        locationColumnIndex = 4; // Country column
    } else if (location.includes(',')) { // City, State format (e.g., "New York, NY")
        sheetName = 'City (US & CA)';
        locationColumnIndex = 2; // City column
    } else { 
        // Single location name - assume it's a State or Province
        // This handles cases like "Alabama", "Alaska", "Ontario", "British Columbia", etc.
        sheetName = 'State & Province';
        locationColumnIndex = 3; // State column
        console.log(`Processing state/province request for: ${location}`);
    }

    console.log(`Request details: location="${location}", sheetName="${sheetName}", columnIndex=${locationColumnIndex}, filterValue="${locationFilterValue}"`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).join(', ')}`);
    }

    const data = aggregateServiceData(sheet, locationColumnIndex, locationFilterValue);

    // Support for JSONP
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + JSON.stringify(data) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    const errorResponse = { error: error.message, stack: error.stack };
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

function aggregateServiceData(sheet, locationColumnIndex, locationFilterValue) {
  if (!sheet) return { topServices: [], newServices: [] };
  const allData = sheet.getDataRange().getValues();
  const headers = allData.shift(); // Get and remove header row

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

  // Filter rows for the selected location
  const data = allData.filter(row => {
      const rowLocation = row[locationColumnIndex];
      const matches = rowLocation && rowLocation.toString().trim().toLowerCase() === locationFilterValue.trim().toLowerCase();
      return matches;
  });

  console.log(`Filtered data: Found ${data.length} rows for location "${locationFilterValue}" in column ${locationColumnIndex}`);
  if (data.length === 0) {
    // Log some sample values to help debug
    const sampleValues = allData.slice(0, 5).map(row => row[locationColumnIndex]).filter(val => val);
    console.log(`No matching data found. Sample values in column ${locationColumnIndex}:`, sampleValues);
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

  // Exit if essential columns are not found
  if (serviceCol === -1 || keywordCol === -1 || !currentMonth) {
    console.error('Essential columns (Service, Keyword) or current month data not found for location: ' + locationFilterValue);
    return { topServices: [], newServices: [] };
  }

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
    
    // Update service-level aggregates
    const competition = parseFloat(row[compCol] || 0);
    const cpc = parseFloat(row[cpcCol] || 0);
    if(!isNaN(competition)) service.totalCompetition += competition;
    if(!isNaN(cpc)) service.totalCpc += cpc;
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