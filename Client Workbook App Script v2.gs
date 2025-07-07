/**
 * Client Workbook App Script
 * Location: Client Workbook Google Sheet
 * 
 * This App Script powers the main dashboard functionality of the SEO tracking web app.
 * It processes and returns data for multiple sections including:
 * - Keywords tracking and analysis
 * - Backlinks monitoring
 * - GeoGrid Maps visualization
 * - Website health metrics
 * - Competitor analysis
 * 
 * The script works with multiple sheets in the workbook to aggregate and process:
 * - On-Page Insights
 * - Keywords data (Client and Competitors)
 * - Backlinks data (Client and Competitors)
 * - GeoGrid Maps data
 * - Census and demographic information
 * - Website technical metrics
 * 
 * For Service Cards visualization on the dashboard,
 * please reference Service Cards (Dashboard) App Script.
 */

/**
 * @license
 * Copyright 2024 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

function doGet(e) {
  try {
    // Read the workbookId from the URL parameter sent by the React app
    const workbookId = e.parameter.workbookId;
    if (!workbookId) {
      throw new Error("Workbook ID is missing.");
    }
    const spreadsheet = SpreadsheetApp.openById(workbookId);
    
    // Aggregate all data from the specified workbook
    const data = {
      dashboard: getDashboardData(spreadsheet),
      websiteStats: getWebsiteStats(spreadsheet),
      geogridData: getGeogridData(spreadsheet),
      backlinksSummary: getBacklinksSummary(spreadsheet),
      backlinksTable: getBacklinksTables(spreadsheet),
      keywordsSummary: getKeywordsSummary(spreadsheet),
      keywordsTable: getKeywordsTables(spreadsheet),
    };

    return ContentService
      .createTextOutput(JSON.stringify(data, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.message, stack: error.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- Main Data Aggregation Functions ---

function getDashboardData(spreadsheet) {
    const overviewData = getDashboardOverview(spreadsheet);
    const censusData = getCensusData(spreadsheet);
    const aiReport = getAiReport(spreadsheet);
    
    const allGeogridData = getGeogridData(spreadsheet);
    console.log('All geogrid keywords:', Object.keys(allGeogridData));
    
    const medspaNearMeData = allGeogridData['medspa near me'] || [];
    console.log('Medspa near me data length:', medspaNearMeData.length);
    
    const latestGeogrid = medspaNearMeData.length > 0 ? medspaNearMeData[0] : null;
    console.log('Latest geogrid data:', latestGeogrid);

    return {
        overviewData,
        censusData,
        aiReport,
        geogridForDashboard: latestGeogrid
    };
}

function getWebsiteStats(spreadsheet) {
    const onPageInsights = getOnPageInsights(spreadsheet);
    const websiteErrors = getWebsiteErrors(spreadsheet);
    const checksData = getChecksData(spreadsheet);
    const clientData = onPageInsights.find(site => site.isClient);

    const healthData = {
        pageScore: clientData ? clientData.pageScore : 0,
        siteSpeed: clientData ? clientData.siteSpeed : '0s',
        brokenLinks: clientData ? clientData.brokenLinks : 0,
        ssl: clientData ? clientData.ssl : false,
        errors: websiteErrors.errors,
        warnings: websiteErrors.warnings,
        checks: checksData,
        aiNotes: clientData ? clientData.aiNotes : null,
        lastUpdated: [
            clientData?.lastModifiedHeader,
            clientData?.lastModifiedSitemap,
            clientData?.lastModifiedMeta
        ].filter(Boolean), // Filter out any null/empty dates
        speedDetails: {
            'Time to Interactive': clientData?.time_to_interactive,
            'DOM Complete': clientData?.dom_complete,
            'Largest Contentful Paint': clientData?.largest_contentful_paint,
            'First Input Delay': clientData?.first_input_delay,
            'Connection Time': clientData?.connection_time,
            'Time to Secure Connection': clientData?.time_to_secure_connection,
            'Request Sent Time': clientData?.request_sent_time,
            'Waiting Time (TTFB)': clientData?.waiting_time,
            'Download Time': clientData?.download_time,
            'Full Page Load': clientData?.duration_time,
        }
    };

    return {
        healthData: healthData,
        competitorPerfData: onPageInsights.map(d => ({
            name: d.name,
            'Site Speed (s)': d.siteSpeed,
            'Referring Domains': d.referringDomains,
            'Total Backlinks': d.totalBacklinks,
            'Est. Monthly Traffic': d.estMonthlyTraffic,
            isClient: d.isClient
        })),
        technicalSeoData: onPageInsights.map(site => ({
            name: site.name,
            url: site.website,
            hasSchema: site.hasMicromarkup,
            h1: site.h1,
            title: site.title,
            description: site.meta,
            hasSSL: site.ssl
        }))
    };
}

// --- Helper Functions ---

function getDashboardOverview(spreadsheet) {
    const clientInfoSheet = spreadsheet.getSheetByName('Client & Competitor Info');
    const onPageSheet = spreadsheet.getSheetByName('On-Page Insights');
    const linksSheet = spreadsheet.getSheetByName('Links');
    
    if (!clientInfoSheet || !onPageSheet || !linksSheet) return [];

    const clientInfoValues = clientInfoSheet.getRange('A2:Q' + clientInfoSheet.getLastRow()).getValues();
    const onPageValues = onPageSheet.getRange('A2:U' + onPageSheet.getLastRow()).getValues();
    
    // Fetch both display values and rich text values to capture hyperlinks
    const linksRange = linksSheet.getRange('A2:U' + linksSheet.getLastRow());
    const linksValues = linksRange.getDisplayValues(); // Use getDisplayValues for consistency
    const linksRichTextValues = linksRange.getRichTextValues();

    return clientInfoValues.map((clientRow, index) => {
        const name = clientRow[1];  // Column B - Clinic Name
        if (!name) return null;

        // Use index-based matching since rows are parallel across sheets.
        const onPageRow = onPageValues[index] || [];
        const linksRow = linksValues[index] || [];
        const richTextLinksRow = linksRichTextValues[index] || [];
        
        // --- More Robust Site Speed Parsing ---
        let speedValue = 'N/A';
        const rawSpeed = onPageRow[10]; // Site Speed is in column K (index 10)
        if (rawSpeed) {
            const speedStr = String(rawSpeed);
            // Find the first sequence of digits and dots
            const matches = speedStr.match(/[\d.]+/); 
            if (matches && matches[0]) {
                const parsedSpeed = parseFloat(matches[0]);
                if (!isNaN(parsedSpeed)) {
                    speedValue = parsedSpeed.toFixed(2) + 's';
                }
            }
        }
        
        const metrics = {
            speed: speedValue,
            kwPos1: parseInt(onPageRow[19]) || 0, // Keywords #1 is col T (19)
            backlinks: parseInt(onPageRow[16]) || 0, // Backlinks is col Q (16)
        };
        
        const reviewScore = parseFloat(clientRow[14]) || 0;
        const reviewCount = parseInt(clientRow[15]) || 0;
        const address = clientRow[2] || '';
        const borough = clientRow[3] || '';
        const city = clientRow[4] || '';
        const website = clientRow[8] || '';
        const hoursRaw = clientRow[13] || 'N/A';
        const hours = formatOpeningHours(hoursRaw);

        // --- Ads Data Extraction with Hyperlinks ---
        const gAdsStatus = (linksRow[18] || 'Unknown').trim();      // Google Ads status (col S)
        const fbAdsStatus = (linksRow[17] || 'Unknown').trim();     // Facebook Ads status (col R)

        // Get link from the hyperlink formula in status cell (R/S), with a fallback to dedicated link columns (T/U)
        const gAdsLink = (richTextLinksRow[18] ? richTextLinksRow[18].getLinkUrl() : null) || (linksRow[20] || '').trim();
        const fbAdsLink = (richTextLinksRow[17] ? richTextLinksRow[17].getLinkUrl() : null) || (linksRow[19] || '').trim();

        return {
            id: index,
            clinicName: name,
            reviewScore: reviewScore.toFixed(1),
            reviewCount: reviewCount,
            address: address,
            borough: borough,
            city: city,
            website: website,
            speed: metrics.speed,
            kwPos1: metrics.kwPos1,
            backlinks: metrics.backlinks,
            hours: hours,
            gAds: gAdsStatus,
            gAdsLink: gAdsLink,
            fbAds: fbAdsStatus,
            fbAdsLink: fbAdsLink,
            isClient: clientRow[0] === 'Client'
        };
    }).filter(row => row);
}

function formatOpeningHours(hoursString) {
  // The frontend now handles the complex parsing and grouping.
  // This function just ensures a clean string is passed.
  if (!hoursString || typeof hoursString !== 'string') return 'N/A';
  return hoursString.trim();
}

function getCensusData(spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Census Info');
    if (!sheet || sheet.getLastRow() < 2) return {};
    
    const values = sheet.getRange('J2:Q2').getValues()[0];
    return {
        'Location': values[0], 'Population': values[1], 'MedSpas (10 Miles)': values[2],
        'Gender': values[3], 'Age Ranges': values[4], 'Ethnicity': values[5],
        'Languages': values[6], 'Median Income': values[7],
    };
}

function getAiReport(spreadsheet) {
    const sheet = spreadsheet.getSheetByName('GBP Insights');
    if (!sheet || sheet.getLastRow() < 2) return "";
    return sheet.getRange('AJ2').getValue();
}

function getOnPageInsights(spreadsheet) {
    const sheetName = 'On-Page Insights';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    // Get values starting from row 2 (where the actual data begins)
    const values = sheet.getRange('A2:DB' + sheet.getLastRow()).getValues(); 
    
    return values.map(row => {
        const isClient = row[0] === 'Client';
        
        // Get AI Notes from the correct column (CP - index 91)
        let aiNotes = '';
        if (isClient && row[91]) {
            try {
                // Get the raw value
                const rawNotes = row[91];
                // Convert to string and clean up
                aiNotes = typeof rawNotes === 'string' ? rawNotes.trim() : String(rawNotes).trim();
                // Log for debugging
                console.log('Raw AI Notes:', rawNotes);
                console.log('Processed AI Notes:', aiNotes);
            } catch (error) {
                console.error('Error processing AI Notes:', error);
            }
        }
        
        return {
            name: row[1], 
            website: row[8], 
            siteSpeed: row[10], 
            title: row[11],
            meta: row[12], 
            h1: row[13], 
            ssl: row[14] === true, 
            referringDomains: row[15],
            totalBacklinks: row[16], 
            estMonthlyTraffic: row[17], 
            kwPos1: row[19], 
            pageScore: row[22], 
            brokenLinks: row[24], 
            isClient: isClient,
            hasMicromarkup: row[42] === true,
            // Last Modified Dates
            lastModifiedHeader: row[29], // Column AD
            lastModifiedSitemap: row[30], // Column AE
            lastModifiedMeta: row[31],    // Column AF
            // AI Notes - Only include if this is the client row
            aiNotes: aiNotes,
            // New Speed Metrics
            time_to_interactive: row[68], // Column CQ
            dom_complete: row[69], // CR
            largest_contentful_paint: row[70], // CS
            first_input_delay: row[71], // CT
            connection_time: row[72], // CU
            time_to_secure_connection: row[73], // CV
            request_sent_time: row[74], // CW
            waiting_time: row[75], // CX
            download_time: row[76], // CY
            duration_time: row[77], // CZ
            fetch_start: row[78], // DA
            fetch_end: row[79] // DB
        };
    }).filter(row => row.name);
}

function getWebsiteErrors(spreadsheet) {
    const sheetName = 'Website Errors';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return { errors: [], warnings: [] };
    const values = sheet.getDataRange().getValues().slice(1);
    const errors = [], warnings = [];
    values.forEach((row, index) => {
        const issueType = row[2], message = row[7];
        if (issueType === 'error') errors.push({ id: `err-${index}`, description: message });
        else if (issueType === 'warning') warnings.push({ id: `warn-${index}`, description: message });
    });
    return { errors, warnings };
}

function getChecksData(spreadsheet) {
    const sheetName = 'On-Page Insights';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return { checksPassed: 0, totalChecks: 0, passedList: [], failedList: [] };
    const headerRow = sheet.getRange('A1:CP1').getValues()[0];
    const clientRow = sheet.getRange('A2:CP' + sheet.getLastRow()).getValues().find(r => r[0] === 'Client');
    if (!clientRow) return { checksPassed: 0, totalChecks: 0, passedList: [], failedList: [] };
    const PASS_WHEN_TRUE = ['is_https', 'has_html_doctype', 'canonical', 'meta_charset_consistency', 'seo_friendly_url', 'seo_friendly_url_characters_check', 'seo_friendly_url_dynamic_check', 'seo_friendly_url_keywords_check', 'seo_friendly_url_relative_length_check', 'has_meta_title'];
    const PASS_WHEN_FALSE = ['no_content_encoding', 'high_loading_time', 'is_redirect', 'is_4xx_code', 'is_5xx_code', 'is_broken', 'is_www', 'is_http', 'high_waiting_time', 'has_micromarkup', 'has_micromarkup_errors', 'no_doctype', 'no_encoding_meta_tag', 'no_h1_tag', 'https_to_http_links', 'size_greater_than_3mb', 'has_meta_refresh_redirect', 'has_render_blocking_resources', 'low_content_rate', 'high_content_rate', 'low_character_count', 'high_character_count', 'small_page_size', 'large_page_size', 'low_readability_rate', 'irrelevant_description', 'irrelevant_title', 'irrelevant_meta_keywords', 'title_too_long', 'title_too_short', 'deprecated_html_tags', 'duplicate_meta_tags', 'duplicate_title_tag', 'no_image_alt', 'no_image_title', 'no_description', 'no_title', 'no_favicon', 'flash', 'frame', 'lorem_ipsum'];
    const allChecks = [...PASS_WHEN_TRUE, ...PASS_WHEN_FALSE];
    let checksPassed = 0;
    const passedList = [], failedList = [];
    headerRow.forEach((header, index) => {
        if(allChecks.includes(header)){
            const value = clientRow[index];
            if ((PASS_WHEN_TRUE.includes(header) && value === true) || (PASS_WHEN_FALSE.includes(header) && value === false)) {
                checksPassed++; passedList.push(header);
            } else {
                failedList.push(header);
            }
        }
    });
    return { checksPassed: checksPassed, totalChecks: allChecks.length, passedList: passedList, failedList: failedList };
}

function getKeywordsSummary(spreadsheet) {
  const sheetName = 'Keywords Summary';
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const values = sheet.getRange('A2:N2').getValues()[0]; 
  return {
    website: values[2], totalKeywords: values[3], pos1: values[4], pos2_3: values[5],
    pos4_10: values[6], pos11_20: values[7], isNew: values[8], isUp: values[9],
    isDown: values[10], isLost: values[11], etv: values[12], estPaidCost: values[13],
  };
}

function getBacklinksSummary(spreadsheet) {
    const sheetName = 'Backlinks Summary';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return {};
    const values = sheet.getRange('A2:N2').getValues()[0];
    return {
        website: values[2], totalBacklinks: values[3], totalDofollow: values[4],
        totalNofollow: values[5], newLinks: values[6], lostLinks: values[7],
        avgSpamScore: values[8], topReferringDomains: values[9], avgRefferRank: values[10],
        titlesCaptured: values[11], backlinksChange: values[12], spamScoreChange: values[13]
    };
}

// Helper function to format website URL into a display name
function formatWebsiteToDisplayName(website) {
  if (!website) return '';
  
  try {
    // Remove protocol and www
    let displayName = website.replace(/^https?:\/\//i, '').replace(/^www\./i, '');
    
    // Remove trailing slash
    displayName = displayName.replace(/\/$/, '');
    
    // Remove everything after the first slash if it exists
    displayName = displayName.split('/')[0];
    
    return displayName;
  } catch (e) {
    console.error('Error formatting website:', e);
    return website; // Return original if formatting fails
  }
}

// Helper function to get website from sheet
function getWebsiteFromSheet(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return '';
  
  // Get all values from the first data row
  const firstDataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the column with header 'website' or containing website URL
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const websiteColumnIndex = headers.findIndex(header => 
    String(header).toLowerCase() === 'website' || 
    String(header).toLowerCase() === 'website url' ||
    String(header).toLowerCase() === 'url'
  );
  
  // If we found the website column, return its value, otherwise return empty string
  return websiteColumnIndex !== -1 ? firstDataRow[websiteColumnIndex] : '';
}

function getKeywordsTables(spreadsheet) {
  const sheetMapping = [
    { sheetName: "Client_kw" }, { sheetName: "Competitor 1_kw" },
    { sheetName: "Competitor 2_kw" }, { sheetName: "Competitor 3_kw" },
    { sheetName: "Competitor 4_kw" },
  ];
  const allTables = {};
  
  sheetMapping.forEach(mapping => {
    const sheet = spreadsheet.getSheetByName(mapping.sheetName);
    if (sheet && sheet.getLastRow() >= 2) {
      // Get the website from the first data row, column C (index 2)
      const website = sheet.getRange('C2').getValue();
      if (website) {
        // Format the website URL to be used as the key
        const displayName = formatWebsiteToDisplayName(website);
        if (displayName) {
          allTables[displayName] = getKeywordTableData(sheet);
        }
      }
    }
  });
  
  return allTables;
}

function getKeywordTableData(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues().slice(1);
  return values.map(row => ({
    keyword: row[3], rank_now: row[4], rank_prev: row[5], 
    is_new: String(row[6]).toLowerCase() === 'true', 
    is_up: String(row[7]).toLowerCase() === 'true',
    is_down: String(row[8]).toLowerCase() === 'true', 
    is_lost: String(row[18]).toLowerCase() === 'true',
    competition_lvl: row[9], cpc_usd: row[10], search_vol: row[11],
    etv: row[12], est_paid_cost: row[13], intent: row[14], keyword_difficulty: row[15],
    ranking_title: row[17], check_url: row[19],
  })).filter(row => row.keyword);
}

function getBacklinksTables(spreadsheet) {
  const sheetMapping = [
    { sheetName: "Client_bl" }, { sheetName: "Competitor 1_bl" },
    { sheetName: "Competitor 2_bl" }, { sheetName: "Competitor 3_bl" },
    { sheetName: "Competitor 4_bl" },
  ];
  const allTables = {};
  
  sheetMapping.forEach(mapping => {
    const sheet = spreadsheet.getSheetByName(mapping.sheetName);
    if (sheet && sheet.getLastRow() >= 2) {
      // Get the website from the first data row, column C (index 2)
      const website = sheet.getRange('C2').getValue();
      if (website) {
        // Format the website URL to be used as the key
        const displayName = formatWebsiteToDisplayName(website);
        if (displayName) {
          allTables[displayName] = getBacklinkTableData(sheet);
        }
      }
    }
  });
  
  return allTables;
}

function getBacklinkTableData(sheet) {
  if (sheet.getLastRow() < 2) return [];
  
  // Get all data rows
  const values = sheet.getDataRange().getValues().slice(1);
  
  // Map and filter the data
  return values.map(row => ({
    website: row[2] || '',      // Add website column
    backlinkUrl: row[3] || '',  // Add fallback empty string
    linkUrl: row[4] || '',      // Add fallback empty string
    isNew: String(row[5]).toLowerCase() === 'true',     // Handle string "TRUE"
    isLost: String(row[6]).toLowerCase() === 'true',    // Handle string "TRUE"
    spamScore: parseFloat(row[7]) || 0,  // Ensure number
    rank: parseFloat(row[8]) || 0,       // Ensure number
    following: String(row[9]).toLowerCase() === 'true', // Handle string "TRUE"
    title: row[10] || ''       // Add fallback empty string
  })).filter(row => row.backlinkUrl);  // Only include rows with backlink URLs
}

function getGeogridData(spreadsheet) {
    const sheetName = 'GeoGrid Maps';  // Case sensitive sheet name
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        console.log('No GeoGrid Maps sheet found or no data');
        return {};
    }

    const values = sheet.getDataRange().getValues().slice(1);
    console.log('Total rows in GeoGrid Maps:', values.length);
    const groupedByKeyword = {};

    values.forEach((row, index) => {
        const rawKeyword = row[7];  // Column H
        if (!rawKeyword || typeof rawKeyword !== 'string') {
            console.log(`Row ${index + 2}: No keyword found or invalid type`);
            return;
        }
        const keyword = rawKeyword.toLowerCase().trim();

        if (!groupedByKeyword[keyword]) {
            groupedByKeyword[keyword] = [];
        }
        
        const runDate = new Date(row[0]);
        const formattedDate = runDate.toLocaleString('default', { month: 'long', year: 'numeric', timeZone: 'UTC' });
        
        const competitors = [];
        // Starting from column K (index 10), each competitor now takes 5 columns
        for (let i = 0; i < 5; i++) {
            const baseIndex = 10 + (i * 5);     // K, P, U, Z, AE
            const name = row[baseIndex];
            const domain = row[baseIndex + 1];
            const rank = row[baseIndex + 2];
            const top5Raw = row[baseIndex + 3];
            const top10Raw = row[baseIndex + 4];
            
            // Ensure a competitor has a name and a valid rank to be included
            if (name && (rank !== undefined && rank !== null && rank !== '')) {
                // More robust parsing for top 5/10 totals
                const top5 = (typeof top5Raw === 'number') ? top5Raw : parseInt(top5Raw, 10);
                const top10 = (typeof top10Raw === 'number') ? top10Raw : parseInt(top10Raw, 10);

                competitors.push({ 
                    name: name,
                    domain: domain,
                    rank: parseFloat(rank),
                    top5Total: !isNaN(top5) ? top5 : 0,
                    top10Total: !isNaN(top10) ? top10 : 0,
                 });
            }
        }

        if (competitors.length === 0) {
            console.log(`Row ${index + 2}: No competitor data found for keyword: ${keyword}`);
        }

        groupedByKeyword[keyword].push({
            date: formattedDate,
            mapLink: row[9],  // Column J
            competitors: competitors,
        });
    });

    // Log the final structure
    console.log('Keywords found:', Object.keys(groupedByKeyword));
    for (const keyword in groupedByKeyword) {
        console.log(`${keyword}: ${groupedByKeyword[keyword].length} entries`);
    }

    for (const keyword in groupedByKeyword) {
        groupedByKeyword[keyword].sort((a, b) => new Date(b.date) - new Date(a.date));
    }

    return groupedByKeyword;
}

function formatDate(date) {
    if (!date) return '';
    try {
        const d = new Date(date);
        return d.toLocaleDateString('en-US', { 
            year: 'numeric', 
            month: 'short', 
            day: 'numeric' 
        });
    } catch (e) {
        return String(date);
    }
}