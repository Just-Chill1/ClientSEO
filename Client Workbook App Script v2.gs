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

// Toggle verbose logging globally (set to false in production)
var DEBUG = true;  // TEMPORARY: Enable for error report debugging

function log(){ if (DEBUG) console.log.apply(console, arguments); }

function doGet(e) {
  try {
    // Read the workbookId from the URL parameter sent by the React app
    const workbookId = e.parameter.workbookId;
    if (!workbookId) {
      throw new Error("Workbook ID is missing.");
    }
    const spreadsheet = SpreadsheetApp.openById(workbookId);
    
    // Optional: sections to fetch (comma-separated). When provided, we only compute those sections.
    // Example: sections=dashboard,webhooks or sections=keywordsSummary,keywordsTable
    const sectionsParam = (e.parameter.sections || '').trim();
    const requestedSections = sectionsParam ? sectionsParam.split(',').map(s => s.trim()).filter(Boolean) : null;
    const cacheBust = String(e.parameter.cacheBust || '').trim() !== '';
    const cache = CacheService.getDocumentCache();
    
    // Aggregate all data from the specified workbook
    // Helper to compute a single section with lightweight caching
    const computeSection = (name) => {
      const key = `cw:${workbookId}:${name}`;
      if (!cacheBust) {
        const cached = cache.get(key);
        if (cached) {
          try {
            return JSON.parse(cached);
          } catch (ignored) {}
        }
      }
      let value;
      switch (name) {
        case 'dashboard': value = getDashboardData(spreadsheet); break;
        case 'websiteStats': value = getWebsiteStats(spreadsheet); break;
        case 'geogridData': value = getGeogridData(spreadsheet); break;
        case 'gbpInsights': value = getGbpInsights(spreadsheet); break;
        case 'backlinksSummary': value = getBacklinksSummary(spreadsheet); break;
        case 'backlinksTable': value = getBacklinksTables(spreadsheet); break;
        case 'backlinksSummaryArchive': value = getBacklinksSummaryArchive(spreadsheet); break;
        case 'keywordsSummary': value = getKeywordsSummary(spreadsheet); break;
        case 'keywordsTable': value = getKeywordsTables(spreadsheet); break;
        case 'keywordsSummaryArchive': value = getKeywordsSummaryArchive(spreadsheet); break;
        case 'webhooks': value = getWebhookUrls(spreadsheet); break;
        default:
          value = null;
      }
      try { cache.put(key, JSON.stringify(value), 300); } catch (ignored) {}
      return value;
    };

    let data;
    if (requestedSections && requestedSections.length > 0) {
      // Only compute requested sections
      data = {};
      requestedSections.forEach(sec => {
        data[sec] = computeSection(sec);
      });
    } else {
      // Backwards-compatible: compute full payload
      data = {
        dashboard: computeSection('dashboard'),
        websiteStats: computeSection('websiteStats'),
        geogridData: computeSection('geogridData'),
        gbpInsights: computeSection('gbpInsights'),
        backlinksSummary: computeSection('backlinksSummary'),
        backlinksTable: computeSection('backlinksTable'),
        backlinksSummaryArchive: computeSection('backlinksSummaryArchive'),
        keywordsSummary: computeSection('keywordsSummary'),
        keywordsTable: computeSection('keywordsTable'),
        keywordsSummaryArchive: computeSection('keywordsSummaryArchive'),
        webhooks: computeSection('webhooks')
      };
    }

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
    
    // Try "medspa near me" first, fallback to "medical spa marketing & advertising" if not found
    let medspaNearMeData = allGeogridData['medspa near me'] || [];
    if (medspaNearMeData.length === 0) {
        medspaNearMeData = allGeogridData['medical spa marketing & advertising'] || [];
        console.log('Medspa near me not found, using medical spa marketing & advertising fallback');
    }
    console.log('Geogrid data length:', medspaNearMeData.length);
    
    const latestGeogrid = medspaNearMeData.length > 0 ? medspaNearMeData[0] : null;
    console.log('Latest geogrid data:', latestGeogrid);
    
    const aiSentimentData = getAiSentimentData(spreadsheet);

    return {
        overviewData,
        censusData,
        aiReport,
        geogridForDashboard: latestGeogrid,
        aiSentimentData
    };
}

function getWebsiteStats(spreadsheet) {
    const onPageInsights = getOnPageInsights(spreadsheet);
    const websiteErrors = getWebsiteErrors(spreadsheet);
    const checksData = getChecksData(spreadsheet);
    const clientData = onPageInsights.find(site => site.isClient);
    
    // NEW: Read crawl summary for score and broken links
    const crawlSummary = getWebsiteCrawlSummary(spreadsheet);
    // NEW: Aggregate checks across all client pages
    const checksAggregate = getWebsiteCrawlPagesAggregate(spreadsheet);
    // NEW: Read AI error report text from Website Crawl Summary (AK2)
    console.log('📊 [WEBSITE STATS] About to call getWebsiteErrorReport...');
    const errorReport = getWebsiteErrorReport(spreadsheet);
    console.log('📊 [WEBSITE STATS] Error report retrieved:', errorReport ? `${errorReport.length} characters` : 'Empty or null');
    


    const healthData = {
        // Prefer crawl summary average score; fallback to legacy clientData.pageScore
        pageScore: crawlSummary ? crawlSummary.averageScore : (clientData ? clientData.pageScore : 0),
        siteSpeed: clientData ? clientData.siteSpeed : '0s',
        brokenLinks: crawlSummary ? crawlSummary.brokenLinks : (clientData ? clientData.brokenLinks : 0),
        ssl: clientData ? clientData.ssl : false,
        errors: websiteErrors.errors,
        warnings: websiteErrors.warnings,
        checks: checksData,
        // NEW aggregate for whole-site technical checks
        checksAggregate: checksAggregate,
        // New AI error report text
        errorReport: errorReport,
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
            'Full Page Load': clientData?.duration_time
        }
    };

    // NEW: Get broken links details for tooltip
    const brokenLinksDetails = getBrokenLinksDetails(spreadsheet);
    
    // DEBUG: Log final healthData before return
    console.log('═══════════════════════════════════════════════════');
    console.log('📊 [WEBSITE STATS] FINAL DATA BEFORE RETURN');
    console.log('═══════════════════════════════════════════════════');
    console.log('🏥 healthData.errorReport:', healthData.errorReport);
    console.log('📏 errorReport length:', healthData.errorReport ? healthData.errorReport.length : 0);
    console.log('📊 errorReport type:', typeof healthData.errorReport);
    if (healthData.errorReport) {
        console.log('📄 First 200 chars:', healthData.errorReport.substring(0, 200));
    }
    console.log('🔑 All healthData keys:', Object.keys(healthData));
    console.log('═══════════════════════════════════════════════════');

    return {
        healthData: healthData,
        brokenLinksDetails: brokenLinksDetails,
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
            clinicName: site.name,
            website: site.website,
            hasSchema: site.hasMicromarkup,
            hasSchemaErrors: site.hasMicromarkupErrors,
            h1: site.h1,
            title: site.title,
            description: site.meta,
            hasSSL: site.ssl,
            isClient: site.isClient
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
    // Also read headers so we can find social columns robustly
    const linksHeaders = linksSheet.getRange(1, 1, 1, linksSheet.getLastColumn()).getValues()[0];
    const linksHeadersLower = linksHeaders.map(function(h){ return String(h).trim().toLowerCase(); });
    const fbColIdx = linksHeadersLower.indexOf('facebook');
    const igColIdx = linksHeadersLower.indexOf('instagram');
    const ytColIdx = linksHeadersLower.indexOf('youtube');

    return clientInfoValues.map((clientRow, index) => {
        const name = clientRow[1];  // Column B - Clinic Name
        if (!name) return null;

        // Use index-based matching since rows are parallel across sheets.
        const onPageRow = onPageValues[index] || [];
        const linksRow = linksValues[index] || [];
        const richTextLinksRow = linksRichTextValues[index] || [];
        
        // --- Site Speed Processing - Keep raw millisecond values ---
        let speedValue = 'N/A';
        const rawSpeed = onPageRow[10]; // Site Speed is in column K (index 10)
        if (rawSpeed) {
            const speedStr = String(rawSpeed);
            // Find the first sequence of digits and dots
            const matches = speedStr.match(/[\d.]+/); 
            if (matches && matches[0]) {
                const parsedSpeed = parseFloat(matches[0]);
                if (!isNaN(parsedSpeed)) {
                    // Keep the raw millisecond value for frontend conversion
                    speedValue = parsedSpeed;
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

        // --- Social Links (Facebook / Instagram / YouTube) ---
        function extractUrl(idx){
            if (idx < 0) return '';
            try {
                var url = (richTextLinksRow[idx] && richTextLinksRow[idx].getLinkUrl()) || (linksRow[idx] || '').toString().trim();
                // Basic validation: must look like a URL
                if (url && /https?:\/\//i.test(url)) return url;
            } catch (ignored) {}
            return '';
        }
        const facebook = extractUrl(fbColIdx);
        const instagram = extractUrl(igColIdx);
        const youtube = extractUrl(ytColIdx);

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
            facebook: facebook,
            instagram: instagram,
            youtube: youtube,
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

// NEW: Expose GBP Insights client mini section
function getGbpInsights(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('GBP Insights');
  if (!sheet || sheet.getLastRow() < 2) return {};

  // Read headers for robust column lookup
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headersLower = headers.map(function(h){ return String(h).trim().toLowerCase(); });
  const idxOf = function(name){ return headersLower.indexOf(String(name).trim().toLowerCase()); };

  // Column names from user sample
  const accountTypeIdx = idxOf('Account Type');
  const primaryCategoryIdx = idxOf('Primary Categorie') >= 0 ? idxOf('Primary Categorie') : idxOf('Primary Category');
  const secondaryCategoriesIdx = idxOf('Secondary Categories');
  const descriptionIdx = idxOf('Description');
  const reviewScoreIdx = idxOf('Review Score');
  const reviewCountIdx = idxOf('Review Count');
  const totalPhotosIdx = idxOf('Total Photos');
  const qaCountIdx = idxOf('Questions & Answers Count');

  const values = sheet.getDataRange().getValues().slice(1);

  // Find client row (Account Type == 'Client')
  const clientRow = values.find(function(row){
    if (accountTypeIdx < 0) return false;
    return String(row[accountTypeIdx]).trim().toLowerCase() === 'client';
  });
  if (!clientRow) return {};

  function getVal(idx){
    if (idx < 0) return '';
    var v = clientRow[idx];
    if (v === null || v === undefined) return '';
    return typeof v === 'string' ? v.trim() : String(v).trim();
  }

  return {
    primaryCategory: getVal(primaryCategoryIdx),
    secondaryCategories: getVal(secondaryCategoriesIdx),
    description: getVal(descriptionIdx),
    reviewScore: getVal(reviewScoreIdx),
    reviewCount: getVal(reviewCountIdx),
    totalPhotos: getVal(totalPhotosIdx),
    qaCount: getVal(qaCountIdx)
  };
}

function getAiSentimentData(spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Ai Sentiment');
    if (!sheet || sheet.getLastRow() < 2) {
        console.log('No Ai Sentiment sheet found or no data');
        return [];
    }
    
    const values = sheet.getRange('A2:AB' + sheet.getLastRow()).getValues();
    console.log('AI Sentiment data rows:', values.length);
    
    return values.map((row, index) => {
        const accountType = row[0];  // Column A - Account Type
        const clinicName = row[1];   // Column B - Clinic Name
        
        // Skip empty rows
        if (!clinicName) return null;
        
        // Parse sentiment values and calculate signed sentiment
        const chatGptSentiment = parseFloat(row[11]) || 0;      // Column L - ChatGPT Sentiment
        const chatGptStrength = parseFloat(row[12]) || 0;       // Column M - ChatGPT Strength
        const chatGptConfidence = parseFloat(row[13]) || 0;     // Column N - ChatGPT Confidence
        const chatGptVisibility = parseFloat(row[14]) || 0;     // Column O - ChatGPT Visibility
        const chatGptAppearance = parseFloat(row[15]) || 0;     // Column P - ChatGPT Appearance
        
        const geminiSentiment = parseFloat(row[17]) || 0;       // Column R - Gemini Sentiment
        const geminiStrength = parseFloat(row[18]) || 0;        // Column S - Gemini Strength
        const geminiConfidence = parseFloat(row[19]) || 0;      // Column T - Gemini Confidence
        const geminiVisibility = parseFloat(row[20]) || 0;      // Column U - Gemini Visibility
        const geminiAppearance = parseFloat(row[21]) || 0;      // Column V - Gemini Appearance
        
        // NEW: Perplexity data - user mentioned these headers: Perplexity Summary, Perplexity Sentiment, Perplexity Strength, Perplexity Confidence, Perplexity Visability, Perplexity Appearance
        // Handle string sentiment values: "Positive" = 1, "Negative" = -1, "Undefined" = 0
        const perplexitySentimentRaw = String(row[23] || '').trim().toLowerCase();
        const perplexitySentiment = perplexitySentimentRaw === 'positive' ? 1 : 
                                   perplexitySentimentRaw === 'negative' ? -1 : 0;
        const perplexityStrength = parseFloat(row[24]) || 0;    // Column Y - Perplexity Strength  
        const perplexityConfidence = parseFloat(row[25]) || 0;  // Column Z - Perplexity Confidence
        const perplexityVisibility = parseFloat(row[26]) || 0;  // Column AA - Perplexity Visibility
        const perplexityAppearance = parseFloat(row[27]) || 0;  // Column AB - Perplexity Appearance
        
        // Calculate signed sentiment (strength × sentimentSign)
        const chatGptSentimentSign = chatGptSentiment >= 0 ? 1 : -1;
        const geminiSentimentSign = geminiSentiment >= 0 ? 1 : -1;
        const perplexitySentimentSign = perplexitySentiment >= 0 ? 1 : -1;
        
        const chatGptSignedSentiment = chatGptStrength * chatGptSentimentSign;
        const geminiSignedSentiment = geminiStrength * geminiSentimentSign;
        const perplexitySignedSentiment = perplexityStrength * perplexitySentimentSign;
        
        // Debug logging for Perplexity data
        if (accountType === 'Client') {
            console.log('Perplexity Debug for Client:', {
                rawSentiment: row[23],
                parsedSentiment: perplexitySentiment,
                strength: perplexityStrength,
                confidence: perplexityConfidence,
                visibility: perplexityVisibility,
                appearance: perplexityAppearance,
                signedSentiment: perplexitySignedSentiment
            });
        }
        
        return {
            accountType: accountType,
            clinicName: clinicName,
            address: row[2] || '',           // Column C - Address
            borough: row[3] || '',           // Column D - Borough
            city: row[4] || '',              // Column E - City
            state: row[5] || '',             // Column F - State
            zip: row[6] || '',               // Column G - Zip
            country: row[7] || '',           // Column H - Country
            website: row[8] || '',           // Column I - Website
            gbpPlaceId: row[9] || '',        // Column J - GBP Place ID
            isClient: accountType === 'Client',
            
            // ChatGPT data
            chatGpt: {
                summary: row[10] || '',                    // Column K - ChatGPT Summary
                sentiment: chatGptSentiment,               // Column L - ChatGPT Sentiment
                strength: chatGptStrength,                 // Column M - ChatGPT Strength
                confidence: chatGptConfidence,             // Column N - ChatGPT Confidence
                visibility: chatGptVisibility,             // Column O - ChatGPT Visibility
                appearance: chatGptAppearance,             // Column P - ChatGPT Appearance
                signedSentiment: chatGptSignedSentiment    // Calculated: strength × sentimentSign
            },
            
            // Gemini data
            gemini: {
                summary: row[16] || '',                    // Column Q - Gemini Summary
                sentiment: geminiSentiment,                // Column R - Gemini Sentiment
                strength: geminiStrength,                  // Column S - Gemini Strength
                confidence: geminiConfidence,              // Column T - Gemini Confidence
                visibility: geminiVisibility,              // Column U - Gemini Visibility
                appearance: geminiAppearance,              // Column V - Gemini Appearance
                signedSentiment: geminiSignedSentiment     // Calculated: strength × sentimentSign
            },
            
            // Perplexity data
            perplexity: {
                summary: row[22] || '',                    // Column W - Perplexity Summary
                sentiment: perplexitySentiment,            // Column X - Perplexity Sentiment
                strength: perplexityStrength,              // Column Y - Perplexity Strength
                confidence: perplexityConfidence,          // Column Z - Perplexity Confidence
                visibility: perplexityVisibility,          // Column AA - Perplexity Visibility
                appearance: perplexityAppearance,          // Column AB - Perplexity Appearance
                signedSentiment: perplexitySignedSentiment // Calculated: strength × sentimentSign
            }
        };
    }).filter(row => row); // Remove null entries
}

function getWebhookUrls(spreadsheet) {
    const sheet = spreadsheet.getSheetByName('Config');
    if (!sheet || sheet.getLastRow() < 2) return {};
    
    // Get all webhook data from Config tab
    const values = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
    const webhooks = {};
    
    values.forEach(row => {
        const action = row[0]; // Column A - Action name
        const tabName = row[1]; // Column B - Tab Name
        const webhookUrl = row[2]; // Column C - Webhook URL
        
        // Include any row that has all three values (action, tabName, webhookUrl)
        if (action && tabName && webhookUrl) {
            console.log(`Found webhook: ${action} -> ${webhookUrl}`);
            webhooks[action] = webhookUrl;
        }
    });
    
    console.log('All webhooks found:', webhooks);
    return webhooks;
}

function getOnPageInsights(spreadsheet) {
    const sheetName = 'On-Page Insights';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    // Get values starting from row 2 (where the actual data begins)
    const values = sheet.getRange('A2:DB' + sheet.getLastRow()).getValues(); 
    
    return values.map(row => {
        const isClient = row[0] === 'Client';
        
        // Get AI Notes from the correct column (CP - index 93)
        let aiNotes = '';
        if (isClient && row[93]) {
            try {
                // Get the raw value
                const rawNotes = row[93];
                // Convert to string and clean up
                aiNotes = typeof rawNotes === 'string' ? rawNotes.trim() : String(rawNotes).trim();
            } catch (error) {
                console.error('App Script - Error processing AI Notes:', error);
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
            hasMicromarkup: row[51] === true, // Column AZ (has_micromarkup)
            hasMicromarkupErrors: row[52] === true, // Column BA (has_micromarkup_errors)
            // Last Modified Dates
            lastModifiedHeader: row[29], // Column AD
            lastModifiedSitemap: row[30], // Column AE
            lastModifiedMeta: row[31],    // Column AF
            // AI Notes - Only include if this is the client row
            aiNotes: aiNotes,
            // New Speed Metrics - Fixed column indices
            time_to_interactive: row[94], // Column CQ
            dom_complete: row[95], // CR
            largest_contentful_paint: row[96], // CS
            first_input_delay: row[97], // CT
            connection_time: row[98], // CU
            time_to_secure_connection: row[99], // CV
            request_sent_time: row[100], // CW
            waiting_time: row[101], // CX
            download_time: row[102], // CY
            duration_time: row[103], // CZ
            fetch_start: row[104], // DA
            fetch_end: row[105] // DB
        };
    }).filter(row => row.name);
}

function getWebsiteErrors(spreadsheet) {
    const sheetName = 'Website Crawl Errors';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return { errors: [], warnings: [] };
    
    // Read headers to locate columns by name for robustness
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const typeIdx = headers.findIndex(h => String(h).toLowerCase() === 'issue_type');
    const msgIdx = headers.findIndex(h => String(h).toLowerCase() === 'message');
    
    const values = sheet.getDataRange().getValues().slice(1);
    const errors = [], warnings = [];
    values.forEach((row, index) => {
        const issueType = String(typeIdx >= 0 ? row[typeIdx] : '').toLowerCase();
        const message = msgIdx >= 0 ? row[msgIdx] : '';
        if (issueType === 'error') errors.push({ id: `err-${index}`, description: message });
        else if (issueType === 'warning') warnings.push({ id: `warn-${index}`, description: message });
    });
    return { errors, warnings };
}

// Read AI error report text from Website Crawl Summary → column AK, row 2
function getWebsiteErrorReport(spreadsheet) {
    console.log('🔍 [ERROR REPORT] Starting getWebsiteErrorReport... VERSION 2.0');
    console.log('🔍 [ERROR REPORT] Spreadsheet ID:', spreadsheet.getId());
    console.log('🔍 [ERROR REPORT] Spreadsheet Name:', spreadsheet.getName());
    
    let sheet = spreadsheet.getSheetByName('Website Crawl Summary');
    console.log('🔍 [ERROR REPORT] Sheet lookup result:', sheet ? 'Found' : 'Not found');
    
    if (!sheet) {
        console.log('❌ [ERROR REPORT] Sheet "Website Crawl Summary" not found');
        return '';
    }
    
    console.log('✅ [ERROR REPORT] Sheet details:', {
        name: sheet.getName(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
    });
    
    if (sheet.getLastRow() < 2) {
        console.log('❌ [ERROR REPORT] Sheet has no data (less than 2 rows)');
        return '';
    }
    
    try {
        // DIRECT READ from AK2 (column 37, index 36)
        const akCell = sheet.getRange('AK2');
        let value = akCell.getDisplayValue();
        
        console.log('🔍 [ERROR REPORT] AK2 value:', value ? `Found ${value.length} characters` : 'Empty or null');
        console.log('🔍 [ERROR REPORT] AK2 preview:', value ? value.substring(0, 100) + '...' : 'N/A');
        
        // Fallback: try getValue() if displayValue is empty
        if (!value) {
            value = akCell.getValue();
            console.log('🔍 [ERROR REPORT] Fallback getValue():', value ? `Found ${String(value).length} characters` : 'Empty or null');
        }
        
        // Clean and return
        const result = typeof value === 'string' ? value.trim() : String(value || '').trim();
        console.log('✅ [ERROR REPORT] Final result:', result ? `${result.length} characters` : 'Empty string');
        
        return result;
    } catch (e) {
        console.log('❌ [ERROR REPORT] Error reading AK2:', e.message);
        return '';
    }
}

function getChecksData(spreadsheet) {
    const sheetName = 'On-Page Insights';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return { checksPassed: 0, totalChecks: 0, passedList: [], failedList: [] };
    const headerRow = sheet.getRange('A1:CP1').getValues()[0];
    const clientRow = sheet.getRange('A2:CP' + sheet.getLastRow()).getValues().find(r => r[0] === 'Client');
    if (!clientRow) return { checksPassed: 0, totalChecks: 0, passedList: [], failedList: [] };
    const PASS_WHEN_TRUE = ['is_https', 'has_html_doctype', 'canonical', 'meta_charset_consistency', 'seo_friendly_url', 'seo_friendly_url_characters_check', 'seo_friendly_url_dynamic_check', 'seo_friendly_url_keywords_check', 'seo_friendly_url_relative_length_check', 'has_meta_title', 'no_duplicate_meta_tags', 'no_duplicate_titles'];
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
  const values = sheet.getRange('A2:P2').getValues()[0]; // Extended to include column P
  return {
    website: values[2], totalKeywords: values[3], pos1: values[4], pos2_3: values[5],
    pos4_10: values[6], pos11_20: values[7], isNew: values[8], isUp: values[9],
    isDown: values[10], isLost: values[11], etv: values[12], estPaidCost: values[13],
    keywordsAiReport: values[15] || '' // Column P - Keywords AI Report
  };
}

function getBacklinksSummary(spreadsheet) {
    const sheetName = 'Backlinks Summary';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        console.log('Backlinks Summary Debug - Sheet not found or empty:', {
            sheetExists: !!sheet,
            lastRow: sheet ? sheet.getLastRow() : 'N/A'
        });
        return {};
    }
    
    // Check if there are multiple rows of data
    const lastRow = sheet.getLastRow();
    console.log('Backlinks Summary Debug - Sheet info:', {
        totalRows: lastRow,
        hasMultipleDataRows: lastRow > 2
    });
    
    // If there are multiple rows, let's see what's in them
    if (lastRow > 2) {
        const allData = sheet.getRange('A2:O' + lastRow).getValues();
        console.log('Backlinks Summary Debug - Multiple rows found:', {
            totalDataRows: allData.length,
            firstRow: allData[0],
            lastRow: allData[allData.length - 1]
        });
    }
    
    // Get the most recent row (last row) instead of just row 2
    // This handles cases where there might be multiple summary entries
    let values;
    if (lastRow > 2) {
        // If multiple rows exist, use the last row (most recent)
        values = sheet.getRange('A' + lastRow + ':O' + lastRow).getValues()[0];
        console.log('Using last row (' + lastRow + ') for most recent data');
    } else {
        // If only one data row, use row 2
        values = sheet.getRange('A2:O2').getValues()[0];
        console.log('Using row 2 for single data row');
    }
    
    // Enhanced debug logging to identify the exact issue
    console.log('Backlinks Summary Debug - Enhanced:', {
        sheetName: 'Backlinks Summary',
        rowRange: 'A2:O2',
        rawRowLength: values.length,
        website: values[2],
        totalBacklinks: values[3],
        totalDofollow: values[4], 
        totalNofollow: values[5],
        newLinks: values[6],
        lostLinks: values[7],
        avgSpamScore: values[8],
        backlinksAiReport: values[14],
        // Check data types
        dataTypes: {
            totalBacklinks: typeof values[3],
            totalDofollow: typeof values[4],
            totalNofollow: typeof values[5],
            newLinks: typeof values[6],
            lostLinks: typeof values[7],
            avgSpamScore: typeof values[8]
        },
        // Check all columns with headers
        allColumns: {
            'A (crawl_date)': values[0], 
            'B (account_type)': values[1], 
            'C (website)': values[2], 
            'D (Total Backlinks)': values[3], 
            'E (Total Dofollow)': values[4], 
            'F (Total Nofollow)': values[5], 
            'G (New Links)': values[6], 
            'H (Lost Links)': values[7], 
            'I (Avg Spam Score)': values[8], 
            'J (Top Referring Domains)': values[9], 
            'K (Avg Reffer Rank)': values[10], 
            'L (Titles Captured)': values[11], 
            'M (Backlinks Change)': values[12], 
            'N (Spam Score Change)': values[13],
            'O (Backlinks AI Report)': values[14]
        }
    });
    
    // Enhanced parsing with better type conversion and validation
    const parseIntSafe = (value) => {
        if (value === null || value === undefined || value === '') return 0;
        const parsed = parseInt(String(value).replace(/[^\d.-]/g, ''));
        return isNaN(parsed) ? 0 : parsed;
    };
    
    const parseFloatSafe = (value) => {
        if (value === null || value === undefined || value === '') return 0;
        // Handle different decimal separators and clean the string
        let cleanValue = String(value).replace(/[^\d.-]/g, '');
        const parsed = parseFloat(cleanValue);
        return isNaN(parsed) ? 0 : parsed;
    };
    
    const parsedData = {
        website: values[2] || '',
        totalBacklinks: parseIntSafe(values[3]),
        totalDofollow: parseIntSafe(values[4]),
        totalNofollow: parseIntSafe(values[5]),
        newLinks: parseIntSafe(values[6]),
        lostLinks: parseIntSafe(values[7]),
        avgSpamScore: parseFloatSafe(values[8]),
        topReferringDomains: parseIntSafe(values[9]),
        avgRefferRank: parseIntSafe(values[10]),
        titlesCaptured: parseIntSafe(values[11]),
        backlinksChange: parseIntSafe(values[12]),
        spamScoreChange: parseFloatSafe(values[13]),
        backlinksAiReport: values[14] || '' // Column O - Backlinks AI Report
    };
    
    // Additional debug logging to see parsed values
    console.log('Backlinks Summary Debug - Parsed values:', parsedData);
    
    // Special debugging for spam score issue
    if (parsedData.avgSpamScore !== parseFloatSafe(values[8])) {
        console.log('Spam Score Parsing Issue:', {
            originalValue: values[8],
            originalType: typeof values[8],
            parsedValue: parsedData.avgSpamScore,
            expectedValue: 1.6 // Based on user's data
        });
    }
    
    return parsedData;
}

function getBacklinksSummaryArchive(spreadsheet) {
    const sheetName = 'Backlinks Summary Archive';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        console.log('No Backlinks Summary Archive sheet found or no data');
        return [];
    }

    const values = sheet.getDataRange().getValues().slice(1); // Skip header row
    console.log('Total rows in Backlinks Summary Archive:', values.length);
    
    const archiveData = [];

    values.forEach((row, index) => {
        // Skip empty rows
        if (!row[0] && !row[2]) return;
        
        const rawDate = row[0]; // crawl_date - Column A
        const accountType = row[1]; // account_type - Column B
        const website = row[2]; // website - Column C
        
        // Only process client data
        if (accountType !== 'Client') return;
        
        let crawlDate;
        let formattedDate;
        
        try {
            // Try parsing the date as-is first
            crawlDate = new Date(rawDate);
            
            // If that fails and it looks like a US date format, try parsing differently
            if (isNaN(crawlDate.getTime()) && typeof rawDate === 'string') {
                // Try MM/DD/YYYY format parsing
                const parts = rawDate.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0]) - 1; // Convert to 0-based month
                    const day = parseInt(parts[1]);
                    const year = parseInt(parts[2]);
                    crawlDate = new Date(year, month, day);
                }
            }
            
            // If we still don't have a valid date, skip this row
            if (isNaN(crawlDate.getTime())) {
                console.log(`Row ${index + 2}: Invalid date format: ${rawDate}, skipping`);
                return;
            }
            
            // Format the date consistently
            formattedDate = crawlDate.toLocaleString('en-US', { 
                month: 'long', 
                year: 'numeric',
                timeZone: 'UTC'
            });
            
            console.log(`Row ${index + 2}: Raw date: ${rawDate}, Parsed: ${crawlDate}, Formatted: ${formattedDate}`);
            
        } catch (e) {
            console.log(`Row ${index + 2}: Error parsing date ${rawDate}:`, e.message);
            return;
        }
        
        archiveData.push({
            date: formattedDate,
            rawDate: rawDate,
            parsedDate: crawlDate.toISOString(),
            website: website,
            totalBacklinks: parseInt(row[3]) || 0,        // Total Backlinks - Column D
            totalDofollow: parseInt(row[4]) || 0,         // Total Dofollow - Column E
            totalNofollow: parseInt(row[5]) || 0,         // Total Nofollow - Column F
            newLinks: parseInt(row[6]) || 0,              // New Links (this run) - Column G
            lostLinks: parseInt(row[7]) || 0,             // Lost Links (this run) - Column H
            avgSpamScore: parseFloat(row[8]) || 0,        // Avg Spam Score - Column I
            topReferringDomains: parseInt(row[9]) || 0,   // Top Referring Domains - Column J
            avgRefferRank: parseInt(row[10]) || 0,        // Avg. Reffer Rank - Column K
            titlesCaptured: parseInt(row[11]) || 0,       // Titles Captured - Column L
            // Calculate changes from previous month (we'll do this after sorting)
            backlinksChange: 0,
            spamScoreChange: 0
        });
    });

    // Sort by parsed date (newest first)
    archiveData.sort((a, b) => {
        const dateA = new Date(a.parsedDate);
        const dateB = new Date(b.parsedDate);
        return dateB.getTime() - dateA.getTime();
    });

    // Calculate month-over-month changes
    for (let i = 0; i < archiveData.length; i++) {
        const current = archiveData[i];
        const previous = archiveData[i + 1]; // Previous month (older)
        
        if (previous) {
            current.backlinksChange = current.totalBacklinks - previous.totalBacklinks;
            current.spamScoreChange = current.avgSpamScore - previous.avgSpamScore;
        }
    }

    console.log('Processed backlinks archive data:', archiveData.length, 'entries');
    return archiveData;
}

function getKeywordsSummaryArchive(spreadsheet) {
    const sheetName = 'Keywords Summary Archive';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        console.log('No Keywords Summary Archive sheet found or no data');
        return [];
    }

    const values = sheet.getDataRange().getValues().slice(1); // Skip header row
    console.log('Total rows in Keywords Summary Archive:', values.length);
    
    const archiveData = [];

    values.forEach((row, index) => {
        // Skip empty rows
        if (!row[0] && !row[2]) return;
        
        const rawDate = row[0]; // crawl_date - Column A
        const accountType = row[1]; // account_type - Column B
        const website = row[2]; // website - Column C
        
        // Only process client data
        if (accountType !== 'Client') return;
        
        let crawlDate;
        let formattedDate;
        
        try {
            // Try parsing the date as-is first
            crawlDate = new Date(rawDate);
            
            // If that fails and it looks like a US date format, try parsing differently
            if (isNaN(crawlDate.getTime()) && typeof rawDate === 'string') {
                // Try MM/DD/YYYY format parsing
                const parts = rawDate.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0]) - 1; // Convert to 0-based month
                    const day = parseInt(parts[1]);
                    const year = parseInt(parts[2]);
                    crawlDate = new Date(year, month, day);
                }
            }
            
            // If we still don't have a valid date, skip this row
            if (isNaN(crawlDate.getTime())) {
                console.log(`Row ${index + 2}: Invalid date format: ${rawDate}, skipping`);
                return;
            }
            
            // Format the date consistently
            formattedDate = crawlDate.toLocaleString('en-US', { 
                month: 'long', 
                year: 'numeric',
                timeZone: 'UTC'
            });
            
            console.log(`Row ${index + 2}: Raw date: ${rawDate}, Parsed: ${crawlDate}, Formatted: ${formattedDate}`);
            
        } catch (e) {
            console.log(`Row ${index + 2}: Error parsing date ${rawDate}:`, e.message);
            return;
        }
        
        archiveData.push({
            date: formattedDate,
            rawDate: rawDate,
            parsedDate: crawlDate.toISOString(),
            website: website,
            totalKeywords: parseInt(row[3]) || 0,         // Total Keywords - Column D
            pos1: parseInt(row[4]) || 0,                  // Position 1 - Column E
            pos2_3: parseInt(row[5]) || 0,                // Position 2-3 - Column F
            pos4_10: parseInt(row[6]) || 0,               // Position 4-10 - Column G
            pos11_20: parseInt(row[7]) || 0,              // Position 11-20 - Column H
            isNew: parseInt(row[8]) || 0,                 // Is New - Column I
            isUp: parseInt(row[9]) || 0,                  // Is Up - Column J
            isDown: parseInt(row[10]) || 0,               // Is Down - Column K
            isLost: parseInt(row[11]) || 0,               // Is Lost - Column L
            etv: parseFloat(row[12]) || 0,                // Etv - Column M
            estPaidCost: parseFloat(row[13]) || 0,        // Estimated Paid Cost - Column N
            // Calculate changes from previous month (we'll do this after sorting)
            totalKeywordsChange: 0,
            pos1Change: 0,
            etvChange: 0,
            estPaidCostChange: 0
        });
    });

    // Sort by parsed date (newest first)
    archiveData.sort((a, b) => {
        const dateA = new Date(a.parsedDate);
        const dateB = new Date(b.parsedDate);
        return dateB.getTime() - dateA.getTime();
    });

    // Calculate month-over-month changes
    for (let i = 0; i < archiveData.length; i++) {
        const current = archiveData[i];
        const previous = archiveData[i + 1]; // Previous month (older)
        
        if (previous) {
            current.totalKeywordsChange = current.totalKeywords - previous.totalKeywords;
            current.pos1Change = current.pos1 - previous.pos1;
            current.etvChange = current.etv - previous.etv;
            current.estPaidCostChange = current.estPaidCost - previous.estPaidCost;
        }
    }

    console.log('Processed keywords archive data:', archiveData.length, 'entries');
    return archiveData;
}

// NEW: Read Website Crawl Summary for overall site metrics
function getWebsiteCrawlSummary(spreadsheet) {
    const sheetName = 'Website Crawl Summary';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return null;
    
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const idxOf = (name) => headers.findIndex(h => String(h).trim().toLowerCase() === name.trim().toLowerCase());
    const averageScoreIdx = idxOf('Average Score');
    const brokenLinksIdx = idxOf('Broken Links');
    
    const averageScore = averageScoreIdx >= 0 ? parseFloat(row[averageScoreIdx]) || 0 : 0;
    const brokenLinks = brokenLinksIdx >= 0 ? parseInt(row[brokenLinksIdx]) || 0 : 0;
    
    return { averageScore, brokenLinks };
}

// NEW: Aggregate technical checks across all pages from 'Website Crawl Pages'
function getWebsiteCrawlPagesAggregate(spreadsheet) {
	const sheetName = 'Website Crawl Pages';
	const sheet = spreadsheet.getSheetByName(sheetName);
	if (!sheet || sheet.getLastRow() < 2) {
		return null;
	}
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
	const headerIndex = {};
	headers.forEach((h, idx) => { headerIndex[h.toLowerCase()] = idx; });
	const values = sheet.getDataRange().getValues().slice(1);
	const totalPages = values.length;

	// Define pass logic similar to getChecksData, but page-wise aggregation
	const PASS_WHEN_TRUE = ['is_https','has_html_doctype','canonical','meta_charset_consistency','seo_friendly_url','seo_friendly_url_characters_check','seo_friendly_url_dynamic_check','seo_friendly_url_keywords_check','seo_friendly_url_relative_length_check','has_meta_title','no_duplicate_meta_tags','no_duplicate_titles'];
	const PASS_WHEN_FALSE = ['no_content_encoding','high_loading_time','is_redirect','is_4xx_code','is_5xx_code','is_broken','is_www','is_http','high_waiting_time','no_doctype','no_encoding_meta_tag','no_h1_tag','https_to_http_links','size_greater_than_3mb','has_meta_refresh_redirect','has_render_blocking_resources','low_content_rate','high_content_rate','low_character_count','high_character_count','small_page_size','large_page_size','low_readability_rate','irrelevant_description','irrelevant_title','irrelevant_meta_keywords','title_too_long','title_too_short','deprecated_html_tags','duplicate_meta_tags','duplicate_title_tag','no_image_alt','no_image_title','no_description','no_title','no_favicon','flash','frame','lorem_ipsum'];

	// Exclude microdata/schema from aggregate per user instruction
	const EXCLUDE_CHECKS = ['has_micromarkup','has_micromarkup_errors'];

	// Aliases to match UI check keys (left) to sheet columns (right)
	// Values can be a string or an array of candidate column names (first present is used)
	const ALIASES = {
		'has_favicon': 'no_favicon', // UI expects has_favicon; sheet provides no_favicon
		'duplicate_meta_tag': 'duplicate_meta_tags', // singular vs plural
		'no_duplicate_meta_tags': 'duplicate_meta_tags', // inverted virtual check
		'no_duplicate_titles': 'duplicate_title_tag', // inverted virtual check
		// Prefer positive columns; fall back to negative equivalents (inverted)
		'has_meta_title': ['has_meta_title','no_title'],
		'has_html_doctype': ['has_html_doctype','no_doctype'],
		// Robust render-blocking/frames aliases (DataForSEO variants)
		'has_render_blocking_resources': ['has_render_blocking_resources','render_blocking_resources','render_blocking_js','render_blocking_css'],
		'frame': ['frame','has_frame','frames','iframes']
	};

	// Build the final set of checks to consider
	const allCheckKeys = Array.from(new Set([...PASS_WHEN_TRUE, ...PASS_WHEN_FALSE, ...Object.keys(ALIASES)]))
		.filter(k => !EXCLUDE_CHECKS.includes(k));

	const perCheck = {};
	const isTruthy = (v) => v === true || String(v).toLowerCase() === 'true' || v === 1;
	const isFalsy = (v) => v === false || String(v).toLowerCase() === 'false' || v === 0 || v === '' || v === null || v === undefined;

	allCheckKeys.forEach((uiKey) => {
		const aliasValue = ALIASES[uiKey] || uiKey;
		// Determine which sheet column(s) to use
		const candidateKeys = Array.isArray(aliasValue) ? aliasValue.map(k => k.toLowerCase()) : [String(aliasValue).toLowerCase()];
		// Choose first available column; for arrays we may need to OR values
		const availableCols = candidateKeys
			.map(key => ({ key, idx: headerIndex[key] }))
			.filter(entry => entry.idx !== undefined);
		if (availableCols.length === 0) {
			return; // No matching columns present
		}
		let passed = 0;
		let total = 0;
		const chosenKey = availableCols[0].key; // primary column actually used
		const isAliasedToDifferentKey = chosenKey !== String(uiKey).toLowerCase();
		for (let i = 0; i < values.length; i++) {
			const row = values[i];
			// If multiple columns available for this check, treat cell as OR across them (any truthy means the condition exists)
			const cellTruthy = availableCols.some(({ idx }) => isTruthy(row[idx]));
			// Inversion for specific UI keys when they are mapped to negative columns
			if (
				uiKey === 'has_favicon' ||
				uiKey === 'no_duplicate_meta_tags' ||
				uiKey === 'no_duplicate_titles' ||
				((uiKey === 'has_meta_title' || uiKey === 'has_html_doctype') && isAliasedToDifferentKey)
			) {
				// Pass when negative column is FALSE
				if (!cellTruthy) passed++;
				total++;
				continue;
			}
			// Determine rule
			const ruleTrue = PASS_WHEN_TRUE.includes(uiKey) || PASS_WHEN_TRUE.includes(chosenKey);
			const ruleFalse = PASS_WHEN_FALSE.includes(uiKey) || PASS_WHEN_FALSE.includes(chosenKey);
			if (ruleTrue) {
				if (cellTruthy) passed++;
				total++;
			} else if (ruleFalse) {
				if (!cellTruthy) passed++;
				total++;
			} else {
				// If not explicitly listed, skip
			}
		}
		if (total > 0) {
			perCheck[uiKey] = { passed, total };
		}
	});

	const totalChecks = Object.values(perCheck).reduce((sum, c) => sum + c.total, 0);
	const totalPassed = Object.values(perCheck).reduce((sum, c) => sum + c.passed, 0);

	// Build a lightweight list of client page paths to use in prompts (limit to avoid huge payloads)
	const urlIdx = headerIndex['url'];
	const pagesList = [];
	if (urlIdx !== undefined) {
		const seen = {};
		for (let i = 0; i < values.length; i++) {
			const fullUrl = values[i][urlIdx];
			if (!fullUrl) continue;
			try {
				let u = String(fullUrl).trim();
				// Normalize and extract path only
				u = u.replace(/^https?:\/\//i, '');
				const slash = u.indexOf('/');
				let path = slash >= 0 ? u.substring(slash) : '/';
				// Ensure leading slash and collapse double slashes
				if (!path.startsWith('/')) path = '/' + path;
				path = path.replace(/\/+/g, '/');
				if (!seen[path]) {
					pagesList.push(path);
					seen[path] = true;
				}
				if (pagesList.length >= 120) break; // keep payload bounded
			} catch (ignored) {}
		}
	}

	return { totalPassed, totalChecks, perCheck, pagesList };
}

// NEW: Get broken links details from Website Crawl Pages
function getBrokenLinksDetails(spreadsheet) {
    const sheetName = 'Website Crawl Pages';
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        return [];
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
    const values = sheet.getDataRange().getValues().slice(1);
    
    // Find the columns we need
    const urlIdx = headers.indexOf('url');
    const isBrokenIdx = headers.indexOf('is_broken');
    const statusCodeIdx = headers.indexOf('status_code');
    
    if (urlIdx === -1 || isBrokenIdx === -1) {
        console.log('Required columns not found for broken links');
        return [];
    }
    
    const brokenLinks = [];
    
    values.forEach((row, index) => {
        const url = row[urlIdx];
        const isBroken = row[isBrokenIdx];
        const statusCode = statusCodeIdx >= 0 ? row[statusCodeIdx] : '';
        
        // Check if this page has broken links (is_broken = true)
        if (isBroken === true || String(isBroken).toLowerCase() === 'true') {
            brokenLinks.push({
                url: url || '',
                statusCode: statusCode || 'Unknown',
                page: url || `Page ${index + 1}`
            });
        }
    });
    
    console.log(`Found ${brokenLinks.length} broken links`);
    return brokenLinks.slice(0, 10); // Limit to first 10 to avoid huge tooltips
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
  
  console.log(`Keyword Table Debug - Sheet has ${values.length} rows of data`);
  
  const processedKeywords = values.map(row => ({
    keyword: row[3], rank_now: row[4], rank_prev: row[5], 
    is_new: String(row[6]).toLowerCase() === 'true', 
    is_up: String(row[7]).toLowerCase() === 'true',
    is_down: String(row[8]).toLowerCase() === 'true', 
    is_lost: String(row[18]).toLowerCase() === 'true',
    competition_lvl: row[9], cpc_usd: row[10], search_vol: row[11],
    etv: row[12], est_paid_cost: row[13], intent: row[14], keyword_difficulty: row[15],
    ranking_title: row[17], check_url: row[19],
  })).filter(row => row.keyword);
  
  console.log(`Keyword Table Debug - After filtering, returning ${processedKeywords.length} keywords`);
  
  return processedKeywords;
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

    // Get client info to identify which competitor is the client
    const clientInfoSheet = spreadsheet.getSheetByName('Client & Competitor Info');
    let clientName = '';
    if (clientInfoSheet && clientInfoSheet.getLastRow() >= 2) {
        const clientInfoValues = clientInfoSheet.getRange('A2:B' + clientInfoSheet.getLastRow()).getValues();
        const clientRow = clientInfoValues.find(row => row[0] === 'Client');
        if (clientRow) {
            clientName = clientRow[1]; // Column B - Clinic Name
        }
    }
    console.log('Client name from Client & Competitor Info:', clientName);

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
        
        // Enhanced date parsing to handle US date formats better
        const rawDate = row[0];
        let runDate;
        let formattedDate;
        
        try {
            // Try parsing the date as-is first
            runDate = new Date(rawDate);
            
            // If that fails and it looks like a US date format, try parsing differently
            if (isNaN(runDate.getTime()) && typeof rawDate === 'string') {
                // Try MM/DD/YYYY format parsing
                const parts = rawDate.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0]) - 1; // Convert to 0-based month
                    const day = parseInt(parts[1]);
                    const year = parseInt(parts[2]);
                    runDate = new Date(year, month, day);
                }
            }
            
            // If we still don't have a valid date, use current date as fallback
            if (isNaN(runDate.getTime())) {
                console.log(`Row ${index + 2}: Invalid date format: ${rawDate}, using current date`);
                runDate = new Date();
            }
            
            // Format the date consistently
            formattedDate = runDate.toLocaleString('en-US', { 
                month: 'long', 
                year: 'numeric',
                timeZone: 'UTC'
            });
            
            console.log(`Row ${index + 2}: Raw date: ${rawDate}, Parsed: ${runDate}, Formatted: ${formattedDate}`);
            
        } catch (e) {
            console.log(`Row ${index + 2}: Error parsing date ${rawDate}:`, e.message);
            runDate = new Date();
            formattedDate = runDate.toLocaleString('en-US', { 
                month: 'long', 
                year: 'numeric',
                timeZone: 'UTC'
            });
        }
        
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

                // Check if this competitor is the client
                const isClient = clientName && name === clientName;
                
                competitors.push({ 
                    name: name,
                    domain: domain,
                    rank: parseFloat(rank),
                    top5Total: !isNaN(top5) ? top5 : 0,
                    top10Total: !isNaN(top10) ? top10 : 0,
                    isClient: isClient
                 });
            }
        }

        if (competitors.length === 0) {
            console.log(`Row ${index + 2}: No competitor data found for keyword: ${keyword}`);
        } else {
            const clientCount = competitors.filter(c => c.isClient).length;
            const competitorCount = competitors.filter(c => !c.isClient).length;
            console.log(`Row ${index + 2}: Found ${competitors.length} total (${clientCount} client, ${competitorCount} competitors)`);
        }

        groupedByKeyword[keyword].push({
            date: formattedDate,
            rawDate: rawDate, // Keep original for debugging
            parsedDate: runDate.toISOString(), // ISO format for consistent parsing
            mapLink: row[9],  // Column J
            competitors: competitors,
        });
    });

    // Log the final structure
    console.log('Keywords found:', Object.keys(groupedByKeyword));
    for (const keyword in groupedByKeyword) {
        console.log(`${keyword}: ${groupedByKeyword[keyword].length} entries`);
    }

    // Sort by parsed date (newest first) to handle missing months properly
    for (const keyword in groupedByKeyword) {
        groupedByKeyword[keyword].sort((a, b) => {
            const dateA = new Date(a.parsedDate);
            const dateB = new Date(b.parsedDate);
            return dateB.getTime() - dateA.getTime();
        });
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