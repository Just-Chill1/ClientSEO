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
    const medspaNearMeData = allGeogridData['medspa near me'] || [];
    const latestGeogrid = medspaNearMeData.length > 0 ? medspaNearMeData[0] : null;

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
        siteSpeed: clientData ? clientData.siteSpeed : 0,
        brokenLinks: clientData ? clientData.brokenLinks : 0,
        ssl: clientData ? clientData.ssl : false,
        errors: websiteErrors.errors,
        warnings: websiteErrors.warnings,
        checks: checksData 
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
            hasTitle: !!site.title,
            hasDescription: !!site.meta,
            hasSSL: site.ssl,
            lastModified: site.lastModified
        }))
    };
}

// --- Helper Functions ---

function getDashboardOverview(spreadsheet) {
    const clientInfoSheet = spreadsheet.getSheetByName('Client & Competitor Info');
    const linksSheet = spreadsheet.getSheetByName('Links');
    const onPageSheet = spreadsheet.getSheetByName('On-Page Insights');
    
    if (!clientInfoSheet || !linksSheet) return [];

    const clientInfoValues = clientInfoSheet.getRange('A2:P' + clientInfoSheet.getLastRow()).getValues();
    const linksValues = linksSheet.getRange('A2:S' + linksSheet.getLastRow()).getValues();
    
    // Get On-Page Insights data
    let onPageData = {};
    if (onPageSheet && onPageSheet.getLastRow() > 1) {
        const onPageValues = onPageSheet.getRange('A2:CP' + onPageSheet.getLastRow()).getValues();
        onPageData = onPageValues.reduce((acc, row) => {
            const name = String(row[1] || '').trim();
            if (name) {
                acc[name] = {
                    speed: row[10] ? row[10].toFixed(4) + ' sec.' : 'N/A',
                    kwPos1: row[19] || 0,
                    backlinks: row[16] || 0
                };
            }
            return acc;
        }, {});
    }

    return clientInfoValues.map((clientRow, index) => {
        const linksRow = linksValues[index] || [];
        const accountType = clientRow[0];
        if (!accountType) return null;

        // Get all the basic info and ensure it's properly formatted
        const clinicName = String(clientRow[1] || '').trim();
        const address = String(clientRow[2] || '').trim();
        const borough = String(clientRow[3] || '').trim();
        const city = String(clientRow[4] || '').trim();
        const website = String(clientRow[8] || '').trim();
        const hours = String(clientRow[13] || '').trim();
        const reviewScore = parseFloat(clientRow[14]) || 0;
        const reviewCount = parseInt(clientRow[15]) || 0;

        // Only return if we have a clinic name
        if (!clinicName) return null;

        // Get matching on-page data
        const onPageMetrics = onPageData[clinicName] || {
            speed: 'N/A',
            kwPos1: 0,
            backlinks: 0
        };

        // Handle Google Ads and Facebook Ads status
        const gAdsStatus = linksRow[17] ? (String(linksRow[17]).toLowerCase() === 'yes' ? 'Multiple Ads Found' : 'No Ads Found') : 'Unknown';
        const fbAdsStatus = linksRow[16] ? (String(linksRow[16]).toLowerCase() === 'yes' ? 'Multiple Ads Found' : 'No Ads Found') : 'Unknown';

        // Create the final object with all data properly formatted
        return {
            id: index,
            clinicName: clinicName,
            address: address,
            borough: borough,
            city: city,
            website: website,
            reviewScore: reviewScore.toFixed(1),
            reviewCount: reviewCount,
            hours: hours || 'N/A',
            gAds: gAdsStatus,
            fbAds: fbAdsStatus,
            gAdsLink: String(linksRow[17]).toLowerCase() === 'yes' ? 'https://ads.google.com' : null,
            fbAdsLink: String(linksRow[16]).toLowerCase() === 'yes' ? 'https://facebook.com/ads' : null,
            isClient: accountType.toLowerCase() === 'client',
            speed: onPageMetrics.speed,
            kwPos1: onPageMetrics.kwPos1,
            backlinks: onPageMetrics.backlinks
        };
    }).filter(row => row && row.clinicName); // Only return rows that have a clinic name
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
    const values = sheet.getRange('A2:CP' + sheet.getLastRow()).getValues(); 
    return values.map(row => ({
        name: row[1], website: row[8], siteSpeed: row[10], title: row[11],
        meta: row[12], h1: row[13], ssl: row[14] === true, referringDomains: row[15],
        totalBacklinks: row[16], estMonthlyTraffic: row[17], kwPos1: row[19], pageScore: row[22], 
        brokenLinks: row[24], lastModified: row[27], isClient: row[0] === 'Client',
        hasMicromarkup: row[42] === true,
    })).filter(row => row.name);
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
  
  // Remove protocol and www
  let displayName = website.replace(/^https?:\/\//i, '').replace(/^www\./i, '');
  
  // Remove trailing slash
  displayName = displayName.replace(/\/$/, '');
  
  // Remove everything after the first slash if it exists
  displayName = displayName.split('/')[0];
  
  return displayName;
}

// Helper function to get website from sheet
function getWebsiteFromSheet(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return '';
  
  // Assuming website is in column C (index 2)
  const websiteCol = 2;
  const website = sheet.getRange(2, websiteCol + 1).getValue();
  return website;
}

function getKeywordsTables(spreadsheet) {
  const sheetMapping = [
    { sheetName: 'Client_kw' },
    { sheetName: 'Competitor 1_kw' },
    { sheetName: 'Competitor 2_kw' },
    { sheetName: 'Competitor 3_kw' },
    { sheetName: 'Competitor 4_kw' }
  ];
  
  const allTables = {};
  
  sheetMapping.forEach(mapping => {
    const sheet = spreadsheet.getSheetByName(mapping.sheetName);
    if (sheet) {
      const website = getWebsiteFromSheet(sheet);
      const displayName = formatWebsiteToDisplayName(website);
      if (displayName) {
        allTables[displayName] = getKeywordTableData(sheet);
      }
    }
  });
  
  return allTables;
}

function getKeywordTableData(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues().slice(1);
  return values.map(row => ({
    keyword: row[3], rank_now: row[4], rank_prev: row[5], is_new: row[6], is_up: row[7],
    is_down: row[8], competition_lvl: row[9], cpc_usd: row[10], search_vol: row[11],
    etv: row[12], est_paid_cost: row[13], intent: row[14], keyword_difficulty: row[15],
    ranking_title: row[17], is_lost: row[18], check_url: row[19],
  })).filter(row => row.keyword);
}

function getBacklinksTables(spreadsheet) {
  const sheetMapping = [
    { sheetName: 'Client_bl' },
    { sheetName: 'Competitor 1_bl' },
    { sheetName: 'Competitor 2_bl' },
    { sheetName: 'Competitor 3_bl' },
    { sheetName: 'Competitor 4_bl' }
  ];
  
  const allTables = {};
  
  sheetMapping.forEach(mapping => {
    const sheet = spreadsheet.getSheetByName(mapping.sheetName);
    if (sheet) {
      const website = getWebsiteFromSheet(sheet);
      const displayName = formatWebsiteToDisplayName(website);
      if (displayName) {
        allTables[displayName] = getBacklinkTableData(sheet);
      }
    }
  });
  
  return allTables;
}

function getBacklinkTableData(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues().slice(1);
  return values.map(row => ({
    backlinkUrl: row[3], linkUrl: row[4], isNew: row[5], isLost: row[6],
    spamScore: row[7], rank: row[8], following: row[9], title: row[10]
  })).filter(row => row.backlinkUrl);
}

function getGeogridData(spreadsheet) {
    const geogridSheet = spreadsheet.getSheetByName('GeoGrid Maps');
    if (!geogridSheet) return {};

    const data = geogridSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows;
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