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