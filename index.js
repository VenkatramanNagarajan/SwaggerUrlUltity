const fs = require('fs');
const csv = require('csv-parser');
const xlsx = require('xlsx');
const path = require('path');

// Input and Output File Paths
const excelFilePath = 'Newset.xlsx'; // Input Excel file with placeholders
const logCsvPath = 'QueryResults-0311b535-d6de-40a1-a010-70b472250752-000000000003.csv'; // Input log CSV file
const outputExcelPath = 'outputdata0311b535_casesen.xlsx'; // Output Excel file

// Read the Excel File and extract sanitized URLs from column index 7
function readExcel(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    return data.map(row => row[7]); // Extract column 7 (index 6)
}

// Read the CSV File and extract unsanitized URLs from column index 14
function readLogCsv(filePath) {
    return new Promise((resolve, reject) => {
        const urls = [];
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', row => {
                urls.push(Object.values(row)[14]); // Extract column 14 (index 13)
            })
            .on('end', () => resolve(urls))
            .on('error', error => reject(error));
    });
}

// Compare and match URLs
function compareUrls(sanitizedUrls, logUrls, csvData) {
    const results = [];

    sanitizedUrls.forEach((sanitizedUrl) => {
        const placeholderRegex = /{{\w+}}/g; // Match all placeholders (e.g., {{client}}, {{id}})
        const sanitizedRegex = sanitizedUrl.replace(placeholderRegex, '(.+?)'); // Create regex for matching actual URLs

        logUrls.forEach((logUrl, logIndex) => {
            const regex = new RegExp(`^${sanitizedRegex}$`, 'i'); // Ensure exact match
            const match = logUrl.match(regex);

            if (match) {
                const placeholders = sanitizedUrl.match(placeholderRegex); // Get all placeholders
                let sanitizedWithParams = sanitizedUrl;

                // Replace each placeholder with its corresponding value
                if (placeholders && placeholders.length > 0) {
                placeholders.forEach((placeholder, index) => {
                    const value = match[index + 1];
                    sanitizedWithParams = sanitizedWithParams.replace(placeholder, value);
                });
              }

                results.push({
                    OriginalURL: sanitizedUrl,
                    TargetURL: logUrl,
                    SanitizedURL: sanitizedWithParams,
                    CsvData: csvData[logIndex] // Include corresponding CSV data
                });
            }
        });
    });

    return results;
}

// Write output to Excel
function writeOutputToExcel(results, logCsvPath, outputPath) {
    const headers = ['Swagger_url', 'Target URL', 'Sanitized URL'];
    const data = results.map(result => [
        result.OriginalURL, 
        result.TargetURL, 
        result.SanitizedURL,
        ...Object.values(result.CsvData) // Include CSV data in the output
    ]);

    // Read the CSV file to get all columns
    const csvData = [];
    fs.createReadStream(logCsvPath)
        .pipe(csv())
        .on('data', row => {
            csvData.push(row);
        })
        .on('end', () => {
            // Add CSV columns to headers
            const csvHeaders = Object.keys(csvData[0]);
            const finalHeaders = [...headers, ...csvHeaders];

            const worksheet = xlsx.utils.aoa_to_sheet([finalHeaders, ...data]);
            const workbook = xlsx.utils.book_new();
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Results');
            xlsx.writeFile(workbook, outputPath);
            console.log(`Output written to ${outputPath}`);
        })
        .on('error', error => {
            console.error('Error reading CSV file:', error);
        });
}

// Main Execution
(async function () {
    try {
        console.log('Reading input files...');
        const sanitizedUrls = readExcel(excelFilePath);
        const logUrls = await readLogCsv(logCsvPath);

        // Read the CSV file to get all columns
        const csvData = [];
        fs.createReadStream(logCsvPath)
            .pipe(csv())
            .on('data', row => {
                csvData.push(row);
            })
            .on('end', async () => {
                console.log('Comparing URLs...');
                const results = compareUrls(sanitizedUrls, logUrls, csvData);

                console.log('Writing output...');
                writeOutputToExcel(results, logCsvPath, outputExcelPath);
            })
            .on('error', error => {
                console.error('Error reading CSV file:', error);
            });
    } catch (error) {
        console.error('Error:', error);
    }
})();