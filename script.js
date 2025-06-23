const XLSX = require('xlsx');
const xml2js = require('xml2js');
const fs = require('fs-extra');
const path = require('path');

// Create output directory if it doesn't exist
const outputDir = path.join(__dirname, 'output');
fs.ensureDirSync(outputDir);

// Read the Excel file
const workbook = XLSX.readFile('Sample_data.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(worksheet);

// Create a map of rows by Index for easier lookup
const rowsByIndex = {};
rows.forEach(row => {
    rowsByIndex[row.Index] = row;
});

// Read the XML template
const xmlTemplate = fs.readFileSync('4299.xml', 'utf8');

// Process each row
async function processRows() {
    for (const row of rows) {
        try {
            const index = row.Index;
            console.log(`Processing Index: ${index}`);
            
            // Parse the XML template
            const parser = new xml2js.Parser();
            const builder = new xml2js.Builder();
            
            const result = await parser.parseStringPromise(xmlTemplate);
            
            // 1. Update recordId with Index
            result.nuds.control[0].recordId[0] = index.toString();
            
            // 2. Update title with the full title from Excel
            console.log(`Setting title for Index ${index}:`, row.Title);
            result.nuds.descMeta[0].title[0] = row.Title;
            
            // 3. Update fromDate
            const startDate = row['http://nomisma.org/id/numismatic_start_date'];
            if (startDate) {
                // Extract the year from the date (assuming format like "BC 225")
                const year = startDate.match(/\d+/)[0];
                const isBC = startDate.includes('BC');
                const standardDate = isBC ? `-${year.padStart(4, '0')}` : year;
                
                // Update the fromDate element
                const fromDateElement = result.nuds.descMeta[0].typeDesc[0].dateRange[0].fromDate[0];
                fromDateElement._ = startDate; // Set the text content
                fromDateElement.$ = { standardDate: standardDate }; // Set the attribute
            }
            
            // 4. Update toDate
            const endDate = row['http://nomisma.org/id/numismatic_end_date'];
            if (endDate) {
                // Extract the year from the date (assuming format like "BC 205")
                const year = endDate.match(/\d+/)[0];
                const isBC = endDate.includes('BC');
                const standardDate = isBC ? `-${year.padStart(4, '0')}` : year;
                
                // Update the toDate element
                const toDateElement = result.nuds.descMeta[0].typeDesc[0].dateRange[0].toDate[0];
                toDateElement._ = endDate; // Set the text content
                toDateElement.$ = { standardDate: standardDate }; // Set the attribute
            }

            // 5. Update denomination
            const denomination = row['http://nomisma.org/id/denomination'];
            if (denomination) {
                // Update the denomination element
                const denominationElement = result.nuds.descMeta[0].typeDesc[0].denomination[0];
                denominationElement.$ = { 
                    'xlink:href': denomination.trim(),
                    'xlink:type': 'simple'
                };
                // Extract the denomination name from the URL (last part after the last slash)
                const denominationName = denomination.split('/').pop();
                denominationElement._ = denominationName;
            }
            
            // Convert back to XML
            const xml = builder.buildObject(result);
            
            // Save the new XML file
            const outputPath = path.join(outputDir, `${index}.xml`);
            fs.writeFileSync(outputPath, xml);
            
            console.log(`Created ${index}.xml with title: ${row.Title}`);
        } catch (error) {
            console.error(`Error processing row ${row.Index}:`, error);
        }
    }
}

// Run the script
processRows().catch(console.error); 