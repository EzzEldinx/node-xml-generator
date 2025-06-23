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

// Read the XML template
const xmlTemplate = fs.readFileSync('4299.xml', 'utf8');

// Process each row
async function processRows() {
    for (const row of rows) {
        try {
            // Parse the XML template
            const parser = new xml2js.Parser();
            const builder = new xml2js.Builder();
            
            const result = await parser.parseStringPromise(xmlTemplate);
            
            // Update recordId
            result.nuds.control[0].recordId[0] = row.Index.toString();
            
            // Update other matching tags
            for (const [key, value] of Object.entries(row)) {
                if (key === 'Index') continue; // Skip Index as we already handled it
                
                // Function to recursively update XML nodes
                function updateNode(obj) {
                    for (const prop in obj) {
                        if (typeof obj[prop] === 'object' && obj[prop] !== null) {
                            if (prop === key && Array.isArray(obj[prop])) {
                                obj[prop][0] = value.toString();
                            }
                            updateNode(obj[prop]);
                        }
                    }
                }
                
                updateNode(result);
            }
            
            // Convert back to XML
            const xml = builder.buildObject(result);
            
            // Save the new XML file
            const outputPath = path.join(outputDir, `${row.Index}.xml`);
            fs.writeFileSync(outputPath, xml);
            
            console.log(`Created ${row.Index}.xml`);
        } catch (error) {
            console.error(`Error processing row ${row.Index}:`, error);
        }
    }
}

// Run the script
processRows().catch(console.error); 