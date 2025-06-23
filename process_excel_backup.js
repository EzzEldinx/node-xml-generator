const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { XMLParser, XMLBuilder } = require('fast-xml-parser');

// Helper function to format date for standardDate attribute
function formatStandardDate(dateString) {
    if (!dateString || typeof dateString !== 'string') {
        return '';
    }
    
    // Remove any extra whitespace
    dateString = dateString.trim();
    
    // Check if it's a BC date (contains "BC")
    if (dateString.toUpperCase().includes('BC')) {
        // Extract the year number
        const yearMatch = dateString.match(/(\d+)/);
        if (yearMatch) {
            const year = parseInt(yearMatch[1]);
            // Format as negative 4-digit year (e.g., -0225 for BC 225)
            return `-${year.toString().padStart(4, '0')}`;
        }
    }
    
    // Check if it's an AD date (contains "AD" or just a number)
    if (dateString.toUpperCase().includes('AD') || /^\d+$/.test(dateString)) {
        // Extract the year number
        const yearMatch = dateString.match(/(\d+)/);
        if (yearMatch) {
            const year = parseInt(yearMatch[1]);
            // Format as 4-digit year (e.g., 0325 for AD 325)
            return year.toString().padStart(4, '0');
        }
    }
    
    // If no pattern matches, return the original string
    return dateString;
}

async function processExcelFile() {
    try {
        // Read the Excel file
        const workbook = xlsx.readFile('DB_SCA_V5.xlsx');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const range = xlsx.utils.decode_range(worksheet['!ref']);

        console.log('Processing Excel file...');

        // Read the template XML file
        const templatePath = 'script.xml';
        if (!fs.existsSync(templatePath)) {
            console.error(`Template file not found: ${templatePath}`);
            return;
        }

        const templateContent = fs.readFileSync(templatePath, 'utf-8');
        const parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: '@_'
        });

        // Add a report object to collect results
        let report = [];

        const outputDir = path.join(__dirname, 'output');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
            console.log('Created output directory:', outputDir);
        } else {
            console.log('Output directory exists:', outputDir);
        }

        // Process each row
        for (let R = range.s.r; R <= range.e.r; R++) {
            let mainRef = null;
            let cealexRef = null;
            let identifiers = null;
            let indexIdentifier = null;
            let coinNumberIdentifier = null;
            let inventoryNumberIdentifier = null;
            let materialLabel = null;
            let mintLabel = null;
            let referenceTitle = null;
            let departmentLabel = null;
          
            // باقي الكود بتاع اللوب هنا...
          

            // Get the Index value from column A (assuming Index is in column A)
            const indexCell = worksheet[xlsx.utils.encode_cell({r: R, c: 0})];
            if (!indexCell) continue;
            
            const index = indexCell.v;
            
            // Skip header rows - only process if index is a number
            if (typeof index !== 'number' || isNaN(index)) {
                console.log(`Skipping row with non-numeric index: ${index}`);
                continue;
            }
            
            const xmlFilePath = path.join(outputDir, `${index}.xml`);
            
            // Parse the template XML FIRST
            const result = parser.parse(templateContent);
            
            // Get values from Excel columns
            const titleCell = worksheet[xlsx.utils.encode_cell({r: R, c: 1})]; // Column B
            const fromDateCell = worksheet[xlsx.utils.encode_cell({r: R, c: 2})]; // Column C
            const toDateCell = worksheet[xlsx.utils.encode_cell({r: R, c: 3})]; // Column D
            const denominationUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 4})]; // Column E
            const denominationNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 5})]; // Column F
            const typeSeriesCell = worksheet[xlsx.utils.encode_cell({r: R, c: 9})]; // Column J
            const materialUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 6})]; // Column G
            const mintUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 10})]; // Column K
            let obverseDescription = worksheet[xlsx.utils.encode_cell({r: R, c: 11})]; // Column L
            const obverseLegendCell = worksheet[xlsx.utils.encode_cell({r: R, c: 12})]; // Column M
            let reverseDescription = worksheet[xlsx.utils.encode_cell({r: R, c: 13})]; // Column N
            const reverseLegendCell = worksheet[xlsx.utils.encode_cell({r: R, c: 14})]; // Column O
            const axisCell = worksheet[xlsx.utils.encode_cell({r: R, c: 16})]; // Column Q
            const weightCell = worksheet[xlsx.utils.encode_cell({r: R, c: 17})]; // Column R
            const diameterCell = worksheet[xlsx.utils.encode_cell({r: R, c: 18})]; // Column S
            const countermarkCell = worksheet[xlsx.utils.encode_cell({r: R, c: 19})]; // Column T
            const referenceInfoCell = worksheet[xlsx.utils.encode_cell({r: R, c: 20})]; // Column U
            const referenceUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 22})]; // Column W
            const stratigraphicUnitCell = worksheet[xlsx.utils.encode_cell({r: R, c: 23})]; // Column X
            const fallsWithinUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 24})]; // Column Y
            const coinNumberCell = worksheet[xlsx.utils.encode_cell({r: R, c: 28})]; // Column AC
            const inventoryNumberCell = worksheet[xlsx.utils.encode_cell({r: R, c: 29})]; // Column AD
            const departmentUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 30})]; // Column AE
            const repositoryCell = worksheet[xlsx.utils.encode_cell({r: R, c: 31})]; // Column AF
            const fileLocationCell = worksheet[xlsx.utils.encode_cell({r: R, c: 32})]; // Column AG
            const fallsWithinNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 33})]; // Column AH

            const title = titleCell ? titleCell.v : '';
            const fromDate = fromDateCell ? fromDateCell.v : '';
            const toDate = toDateCell ? toDateCell.v : '';
            const denominationUrl = denominationUrlCell ? denominationUrlCell.v : '';
            const denominationName = denominationNameCell ? denominationNameCell.v : '';
            const typeSeries = typeSeriesCell ? typeSeriesCell.v : '';
            const materialUrl = materialUrlCell ? materialUrlCell.v : '';
            const mintUrl = mintUrlCell ? mintUrlCell.v : '';
            obverseDescription = obverseDescription ? obverseDescription.v : '';
            const obverseLegendText = obverseLegendCell ? obverseLegendCell.v : '';
            reverseDescription = reverseDescription ? reverseDescription.v : '';
            const reverseLegendText = reverseLegendCell ? reverseLegendCell.v : '';
            const axisValue = axisCell ? axisCell.v : '';
            const weightValue = weightCell ? weightCell.v : '';
            const diameterValue = diameterCell ? diameterCell.v : '';
            const countermarkValue = countermarkCell ? countermarkCell.v : '';
            const referenceInfo = referenceInfoCell ? referenceInfoCell.v : '';
            const referenceUrl = referenceUrlCell ? referenceUrlCell.v : '';
            const stratigraphicUnit = stratigraphicUnitCell ? stratigraphicUnitCell.v : '';
            const fallsWithinUrl = fallsWithinUrlCell ? fallsWithinUrlCell.v : '';
            const coinNumber = coinNumberCell ? coinNumberCell.v : '';
            const inventoryNumber = inventoryNumberCell ? inventoryNumberCell.v : '';
            const departmentUrl = departmentUrlCell ? departmentUrlCell.v : '';
            const repository = repositoryCell ? repositoryCell.v : '';
            const fileLocation = fileLocationCell ? fileLocationCell.v : '';
            const fallsWithinName = fallsWithinNameCell ? fallsWithinNameCell.v : '';

            // Debug: Show what values are being read
            console.log(`Debug - Index ${index}: Column Q value = "${axisValue}" (type: ${typeof axisValue})`);
            console.log(`Debug - Index ${index}: Column B (Title) = "${title}"`);
            console.log(`Debug - Index ${index}: Column C (FromDate) = "${fromDate}"`);
            console.log(`Debug - Index ${index}: Column D (ToDate) = "${toDate}"`);
            console.log(`Debug - Index ${index}: Column E (Denomination URL) = "${denominationUrl}"`);
            console.log(`Debug - Index ${index}: Column F (Denomination Name) = "${denominationName}"`);
            console.log(`Debug - Index ${index}: Column J (TypeSeries) = "${typeSeries}"`);
            console.log(`Debug - Index ${index}: Column G (Material URL) = "${materialUrl}"`);
            console.log(`Debug - Index ${index}: Column K (Mint URL) = "${mintUrl}"`);
            console.log(`Debug - Index ${index}: Column L (Obverse Description) = "${obverseDescription}"`);
            console.log(`Debug - Index ${index}: Column M (Obverse Legend) = "${obverseLegendText}"`);
            console.log(`Debug - Index ${index}: Column N (Reverse Description) = "${reverseDescription}"`);
            console.log(`Debug - Index ${index}: Column O (Reverse Legend) = "${reverseLegendText}"`);
            console.log(`Debug - Index ${index}: Column R (Weight) = "${weightValue}"`);
            console.log(`Debug - Index ${index}: Column S (Diameter) = "${diameterValue}"`);
            console.log(`Debug - Index ${index}: Column T (Countermark) = "${countermarkValue}"`);
            console.log(`Debug - Index ${index}: Column U (Reference Info) = "${referenceInfo}"`);
            console.log(`Debug - Index ${index}: Column W (Reference URL) = "${referenceUrl}"`);
            console.log(`Debug - Index ${index}: Column X (Stratigraphic Unit) = "${stratigraphicUnit}"`);
            console.log(`Debug - Index ${index}: Column Y (Falls Within URL) = "${fallsWithinUrl}"`);
            console.log(`Debug - Index ${index}: Column AC (Coin Number) = "${coinNumber}"`);
            console.log(`Debug - Index ${index}: Column AD (Inventory Number) = "${inventoryNumber}"`);
            console.log(`Debug - Index ${index}: Column AE (Department URL) = "${departmentUrl}"`);
            console.log(`Debug - Index ${index}: Column AF (Repository) = "${repository}"`);
            console.log(`Debug - Index ${index}: Column AG (File Location) = "${fileLocation}"`);
            console.log(`Debug - Index ${index}: Column AH (Falls Within Name) = "${fallsWithinName}"`);
            
            // Check what's in column J specifically
            const columnJCell = worksheet[xlsx.utils.encode_cell({r: R, c: 9})];
            const columnJValue = columnJCell ? columnJCell.v : '';
            console.log(`Debug - Index ${index}: Column J raw value = "${columnJValue}"`);

            // Get date value from column AA (column 26, zero-based)
            const dateCell = worksheet[xlsx.utils.encode_cell({r: R, c: 26})]; // Column AA
            const rawDate = dateCell ? dateCell.v : '';
            let isoDate = '';
            let readableDate = '';
            if (rawDate) {
                // Try to parse and format the date
                const dateObj = new Date(rawDate);
                if (!isNaN(dateObj)) {
                    // ISO 8601 format
                    isoDate = dateObj.toISOString().slice(0, 10);
                    // Readable format (e.g., December 02, 2002)
                    const options = { year: 'numeric', month: 'long', day: '2-digit' };
                    readableDate = dateObj.toLocaleDateString('en-US', options);
                } else {
                    // If parsing fails, use the raw value as fallback
                    isoDate = rawDate;
                    readableDate = rawDate;
                }
            }
            // Insert the date tag and comment if a date is present
            if (isoDate && readableDate) {
                // Find the correct place in the XML structure (findspotDesc > findspot)
                if (result.nuds.descMeta.findspotDesc && result.nuds.descMeta.findspotDesc.findspot) {
                    let findspot = result.nuds.descMeta.findspotDesc.findspot;
                    // Insert the comment and date tag
                    findspot.date = {
                        '@_standardDate': isoDate,
                        '#text': readableDate
                    };
                    // Add a comment above the date tag (not all XML builders support comments, but we can try)
                    findspot['#comment'] = `ISO 8601: ${isoDate}`;
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            mainRef['#text'] = '';
                            console.log(`Updated reference URL for Index ${index}: ${referenceUrl}`);
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.length > 0) {
                                        const title = h4.text().trim();
                                        mainRef['#text'] = title;
                                        console.log(`Special case - Updated reference for Index ${index}: "${title}" with URL: ${referenceUrl}`);
                                    } else {
                                        mainRef['#text'] = '';
                                        console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (h4 not found)`);
                                    }
                                } catch (error) {
                                    mainRef['#text'] = '';
                                    console.log(`Special case - Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed: ${error.message})`);
                                }
                            } else {
                                // Regular logic for all other indices
                                referenceTitle = await getReferenceTitle(referenceUrl.toString().trim());
                                if (referenceTitle) {
                                    mainRef['#text'] = referenceTitle;
                                    console.log(`Updated reference for Index ${index}: "${referenceTitle}" with URL: ${referenceUrl}`);
                                } else {
                                    // Clear the text content if fetch fails
                                    mainRef['#text'] = '';
                                    console.log(`Updated reference URL for Index ${index}: ${referenceUrl} (fetch failed)`);
                                }
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference for Index ${index}:`, error.message);
                }
            } else {
                // Remove reference tag if URL is empty
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        // Remove the first reference element (main reference)
                        refDesc.reference.splice(0, 1);
                        console.log(`Removed reference for Index ${index} (empty URL in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing reference for Index ${index}:`, error.message);
                }
            }

            console.log(`[${index}] Before parsing template XML`);
            console.log(`[${index}] After parsing template XML`);

            // Update recordId
            result.nuds.control.recordId = index.toString();

            // Update title if provided
            if (title) {
                result.nuds.descMeta.title = title;
                console.log(`Updated title for Index ${index}: "${title}"`);
            }

            // Update date range if provided
            if (fromDate) {
                result.nuds.descMeta.typeDesc.dateRange.fromDate['#text'] = fromDate;
                result.nuds.descMeta.typeDesc.dateRange.fromDate['@_standardDate'] = formatStandardDate(fromDate);
                console.log(`Updated fromDate for Index ${index}: "${fromDate}"`);
            }

            if (toDate) {
                result.nuds.descMeta.typeDesc.dateRange.toDate['#text'] = toDate;
                result.nuds.descMeta.typeDesc.dateRange.toDate['@_standardDate'] = formatStandardDate(toDate);
                console.log(`Updated toDate for Index ${index}: "${toDate}"`);
            }

            // Update denomination if provided
            if (denominationUrl || denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denominationUrl) {
                    denomination['@_xlink:href'] = denominationUrl;
                }
                if (denominationName) {
                    denomination['#text'] = denominationName;
                }
                console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
                // Create typeSeries tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.typeSeries) {
                    result.nuds.descMeta.typeDesc.typeSeries = {
                        '#text': typeSeries,
                        '@_xml:lang': 'en'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.typeSeries['#text'] = typeSeries;
                }
                console.log(`Updated typeSeries for Index ${index}: "${typeSeries}"`);
            } else {
                // Remove the typeSeries tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('typeSeries')) {
                    delete result.nuds.descMeta.typeDesc.typeSeries;
                    console.log(`Removed typeSeries for Index ${index} (empty in DB)`);
                }
            }

            // Update material if URL exists
            if (materialUrl && materialUrl.toString().trim() !== '') {
                // Create material tag if it doesn't exist
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple'
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                }
                console.log(`Updated material URL for Index ${index}: ${materialUrl}`);
            } else {
                // Remove the material tag if the cell is empty
                if (result.nuds.descMeta.typeDesc.hasOwnProperty('material')) {
                    delete result.nuds.descMeta.typeDesc.material;
                    console.log(`Removed material for Index ${index} (empty in DB)`);
                }
            }

            // Update mint information if URL exists
            if (mintUrl) {
                // Update the geogname element with mint role
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        console.log(`Updated mint URL for Index ${index}: ${mintUrl}`);
                    }
                }
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                obverseDescription = obverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        obverseType.description['#text'] = obverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        obverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    } else {
                        console.warn(`No obverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else if (obverse && obverse.hasOwnProperty('legend')) {
                delete obverse.legend;
                console.log(`Removed obverse legend for Index ${index} (empty in DB)`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                // Escape XML apostrophe for corne d'abondance
                reverseDescription = reverseDescription.replace(/corne d'abondance/g, "corne d&apos;abondance");
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        // Preserve the xml:lang attribute
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        
                        // Update the description text
                        reverseType.description['#text'] = reverseDescription;
                        
                        // Ensure the xml:lang attribute is preserved
                        reverseType.description['@_xml:lang'] = langAttribute;
                        
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    } else {
                        console.warn(`No reverse description tag found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else if (reverse && reverse.hasOwnProperty('legend')) {
                delete reverse.legend;
                console.log(`Removed reverse legend for Index ${index} (empty in DB)`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Check if axis property exists (even if empty)
                        if (physDesc.hasOwnProperty('axis')) {
                            // Update the axis value regardless of whether it's empty or not
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            // If axis tag doesn't exist, create it
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        // Update weight if provided
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                // Create weight tag if it doesn't exist
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        }
                        
                        // Update diameter if provided
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                // Create diameter tag if it doesn't exist
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        }
                    } else {
                        console.warn(`No measurementsSet section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    // Try both possible locations for physDesc
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        // Update or create countermark tag with raw text (no HTML encoding)
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    } else {
                        console.warn(`No physDesc section found for Index ${index}`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                // Remove countermark tag if cell is empty
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.hasOwnProperty('countermark')) {
                        delete physDesc.countermark;
                        console.log(`Removed countermark for Index ${index} (empty in DB)`);
                    }
                } catch (error) {
                    console.warn(`Error removing countermark for Index ${index}:`, error.message);
                }
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Update reference info from column U
            if (referenceInfo && referenceInfo.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        cealexRef = refDesc.reference[1];
                        if (cealexRef) {
                            // Parse the reference info: "Picard, Faucher 2012, 367"
                            const refText = referenceInfo.toString().trim();
                            
                            // Extract the fixed part and number
                            const parts = refText.split(',');
                            if (parts.length >= 2) {
                                const fixedPart = parts[0] + ', ' + parts[1].trim(); // "Picard, Faucher 2012"
                                const numberPart = parts[2] ? parts[2].trim() : ''; // "367"
                                
                                // Update the CEAlex tag
                                if (cealexRef['tei:CEAlex']) {
                                    cealexRef['tei:CEAlex'] = fixedPart;
                                }
                                
                                // Update the idno tag
                                if (cealexRef['tei:idno']) {
                                    cealexRef['tei:idno'] = numberPart;
                                }
                                
                                console.log(`Updated reference info for Index ${index}: "${fixedPart}" with idno: "${numberPart}"`);
                            }
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating reference info for Index ${index}:`, error.message);
                }
            }

            // Update reference URL and title from column W
            referenceTitle = null;
            if (referenceUrl && referenceUrl.toString().trim() !== '') {
                try {
                    const refDesc = result.nuds.descMeta.refDesc;
                    if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                        mainRef = refDesc.reference[0];
                        if (mainRef) {
                            // Update the xlink:href attribute
                            mainRef['@_xlink:href'] = referenceUrl.toString().trim();
                            
                            // Special case for Index 3476 - extract from h4 class="text-center"
                            if (index === 3476) {
                                try {
                                    const response = await axios.get(referenceUrl.toString().trim(), {
                                        headers: {
                                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                                        }
                                    });
                                    const $ = cheerio.load(response.data);
                                    
                                    // Get the h4 element with class="text-center"
                                    const h4 = $('h4.text-center');
                                    if (h4.