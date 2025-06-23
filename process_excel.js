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

// Helper function to format discovery date
function formatDiscoveryDate(dateInput) {
    if (!dateInput) {
        return null;
    }

    const dateStr = dateInput.toString().trim();
    let dateObj;

    // Try parsing DD.MM.YYYY, YYYY-MM-DD, or other standard formats
    const dotMatch = dateStr.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (dotMatch) {
        // Handles DD.MM.YYYY
        dateObj = new Date(`${dotMatch[3]}-${dotMatch[2]}-${dotMatch[1]}T00:00:00Z`);
    } else {
        // Handles YYYY-MM-DD and other formats recognized by new Date()
        dateObj = new Date(dateStr);
    }
    
    // Check for invalid date
    if (isNaN(dateObj.getTime())) {
         // Handle year-only case
        if (/^\d{4}$/.test(dateStr)) {
            return {
                standard: dateStr,
                readable: dateStr
            };
        }
        console.warn(`Could not parse invalid date: "${dateStr}"`);
        return null;
    }

    // Ensure we handle the timezone correctly to avoid off-by-one day errors
    const year = dateObj.getUTCFullYear();
    const month = (dateObj.getUTCMonth() + 1);
    const day = dateObj.getUTCDate();

    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June', 
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    return {
        standard: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
        readable: `${monthNames[month - 1]} ${String(day).padStart(2, '0')}, ${year}`
    };
}

async function processExcelFile() {
    try {
        // Read the Excel file
        const workbook = xlsx.readFile('Sample_data.xlsx');
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

        const outputDir = path.join(__dirname, 'output');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
            console.log('Created output directory:', outputDir);
        } else {
            console.log('Output directory exists:', outputDir);
        }

        // Process each row
        for (let R = range.s.r; R <= range.e.r; R++) {
            // Declare all variables at the top of the loop to avoid redeclaration issues
            let mainRef = null;
            let cealexRef = null;
            let identifiers = null;
            let indexIdentifier = null;
            let coinNumberIdentifier = null;
            let inventoryNumberIdentifier = null;
            let discoveryDateUpdated = false;
            let shouldCommentDenomination = false;
            let tagsToComment = [];
            
            // Get the Index value from column A
            const indexCell = worksheet[xlsx.utils.encode_cell({r: R, c: 0})];
            if (!indexCell) continue;
            
            const index = indexCell.v;
            
            // Skip header rows - only process if index is a number
            if (typeof index !== 'number' || isNaN(index)) {
                console.log(`Skipping row with non-numeric index: ${index}`);
                continue;
            }
            
            const xmlFilePath = path.join(outputDir, `${index}.xml`);
            
            // Parse the template XML
            const result = parser.parse(templateContent);
            
            // Get values from Excel columns
            const titleCell = worksheet[xlsx.utils.encode_cell({r: R, c: 1})]; // Column B
            const fromDateCell = worksheet[xlsx.utils.encode_cell({r: R, c: 2})]; // Column C
            const toDateCell = worksheet[xlsx.utils.encode_cell({r: R, c: 3})]; // Column D
            const denominationUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 5})]; // Column F
            const denominationNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 35})]; // Column AJ
            const materialUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 6})]; // Column G
            const materialNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 36})]; // Column AK
            const authorityUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 7})]; // Column H
            const authorityNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 37})]; // Column AL
            const typeSeriesCell = worksheet[xlsx.utils.encode_cell({r: R, c: 9})]; // Column J
            const mintUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 10})]; // Column K
            const mintNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 38})]; // Column AM
            let obverseDescription = worksheet[xlsx.utils.encode_cell({r: R, c: 11})]; // Column L
            const obverseLegendCell = worksheet[xlsx.utils.encode_cell({r: R, c: 12})]; // Column M
            let reverseDescription = worksheet[xlsx.utils.encode_cell({r: R, c: 13})]; // Column N
            const reverseLegendCell = worksheet[xlsx.utils.encode_cell({r: R, c: 14})]; // Column O
            const symbolUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 15})]; // Column P
            const symbolNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 39})]; // Column AN
            const axisCell = worksheet[xlsx.utils.encode_cell({r: R, c: 16})]; // Column Q
            const weightCell = worksheet[xlsx.utils.encode_cell({r: R, c: 17})]; // Column R
            const diameterCell = worksheet[xlsx.utils.encode_cell({r: R, c: 18})]; // Column S
            const countermarkCell = worksheet[xlsx.utils.encode_cell({r: R, c: 19})]; // Column T
            const referenceTextCell = worksheet[xlsx.utils.encode_cell({r: R, c: 21})]; // Column V
            const referenceUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 22})]; // Column W
            const cealexReferenceCell = worksheet[xlsx.utils.encode_cell({r: R, c: 20})]; // Column U
            const stratigraphicUnitCell = worksheet[xlsx.utils.encode_cell({r: R, c: 23})]; // Column X
            const fallsWithinUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 24})]; // Column Y
            const hoardUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 25})]; // Column Z
            const hoardNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 40})]; // Column AO
            const discoveryDateCell = worksheet[xlsx.utils.encode_cell({r: R, c: 26})]; // Column AA
            const coinNumberCell = worksheet[xlsx.utils.encode_cell({r: R, c: 28})]; // Column AC
            const inventoryNumberCell = worksheet[xlsx.utils.encode_cell({r: R, c: 29})]; // Column AD
            const departmentUrlCell = worksheet[xlsx.utils.encode_cell({r: R, c: 30})]; // Column AE
            const departmentNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 41})]; // Column AP
            const repositoryCell = worksheet[xlsx.utils.encode_cell({r: R, c: 31})]; // Column AF
            const fileLocationCell = worksheet[xlsx.utils.encode_cell({r: R, c: 32})]; // Column AG
            const fallsWithinNameCell = worksheet[xlsx.utils.encode_cell({r: R, c: 33})]; // Column AH

            // Helper function to check if a cell is empty
            function isEmptyCell(cell) {
                return !cell || 
                       cell.v === undefined || 
                       cell.v === null || 
                       cell.v === '' || 
                       (typeof cell.v === 'string' && cell.v.trim() === '') ||
                       (typeof cell.v === 'number' && isNaN(cell.v));
            }

            // Helper function to get cell value safely
            function getCellValue(cell) {
                if (isEmptyCell(cell)) return '';
                return cell.v.toString().trim();
            }

            const title = getCellValue(titleCell);
            const fromDate = getCellValue(fromDateCell);
            const toDate = getCellValue(toDateCell);
            const denominationUrl = getCellValue(denominationUrlCell);
            const denominationName = getCellValue(denominationNameCell);
            const materialUrl = getCellValue(materialUrlCell);
            const materialName = getCellValue(materialNameCell);
            const authorityUrl = getCellValue(authorityUrlCell);
            const authorityName = getCellValue(authorityNameCell);
            const typeSeries = getCellValue(typeSeriesCell);
            const mintUrl = getCellValue(mintUrlCell);
            const mintName = getCellValue(mintNameCell);
            obverseDescription = getCellValue(obverseDescription);
            const obverseLegendText = getCellValue(obverseLegendCell);
            reverseDescription = getCellValue(reverseDescription);
            const reverseLegendText = getCellValue(reverseLegendCell);
            const symbolUrl = getCellValue(symbolUrlCell);
            const symbolName = getCellValue(symbolNameCell);
            const axisValue = getCellValue(axisCell);
            const weightValue = getCellValue(weightCell);
            const diameterValue = getCellValue(diameterCell);
            const countermarkValue = getCellValue(countermarkCell);
            const referenceInfo = getCellValue(referenceTextCell);
            const referenceUrl = getCellValue(referenceUrlCell);
            const stratigraphicUnit = getCellValue(stratigraphicUnitCell);
            const fallsWithinUrl = getCellValue(fallsWithinUrlCell);
            const hoardUrl = getCellValue(hoardUrlCell);
            const hoardName = getCellValue(hoardNameCell);
            const discoveryDate = getCellValue(discoveryDateCell);
            const coinNumber = getCellValue(coinNumberCell);
            const inventoryNumber = getCellValue(inventoryNumberCell);
            const departmentUrl = getCellValue(departmentUrlCell);
            const departmentName = getCellValue(departmentNameCell);
            const repository = getCellValue(repositoryCell);
            const fileLocation = getCellValue(fileLocationCell);
            const fallsWithinName = getCellValue(fallsWithinNameCell);

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

            // Update denomination logic
            if (denominationUrl && denominationName) {
                let denomination = result.nuds.descMeta.typeDesc.denomination;
                if (denomination) {
                    denomination['@_xlink:href'] = denominationUrl;
                    denomination['#text'] = denominationName;
                    console.log(`Updated denomination for Index ${index}: "${denominationName}" with URL: ${denominationUrl}`);
                }
            } else {
                shouldCommentDenomination = true;
                console.log(`Denomination data for Index ${index} is incomplete. URL: "${denominationUrl}", Name: "${denominationName}". The tag will be commented out.`);
            }

            // Update typeSeries if provided
            if (typeSeries && typeSeries.toString().trim() !== '') {
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
                tagsToComment.push('typeSeries');
                console.log(`TypeSeries for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update material if URL and name exist
            if (materialUrl && materialName) {
                if (!result.nuds.descMeta.typeDesc.material) {
                    result.nuds.descMeta.typeDesc.material = {
                        '@_xlink:href': materialUrl,
                        '@_xlink:type': 'simple',
                        '#text': materialName
                    };
                } else {
                    result.nuds.descMeta.typeDesc.material['@_xlink:href'] = materialUrl;
                    result.nuds.descMeta.typeDesc.material['#text'] = materialName;
                }
                console.log(`Updated material for Index ${index}: "${materialName}" with URL: ${materialUrl}`);
            } else {
                tagsToComment.push('material');
                console.log(`Material data for Index ${index} is incomplete. URL: "${materialUrl}", Name: "${materialName}". The tag will be commented out.`);
            }

            // Update authority/persname - always use Column H for xlink:href and Column AL for inner text
            try {
                if (result.nuds.descMeta.typeDesc.authority && result.nuds.descMeta.typeDesc.authority.persname) {
                    const persname = result.nuds.descMeta.typeDesc.authority.persname;
                    if (authorityUrl) {
                        persname['@_xlink:href'] = authorityUrl;
                    }
                    persname['#text'] = authorityName || '';
                    console.log(`Updated authority/persname for Index ${index}: "${authorityName || 'empty'}" with URL: ${authorityUrl || 'none'}`);
                }
            } catch (error) {
                console.warn(`Error updating authority/persname for Index ${index}:`, error.message);
                tagsToComment.push('authority');
            }

            // Update mint information if URL and name exist
            if (mintUrl && mintName) {
                const geognameElements = result.nuds.descMeta.typeDesc.geographic.geogname;
                if (Array.isArray(geognameElements)) {
                    const mintElement = geognameElements.find(el => el['@_xlink:role'] === 'mint');
                    if (mintElement) {
                        mintElement['@_xlink:href'] = mintUrl;
                        mintElement['#text'] = mintName;
                        console.log(`Updated mint for Index ${index}: "${mintName}" with URL: ${mintUrl}`);
                    }
                }
            } else {
                tagsToComment.push('mint');
                console.log(`Mint data for Index ${index} is incomplete. URL: "${mintUrl}", Name: "${mintName}". The tag will be commented out.`);
            }

            // Update obverse description if value exists
            if (obverseDescription) {
                try {
                    const obverseType = result.nuds.descMeta.typeDesc.obverse.type;
                    if (obverseType && obverseType.description) {
                        const langAttribute = obverseType.description['@_xml:lang'] || 'fr';
                        obverseType.description['#text'] = obverseDescription;
                        obverseType.description['@_xml:lang'] = langAttribute;
                        console.log(`Updated obverse description for Index ${index}: "${obverseDescription}"`);
                    }
                } catch (error) {
                    console.warn(`Error updating obverse description for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('obverseDescription');
                console.log(`Obverse description for Index ${index} is empty. The tag will be commented out.`);
            }

            // Handle obverse legend from column M
            const obverse = result.nuds.descMeta.typeDesc.obverse;
            if (obverseLegendText && obverseLegendText.trim() !== '') {
                obverse.legend = obverseLegendText.trim();
                console.log(`Added obverse legend for Index ${index}: "${obverseLegendText.trim()}"`);
            } else {
                // Remove the legend property if it exists and is empty
                if (obverse.legend) {
                    delete obverse.legend;
                }
                console.log(`Obverse legend for Index ${index} is empty. The tag will not be included.`);
            }

            // Update reverse description if value exists
            if (reverseDescription) {
                try {
                    const reverseType = result.nuds.descMeta.typeDesc.reverse.type;
                    if (reverseType && reverseType.description) {
                        const langAttribute = reverseType.description['@_xml:lang'] || 'fr';
                        reverseType.description['#text'] = reverseDescription;
                        reverseType.description['@_xml:lang'] = langAttribute;
                        console.log(`Updated reverse description for Index ${index}: "${reverseDescription}"`);
                    }
                } catch (error) {
                    console.warn(`Error updating reverse description for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('reverseDescription');
                console.log(`Reverse description for Index ${index} is empty. The tag will be commented out.`);
            }

            // Handle reverse legend from column O
            const reverse = result.nuds.descMeta.typeDesc.reverse;
            if (reverseLegendText && reverseLegendText.trim() !== '') {
                reverse.legend = reverseLegendText.trim();
                console.log(`Added reverse legend for Index ${index}: "${reverseLegendText.trim()}"`);
            } else {
                // Remove the legend property if it exists and is empty
                if (reverse.legend) {
                    delete reverse.legend;
                }
                console.log(`Reverse legend for Index ${index} is empty. The tag will not be included.`);
            }

            // Handle symbol from column P and AN
            if (symbolUrl && symbolName) {
                try {
                    const typeDesc = result.nuds.descMeta.typeDesc;
                    if (typeDesc && typeDesc.symbol) {
                        typeDesc.symbol['@_xlink:href'] = symbolUrl;
                        typeDesc.symbol['#text'] = symbolName;
                        console.log(`Updated symbol for Index ${index}: "${symbolName}" with URL: ${symbolUrl}`);
                    }
                } catch (error) {
                    console.warn(`Error updating symbol for Index ${index}:`, error.message);
                    tagsToComment.push('symbol');
                }
            } else {
                tagsToComment.push('symbol');
                console.log(`Symbol data for Index ${index} is incomplete. URL: "${symbolUrl}", Name: "${symbolName}". The tag will be commented out.`);
            }

            // Update axis value from column Q
            if (axisValue && axisValue.toString().trim() !== '') {
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        if (physDesc.hasOwnProperty('axis')) {
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Updated axis for Index ${index}: "${axisValue}"`);
                        } else {
                            physDesc.axis = axisValue.toString().trim();
                            console.log(`Created axis tag for Index ${index}: "${axisValue}"`);
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating axis for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('axis');
                console.log(`Axis value for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update weight and diameter values from columns R and S
            if (weightValue || diameterValue) {
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc && physDesc.measurementsSet) {
                        if (weightValue && weightValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.weight) {
                                physDesc.measurementsSet.weight['#text'] = weightValue.toString().trim();
                                console.log(`Updated weight for Index ${index}: "${weightValue}"`);
                            } else {
                                physDesc.measurementsSet.weight = {
                                    '#text': weightValue.toString().trim(),
                                    '@_units': 'g'
                                };
                                console.log(`Created weight tag for Index ${index}: "${weightValue}"`);
                            }
                        } else {
                            tagsToComment.push('weight');
                            console.log(`Weight value for Index ${index} is empty. The tag will be commented out.`);
                        }
                        
                        if (diameterValue && diameterValue.toString().trim() !== '') {
                            if (physDesc.measurementsSet.diameter) {
                                physDesc.measurementsSet.diameter['#text'] = diameterValue.toString().trim();
                                console.log(`Updated diameter for Index ${index}: "${diameterValue}"`);
                            } else {
                                physDesc.measurementsSet.diameter = {
                                    '#text': diameterValue.toString().trim(),
                                    '@_units': 'mm'
                                };
                                console.log(`Created diameter tag for Index ${index}: "${diameterValue}"`);
                            }
                        } else {
                            tagsToComment.push('diameter');
                            console.log(`Diameter value for Index ${index} is empty. The tag will be commented out.`);
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating measurements for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('measurements');
                console.log(`Measurement values for Index ${index} are empty. The tags will be commented out.`);
            }

            // Update countermark value from column T
            if (countermarkValue && countermarkValue.toString().trim() !== '') {
                try {
                    let physDesc = result.nuds.descMeta.physDesc;
                    if (!physDesc) {
                        physDesc = result.nuds.physDesc;
                    }
                    
                    if (physDesc) {
                        physDesc.countermark = {
                            '#text': countermarkValue.toString().trim()
                        };
                        console.log(`Updated countermark for Index ${index}: "${countermarkValue}"`);
                    }
                } catch (error) {
                    console.warn(`Error updating countermark for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('countermark');
                console.log(`Countermark value for Index ${index} is empty. The tag will be commented out.`);
            }

            // Special case: Update reference tags - always generate, use Column V for inner text and Column W for xlink:href
            try {
                const refDesc = result.nuds.descMeta.refDesc;
                if (refDesc && refDesc.reference && Array.isArray(refDesc.reference)) {
                    // Update first reference (main reference)
                    const mainRef = refDesc.reference[0];
                    if (mainRef) {
                        mainRef['@_xlink:href'] = referenceUrl || '';
                        mainRef['#text'] = referenceInfo || '';
                        console.log(`Updated main reference for Index ${index}: text="${referenceInfo || 'empty'}", url="${referenceUrl || 'empty'}"`);
                    }
                    // Update second reference (CEAlex reference) if it exists
                    if (refDesc.reference.length > 1) {
                        const cealexRef = refDesc.reference[1];
                        // Extract number after last comma from Column U
                        let idnoValue = '';
                        if (cealexReferenceCell && cealexReferenceCell.v) {
                            const cealexText = cealexReferenceCell.v.toString();
                            const parts = cealexText.split(',');
                            if (parts.length > 0) {
                                idnoValue = parts[parts.length - 1].trim();
                            }
                        }
                        if (cealexRef) {
                            // If tei:idno is an object
                            if (cealexRef['tei:idno'] && typeof cealexRef['tei:idno'] === 'object') {
                                cealexRef['tei:idno']['#text'] = idnoValue;
                                console.log(`Updated CEAlex reference idno for Index ${index}: "${idnoValue}" (object)`);
                            } else if (Array.isArray(cealexRef['tei:idno'])) {
                                // If tei:idno is an array, update all
                                cealexRef['tei:idno'].forEach((idno, idx) => {
                                    if (typeof idno === 'object') idno['#text'] = idnoValue;
                                });
                                console.log(`Updated CEAlex reference idno for Index ${index}: "${idnoValue}" (array)`);
                            } else {
                                // If tei:idno is a string or missing, set as object
                                cealexRef['tei:idno'] = {'#text': idnoValue};
                                console.log(`Set CEAlex reference idno for Index ${index}: "${idnoValue}" (created)`);
                            }
                        }
                    }
                }
            } catch (error) {
                console.warn(`Error updating references for Index ${index}:`, error.message);
            }

            // Update stratigraphic unit from column X
            if (stratigraphicUnit && stratigraphicUnit.toString().trim() !== '') {
                try {
                    const findspotDesc = result.nuds.descMeta.findspotDesc;
                    if (findspotDesc && findspotDesc.findspot) {
                        const findspot = findspotDesc.findspot;
                        
                        if (findspot.geogname && Array.isArray(findspot.geogname)) {
                            const stratUnitElement = findspot.geogname.find(el => el['@_xlink:role'] === 'stratigraphicUnit');
                            if (stratUnitElement) {
                                stratUnitElement['#text'] = stratigraphicUnit.toString().trim();
                                console.log(`Updated stratigraphic unit for Index ${index}: "${stratigraphicUnit}"`);
                            }
                        } else if (findspot.geogname && findspot.geogname['@_xlink:role'] === 'stratigraphicUnit') {
                            findspot.geogname['#text'] = stratigraphicUnit.toString().trim();
                            console.log(`Updated stratigraphic unit for Index ${index}: "${stratigraphicUnit}"`);
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating stratigraphic unit for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('stratigraphicUnit');
                console.log(`Stratigraphic unit for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update fallsWithin geogname from columns Y (URL) and AH (name)
            if ((fallsWithinUrl && fallsWithinUrl.toString().trim() !== '') || (fallsWithinName && fallsWithinName.toString().trim() !== '')) {
                try {
                    const findspotDesc = result.nuds.descMeta.findspotDesc;
                    if (findspotDesc && findspotDesc.findspot) {
                        const findspot = findspotDesc.findspot;
                        if (!findspot.fallsWithin) {
                            findspot.fallsWithin = { geogname: {} };
                        }
                        let geogname = Array.isArray(findspot.fallsWithin.geogname) ? findspot.fallsWithin.geogname[0] : findspot.fallsWithin.geogname;
                        if (!geogname) {
                            geogname = {};
                        }
                        if (fallsWithinUrl && fallsWithinUrl.toString().trim() !== '') {
                            geogname['@_xlink:type'] = 'simple';
                            geogname['@_xlink:href'] = fallsWithinUrl.toString().trim();
                        }
                        geogname['@_xlink:role'] = 'findspot';
                        if (fallsWithinName && fallsWithinName.toString().trim() !== '') {
                            geogname['#text'] = fallsWithinName.toString().trim();
                        }
                        findspot.fallsWithin.geogname = geogname;
                        console.log(`Updated fallsWithin for Index ${index}: name="${fallsWithinName}", url="${fallsWithinUrl}"`);
                    }
                } catch (error) {
                    console.warn(`Error updating fallsWithin for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('fallsWithin');
                console.log(`FallsWithin data for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update Coin Number from column AC
            if (coinNumber && coinNumber.toString().trim() !== '') {
                try {
                    if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                        const identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                            ? result.nuds.descMeta.adminDesc.identifier 
                            : [result.nuds.descMeta.adminDesc.identifier];
                        
                        const coinNumberIdentifier = identifiers.find(id => id['@_localType'] === 'Coin Number');
                        if (coinNumberIdentifier) {
                            coinNumberIdentifier['#text'] = coinNumber.toString().trim();
                            console.log(`Updated Coin Number for Index ${index}: "${coinNumber}"`);
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating Coin Number for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('coinNumber');
                console.log(`Coin Number for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update Inventory Number from column AD
            if (inventoryNumber && inventoryNumber.toString().trim() !== '') {
                try {
                    if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                        const identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                            ? result.nuds.descMeta.adminDesc.identifier 
                            : [result.nuds.descMeta.adminDesc.identifier];
                        
                        const inventoryNumberIdentifier = identifiers.find(id => id['@_localType'] === 'Inventory Number');
                        if (inventoryNumberIdentifier) {
                            inventoryNumberIdentifier['#text'] = inventoryNumber.toString().trim();
                            console.log(`Updated Inventory Number for Index ${index}: "${inventoryNumber}"`);
                        }
                    }
                } catch (error) {
                    console.warn(`Error updating Inventory Number for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('inventoryNumber');
                console.log(`Inventory Number for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update Department from column AE and AP
            if (departmentUrl && departmentName) {
                try {
                    if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.department) {
                        const department = result.nuds.descMeta.adminDesc.department;
                        department['@_xlink:href'] = departmentUrl;
                        department['#text'] = departmentName;
                        console.log(`Updated department for Index ${index}: "${departmentName}" with URL: ${departmentUrl}`);
                    }
                } catch (error) {
                    console.warn(`Error updating department for Index ${index}:`, error.message);
                    tagsToComment.push('department');
                }
            } else {
                tagsToComment.push('department');
                console.log(`Department data for Index ${index} is incomplete. URL: "${departmentUrl}", Name: "${departmentName}". The tag will be commented out.`);
            }

            // Update Repository from column AF
            if (repository && repository.toString().trim() !== '') {
                try {
                    if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.repository) {
                        const repositoryElement = result.nuds.descMeta.adminDesc.repository;
                        repositoryElement['#text'] = repository.toString().trim();
                        console.log(`Updated repository for Index ${index}: "${repository}"`);
                    }
                } catch (error) {
                    console.warn(`Error updating repository for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('repository');
                console.log(`Repository for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update File Location from column AG
            if (fileLocation && fileLocation.toString().trim() !== '') {
                try {
                    if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.fileLocation) {
                        const fileLocationElement = result.nuds.descMeta.adminDesc.fileLocation;
                        fileLocationElement['#text'] = fileLocation.toString().trim();
                        console.log(`Updated file location for Index ${index}: "${fileLocation}"`);
                    }
                } catch (error) {
                    console.warn(`Error updating file location for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('fileLocation');
                console.log(`File location for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update discovery date from column AA
            if (discoveryDate && discoveryDate.toString().trim() !== '') {
                try {
                    const findspotDesc = result.nuds.descMeta.findspotDesc;
                    const formattedDate = formatDiscoveryDate(discoveryDate);

                    if (findspotDesc && findspotDesc.discovery && findspotDesc.discovery.date && formattedDate) {
                        const dateElement = findspotDesc.discovery.date;
                        
                        dateElement['@_standardDate'] = formattedDate.standard;
                        dateElement['#text'] = formattedDate.readable;
                        discoveryDateUpdated = true;

                        console.log(`Updated discovery date for Index ${index}: "${formattedDate.readable}" (standard: ${formattedDate.standard})`);
                    }
                } catch (error) {
                    console.warn(`Error updating discovery date for Index ${index}:`, error.message);
                }
            } else {
                tagsToComment.push('discoveryDate');
                console.log(`Discovery date for Index ${index} is empty. The tag will be commented out.`);
            }

            // Update the identifier in adminDesc
            if (result.nuds.descMeta.adminDesc && result.nuds.descMeta.adminDesc.identifier) {
                const identifiers = Array.isArray(result.nuds.descMeta.adminDesc.identifier) 
                    ? result.nuds.descMeta.adminDesc.identifier 
                    : [result.nuds.descMeta.adminDesc.identifier];
                
                const indexIdentifier = identifiers.find(id => id['@_localType'] === 'Index');
                if (indexIdentifier) {
                    indexIdentifier['#text'] = index.toString();
                }
            }

            // Handle hoard from column Z and AO
            if (hoardUrl && hoardName) {
                try {
                    const findspotDesc = result.nuds.descMeta.findspotDesc;
                    if (findspotDesc && findspotDesc.findspot) {
                        const findspot = findspotDesc.findspot;
                        if (!findspot.hoard) {
                            findspot.hoard = {
                                '@_xlink:href': hoardUrl,
                                '@_xlink:type': 'simple',
                                '#text': hoardName
                            };
                        } else {
                            findspot.hoard['@_xlink:href'] = hoardUrl;
                            findspot.hoard['#text'] = hoardName;
                        }
                        console.log(`Updated hoard for Index ${index}: "${hoardName}" with URL: ${hoardUrl}`);
                    }
                } catch (error) {
                    console.warn(`Error updating hoard for Index ${index}:`, error.message);
                    tagsToComment.push('hoard');
                }
            } else {
                tagsToComment.push('hoard');
                console.log(`Hoard data for Index ${index} is incomplete. URL: "${hoardUrl}", Name: "${hoardName}". The tag will be commented out.`);
            }

            // Escape XML apostrophe for Centre d'Études Alexandrines
            const builder = new XMLBuilder({
                ignoreAttributes: false,
                attributeNamePrefix: '@_',
                format: true,
                suppressEmptyNode: true,
                unpairedTags: ["hr", "br", "link", "meta"],
                processEntities: false,
            });

            let xmlContent = builder.build(result);
            
            // Update mets:FLocat with URL from column AG
            if (fileLocation && fileLocation.toString().trim() !== '') {
                const flocatRegex = /<mets:FLocat LOCYPE="URL" xlink:href="[^"]*"/;
                const match = xmlContent.match(flocatRegex);
                if (match) {
                    xmlContent = xmlContent.replace(flocatRegex, `<mets:FLocat LOCYPE="URL" xlink:href="${fileLocation.toString().trim()}"`);
                    console.log(`Updated mets:FLocat URL for Index ${index}: ${fileLocation}`);
                }
            } else {
                // Comment out mets:FLocat if fileLocation is empty
                const flocatRegex = /<mets:FLocat[\s\S]*?\/>/;
                const match = xmlContent.match(flocatRegex);
                if (match) {
                    xmlContent = xmlContent.replace(flocatRegex, `<!-- ${match[0]} -->`);
                    console.log(`Commented out mets:FLocat for Index ${index} (empty fileLocation)`);
                }
            }
            
            // If denomination URL was empty, comment out the denomination tag
            if (shouldCommentDenomination) {
                const denominationTagRegex = /<denomination[\s\S]*?<\/denomination>/;
                const match = xmlContent.match(denominationTagRegex);
                if (match) {
                    xmlContent = xmlContent.replace(denominationTagRegex, `<!-- ${match[0]} -->`);
                }
            }

            // Apply global commenting logic for all tracked tags
            tagsToComment.forEach(tagName => {
                let tagRegex;
                let shouldSkipDefault = false;
                
                switch (tagName) {
                    case 'typeSeries':
                        tagRegex = /<typeSeries[\s\S]*?<\/typeSeries>/;
                        break;
                    case 'material':
                        tagRegex = /<material[\s\S]*?\/>/;
                        break;
                    case 'authority':
                        tagRegex = /<authority[\s\S]*?<\/authority>/;
                        break;
                    case 'mint':
                        // Comment out the entire geographic section when mint data is missing
                        const geographicRegex = /<geographic>[\s\S]*?<\/geographic>/;
                        const match = xmlContent.match(geographicRegex);
                        if (match) {
                            xmlContent = xmlContent.replace(geographicRegex, `<!-- ${match[0]} -->`);
                            console.log(`Commented out geographic section for Index ${index} (missing mint data)`);
                        }
                        shouldSkipDefault = true;
                        break;
                    case 'obverseDescription':
                        // Only comment out description tags that are within obverse sections and are empty
                        const obverseMatches = xmlContent.match(/<obverse>[\s\S]*?<\/obverse>/g);
                        if (obverseMatches) {
                            obverseMatches.forEach(match => {
                                const descriptionMatch = match.match(/<description[\s\S]*?xml:lang="fr"[\s\S]*?<\/description>/);
                                if (descriptionMatch) {
                                    // Check if the description content is empty or just whitespace
                                    const contentMatch = descriptionMatch[0].match(/<description[^>]*>([\s\S]*?)<\/description>/);
                                    if (contentMatch && (!contentMatch[1] || contentMatch[1].trim() === '')) {
                                        xmlContent = xmlContent.replace(descriptionMatch[0], `<!-- ${descriptionMatch[0]} -->`);
                                        console.log(`Commented out empty obverse description for Index ${index}`);
                                    }
                                }
                            });
                        }
                        shouldSkipDefault = true;
                        break;
                    case 'obverseLegend':
                        tagRegex = /<legend[\s\S]*?<\/legend>/;
                        break;
                    case 'reverseDescription':
                        // Only comment out description tags that are within reverse sections and are empty
                        const reverseMatches = xmlContent.match(/<reverse>[\s\S]*?<\/reverse>/g);
                        if (reverseMatches) {
                            reverseMatches.forEach(match => {
                                const descriptionMatch = match.match(/<description[\s\S]*?xml:lang="fr"[\s\S]*?<\/description>/);
                                if (descriptionMatch) {
                                    // Check if the description content is empty or just whitespace
                                    const contentMatch = descriptionMatch[0].match(/<description[^>]*>([\s\S]*?)<\/description>/);
                                    if (contentMatch && (!contentMatch[1] || contentMatch[1].trim() === '')) {
                                        xmlContent = xmlContent.replace(descriptionMatch[0], `<!-- ${descriptionMatch[0]} -->`);
                                        console.log(`Commented out empty reverse description for Index ${index}`);
                                    }
                                }
                            });
                        }
                        shouldSkipDefault = true;
                        break;
                    case 'reverseLegend':
                        tagRegex = /<legend[\s\S]*?<\/legend>/;
                        break;
                    case 'axis':
                        tagRegex = /<axis[\s\S]*?<\/axis>/;
                        break;
                    case 'weight':
                        tagRegex = /<weight[\s\S]*?<\/weight>/;
                        break;
                    case 'diameter':
                        tagRegex = /<diameter[\s\S]*?<\/diameter>/;
                        break;
                    case 'countermark':
                        tagRegex = /<countermark[\s\S]*?<\/countermark>/;
                        break;
                    case 'stratigraphicUnit':
                        tagRegex = /<geogname[\s\S]*?xlink:role="stratigraphicUnit"[\s\S]*?\/>/;
                        break;
                    case 'fallsWithin':
                        tagRegex = /<fallsWithin[\s\S]*?<\/fallsWithin>/;
                        break;
                    case 'hoard':
                        tagRegex = /<hoard[\s\S]*?<\/hoard>/;
                        break;
                    case 'coinNumber':
                        tagRegex = /<identifier[\s\S]*?localType="Coin Number"[\s\S]*?<\/identifier>/;
                        break;
                    case 'inventoryNumber':
                        tagRegex = /<identifier[\s\S]*?localType="Inventory Number"[\s\S]*?<\/identifier>/;
                        break;
                    case 'department':
                        tagRegex = /<department[\s\S]*?<\/department>/;
                        break;
                    case 'repository':
                        tagRegex = /<repository[\s\S]*?<\/repository>/;
                        break;
                    case 'fileLocation':
                        tagRegex = /<fileLocation[\s\S]*?<\/fileLocation>/;
                        break;
                    case 'discoveryDate':
                        tagRegex = /<date[\s\S]*?<\/date>/;
                        break;
                    case 'symbol':
                        tagRegex = /<symbol[\s\S]*?<\/symbol>/;
                        break;
                    default:
                        return; // Skip unknown tags
                }
                
                if (!shouldSkipDefault && tagRegex) {
                    const match = xmlContent.match(tagRegex);
                    if (match) {
                        xmlContent = xmlContent.replace(tagRegex, `<!-- ${match[0]} -->`);
                        console.log(`Commented out ${tagName} tag for Index ${index}`);
                    }
                }
            });

            // Add discovery date comment if the date was updated
            if (discoveryDateUpdated) {
                const discoveryTagRegex = /(<discovery>[\s\S]*?<\/date>)/;
                const match = xmlContent.match(discoveryTagRegex);
                if (match) {
                    xmlContent = xmlContent.replace(discoveryTagRegex, `${match[0]}\n        <!-- ISO 8601: yyyy-mm-dd -->`);
                }
            }

            // Write the XML file
            fs.writeFileSync(xmlFilePath, xmlContent, 'utf-8');
            console.log(`Generated XML for Index ${index}: ${xmlFilePath}`);
        }

        console.log('Excel processing completed successfully!');
    } catch (error) {
        console.error('Error processing Excel file:', error);
    }
}

// Run the script
processExcelFile();