const fs = require('fs-extra');
const path = require('path');
const xml2js = require('xml2js');
const axios = require('axios');
const cheerio = require('cheerio');

// Create output directory if it doesn't exist
const outputDir = path.join(__dirname, 'output');

async function updateDenominations() {
    try {
        // Get all XML files in the output directory
        const files = await fs.readdir(outputDir);
        const xmlFiles = files.filter(file => file.endsWith('.xml'));

        for (const file of xmlFiles) {
            console.log(`Processing ${file}...`);
            
            // Read the XML file
            const xmlPath = path.join(outputDir, file);
            const xmlContent = await fs.readFile(xmlPath, 'utf8');
            
            // Parse the XML
            const parser = new xml2js.Parser();
            const result = await parser.parseStringPromise(xmlContent);
            
            // Find the denomination element
            const denominationElement = result.nuds.descMeta[0].typeDesc[0].denomination[0];
            if (denominationElement && denominationElement.$ && denominationElement.$['xlink:href']) {
                const url = denominationElement.$['xlink:href'];
                console.log(`Found denomination URL: ${url}`);
                
                try {
                    // Fetch the webpage with proper headers
                    const response = await axios.get(url, {
                        headers: {
                            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                        }
                    });
                    const html = response.data;
                    
                    // Parse the HTML
                    const $ = cheerio.load(html);
                    
                    // Extract only the direct text of the h2 (excluding child tags)
                    const h2 = $('h2').first();
                    let h2Content = '';
                    h2.contents().each(function() {
                        if (this.type === 'text') {
                            h2Content += $(this).text();
                        }
                    });
                    h2Content = h2Content.trim();
                    console.log(`Extracted h2 content: ${h2Content}`);
                    
                    // Update the denomination text while keeping attributes
                    denominationElement._ = h2Content;
                    
                    // Convert back to XML
                    const builder = new xml2js.Builder();
                    const updatedXml = builder.buildObject(result);
                    
                    // Save the updated XML
                    await fs.writeFile(xmlPath, updatedXml);
                    console.log(`Updated ${file} with new denomination text`);
                } catch (error) {
                    console.error(`Error processing URL ${url}:`, error.message);
                }
            }
        }
        
        console.log('Finished processing all files');
    } catch (error) {
        console.error('Error:', error);
    }
}

// Run the script
updateDenominations().catch(console.error); 