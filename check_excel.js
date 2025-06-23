const xlsx = require('xlsx');

try {
    // Read the Excel file
    const workbook = xlsx.readFile('Sample_data.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    console.log('Sheet name:', sheetName);
    console.log('Range:', worksheet['!ref']);
    
    // Check columns Y (24) and AH (33) for more rows
    console.log('\n=== Column Y (URLs) and Column AH (Names) ===');
    console.log('Row | Column Y (URL) | Column AH (Name)');
    console.log('----|----------------|------------------');
    
    for (let row = 1; row <= 20; row++) {
        const urlCell = worksheet[xlsx.utils.encode_cell({r: row, c: 24})];
        const nameCell = worksheet[xlsx.utils.encode_cell({r: row, c: 33})];
        const url = urlCell ? urlCell.v : 'empty';
        const name = nameCell ? nameCell.v : 'empty';
        
        if (url !== 'empty' || name !== 'empty') {
            console.log(`${row.toString().padStart(3)} | ${url.substring(0, 50)}... | ${name}`);
        }
    }
    
    // Check if there are different values
    console.log('\n=== Unique URLs in Column Y ===');
    const uniqueUrls = new Set();
    for (let row = 1; row <= 50; row++) {
        const cell = worksheet[xlsx.utils.encode_cell({r: row, c: 24})];
        if (cell && cell.v) {
            uniqueUrls.add(cell.v);
        }
    }
    uniqueUrls.forEach(url => console.log(url));
    
    console.log('\n=== Unique Names in Column AH ===');
    const uniqueNames = new Set();
    for (let row = 1; row <= 50; row++) {
        const cell = worksheet[xlsx.utils.encode_cell({r: row, c: 33})];
        if (cell && cell.v) {
            uniqueNames.add(cell.v);
        }
    }
    uniqueNames.forEach(name => console.log(name));
    
} catch (error) {
    console.error('Error reading Excel file:', error.message);
} 