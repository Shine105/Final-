const XLSX = require('xlsx');

// Specify the path to the XLS file
const xlsFilePath = 'AHERI_220_14_11_2023.xls';

// Read the XLS file
const workbook = XLSX.readFile(xlsFilePath);

// Get the first sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

// Get the station name from cell 'A1'
const stationName = worksheet['A1'].v;

// Extract the initial date from cell 'B6' (offset index 4005)
const initialDateCell = worksheet['B6'];
let lastDate = '';
if (initialDateCell && typeof initialDateCell.v === 'number') {
    const dateObj = XLSX.SSF.parse_date_code(initialDateCell.v);
    lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
    console.log(`Initial date from B6 parsed as ${lastDate}`);
} else {
    console.error('Initial date in cell B6 is missing or not a number.');
}

// Extract the relevant range (4000 to 8000)
const limitedData = jsonData.slice(4000, 8001); // 8001 because slice end index is exclusive

const data = [];

// Iterate through the rows in the limited range
for (let rowIndex = 0; rowIndex < limitedData.length; rowIndex++) {
    const row = limitedData[rowIndex];
    
    // Extract date from column 'B' (index 1)
    const dateValue = row[1];
    if (typeof dateValue === 'number') {
        const dateObj = XLSX.SSF.parse_date_code(dateValue);
        lastDate = new Date(dateObj.y, dateObj.m - 1, dateObj.d).toISOString().split('T')[0];
        console.log(`Row ${rowIndex + 4000}: Date parsed as ${lastDate}`);
    }

    // Extract feeder names with 'F' prefix
    row.forEach(cell => {
        if (typeof cell === 'string' && cell.startsWith('F')) {
            if (lastDate === '') {
                console.error(`Row ${rowIndex + 4000}: Found feeder name ${cell} but no date has been set yet.`);
            } else {
                console.log(`Row ${rowIndex + 4000}: Found feeder name ${cell}`);
                data.push({
                    'Date': lastDate,
                    'Name of Station': stationName,
                    'Name of Feeder': cell
                });
            }
        }
    });
}

// Filter out entries where the date is still empty
const filteredData = data.filter(entry => entry.Date !== '');

if (filteredData.length === 0) {
    console.error('No valid data entries found. Please check the date extraction logic.');
}

// Output the data as a table
console.table(filteredData);
